"""Matching logic shared identically in shape by 家电 and 数码 coupon rows.

Both projects' output headers start with the same ten columns (单据号,
单据日期, 商品名称, 品牌, 财务大类, 明细摘要, <subsidy header>, 备注,
详细情况, 回款情况) — only the subsidy header's text differs, never its
position. The row-layout indexes below are therefore derived from the two
profiles' output headers rather than written as positional literals: if a
profile ever gains, loses or reorders a column, the derivation fails loudly
at import time instead of silently misreading every row. Both projects
exclude the trailing receipt-remark block already pinned to the bottom by an
earlier pass, instead of maintaining separate matching implementations.
"""

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import NamedTuple

from processors.common.dates import (
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.common.references import REFERENCE_RE
from processors.coupons import sources


def shared_index(label: str) -> int:
    """Resolve one row index that both profiles' output headers agree on.

    Every column except the subsidy one shares a header text across the two
    profiles, so a label that either profile lacks — or that the two place
    differently — is a layout drift that would otherwise silently misread
    every row of one project.
    """
    indexes: set[int] = set()
    for profile in (sources.APPLIANCE_PROFILE, sources.DIGITAL_PROFILE):
        if label not in profile.output_header:
            raise RuntimeError(
                f"{label}在{profile.name}的行布局中缺失"
            )
        indexes.add(profile.output_header.index(label))
    if len(indexes) != 1:
        raise RuntimeError(f"{label}在家电和数码行布局中的位置不一致")
    return next(iter(indexes))


def shared_profile_field_index(appliance_label: str, digital_label: str) -> int:
    """Resolve one index whose header text differs between the two profiles.

    The subsidy header texts are pinned to the external export's column
    names, so the two profiles look the column up under different labels;
    the position itself must still agree, or one project's rows would be
    read under the other's columns.
    """
    positions: list[int] = []
    for profile, label in (
        (sources.APPLIANCE_PROFILE, appliance_label),
        (sources.DIGITAL_PROFILE, digital_label),
    ):
        if label not in profile.output_header:
            raise RuntimeError(f"{label}在{profile.name}的行布局中缺失")
        positions.append(profile.output_header.index(label))
    if positions[0] != positions[1]:
        raise RuntimeError(
            f"{appliance_label}（家电）与{digital_label}（数码）"
            "在行布局中的位置不一致"
        )
    return positions[0]


DOCUMENT_INDEX = shared_index("单据号")
DATE_INDEX = shared_index("单据日期")
SUMMARY_INDEX = shared_index("明细摘要")
SUBSIDY_INDEX = shared_profile_field_index(
    sources.COUPON_FAMILY_SUBSIDY_HEADER,
    sources.COUPON_DIGITAL_SUBSIDY_HEADER,
)
REMARK_INDEX = shared_index("备注")
DETAIL_INDEX = shared_index("详细情况")
PAYMENT_STATUS_INDEX = shared_index("回款情况")
UPLOADED_REMARK = "已上传"
PAID_STATUS = "已回款"

REFERENCE_REPORT_CORRECTED = "已自动纠正"
REFERENCE_REPORT_UNRESOLVED = "无唯一候选"
REFERENCE_REPORT_COLLISION = "目标冲突"
REFERENCE_REPORT_ORDER = {
    REFERENCE_REPORT_CORRECTED: 0,
    REFERENCE_REPORT_COLLISION: 1,
    REFERENCE_REPORT_UNRESOLVED: 2,
}


class ReferenceDecision(NamedTuple):
    outcome: str
    document_number: str
    document_date: object
    raw_reference: str
    note: str

# ---------------------------------------------------------------------------
# 参考号纠正（明细摘要 → 已上传宇宙中的 \d{11}N）
#
# 流水线（家电在此之前还有补充文件一步，见 appliance.fill_coupon_reference_supplement）：
#   1. 已在宇宙 → 跳过
#   2. 分段求候选（见 reference_correction_candidates）
#   3. 唯一候选且无冲突 → 写回；否则保留原值并记决策
#
# 合法格式：11 位数字 + 大写 N（processors.common.references.REFERENCE_RE）。
# ---------------------------------------------------------------------------

# 连续非字母数字（空格、标点、中文等）视为「两种数据粘在一格」的分隔。
_SUMMARY_FIELD_SPLIT_RE = re.compile(r"[^0-9A-Za-z]+")
# 从大写 N 往前取 11 位数字；N 后可继续粘型号等，故不加结尾边界。
_EMBEDDED_REFERENCE_RE = re.compile(r"(?<!\d)(\d{11}N)")


def as_currency(amount: Decimal) -> Decimal:
    """Round monetary comparisons to cents."""
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


@dataclass(frozen=True)
class ReferenceCorrectionIndex:
    """Precomputed lookup tables over the uploaded-reference universe."""

    references: frozenset[str]
    digit_prefixes: dict[str, frozenset[str]]
    deletion_variants: dict[str, frozenset[str]]
    substitution_variants: dict[tuple[int, str], frozenset[str]]


def build_reference_correction_index(
    reference_universe: set[str],
) -> ReferenceCorrectionIndex:
    """Index every legal uploaded reference for O(1)-ish dirty matching.

    digit_prefixes: 11 digits → {that digits + N}
    deletion_variants: reference with one char removed → originals (len 11 path)
    substitution_variants: (pos, reference-without-char-at-pos) → originals
        (used for len-12 dirty strings: wrong letter, extra char, etc.)
    """
    references = frozenset(reference.upper() for reference in reference_universe)
    invalid_references = sorted(
        reference
        for reference in references
        if not REFERENCE_RE.fullmatch(reference)
    )
    if invalid_references:
        raise ValueError(
            "参考号格式应为11位数字后跟大写字母 N，"
            f"实际存在无效参考号：{invalid_references[:5]}"
        )
    digit_prefixes: dict[str, set[str]] = {}
    deletion_variants: dict[str, set[str]] = {}
    substitution_variants: dict[tuple[int, str], set[str]] = {}
    for reference in references:
        digit_prefixes.setdefault(reference[:11], set()).add(reference)
        for index in range(len(reference)):
            without_character = reference[:index] + reference[index + 1:]
            deletion_variants.setdefault(without_character, set()).add(
                reference
            )
            substitution_variants.setdefault(
                (index, without_character),
                set(),
            ).add(reference)
    return ReferenceCorrectionIndex(
        references=references,
        digit_prefixes={
            prefix: frozenset(candidates)
            for prefix, candidates in digit_prefixes.items()
        },
        deletion_variants={
            variant: frozenset(candidates)
            for variant, candidates in deletion_variants.items()
        },
        substitution_variants={
            variant: frozenset(candidates)
            for variant, candidates in substitution_variants.items()
        },
    )


def _resolve_correction_index(
    reference_universe: set[str] | ReferenceCorrectionIndex,
) -> ReferenceCorrectionIndex:
    if isinstance(reference_universe, ReferenceCorrectionIndex):
        return reference_universe
    return build_reference_correction_index(reference_universe)


def _split_summary_fields(raw_reference: str) -> list[str]:
    """Split a 明细摘要 cell on non-alphanumeric runs; drop empty parts."""
    return [part for part in _SUMMARY_FIELD_SPLIT_RE.split(raw_reference) if part]


def _candidates_embedded_legal_reference(
    upper_text: str,
    references: frozenset[str],
) -> set[str]:
    """Embedded \\d{11}N inside longer text (model glued after N, Chinese prefix…)."""
    return {
        token
        for token in _EMBEDDED_REFERENCE_RE.findall(upper_text)
        if token in references
    }


def _candidates_exact_cleaned(
    upper_text: str,
    references: frozenset[str],
) -> set[str]:
    """Whole field after stripping non [0-9A-Z] equals a universe member."""
    cleaned = re.sub(r"[^0-9A-Z]", "", upper_text)
    if cleaned in references:
        return {cleaned}
    return set()


def _candidates_from_compact_edits(
    compact: str,
    index: ReferenceCorrectionIndex,
) -> set[str]:
    """Match compact (whitespace-stripped) text via length-11/12/13 edit paths.

    - exactly 11 digits → digit_prefixes (missing trailing N)
    - length 11 → one deletion away from a universe member
    - length 12 → one substitution/deletion (wrong suffix letter, etc.)
    - length 13 → drop one character
    """
    candidates: set[str] = set()
    if re.fullmatch(r"\d{11}", compact):
        candidates.update(index.digit_prefixes.get(compact, ()))
    if len(compact) == 11:
        candidates.update(index.deletion_variants.get(compact, ()))
    elif len(compact) == 13:
        for character_index in range(13):
            candidate = compact[:character_index] + compact[character_index + 1:]
            if candidate in index.references:
                candidates.add(candidate)
    elif len(compact) == 12:
        for character_index in range(12):
            without_character = (
                compact[:character_index] + compact[character_index + 1:]
            )
            candidates.update(
                index.substitution_variants.get(
                    (character_index, without_character),
                    (),
                )
            )
    return candidates


def _candidates_for_single_field(
    field: str,
    index: ReferenceCorrectionIndex,
) -> set[str]:
    """All candidate uploaded references for one fragment (no further split)."""
    upper_text = field.upper()
    compact = re.sub(r"\s+", "", upper_text)
    return (
        _candidates_embedded_legal_reference(upper_text, index.references)
        | _candidates_exact_cleaned(upper_text, index.references)
        | _candidates_from_compact_edits(compact, index)
    )


def reference_correction_candidates(
    raw_reference: str,
    reference_universe: set[str] | ReferenceCorrectionIndex,
) -> set[str]:
    """Uploaded references that a dirty 明细摘要 cell could mean.

    1. Split the cell on non-alphanumeric runs (comma, space, Chinese, …).
    2. Score each fragment with _candidates_for_single_field.
    3. If there were multiple fragments, also score the unsplit cell so a
       single reference written with internal spaces ("12345 678901 N") still
       matches after compacting.
    """
    index = _resolve_correction_index(reference_universe)
    fields = _split_summary_fields(raw_reference)
    if not fields:
        return set()
    if len(fields) == 1:
        return _candidates_for_single_field(fields[0], index)

    candidates: set[str] = set()
    for field in fields:
        candidates.update(_candidates_for_single_field(field, index))
    candidates.update(_candidates_for_single_field(raw_reference, index))
    return candidates


def coupon_data_rows(
    rows: list[list[object]],
    excluded_bottom_rows: int,
) -> list[list[object]]:
    """Return the coupon detail rows, excluding the header and any trailing
    rows already matched (and pinned to the bottom) by an earlier pass."""
    return rows[1:-excluded_bottom_rows] if excluded_bottom_rows > 0 else rows[1:]


def fill_coupon_remarks(
    rows: list[list[object]],
    remark_lookup: dict[tuple[str, date], str],
    subsidy_label: str,
) -> tuple[int, Decimal]:
    """Fill 备注 from the receipt remark lookup and split rows into
    unmatched/matched order (matched rows sink to the bottom, to be pinned
    pink by the caller).
    """
    matched_rows: list[list[object]] = []
    unmatched_rows: list[list[object]] = []
    matched_subsidy_total = Decimal("0")
    for row in rows[1:]:
        document_date = row[DATE_INDEX]
        if not isinstance(document_date, date):
            raise ValueError(f"行单据日期必须为 date 类型，实际为 {document_date!r}")
        key = (normalize_document_number(row[DOCUMENT_INDEX]), document_date)
        remark = remark_lookup.get(key, "")
        row[REMARK_INDEX] = remark
        if remark:
            matched_rows.append(row)
            subsidy = row[SUBSIDY_INDEX]
            if subsidy not in (None, ""):
                try:
                    matched_subsidy_total += Decimal(str(subsidy))
                except InvalidOperation as error:
                    raise ValueError(
                        f"组合键 {key} 的{subsidy_label}金额无效：{subsidy!r}"
                    ) from error
        else:
            unmatched_rows.append(row)
    rows[1:] = [*unmatched_rows, *matched_rows]
    return len(matched_rows), matched_subsidy_total


def reference_decision(
    outcome: str,
    row: list[object],
    raw_reference: str,
    note: str,
) -> ReferenceDecision:
    """Identify a decision by 单据号 + 单据日期, not by row position.

    The detail rows get re-sorted after corrections are applied, so a row
    number recorded here would point at the wrong row in the saved sheet.
    """
    return ReferenceDecision(
        outcome=outcome,
        document_number=normalize_document_number(row[DOCUMENT_INDEX]),
        document_date=row[DATE_INDEX],
        raw_reference=raw_reference,
        note=note,
    )


def _propose_corrections(
    included_rows: list[list[object]],
    correction_index: ReferenceCorrectionIndex,
    reference_universe: set[str],
    protected: set[int],
) -> tuple[dict[int, str], Counter[str], list[ReferenceDecision]]:
    """Phase 1: propose a unique correction candidate per dirty row."""
    proposed: dict[int, str] = {}
    target_counts: Counter[str] = Counter()
    decisions: list[ReferenceDecision] = []

    for row_index, row in enumerate(included_rows, start=1):
        if id(row) in protected:
            continue
        raw_reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        if not raw_reference or raw_reference in reference_universe:
            continue

        candidates = reference_correction_candidates(
            raw_reference, correction_index
        )
        if len(candidates) != 1:
            # Legal-but-missing refs are left off the report (see docstring).
            if candidates or not REFERENCE_RE.fullmatch(raw_reference):
                decisions.append(
                    reference_decision(
                        REFERENCE_REPORT_UNRESOLVED,
                        row,
                        raw_reference,
                        f"候选数量 {len(candidates)}，"
                        f"未在已上传数据中找到唯一匹配，保留原值",
                    )
                )
            continue

        target = next(iter(candidates))
        proposed[row_index] = target
        target_counts[target] += 1

    return proposed, target_counts, decisions


def _apply_corrections(
    rows: list[list[object]],
    proposed: dict[int, str],
    target_counts: Counter[str],
    existing_counts: Counter[str],
) -> list[ReferenceDecision]:
    """Phase 2: apply collision-free proposals, record collisions."""
    decisions: list[ReferenceDecision] = []
    for row_index, target in proposed.items():
        row = rows[row_index]
        raw_reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        if existing_counts[target] > 0 or target_counts[target] > 1:
            decisions.append(
                reference_decision(
                    REFERENCE_REPORT_COLLISION,
                    row,
                    raw_reference,
                    f"目标参考号 {target} 已被其他行占用，未纠正",
                )
            )
            continue
        row[SUMMARY_INDEX] = target
        decisions.append(
            reference_decision(
                REFERENCE_REPORT_CORRECTED,
                row,
                raw_reference,
                f"已自动纠正为 {target}，请人工复核",
            )
        )
    return decisions


def correct_coupon_references(
    rows: list[list[object]],
    reference_universe: set[str],
    excluded_bottom_rows: int = 0,
    protected_row_ids: set[int] | None = None,
) -> list[ReferenceDecision]:
    """Rewrite dirty 明细摘要 values to uploaded references; return audit trail.

    Two phases:
      propose — unique candidate per row (skip protected / already in universe)
      apply   — write only if the target is not already used by another row

    Report policy for non-unique candidates:
      - multiple candidates, or zero with a non-legal raw string → 无唯一候选
      - zero candidates but raw is already legal \\d{11}N → unsubmitted only
        (omitted from the report sheet)
    """
    included_end = len(rows) - excluded_bottom_rows
    included_rows = rows[1:included_end]
    existing_counts = Counter(
        normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        for row in included_rows
        if normalize_receipt_identifier(row[SUMMARY_INDEX])
    )
    correction_index = build_reference_correction_index(reference_universe)
    protected = protected_row_ids or set()

    proposed, target_counts, propose_decisions = _propose_corrections(
        included_rows, correction_index, reference_universe, protected
    )
    apply_decisions = _apply_corrections(
        rows, proposed, target_counts, existing_counts
    )

    decisions = propose_decisions + apply_decisions
    decisions.sort(key=lambda d: REFERENCE_REPORT_ORDER[d.outcome])
    return decisions


def fill_upload_statuses(
    rows: list[list[object]],
    detail_lookup: dict[str, str],
    excluded_bottom_rows: int = 0,
) -> tuple[int, int]:
    """Fill 已上传/未上传 and uploaded detail without considering payment."""
    uploaded_count = 0
    unmatched_count = 0
    for row in coupon_data_rows(rows, excluded_bottom_rows):
        reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        detail = detail_lookup.get(reference, "")
        row[DETAIL_INDEX] = detail
        if detail:
            row[REMARK_INDEX] = UPLOADED_REMARK
            uploaded_count += 1
        else:
            row[REMARK_INDEX] = "未上传"
            unmatched_count += 1
    return uploaded_count, unmatched_count


def fill_payment_statuses(
    rows: list[list[object]],
    payment_references: set[str] | frozenset[str],
    excluded_bottom_rows: int = 0,
) -> int:
    """Match payment only for rows already marked 已上传."""
    paid_count = 0
    for row in coupon_data_rows(rows, excluded_bottom_rows):
        row[PAYMENT_STATUS_INDEX] = None
        if row[REMARK_INDEX] != UPLOADED_REMARK:
            continue
        reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        if reference in payment_references:
            row[PAYMENT_STATUS_INDEX] = PAID_STATUS
            paid_count += 1
    return paid_count
