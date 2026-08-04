"""Matching logic shared identically in shape by 家电 and 数码 coupon rows.

Both projects' COUPON_OUTPUT_HEADER starts with the same ten columns
(单据号, 单据日期, 商品名称, 品牌, 财务大类, 明细摘要, <subsidy header>, 备注,
详细情况, 回款情况) — only the subsidy header's text differs, never its
position — so every function here indexes rows positionally rather than
taking a header tuple. Both projects exclude the trailing receipt-remark
block already pinned to the bottom by an earlier pass, instead of maintaining
separate matching implementations.
"""

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from processors.common.dates import (
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.common.references import REFERENCE_RE

DOCUMENT_INDEX = 0
DATE_INDEX = 1
SUMMARY_INDEX = 5
SUBSIDY_INDEX = 6
REMARK_INDEX = 7
DETAIL_INDEX = 8
PAYMENT_STATUS_INDEX = 9
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
ReferenceDecision = tuple[str, str, object, str, str]


def as_currency(amount: Decimal) -> Decimal:
    """Round monetary comparisons to cents."""
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


@dataclass(frozen=True)
class ReferenceCorrectionIndex:
    references: frozenset[str]
    digit_prefixes: dict[str, frozenset[str]]
    deletion_variants: dict[str, frozenset[str]]
    substitution_variants: dict[tuple[int, str], frozenset[str]]


def build_reference_correction_index(
    reference_universe: set[str],
) -> ReferenceCorrectionIndex:
    references = frozenset(reference.upper() for reference in reference_universe)
    invalid_references = sorted(
        reference
        for reference in references
        if not REFERENCE_RE.fullmatch(reference)
    )
    if invalid_references:
        raise ValueError(
            "参考号格式应为11位数字后跟一个大写字母，"
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


def reference_correction_candidates(
    raw_reference: str,
    reference_universe: set[str] | ReferenceCorrectionIndex,
) -> set[str]:
    index = (
        reference_universe
        if isinstance(reference_universe, ReferenceCorrectionIndex)
        else build_reference_correction_index(reference_universe)
    )
    references = index.references
    candidates: set[str] = set()
    upper_reference = raw_reference.upper()
    compact = re.sub(r"\s+", "", upper_reference)
    cleaned = re.sub(r"[^0-9A-Z]", "", upper_reference)

    for token in re.findall(
        r"(?<!\d)(\d{11}[A-Z])(?![A-Z0-9])",
        upper_reference,
    ):
        if token in references:
            candidates.add(token)
    if cleaned in references:
        candidates.add(cleaned)
    if re.fullmatch(r"\d{11}", compact):
        candidates.update(index.digit_prefixes.get(compact, ()))
    if len(compact) == 11:
        candidates.update(index.deletion_variants.get(compact, ()))
    elif len(compact) == 13:
        for character_index in range(13):
            candidate = (
                compact[:character_index] + compact[character_index + 1:]
            )
            if candidate in references:
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
        key = (normalize_document_number(row[DOCUMENT_INDEX]), row[DATE_INDEX])
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
    return (
        outcome,
        normalize_document_number(row[DOCUMENT_INDEX]),
        row[DATE_INDEX],
        raw_reference,
        note,
    )


def correct_coupon_references(
    rows: list[list[object]],
    reference_universe: set[str],
    excluded_bottom_rows: int = 0,
    protected_row_ids: set[int] | None = None,
) -> tuple[int, int, int, list[ReferenceDecision]]:
    """Correct references and record every decision for the processing report.

    The universe is built from submitted data only, so an operator has to be
    able to review each applied correction, not just the counts.

    A well-formed reference with no candidate at all is simply absent from the
    submitted data — the detail row already carries a 未上传 remark, so it is
    counted but kept out of the report. Only malformed references and genuine
    ambiguities are reported, which is what an operator can actually act on.
    """
    included_end = len(rows) - excluded_bottom_rows
    included_rows = rows[1:included_end]
    existing_counts = Counter(
        normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        for row in included_rows
        if normalize_receipt_identifier(row[SUMMARY_INDEX])
    )
    proposed: dict[int, str] = {}
    target_counts: Counter[str] = Counter()
    unresolved_count = 0
    decisions: list[ReferenceDecision] = []
    correction_index = build_reference_correction_index(reference_universe)

    for row_index, row in enumerate(included_rows, start=1):
        if protected_row_ids is not None and id(row) in protected_row_ids:
            continue
        raw_reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        if not raw_reference or raw_reference in reference_universe:
            continue
        candidates = reference_correction_candidates(
            raw_reference,
            correction_index,
        )
        if len(candidates) != 1:
            unresolved_count += 1
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

    corrected_count = 0
    collision_count = 0
    for row_index, target in proposed.items():
        row = rows[row_index]
        raw_reference = normalize_receipt_identifier(row[SUMMARY_INDEX]).upper()
        if existing_counts[target] > 0 or target_counts[target] > 1:
            collision_count += 1
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
        corrected_count += 1
        decisions.append(
            reference_decision(
                REFERENCE_REPORT_CORRECTED,
                row,
                raw_reference,
                f"已自动纠正为 {target}，请人工复核",
            )
        )

    decisions.sort(key=lambda decision: REFERENCE_REPORT_ORDER[decision[0]])
    return corrected_count, unresolved_count, collision_count, decisions


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
