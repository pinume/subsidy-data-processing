"""家电 side of 销售用券情况统计 processing: the superset report.

家电 is a strict superset of 数码's report: it additionally reads an optional
reference-supplement file, builds 财务大类/品牌 group sheets, and closes
数据汇总 with a five-column (财务大类, 品牌, 上传状态, 数量, 2026国补金额)
summary instead
of 数码's three-column one. Those are real differences in what gets built,
not just different constants, so 数码 keeps its own module
(processors/coupons/digital.py) rather than being forced through this one
with a bundle of feature flags.
"""

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from python_calamine import CalamineWorkbook

from processors.common.dates import (
    normalize_coupon_date,
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.common.excel import (
    calamine_rows,
)
from processors.common.references import validated_reference
from processors.receipts import OUTPUT_FILE as RECEIPTS_OUTPUT_FILE
from processors.submitted import PROFILES as SUBMITTED_PROFILES

from . import matching, sources
from .matching import (
    as_currency,
    coupon_data_rows,
)
from .report_contract import SUMMARY_HEADER, SUMMARY_SHEET_NAME
from .sources import load_coupon_remark_lookup, load_uploaded_summary
from .validation import (
    validate_detail_rows_shape,
    validate_document_and_date_values,
    validate_matched_subsidy_total,
    validate_payment_statuses,
    validate_remark_and_detail_values,
    validate_uploaded_and_unmatched_counts,
)

DETAILS_SHEET_NAME = "家电-明细总表"
# Owned here only to keep this module's own group-sheet titles from
# colliding with processors.coupons.digital's / coupon_report.py's sheets.
DIGITAL_DETAILS_SHEET_NAME = "数码-明细总表"
REFERENCE_REPORT_SHEET_NAME = "Processing Report"

COUPON_SUBSIDY_HEADER = sources.COUPON_FAMILY_SUBSIDY_HEADER
COUPON_REMARK_SOURCE_FILE = RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = SUBMITTED_PROFILES["家电"].output_file

# Rows misclassified as digital are excluded from the Summary aggregation
# and never get their own category-brand sheet, though they still appear
# in the Details sheet.
COUPON_EXCLUDED_CATEGORY = "数码"
COUPON_OUTPUT_HEADER = sources.APPLIANCE_PROFILE.output_header
COUPON_GROUP_HEADER = (
    "单据号",
    "单据日期",
    "商品名称",
    COUPON_SUBSIDY_HEADER,
    "备注",
    "详细情况",
    "回款情况",
)
COUPON_GROUP_COLUMN_INDEXES = tuple(
    COUPON_OUTPUT_HEADER.index(header) for header in COUPON_GROUP_HEADER
)
COUPON_DATE_INDEX = COUPON_OUTPUT_HEADER.index("单据日期")
COUPON_PRODUCT_NAME_INDEX = COUPON_OUTPUT_HEADER.index("商品名称")
COUPON_BRAND_INDEX = COUPON_OUTPUT_HEADER.index("品牌")
COUPON_CATEGORY_INDEX = COUPON_OUTPUT_HEADER.index("财务大类")
COUPON_SUMMARY_COLUMN_INDEX = COUPON_OUTPUT_HEADER.index("明细摘要")
COUPON_SUBSIDY_INDEX = COUPON_OUTPUT_HEADER.index(COUPON_SUBSIDY_HEADER)
COUPON_REMARK_INDEX = COUPON_OUTPUT_HEADER.index("备注")
COUPON_DETAIL_INDEX = COUPON_OUTPUT_HEADER.index("详细情况")
COUPON_MATCH_FILL_COLOR = "FFC7CE"
COUPON_BRAND_REPLACEMENTS = sources.COUPON_BRAND_REPLACEMENTS
# 数据汇总 covers both projects with the generic 2026国补金额 label, so its
# header is not derived from this project's subsidy header.
COUPON_SUMMARY_HEADER = SUMMARY_HEADER
# 财务大类 for this project's 已上传/未上传/合计 block at the foot of 数据汇总.
COUPON_SUMMARY_PROJECT_LABEL = "家电"
COUPON_REMARK_SORT_PRIORITY = {
    "未上传": 0,
    "已上传": 1,
}
INVALID_SHEET_TITLE_RE = re.compile(r"[\[\]:*?/\\]")
COUPON_REFERENCE_SUPPLEMENT_HEADER = ("参考号", "单据号", "单据日期")


def normalize_coupon_reference_supplement_header(value: object) -> str:
    text = str(value or "").strip()
    if text in COUPON_REFERENCE_SUPPLEMENT_HEADER:
        return text
    try:
        return text.encode("cp949").decode("gb18030")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def load_coupon_reference_supplement(
    source: Path,
) -> dict[tuple[str, date], frozenset[str]]:
    if not source.exists():
        print(f"Optional reference supplement file not found; skipping: {source}")
        return {}

    workbook = CalamineWorkbook.from_path(str(source))
    try:
        sheet = workbook.get_sheet_by_index(0)
        rows_iter = calamine_rows(sheet)
        header = tuple(
            normalize_coupon_reference_supplement_header(value)
            for value in next(rows_iter, [])
        )
        if header != COUPON_REFERENCE_SUPPLEMENT_HEADER:
            raise ValueError(
                f"{source.name} 字段标题不符合要求：实际为 {header}"
            )

        references_by_key: dict[tuple[str, date], set[str]] = {}
        for row_number, row in enumerate(rows_iter, start=2):
            if all(value in (None, "") for value in row):
                continue
            reference = validated_reference(
                row[0], f"{source.name} 第 {row_number} 行"
            )
            document_number = normalize_document_number(row[1])
            if not document_number:
                raise ValueError(
                    f"{source.name} 第 {row_number} 行单据号为空"
                )
            document_date = normalize_coupon_date(
                row[2], row_number, source.name
            )
            key = (document_number, document_date)
            references_by_key.setdefault(key, set()).add(reference)
        return {
            key: frozenset(references)
            for key, references in references_by_key.items()
        }
    finally:
        workbook.close()


def fill_coupon_reference_supplement(
    rows: list[list[object]],
    reference_lookup: dict[tuple[str, date], frozenset[str]],
    reference_universe: set[str],
    excluded_bottom_rows: int,
) -> tuple[int, int, set[int], Counter[tuple[str, date, str]]]:
    summary_index = COUPON_SUMMARY_COLUMN_INDEX
    included_rows = coupon_data_rows(rows, excluded_bottom_rows)
    matched_count = 0
    ambiguous_count = 0
    matched_row_ids: set[int] = set()
    matched_values: Counter[tuple[str, date, str]] = Counter()
    for row in included_rows:
        current_reference = normalize_receipt_identifier(
            row[summary_index]
        ).upper()
        if current_reference in reference_universe:
            continue
        key = (normalize_document_number(row[0]), row[1])
        references = reference_lookup.get(key)
        if references is None:
            continue
        if len(references) == 1:
            reference = next(iter(references))
        elif current_reference in references:
            reference = current_reference
        else:
            ambiguous_count += 1
            continue
        row[summary_index] = reference
        matched_count += 1
        matched_row_ids.add(id(row))
        matched_values[(key[0], key[1], reference)] += 1
    return (
        matched_count,
        ambiguous_count,
        matched_row_ids,
        matched_values,
    )


def coupon_text_sort_value(value: object) -> str:
    return str(value or "").strip()


def coupon_date_sort_value(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.max


def coupon_pink_sort_key(row: list[object] | tuple[object, ...]) -> tuple:
    return (
        coupon_text_sort_value(row[COUPON_CATEGORY_INDEX]),
        coupon_text_sort_value(row[COUPON_BRAND_INDEX]),
        coupon_text_sort_value(row[COUPON_PRODUCT_NAME_INDEX]),
        coupon_date_sort_value(row[COUPON_DATE_INDEX]),
    )


def coupon_regular_sort_key(row: list[object] | tuple[object, ...]) -> tuple:
    remark = coupon_text_sort_value(row[COUPON_REMARK_INDEX])
    return (
        COUPON_REMARK_SORT_PRIORITY.get(remark, 2),
        remark,
        coupon_text_sort_value(row[COUPON_CATEGORY_INDEX]),
        coupon_text_sort_value(row[COUPON_BRAND_INDEX]),
        coupon_text_sort_value(row[COUPON_DETAIL_INDEX]),
        coupon_date_sort_value(row[COUPON_DATE_INDEX]),
        coupon_text_sort_value(row[COUPON_PRODUCT_NAME_INDEX]),
    )


def coupon_group_regular_sort_key(
    row: list[object] | tuple[object, ...],
) -> tuple:
    remark = coupon_text_sort_value(row[COUPON_REMARK_INDEX])
    priority = {"已上传": 0, "未上传": 1}.get(remark, 2)
    return (
        priority,
        remark,
        coupon_text_sort_value(row[COUPON_DETAIL_INDEX]),
        coupon_date_sort_value(row[COUPON_DATE_INDEX]),
        coupon_text_sort_value(row[COUPON_PRODUCT_NAME_INDEX]),
    )


def select_coupon_group_columns(
    row: list[object] | tuple[object, ...],
) -> tuple[object, ...]:
    return tuple(row[index] for index in COUPON_GROUP_COLUMN_INDEXES)


def sort_coupon_detail_rows(
    rows: list[list[object]],
    pink_row_count: int,
) -> None:
    pink_start = len(rows) - pink_row_count
    regular_rows = rows[1:pink_start] if pink_row_count else rows[1:]
    pink_rows = rows[pink_start:] if pink_row_count else []
    regular_rows.sort(key=coupon_regular_sort_key)
    pink_rows.sort(key=coupon_pink_sort_key)
    rows[1:] = [*regular_rows, *pink_rows]


def coupon_subsidy_count(amount: Decimal) -> int:
    """Count a coupon row as +1, or as -1 when its 国补 is a reversal.

    A return or exchange is recorded as a negative amount that cancels the
    original row, so counting it as -1 keeps the count consistent with the
    total. A zero amount should not occur; it is bad source data and is
    counted as 0 rather than silently inflating the count.
    """
    if amount > 0:
        return 1
    if amount < 0:
        return -1
    return 0


def build_coupon_summary(
    rows: list[list[object]],
    excluded_bottom_rows: int,
    uploaded_subsidy_count: int,
    uploaded_subsidy_total: Decimal,
) -> tuple[list[tuple[object, ...]], int]:
    """Build the 数据汇总 table.

    The table lists 财务大类 / 品牌 / 备注 groups, then closes with 已上传 /
    未上传 / 合计: 已上传 comes from the generated 已上传 workbook's 补贴金额,
    合计 is this coupon file's own 国补 总, and 未上传 is the difference —
    the same three-way split digital already reports.

    Also returns how many rows carried a zero 国补, which is invalid source
    data worth telling the operator about.
    """
    category_index = COUPON_CATEGORY_INDEX
    brand_index = COUPON_BRAND_INDEX
    remark_index = COUPON_REMARK_INDEX
    subsidy_index = COUPON_SUBSIDY_INDEX
    included_rows = coupon_data_rows(rows, excluded_bottom_rows)
    grouped_counts: Counter[tuple[str, str, str]] = Counter()
    grouped_totals: dict[tuple[str, str, str], Decimal] = {}
    coupon_count = 0
    coupon_total = Decimal("0")
    zero_subsidy_count = 0
    for row in included_rows:
        category = str(row[category_index] or "").strip()
        if category == COUPON_EXCLUDED_CATEGORY:
            continue
        brand = str(row[brand_index] or "").strip()
        remark = str(row[remark_index] or "").strip()
        key = (category, brand, remark)
        grouped_counts[key] += 1
        grouped_totals.setdefault(key, Decimal("0"))
        subsidy = row[subsidy_index]
        if subsidy not in (None, ""):
            try:
                amount = Decimal(str(subsidy))
            except InvalidOperation as error:
                raise ValueError(
                    f"{key!r} 的2026家电国补金额无效：{subsidy!r}"
                ) from error
            grouped_totals[key] += amount
            coupon_total += amount
            coupon_count += coupon_subsidy_count(amount)
            if amount == 0:
                zero_subsidy_count += 1

    summary_rows = [
        (
            *key,
            grouped_counts[key],
            float(grouped_totals[key].quantize(Decimal("0.01"))),
        )
        for key in sorted(grouped_counts)
    ]
    # Labelled 财务大类=家电 with the status in 备注, so this block reads the
    # same way as the 数码 block appended after it in the merged 审核明细
    # workbook (see processors/coupon_report.py).
    summary_rows.extend(
        (
            (
                COUPON_SUMMARY_PROJECT_LABEL,
                None,
                "已上传",
                uploaded_subsidy_count,
                float(as_currency(uploaded_subsidy_total)),
            ),
            (
                COUPON_SUMMARY_PROJECT_LABEL,
                None,
                "未上传",
                coupon_count - uploaded_subsidy_count,
                float(as_currency(coupon_total - uploaded_subsidy_total)),
            ),
            (
                COUPON_SUMMARY_PROJECT_LABEL,
                None,
                "合计",
                coupon_count,
                float(as_currency(coupon_total)),
            ),
        )
    )
    return summary_rows, zero_subsidy_count


def coupon_group_sheet_title(
    category: str,
    brand: str,
    used_titles: set[str],
) -> str:
    base_title = f"{category or '未分类'}-{brand or '未品牌'}"
    base_title = INVALID_SHEET_TITLE_RE.sub("_", base_title).strip("'")
    base_title = base_title[:31] or "未分类-未品牌"
    title = base_title
    suffix = 2
    used_title_keys = {used_title.casefold() for used_title in used_titles}
    while title.casefold() in used_title_keys:
        suffix_text = f"-{suffix}"
        title = f"{base_title[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_titles.add(title)
    return title


def build_coupon_group_sheets(
    rows: list[list[object]],
    excluded_bottom_rows: int,
) -> list[tuple[str, str, str, list[tuple[list[object], bool]]]]:
    category_index = COUPON_CATEGORY_INDEX
    brand_index = COUPON_BRAND_INDEX
    first_pink_index = len(rows) - excluded_bottom_rows
    groups: dict[
        tuple[str, str],
        list[tuple[list[object], bool]],
    ] = {}
    for row_index, row in enumerate(rows[1:], start=1):
        category = str(row[category_index] or "").strip()
        if category == COUPON_EXCLUDED_CATEGORY:
            continue
        brand = str(row[brand_index] or "").strip()
        groups.setdefault((category, brand), []).append(
            (row, row_index >= first_pink_index)
        )

    for grouped_rows in groups.values():
        regular_rows = [item for item in grouped_rows if not item[1]]
        pink_rows = [item for item in grouped_rows if item[1]]
        regular_rows.sort(key=lambda item: coupon_group_regular_sort_key(item[0]))
        pink_rows.sort(key=lambda item: coupon_pink_sort_key(item[0]))
        grouped_rows[:] = [*regular_rows, *pink_rows]

    used_titles = {
        SUMMARY_SHEET_NAME,
        DETAILS_SHEET_NAME,
        DIGITAL_DETAILS_SHEET_NAME,
        REFERENCE_REPORT_SHEET_NAME,
    }
    return [
        (
            coupon_group_sheet_title(category, brand, used_titles),
            category,
            brand,
            grouped_rows,
        )
        for (category, brand), grouped_rows in sorted(groups.items())
    ]


def coupon_summary_group_merges(
    rows: list[tuple[object, ...]],
    row_count: int,
    header_rows: int = 1,
) -> list[tuple[int, int, int]]:
    """Vertical runs of 财务大类 and 品牌 worth merging.

    Returns (first_row, last_row, column) in 1-based sheet coordinates, for
    runs longer than one row only. Computing the ranges instead of merging in
    place lets a writer that cannot read back what it has written — XlsxWriter
    — produce the same sheet as openpyxl did.
    """
    merges: list[tuple[int, int, int]] = []
    category_column = 1
    brand_column = 2
    for column in (brand_column, category_column):
        index = 0
        while index < row_count:
            category = rows[index][0]
            value = rows[index][column - 1]
            end = index
            while (
                end + 1 < row_count
                and rows[end + 1][0] == category
                and rows[end + 1][column - 1] == value
            ):
                end += 1
            if end > index:
                merges.append(
                    (index + 1 + header_rows, end + 1 + header_rows, column)
                )
            index = end + 1
    return merges


def project_summary_blocks(
    rows: list[tuple[object, ...]],
) -> list[tuple[int, int]]:
    """Locate the trailing per-project 已上传/未上传/合计 blocks.

    They are the rows carrying a 财务大类 but no 品牌 (家电 first, then any
    project appended by processors/coupon_report.py). Returned as inclusive
    0-based (start, end) spans so the caller can merge 财务大类 and 品牌 into
    one cell across each block — the split is meaningless there, and leaving
    an empty 品牌 column beside the label reads as a missing value.
    """
    blocks: list[tuple[int, int]] = []
    start: int | None = None
    for index, row in enumerate(rows):
        if row[1] is None:
            if start is None:
                start = index
            elif rows[start][0] != row[0]:
                blocks.append((start, index - 1))
                start = index
        elif start is not None:
            blocks.append((start, index - 1))
            start = None
    if start is not None:
        blocks.append((start, len(rows) - 1))
    return blocks


def coupon_summary_project_merges(
    blocks: list[tuple[int, int]],
    header_rows: int = 1,
) -> list[tuple[int, int, int, int]]:
    """Each project block as (first_row, last_row, first_column, last_column).

    The 财务大类/品牌 split is meaningless on these trailing 已上传/未上传/合计
    rows, so both columns become one cell.
    """
    return [
        (start + 1 + header_rows, end + 1 + header_rows, 1, 2)
        for start, end in blocks
    ]


def merged_coupon_summary_values(
    rows: list[tuple[object, ...]],
) -> list[tuple[object, ...]]:
    merged_rows: list[tuple[object, ...]] = []
    previous_category = None
    previous_brand = None
    for row in rows:
        if row[0] == "合计":
            merged_rows.append(row)
            continue
        category, brand, *remaining = row
        displayed_category = (
            None if category == previous_category else category
        )
        displayed_brand = (
            None
            if category == previous_category and brand == previous_brand
            else brand
        )
        merged_rows.append(
            (displayed_category, displayed_brand, *remaining)
        )
        previous_category = category
        previous_brand = brand
    return merged_rows


@dataclass
class CouponComputation:
    """Everything computed for the 家电 coupon data, before it is written to
    a sheet. Kept separate from workbook assembly so the merged 审核明细
    workbook (built in processors/coupon_report.py) can interleave a 数码
    sheet between this project's own sheets without either project knowing
    about the other at import time."""

    rows: list[list[object]]
    data_row_count: int
    matched_count: int
    matched_subsidy_total: Decimal
    remark_lookup: dict[tuple[str, date], str]
    detail_lookup: dict[str, str]
    reference_universe: set[str]
    payment_references: frozenset[str]
    payment_match_count: int
    reference_supplement_count: int
    ambiguous_reference_supplement_count: int
    reference_supplement_matches: Counter[tuple[str, date, str]]
    corrected_count: int
    unresolved_count: int
    correction_collision_count: int
    reference_decisions: list[matching.ReferenceDecision]
    final_unresolved_reference_count: int
    uploaded_count: int
    unmatched_count: int
    excluded_category_row_count: int
    uploaded_subsidy_count: int
    uploaded_subsidy_total: Decimal
    zero_subsidy_count: int
    source_total: Decimal | None
    computed_total: Decimal
    summary_rows: list[tuple[object, ...]]
    group_sheets: list[tuple[str, str, str, list[tuple[list[object], bool]]]]


_SOURCE_TOTAL_UNSET = object()


def compute_coupon_data(
    *,
    rows: list[list[object]] | None = None,
    remark_lookup: dict[tuple[str, date], str] | None = None,
    payment_reference_locations: dict[str, str] | None = None,
    source_total: Decimal | None | object = _SOURCE_TOTAL_UNSET,
) -> CouponComputation:
    if sources.COUPON_SOURCE_FILE is None:
        raise FileNotFoundError(
            f"未在 {sources.DATA_DIR} 中找到文件名包含"
            f"“{sources.COUPON_STATISTICS_KEYWORD}”且表头为"
            f"“{COUPON_SUBSIDY_HEADER}”的 .XLSX 文件"
        )

    if remark_lookup is None:
        remark_lookup = load_coupon_remark_lookup(COUPON_REMARK_SOURCE_FILE)
    if rows is None:
        rows = sources.read_coupon_rows(
            sources.COUPON_SOURCE_FILE,
            sources.APPLIANCE_PROFILE,
            remark_lookup=remark_lookup,
        )
    matched_count, matched_subsidy_total = matching.fill_coupon_remarks(
        rows,
        remark_lookup,
        "2026家电国补",
    )
    detail_lookup, uploaded_subsidy_count, uploaded_subsidy_total = (
        load_uploaded_summary(COUPON_UPLOADED_SOURCE_FILE)
    )
    # Unsubmitted data is no longer supplied, so submitted data is the only
    # source of valid references.
    reference_universe = set(detail_lookup)
    payment_reference_locations = payment_reference_locations or {}
    sources.validate_payment_reference_subset(
        "家电",
        payment_reference_locations,
        reference_universe,
    )
    payment_references = frozenset(payment_reference_locations)
    reference_supplement_lookup = load_coupon_reference_supplement(
        sources.COUPON_REFERENCE_SUPPLEMENT_FILE
    )
    (
        reference_supplement_count,
        ambiguous_reference_supplement_count,
        reference_supplement_row_ids,
        reference_supplement_matches,
    ) = fill_coupon_reference_supplement(
        rows,
        reference_supplement_lookup,
        reference_universe,
        matched_count,
    )
    (
        corrected_count,
        unresolved_count,
        correction_collision_count,
        reference_decisions,
    ) = matching.correct_coupon_references(
        rows,
        reference_universe,
        matched_count,
        reference_supplement_row_ids,
    )
    summary_index = COUPON_SUMMARY_COLUMN_INDEX
    regular_rows = coupon_data_rows(rows, matched_count)
    final_unresolved_reference_count = sum(
        reference not in reference_universe
        for reference in (
            normalize_receipt_identifier(row[summary_index]).upper()
            for row in regular_rows
        )
        if reference
    )
    uploaded_count, unmatched_count = matching.fill_upload_statuses(
        rows,
        detail_lookup,
        matched_count,
    )
    payment_match_count = matching.fill_payment_statuses(
        rows,
        payment_references,
        matched_count,
    )
    sort_coupon_detail_rows(rows, matched_count)
    excluded_category_row_count = sum(
        1
        for row in coupon_data_rows(rows, matched_count)
        if str(row[COUPON_CATEGORY_INDEX] or "").strip()
        == COUPON_EXCLUDED_CATEGORY
    )
    summary_rows, zero_subsidy_count = build_coupon_summary(
        rows,
        matched_count,
        uploaded_subsidy_count,
        uploaded_subsidy_total,
    )
    group_sheets = build_coupon_group_sheets(rows, matched_count)
    if source_total is _SOURCE_TOTAL_UNSET:
        source_total = sources.read_coupon_source_total(
            sources.COUPON_SOURCE_FILE,
            remark_lookup=remark_lookup,
        )
    computed_total = Decimal(str(summary_rows[-1][4]))

    return CouponComputation(
        rows=rows,
        data_row_count=max(len(rows) - 1, 0),
        matched_count=matched_count,
        matched_subsidy_total=matched_subsidy_total,
        remark_lookup=remark_lookup,
        detail_lookup=detail_lookup,
        reference_universe=reference_universe,
        payment_references=payment_references,
        payment_match_count=payment_match_count,
        reference_supplement_count=reference_supplement_count,
        ambiguous_reference_supplement_count=(
            ambiguous_reference_supplement_count
        ),
        reference_supplement_matches=reference_supplement_matches,
        corrected_count=corrected_count,
        unresolved_count=unresolved_count,
        correction_collision_count=correction_collision_count,
        reference_decisions=reference_decisions,
        final_unresolved_reference_count=final_unresolved_reference_count,
        uploaded_count=uploaded_count,
        unmatched_count=unmatched_count,
        excluded_category_row_count=excluded_category_row_count,
        uploaded_subsidy_count=uploaded_subsidy_count,
        uploaded_subsidy_total=uploaded_subsidy_total,
        zero_subsidy_count=zero_subsidy_count,
        source_total=source_total,
        computed_total=computed_total,
        summary_rows=summary_rows,
        group_sheets=group_sheets,
    )


def validate_computation(
    computation: CouponComputation,
    extra_summary_rows: list[tuple[object, ...]] = (),
) -> None:
    """Validate business invariants before serializing the workbook."""
    expected_data_rows = computation.data_row_count
    expected_matched_rows = computation.matched_count
    remark_lookup = computation.remark_lookup
    expected_reference_supplement_matches = (
        computation.reference_supplement_matches
    )
    expected_matched_subsidy_total = computation.matched_subsidy_total
    detail_lookup = computation.detail_lookup
    expected_uploaded_rows = computation.uploaded_count
    reference_universe = computation.reference_universe
    payment_references = computation.payment_references
    expected_paid_rows = computation.payment_match_count
    expected_unresolved_rows = computation.final_unresolved_reference_count
    expected_unmatched_rows = computation.unmatched_count
    expected_excluded_category_rows = computation.excluded_category_row_count
    expected_summary_rows = computation.summary_rows
    combined_summary_rows = [*expected_summary_rows, *extra_summary_rows]

    rows = computation.rows
    validate_detail_rows_shape(rows, COUPON_OUTPUT_HEADER, expected_data_rows)
    validate_payment_statuses(
        rows,
        COUPON_OUTPUT_HEADER,
        payment_references,
        expected_matched_rows,
        expected_paid_rows,
    )
    detail_column, remark_column = validate_uploaded_and_unmatched_counts(
        rows, COUPON_OUTPUT_HEADER, expected_uploaded_rows, expected_unmatched_rows
    )
    matched_start = len(rows) - expected_matched_rows
    summary_column = COUPON_OUTPUT_HEADER.index("明细摘要")
    actual_unresolved_rows = sum(
        reference not in reference_universe
        for reference in (
            normalize_receipt_identifier(row[summary_column]).upper()
            for row in rows[1:matched_start]
        )
        if reference
    )
    if actual_unresolved_rows != expected_unresolved_rows:
        raise RuntimeError(
            f"销售用券未解决参考号数量校验失败：预期 "
            f"{expected_unresolved_rows} 条，实际 "
            f"{actual_unresolved_rows} 条"
        )
    actual_regular_rows = rows[1:matched_start]
    if actual_regular_rows != sorted(
        actual_regular_rows,
        key=coupon_regular_sort_key,
    ):
        raise RuntimeError("销售用券明细非粉色区域排序校验失败")
    matched_partition_rows = rows[matched_start:]
    if matched_partition_rows != sorted(
        matched_partition_rows,
        key=coupon_pink_sort_key,
    ):
        raise RuntimeError("销售用券明细匹配分区排序校验失败")
    brand_column = COUPON_OUTPUT_HEADER.index("品牌")
    remaining_source_brands = {
        str(row[brand_column] or "").strip() for row in rows[1:]
    } & COUPON_BRAND_REPLACEMENTS.keys()
    if remaining_source_brands:
        raise RuntimeError(
            "销售用券品牌替换校验失败："
            f"{sorted(remaining_source_brands)}"
        )
    actual_matched_subsidy_total = Decimal("0")
    actual_reference_supplement_matches: Counter[
        tuple[str, date, str]
    ] = Counter()
    subsidy_column = COUPON_OUTPUT_HEADER.index(COUPON_SUBSIDY_HEADER)
    for row_number, row in enumerate(rows[1:], start=2):
        document = row[0]
        document_date = validate_document_and_date_values(
            document, row[1], row_number
        )
        receipt_remark = remark_lookup.get(
            (
                normalize_document_number(document),
                document_date,
            ),
            "",
        )
        reference = normalize_receipt_identifier(
            row[summary_column]
        ).upper()
        in_matched_partition = (
            expected_matched_rows > 0
            and row_number - 1 >= matched_start
        )
        supplement_match = (
            normalize_document_number(document),
            document_date,
            reference,
        )
        if (
            not in_matched_partition
            and supplement_match
            in expected_reference_supplement_matches
        ):
            actual_reference_supplement_matches[supplement_match] += 1
        if in_matched_partition:
            expected_detail = ""
            expected_remark = receipt_remark
        else:
            expected_detail = detail_lookup.get(reference, "")
            if expected_detail:
                expected_remark = "已上传"
            elif reference not in reference_universe:
                expected_remark = "未上传"
            else:
                expected_remark = receipt_remark
        validate_remark_and_detail_values(
            row[remark_column],
            row[detail_column],
            expected_remark,
            expected_detail,
            row_number,
        )
        if in_matched_partition:
            subsidy = row[subsidy_column]
            if subsidy not in (None, ""):
                actual_matched_subsidy_total += Decimal(str(subsidy))

    if any(
        actual_reference_supplement_matches[match] < expected_count
        for match, expected_count
        in expected_reference_supplement_matches.items()
    ):
        raise RuntimeError(
            "销售用券补充参考号逐行匹配结果校验失败"
        )
    validate_matched_subsidy_total(
        actual_matched_subsidy_total, expected_matched_subsidy_total
    )
    actual_summary_rows = combined_summary_rows
    # The 家电 portion ends in 已上传 / 未上传 / 合计 (in 上传状态, with
    # 财务大类 merged into a single 家电 cell); anything appended after it
    # (digital's rows) is covered by the row-for-row equality check above.
    remark_column = COUPON_SUMMARY_HEADER.index("上传状态")
    tail_start = len(expected_summary_rows) - 3
    if tail_start < 0 or [
        row[remark_column]
        for row in actual_summary_rows[tail_start:tail_start + 3]
    ] != ["已上传", "未上传", "合计"]:
        raise RuntimeError("销售用券汇总缺少已上传/未上传/合计三行")
    brand_summary_rows = actual_summary_rows[:tail_start]
    uploaded_row, unuploaded_row, total_row = actual_summary_rows[
        tail_start:tail_start + 3
    ]
    if sum(row[3] for row in brand_summary_rows) != (
        expected_data_rows
        - expected_matched_rows
        - expected_excluded_category_rows
    ):
        raise RuntimeError("销售用券汇总包含匹配分区数据或数量不完整")
    # 已上传 is measured from the generated 已上传 workbook, so it must equal
    # what that file reports rather than anything derived from the coupon rows.
    if (
        uploaded_row[3] != computation.uploaded_subsidy_count
        or Decimal(str(uploaded_row[4]))
        != as_currency(computation.uploaded_subsidy_total)
    ):
        raise RuntimeError(
            "销售用券汇总“已上传”未与家电_已上传的补贴金额一致："
            f"{uploaded_row[3]} / {uploaded_row[4]}"
        )
    if uploaded_row[3] + unuploaded_row[3] != total_row[3]:
        raise RuntimeError("销售用券汇总合计数量校验失败")
    if Decimal(str(uploaded_row[4])) + Decimal(str(unuploaded_row[4])) != (
        Decimal(str(total_row[4]))
    ):
        raise RuntimeError("销售用券汇总合计金额校验失败")
    # 合计 must be this coupon file's own 国补 total, counted with reversals
    # as -1 so a return and its original cancel out.
    expected_total = Decimal("0")
    expected_count = 0
    for row in coupon_data_rows(computation.rows, expected_matched_rows):
        if str(row[COUPON_CATEGORY_INDEX] or "").strip() == (
            COUPON_EXCLUDED_CATEGORY
        ):
            continue
        subsidy = row[COUPON_SUBSIDY_INDEX]
        if subsidy in (None, ""):
            continue
        amount = Decimal(str(subsidy))
        expected_total += amount
        expected_count += coupon_subsidy_count(amount)
    if total_row[3] != expected_count or Decimal(str(total_row[4])) != (
        as_currency(expected_total)
    ):
        raise RuntimeError(
            "销售用券汇总“合计”未与明细总表的国补合计一致："
            f"预期 {expected_count} / {as_currency(expected_total)}，"
            f"实际 {total_row[3]} / {total_row[4]}"
        )
