"""数码 side of 销售用券情况统计 processing.

数码 has no group sheets, reference-supplement file, or approved-detail
panel, so its summary is a simpler three-column (备注, 数量, 合计) table
instead of 家电's five-column one — see processors/coupons/appliance.py's
module docstring for why this stays a separate module rather than a
feature-flagged variant of that one.
"""

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from processors.common.coupons import (
    as_currency,
    load_coupon_remark_lookup,
    load_uploaded_detail_lookup,
)
from processors.common.excel import (
    format_sheet,
    load_measurement_font,
    load_uploaded_subsidy_stats,
    resolve_font,
)
from processors.common.dates import (
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.receipts import OUTPUT_FILE as RECEIPTS_OUTPUT_FILE
from processors.submitted import PROFILES as SUBMITTED_PROFILES

from . import matching, sources


COUPON_SUBSIDY_HEADER = sources.COUPON_DIGITAL_SUBSIDY_HEADER
COUPON_REMARK_SOURCE_FILE = RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = SUBMITTED_PROFILES["数码"].output_file
COUPON_OUTPUT_HEADER = sources.DIGITAL_PROFILE.output_header
COUPON_MATCH_FILL_COLOR = "FFC7CE"
COUPON_SUMMARY_HEADER = (
    "备注",
    "数量",
    f"{COUPON_SUBSIDY_HEADER}合计",
)
DETAILS_SHEET_NAME = "数码-明细总表"


def build_coupon_summary(
    rows: list[list[object]],
    uploaded_count: int,
    uploaded_subsidy_total: Decimal,
) -> list[tuple[object, ...]]:
    subsidy_index = matching.SUBSIDY_INDEX
    coupon_count = 0
    coupon_total = Decimal("0")
    for row_number, row in enumerate(rows[1:], start=2):
        subsidy = row[subsidy_index]
        if subsidy in (None, ""):
            continue
        try:
            subsidy_amount = Decimal(str(subsidy))
            coupon_total += subsidy_amount
            if subsidy_amount < 0:
                coupon_count -= 1
            elif subsidy_amount > 0:
                coupon_count += 1
        except InvalidOperation as error:
            raise ValueError(
                f"第 {row_number} 行的2026数码国补金额无效：{subsidy!r}"
            ) from error

    unuploaded_count = coupon_count - uploaded_count
    unuploaded_total = coupon_total - uploaded_subsidy_total
    return [
        (
            "已上传",
            uploaded_count,
            float(as_currency(uploaded_subsidy_total)),
        ),
        (
            "未上传",
            unuploaded_count,
            float(as_currency(unuploaded_total)),
        ),
        ("合计", coupon_count, float(as_currency(coupon_total))),
    ]


@dataclass
class CouponComputation:
    """Everything computed for the 数码 coupon data, before it is written to
    a sheet. Digital has no group sheets, approved-detail panel, or
    reference-supplement file, so this is smaller than the 家电 equivalent
    in processors/coupons/appliance.py."""

    rows: list[list[object]]
    data_row_count: int
    matched_count: int
    matched_subsidy_total: Decimal
    remark_lookup: dict[tuple[str, date], str]
    detail_lookup: dict[str, str]
    reference_universe: set[str]
    uploaded_subsidy_count: int
    uploaded_subsidy_total: Decimal
    uploaded_match_count: int
    unmatched_count: int
    corrected_count: int
    unresolved_count: int
    correction_collision_count: int
    reference_decisions: list[tuple[str, str, str, str]]
    summary_rows: list[tuple[object, ...]]


def compute_coupon_data(
    *,
    rows: list[list[object]] | None = None,
    remark_lookup: dict[tuple[str, date], str] | None = None,
) -> CouponComputation:
    if sources.COUPON_SOURCE_FILE is None:
        raise FileNotFoundError(
            f"未在 {sources.DATA_DIR} 中找到文件名包含"
            f"“{sources.COUPON_STATISTICS_KEYWORD}”且表头为"
            f"“{COUPON_SUBSIDY_HEADER}”的 .XLSX 文件"
        )

    if rows is None:
        rows = sources.read_coupon_rows(
            sources.COUPON_SOURCE_FILE, sources.DIGITAL_PROFILE
        )
    if remark_lookup is None:
        remark_lookup = load_coupon_remark_lookup(COUPON_REMARK_SOURCE_FILE)
    matched_count, matched_subsidy_total, _receipt_remark_count = (
        matching.fill_coupon_remarks(rows, remark_lookup, "2026数码国补")
    )
    detail_lookup = load_uploaded_detail_lookup(COUPON_UPLOADED_SOURCE_FILE)
    uploaded_subsidy_count, uploaded_subsidy_total = (
        load_uploaded_subsidy_stats(COUPON_UPLOADED_SOURCE_FILE)
    )
    # Unsubmitted data is no longer supplied, so submitted data is the only
    # source of valid references.
    reference_universe = set(detail_lookup)
    (
        corrected_count,
        unresolved_count,
        correction_collision_count,
        reference_decisions,
    ) = matching.correct_coupon_references(rows, reference_universe)
    uploaded_match_count = matching.fill_uploaded_details(rows, detail_lookup)
    unmatched_count = matching.fill_unmatched_remarks(rows, reference_universe)
    summary_rows = build_coupon_summary(
        rows,
        uploaded_subsidy_count,
        uploaded_subsidy_total,
    )
    if as_currency(matched_subsidy_total) != Decimal("0.00"):
        raise ValueError(
            f"备注匹配行的{COUPON_SUBSIDY_HEADER}合计不为 0："
            f"{matched_subsidy_total}"
        )

    return CouponComputation(
        rows=rows,
        data_row_count=max(len(rows) - 1, 0),
        matched_count=matched_count,
        matched_subsidy_total=matched_subsidy_total,
        remark_lookup=remark_lookup,
        detail_lookup=detail_lookup,
        reference_universe=reference_universe,
        uploaded_subsidy_count=uploaded_subsidy_count,
        uploaded_subsidy_total=uploaded_subsidy_total,
        uploaded_match_count=uploaded_match_count,
        unmatched_count=unmatched_count,
        corrected_count=corrected_count,
        unresolved_count=unresolved_count,
        correction_collision_count=correction_collision_count,
        reference_decisions=reference_decisions,
        summary_rows=summary_rows,
    )


def build_detail_sheet(
    workbook: Workbook,
    computation: CouponComputation,
) -> tuple[str, object, PatternFill]:
    """Append 数码-明细总表 to the (already partially built) 审核明细
    workbook. Returns (font_name, measurement_font, matched_fill) purely for
    symmetry with the 家电 builder; callers processing only digital's sheet
    can ignore them."""
    sheet = workbook.create_sheet(DETAILS_SHEET_NAME)
    for row in computation.rows:
        sheet.append(row)

    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    format_sheet(
        sheet,
        font_name,
        measurement_font,
        ("商品名称",),
    )
    for cell in sheet["A"][1:]:
        cell.number_format = "@"
    for cell in sheet["B"][1:]:
        if cell.value not in (None, ""):
            cell.number_format = "yyyy-mm-dd"
    matched_fill = PatternFill("solid", fgColor=COUPON_MATCH_FILL_COLOR)
    matched_start_row = sheet.max_row - computation.matched_count + 1
    for row in sheet.iter_rows(
        min_row=matched_start_row,
        max_row=sheet.max_row,
    ):
        for cell in row:
            cell.fill = matched_fill

    return font_name, measurement_font, matched_fill


def validate_detail_sheet(
    workbook: Workbook,
    computation: CouponComputation,
) -> None:
    expected_data_rows = computation.data_row_count
    expected_matched_rows = computation.matched_count
    remark_lookup = computation.remark_lookup
    expected_matched_subsidy_total = computation.matched_subsidy_total
    detail_lookup = computation.detail_lookup
    expected_uploaded_rows = computation.uploaded_match_count
    reference_universe = computation.reference_universe
    expected_unmatched_rows = computation.unmatched_count

    sheet = workbook[DETAILS_SHEET_NAME]
    header = tuple(cell.value for cell in sheet[1])
    if header != COUPON_OUTPUT_HEADER:
        raise RuntimeError(
            f"销售用券字段标题校验失败：实际为 {header}"
        )
    if sheet.max_column != len(COUPON_OUTPUT_HEADER):
        raise RuntimeError(
            f"销售用券列数校验失败：实际为 {sheet.max_column}"
        )
    actual_data_rows = max(sheet.max_row - 1, 0)
    if actual_data_rows != expected_data_rows:
        raise RuntimeError(
            f"销售用券行数校验失败：预期 {expected_data_rows} 条，"
            f"实际 {actual_data_rows} 条"
        )
    detail_column = COUPON_OUTPUT_HEADER.index("详细情况") + 1
    actual_uploaded_rows = sum(
        sheet.cell(row_number, detail_column).value not in (None, "")
        for row_number in range(2, sheet.max_row + 1)
    )
    if actual_uploaded_rows != expected_uploaded_rows:
        raise RuntimeError(
            f"销售用券已上传匹配数校验失败：预期 "
            f"{expected_uploaded_rows} 条，实际 {actual_uploaded_rows} 条"
        )
    remark_column = COUPON_OUTPUT_HEADER.index("备注") + 1
    actual_unuploaded_remark_rows = sum(
        sheet.cell(row_number, remark_column).value == "未上传"
        for row_number in range(2, sheet.max_row + 1)
    )
    expected_unuploaded_remark_rows = expected_unmatched_rows
    if actual_unuploaded_remark_rows != expected_unuploaded_remark_rows:
        raise RuntimeError(
            f"销售用券“未上传”备注数量校验失败：预期 "
            f"{expected_unuploaded_remark_rows} 条，"
            f"实际 {actual_unuploaded_remark_rows} 条"
        )
    summary_column = COUPON_OUTPUT_HEADER.index("明细摘要") + 1
    product_name_column = COUPON_OUTPUT_HEADER.index("商品名称") + 1
    if any(
        sheet.cell(row_number, product_name_column).alignment.horizontal
        != "left"
        for row_number in range(2, sheet.max_row + 1)
    ):
        raise RuntimeError("销售用券商品名称列左对齐校验失败")
    matched_start_row = sheet.max_row - expected_matched_rows + 1
    actual_matched_subsidy_total = Decimal("0")
    subsidy_column = COUPON_OUTPUT_HEADER.index(COUPON_SUBSIDY_HEADER) + 1
    for row_number in range(2, sheet.max_row + 1):
        document_cell = sheet.cell(row_number, 1)
        date_cell = sheet.cell(row_number, 2)
        remark_cell = sheet.cell(row_number, remark_column)
        detail_cell = sheet.cell(row_number, detail_column)
        if "收款" in str(document_cell.value or ""):
            raise RuntimeError(
                f"销售用券第 {row_number} 行单据号仍包含“收款”"
            )
        if document_cell.number_format != "@":
            raise RuntimeError(
                f"销售用券第 {row_number} 行单据号不是文本格式"
            )
        if (
            date_cell.value not in (None, "")
            and date_cell.number_format != "yyyy-mm-dd"
        ):
            raise RuntimeError(
                f"销售用券第 {row_number} 行单据日期格式错误"
            )
        if date_cell.value not in (None, "") and not isinstance(
            date_cell.value,
            (date, datetime),
        ):
            raise RuntimeError(
                f"销售用券第 {row_number} 行单据日期不是日期值"
            )
        document_date = (
            date_cell.value.date()
            if isinstance(date_cell.value, datetime)
            else date_cell.value
        )
        receipt_remark = remark_lookup.get(
            (
                normalize_document_number(document_cell.value),
                document_date,
            ),
            "",
        )
        reference = normalize_receipt_identifier(
            sheet.cell(row_number, summary_column).value
        ).upper()
        expected_detail = detail_lookup.get(reference, "")
        if expected_detail:
            expected_remark = "已上传"
        elif reference not in reference_universe:
            expected_remark = "未上传"
        else:
            expected_remark = receipt_remark
        actual_remark = str(remark_cell.value or "")
        if actual_remark != expected_remark:
            raise RuntimeError(
                f"销售用券第 {row_number} 行备注匹配校验失败"
            )
        if str(detail_cell.value or "") != expected_detail:
            raise RuntimeError(
                f"销售用券第 {row_number} 行详细情况匹配校验失败"
            )
        expected_pink = (
            expected_matched_rows > 0
            and row_number >= matched_start_row
        )
        is_pink = all(
            sheet.cell(row_number, column).fill.fill_type == "solid"
            and sheet.cell(row_number, column).fill.fgColor.rgb
            in {
                COUPON_MATCH_FILL_COLOR,
                f"00{COUPON_MATCH_FILL_COLOR}",
                f"FF{COUPON_MATCH_FILL_COLOR}",
            }
            for column in range(1, len(COUPON_OUTPUT_HEADER) + 1)
        )
        if is_pink != expected_pink:
            raise RuntimeError(
                f"销售用券第 {row_number} 行粉色填充或位置校验失败"
            )
        if expected_pink:
            subsidy = sheet.cell(row_number, subsidy_column).value
            if subsidy not in (None, ""):
                actual_matched_subsidy_total += Decimal(str(subsidy))

    if as_currency(actual_matched_subsidy_total) != as_currency(
        expected_matched_subsidy_total
    ):
        raise RuntimeError(
            "销售用券匹配行国补合计校验失败："
            f"预期 {expected_matched_subsidy_total}，"
            f"实际 {actual_matched_subsidy_total}"
        )
    if as_currency(actual_matched_subsidy_total) != Decimal("0.00"):
        raise RuntimeError(
            f"销售用券匹配行的{COUPON_SUBSIDY_HEADER}合计不为 0："
            f"{actual_matched_subsidy_total}"
        )
