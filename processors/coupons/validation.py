"""Detail-sheet validation checks shared by appliance.py and digital.py.

Both projects' validate_*_sheet() functions re-check the same detail-sheet
shape, uploaded/unuploaded counts, per-row 单据号/单据日期 formatting, and
pink-fill position/subsidy accumulation — byte-identical code that used to be
duplicated in each module. Only those exactly-identical checks live here;
the surrounding business logic (家电's reference-supplement matching, group
sheets, 数据汇总 five-column summary vs 数码's three-column one) differs
enough between the two projects that forcing it through one shared,
flag-driven validator would trade duplication for a worse kind of coupling.
"""

from datetime import date, datetime

from .matching import as_currency


def validate_detail_sheet_shape(
    sheet,
    expected_header: tuple[str, ...],
    expected_data_rows: int,
) -> None:
    header = tuple(cell.value for cell in sheet[1])
    if header != expected_header:
        raise RuntimeError(f"销售用券字段标题校验失败：实际为 {header}")
    if sheet.max_column != len(expected_header):
        raise RuntimeError(f"销售用券列数校验失败：实际为 {sheet.max_column}")
    actual_data_rows = max(sheet.max_row - 1, 0)
    if actual_data_rows != expected_data_rows:
        raise RuntimeError(
            f"销售用券行数校验失败：预期 {expected_data_rows} 条，"
            f"实际 {actual_data_rows} 条"
        )


def validate_uploaded_and_unmatched_counts(
    sheet,
    header: tuple[str, ...],
    expected_uploaded_rows: int,
    expected_unmatched_rows: int,
) -> tuple[int, int]:
    """Check the 已上传/未上传 row counts; return (详细情况, 备注) 1-based columns."""
    detail_column = header.index("详细情况") + 1
    actual_uploaded_rows = sum(
        sheet.cell(row_number, detail_column).value not in (None, "")
        for row_number in range(2, sheet.max_row + 1)
    )
    if actual_uploaded_rows != expected_uploaded_rows:
        raise RuntimeError(
            f"销售用券已上传匹配数校验失败：预期 "
            f"{expected_uploaded_rows} 条，实际 {actual_uploaded_rows} 条"
        )
    remark_column = header.index("备注") + 1
    actual_unuploaded_remark_rows = sum(
        sheet.cell(row_number, remark_column).value == "未上传"
        for row_number in range(2, sheet.max_row + 1)
    )
    if actual_unuploaded_remark_rows != expected_unmatched_rows:
        raise RuntimeError(
            f"销售用券“未上传”备注数量校验失败：预期 "
            f"{expected_unmatched_rows} 条，"
            f"实际 {actual_unuploaded_remark_rows} 条"
        )
    return detail_column, remark_column


def validate_left_aligned_column(sheet, column_number: int, column_name: str) -> None:
    if any(
        sheet.cell(row_number, column_number).alignment.horizontal != "left"
        for row_number in range(2, sheet.max_row + 1)
    ):
        raise RuntimeError(f"销售用券{column_name}列左对齐校验失败")


def validate_document_and_date_cells(document_cell, date_cell, row_number: int):
    """Check 单据号/单据日期 formatting; return the row's date value (or None)."""
    if "收款" in str(document_cell.value or ""):
        raise RuntimeError(f"销售用券第 {row_number} 行单据号仍包含“收款”")
    if document_cell.number_format != "@":
        raise RuntimeError(f"销售用券第 {row_number} 行单据号不是文本格式")
    if date_cell.value not in (None, "") and date_cell.number_format != "yyyy-mm-dd":
        raise RuntimeError(f"销售用券第 {row_number} 行单据日期格式错误")
    if date_cell.value not in (None, "") and not isinstance(
        date_cell.value, (date, datetime)
    ):
        raise RuntimeError(f"销售用券第 {row_number} 行单据日期不是日期值")
    return date_cell.value.date() if isinstance(date_cell.value, datetime) else date_cell.value


def validate_remark_and_detail(
    remark_cell,
    detail_cell,
    expected_remark: str,
    expected_detail: str,
    row_number: int,
) -> None:
    actual_remark = str(remark_cell.value or "")
    if actual_remark != expected_remark:
        raise RuntimeError(f"销售用券第 {row_number} 行备注匹配校验失败")
    if str(detail_cell.value or "") != expected_detail:
        raise RuntimeError(f"销售用券第 {row_number} 行详细情况匹配校验失败")


def is_pink_row(sheet, row_number: int, column_count: int, fill_color: str) -> bool:
    return all(
        sheet.cell(row_number, column).fill.fill_type == "solid"
        and sheet.cell(row_number, column).fill.fgColor.rgb
        in {fill_color, f"00{fill_color}", f"FF{fill_color}"}
        for column in range(1, column_count + 1)
    )


def validate_pink_position(is_pink: bool, expected_pink: bool, row_number: int) -> None:
    if is_pink != expected_pink:
        raise RuntimeError(f"销售用券第 {row_number} 行粉色填充或位置校验失败")


def validate_matched_subsidy_total(actual_total, expected_total) -> None:
    if as_currency(actual_total) != as_currency(expected_total):
        raise RuntimeError(
            "销售用券匹配行国补合计校验失败："
            f"预期 {expected_total}，"
            f"实际 {actual_total}"
        )
