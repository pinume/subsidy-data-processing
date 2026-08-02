"""Pure-data validation shared by the two coupon computations.

Business rules are checked before the workbook is written.  The written file
is validated separately with calamine, so these checks do not depend on cell
styles or on an Excel reader's worksheet API.
"""

from datetime import date, datetime

from processors.common.dates import normalize_receipt_identifier

from .matching import UPLOADED_REMARK, as_currency


def validate_detail_rows_shape(
    rows: list[list[object]],
    expected_header: tuple[str, ...],
    expected_data_rows: int,
) -> None:
    header = tuple(rows[0]) if rows else ()
    if header != expected_header:
        raise RuntimeError(f"销售用券字段标题校验失败：实际为 {header}")
    actual_data_rows = max(len(rows) - 1, 0)
    if actual_data_rows != expected_data_rows:
        raise RuntimeError(
            f"销售用券行数校验失败：预期 {expected_data_rows} 条，"
            f"实际 {actual_data_rows} 条"
        )
    if any(len(row) != len(expected_header) for row in rows):
        raise RuntimeError("销售用券列数校验失败")


def validate_uploaded_and_unmatched_counts(
    rows: list[list[object]],
    header: tuple[str, ...],
    expected_uploaded_rows: int,
    expected_unmatched_rows: int,
) -> tuple[int, int]:
    """Check counts; return the 0-based (详细情况, 备注) columns."""
    detail_column = header.index("详细情况")
    actual_uploaded_rows = sum(
        row[detail_column] not in (None, "") for row in rows[1:]
    )
    if actual_uploaded_rows != expected_uploaded_rows:
        raise RuntimeError(
            "销售用券已上传匹配数校验失败：预期 "
            f"{expected_uploaded_rows} 条，实际 {actual_uploaded_rows} 条"
        )
    remark_column = header.index("备注")
    actual_unuploaded_remark_rows = sum(
        row[remark_column] == "未上传" for row in rows[1:]
    )
    if actual_unuploaded_remark_rows != expected_unmatched_rows:
        raise RuntimeError(
            "销售用券“未上传”备注数量校验失败：预期 "
            f"{expected_unmatched_rows} 条，实际 "
            f"{actual_unuploaded_remark_rows} 条"
        )
    return detail_column, remark_column


def validate_document_and_date_values(
    document: object,
    document_date: object,
    row_number: int,
) -> date | None:
    """Validate 单据号/单据日期 values and return a date or None."""
    if "收款" in str(document or ""):
        raise RuntimeError(f"销售用券第 {row_number} 行单据号仍包含“收款”")
    if document_date not in (None, "") and not isinstance(
        document_date, (date, datetime)
    ):
        raise RuntimeError(f"销售用券第 {row_number} 行单据日期不是日期值")
    if isinstance(document_date, datetime):
        return document_date.date()
    return document_date


def validate_remark_and_detail_values(
    actual_remark: object,
    actual_detail: object,
    expected_remark: str,
    expected_detail: str,
    row_number: int,
) -> None:
    if str(actual_remark or "") != expected_remark:
        raise RuntimeError(f"销售用券第 {row_number} 行备注匹配校验失败")
    if str(actual_detail or "") != expected_detail:
        raise RuntimeError(f"销售用券第 {row_number} 行详细情况匹配校验失败")


def validate_matched_subsidy_total(actual_total, expected_total) -> None:
    if as_currency(actual_total) != as_currency(expected_total):
        raise RuntimeError(
            "销售用券匹配行国补合计校验失败："
            f"预期 {expected_total}，实际 {actual_total}"
        )


def validate_payment_statuses(
    rows: list[list[object]],
    header: tuple[str, ...],
    payment_references: set[str] | frozenset[str],
    excluded_bottom_rows: int,
    expected_paid_rows: int,
) -> None:
    summary_column = header.index("明细摘要")
    remark_column = header.index("备注")
    payment_status_column = header.index("回款情况")
    included_end = len(rows) - excluded_bottom_rows
    actual_paid_rows = 0
    for row_index, row in enumerate(rows[1:], start=1):
        reference = normalize_receipt_identifier(
            row[summary_column]
        ).upper()
        expected_status = (
            "已回款"
            if (
                row_index < included_end
                and row[remark_column] == UPLOADED_REMARK
                and reference in payment_references
            )
            else ""
        )
        actual_status = str(row[payment_status_column] or "")
        if actual_status != expected_status:
            raise RuntimeError(
                f"销售用券第 {row_index + 1} 行回款情况匹配校验失败"
            )
        actual_paid_rows += int(actual_status == "已回款")
    if actual_paid_rows != expected_paid_rows:
        raise RuntimeError(
            "销售用券已回款匹配数校验失败："
            f"预期 {expected_paid_rows} 条，实际 {actual_paid_rows} 条"
        )
