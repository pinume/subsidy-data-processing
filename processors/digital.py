
import re
from collections import Counter
from collections.abc import Callable
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

from processors.common.excel import (
    create_sheet_styles,
    format_sheet,
    load_measurement_font,
    pixels_to_excel_width,
    read_rows,
    resolve_font,
    save_workbook_atomically,
    width_measurer,
)
from processors.common.dates import (
    is_valid_original_invoice_number,
    normalize_coupon_date,
    normalize_document_number,
    normalize_receipt_date,
    normalize_receipt_identifier,
    receipt_match_key,
)
from processors.common.paths import resolve_existing_data_file


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output" / "digital"
DATA_DIR: Path
INPUT_DIR: Path
OUTPUT_FILE = OUTPUT_DIR / "已上传.xlsx"
RECEIPTS_SOURCE_FILE: Path
RECEIPTS_OUTPUT_FILE = OUTPUT_DIR / "收款单统计.xlsx"
COUPON_SOURCE_FILE: Path
COUPON_OUTPUT_FILE = OUTPUT_DIR / "销售用券情况统计.xlsx"
COUPON_REMARK_SOURCE_FILE = RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = OUTPUT_FILE


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global INPUT_DIR
    global RECEIPTS_SOURCE_FILE
    global COUPON_SOURCE_FILE

    DATA_DIR = data_dir
    INPUT_DIR = DATA_DIR / "submitted"
    RECEIPTS_SOURCE_FILE = resolve_existing_data_file(
        DATA_DIR,
        (
            Path("receipt_statistics") / "receipt_statistics.XLS",
            Path("收款单统计") / "收款单统计.XLS",
        ),
    )
    COUPON_SOURCE_FILE = resolve_existing_data_file(
        DATA_DIR,
        (
            Path("subsidy_coupon_statistics") / "subsidy_coupon_statistics.XLS",
            Path("销售用券情况统计") / "销售用券情况统计.XLS",
        ),
    )

# Keep the required source columns in their original order.
KEPT_SOURCE_COLUMNS = ("D", "E", "F", "G", "I", "J", "Q", "S", "U", "W", "X")
KEPT_COLUMN_INDEXES = tuple(
    column_index_from_string(column) for column in KEPT_SOURCE_COLUMNS
)
# "补贴金额" is inserted by add_subsidy_column, so only source fields are listed.
REQUIRED_SUBMITTED_HEADERS = ("状态", "描述", "交易金额")
STATUS_ORDER = (
    "核销失败",
    "审核失败",
    "暂存",
    "同步(已上送)",
    "待审核",
    "审核通过",
)
STATUS_PRIORITY = {status: index for index, status in enumerate(STATUS_ORDER)}
RECEIPTS_SOURCE_HEADER = ("单据号", "日期", "原票号", "摘要", "商品名称")
RECEIPTS_OUTPUT_HEADER = (*RECEIPTS_SOURCE_HEADER, "备注")
RECEIPTS_REMARK_RETURN = "退换货/倒票（退单）"
RECEIPTS_REMARK_ORIGINAL = "退换货/倒票（原单）"
RECEIPTS_REMARK_BOTH = "退换货/倒票（退单及原单）"
RECEIPTS_REPORT_HEADER = ("异常类型", "源文件行", "匹配值", "说明")
RECEIPTS_ROW_HEIGHT = 20
RECEIPTS_DUPLICATE_FILL_COLOR = "FFC7CE"
RECEIPTS_EXCLUDED_PRODUCT_KEYWORD = "北国"
COUPON_KEPT_SOURCE_COLUMNS = (3, 4, 6, 8, 15, 18, 26)
COUPON_OUTPUT_HEADER = (
    "单据号",
    "单据日期",
    "商品名称",
    "品牌",
    "财务大类",
    "明细摘要",
    "2026数码国补（计入收入）",
    "备注",
    "详细情况",
)
REFERENCE_REPORT_CORRECTED = "已自动纠正"
REFERENCE_REPORT_UNRESOLVED = "无唯一候选"
REFERENCE_REPORT_COLLISION = "目标冲突"
REFERENCE_REPORT_ORDER = {
    REFERENCE_REPORT_CORRECTED: 0,
    REFERENCE_REPORT_COLLISION: 1,
    REFERENCE_REPORT_UNRESOLVED: 2,
}
REFERENCE_REPORT_HEADER = ("处理结果", "输出行号", "原参考号", "说明")
REFERENCE_REPORT_SHEET = "Processing Report"
COUPON_MATCH_FILL_COLOR = "FFC7CE"
COUPON_SUMMARY_HEADER = (
    "备注",
    "数量",
    "2026数码国补（计入收入）合计",
)


def as_currency(amount: Decimal) -> Decimal:
    """Round to cents before comparing.

    Source amounts arrive as floats, so Decimal(str(value)) carries binary
    noise such as 1234.5600000000001 into the running total. Totals are money
    and are only ever meaningful to two decimal places.
    """
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def select_columns(row: list[object]) -> list[object]:
    return [
        row[column - 1] if column <= len(row) else None
        for column in KEPT_COLUMN_INDEXES
    ]


def add_subsidy_column(
    row: list[object],
    *,
    is_header: bool = False,
    source_name: str | None = None,
    source_row: int | None = None,
) -> list[object]:
    result = list(row)
    if is_header:
        result.insert(3, "补贴金额")
        return result

    amount = result[2]
    if amount in (None, ""):
        subsidy = None
    else:
        try:
            calculated = Decimal(str(amount)) * Decimal("0.15")
            subsidy = float(
                min(calculated, Decimal("500")).quantize(
                    Decimal("0.01"),
                    rounding=ROUND_HALF_UP,
                )
            )
        except (InvalidOperation, ValueError) as error:
            location = (
                f"{source_name} 第 {source_row} 行"
                if source_name is not None and source_row is not None
                else "数据"
            )
            raise ValueError(f"{location}的交易金额无效：{amount!r}") from error

    result.insert(3, subsidy)
    return result


def build_workbook() -> tuple[Workbook, int, int]:
    files = sorted(
        path
        for path in INPUT_DIR.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"未在 {INPUT_DIR} 中找到 .xlsx 文件")

    workbook = Workbook(write_only=False)
    sheet = workbook.active
    sheet.title = "Summary"

    expected_header: list[object] | None = None
    output_header: list[object] | None = None
    data_row_count = 0
    data_rows: list[list[object]] = []

    for path in files:
        rows = read_rows(path)
        title = next(rows, None)
        header = next(rows, None)
        if title is None or header is None:
            raise ValueError(f"{path.name} 缺少标题行或表头行")

        if expected_header is None:
            expected_header = header
            # Drop the source title row and use the selected headers as row 1.
            output_header = add_subsidy_column(
                select_columns(header),
                is_header=True,
            )
            # Reject a wrong export before parsing any rows, so the operator sees
            # the missing fields instead of a downstream value error.
            missing_headers = [
                required
                for required in REQUIRED_SUBMITTED_HEADERS
                if required not in output_header
            ]
            if missing_headers:
                raise ValueError(
                    f"{path.name} 不是已上传数据的导出格式，缺少字段："
                    f"{'、'.join(missing_headers)}；"
                    f"实际字段 {tuple(output_header)}"
                )
            sheet.append(output_header)
        elif header != expected_header:
            raise ValueError(f"{path.name} 的表头与第一个文件不一致")

        for source_row, row in enumerate(rows, start=3):
            if any(value not in (None, "") for value in row):
                data_rows.append(
                    add_subsidy_column(
                        select_columns(row),
                        source_name=path.name,
                        source_row=source_row,
                    )
                )
                data_row_count += 1

    if output_header is None:
        raise RuntimeError("未能生成输出表头")

    status_column_index = output_header.index("状态")
    data_rows.sort(
        key=lambda row: STATUS_PRIORITY.get(
            str(row[status_column_index]) if row[status_column_index] is not None else "",
            len(STATUS_ORDER),
        )
    )
    for row in data_rows:
        sheet.append(row)

    description_column_index = output_header.index("描述")
    rows_by_status: dict[str, list[list[object]]] = {
        status: [] for status in STATUS_ORDER
    }
    for row in data_rows:
        status = str(row[status_column_index] or "")
        if status in rows_by_status:
            rows_by_status[status].append(row)

    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    format_sheet(sheet, font_name, measurement_font)
    for status in STATUS_ORDER:
        status_sheet = workbook.create_sheet(title=status)
        status_sheet.append(output_header)
        status_rows = rows_by_status[status]
        status_rows.sort(
            key=lambda row: (
                row[description_column_index] not in (None, ""),
                str(row[description_column_index] or ""),
            ),
            reverse=True,
        )
        for row in status_rows:
            status_sheet.append(row)
        format_sheet(status_sheet, font_name, measurement_font)

    return workbook, len(files), data_row_count


def validate_output(path: Path, expected_data_rows: int) -> None:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        expected_sheet_names = ["Summary", *STATUS_ORDER]
        if workbook.sheetnames != expected_sheet_names:
            raise RuntimeError(
                f"工作表校验失败：预期 {expected_sheet_names}，"
                f"实际 {workbook.sheetnames}"
            )

        sheet = workbook["Summary"]
        actual_data_rows = max(sheet.max_row - 1, 0)
        if actual_data_rows != expected_data_rows:
            raise RuntimeError(
                f"输出校验失败：预期 {expected_data_rows} 条，实际 {actual_data_rows} 条"
            )
        expected_columns = len(KEPT_SOURCE_COLUMNS) + 1
        if sheet.max_column != expected_columns:
            raise RuntimeError(
                f"输出校验失败：预期 {expected_columns} 列，"
                f"实际 {sheet.max_column} 列"
            )

        header = tuple(cell.value for cell in next(sheet.iter_rows(max_row=1)))
        status_column = header.index("状态")
        description_column = header.index("描述")
        amount_column = header.index("交易金额")
        subsidy_column = header.index("补贴金额")

        status_total = sum(
            max(workbook[status].max_row - 1, 0)
            for status in STATUS_ORDER
        )
        known_status_total = sum(
            1
            for row in sheet.iter_rows(min_row=2, values_only=True)
            if row[status_column] in STATUS_ORDER
        )
        if status_total != known_status_total:
            raise RuntimeError(
                f"状态工作表校验失败：预期共 {known_status_total} 条，"
                f"实际共 {status_total} 条"
            )

        for status in STATUS_ORDER:
            status_sheet = workbook[status]
            status_header = tuple(
                cell.value for cell in next(status_sheet.iter_rows(max_row=1))
            )
            if status_header != header:
                raise RuntimeError(f"{status}工作表的标题行与汇总表不一致")

            descriptions: list[str] = []
            blank_description_found = False
            for row in status_sheet.iter_rows(min_row=2, values_only=True):
                if row[status_column] != status:
                    raise RuntimeError(f"{status}工作表中存在其他状态的数据")

                description = row[description_column]
                if description in (None, ""):
                    blank_description_found = True
                else:
                    if blank_description_found:
                        raise RuntimeError(
                            f"{status}工作表的空白描述未全部排在末尾"
                        )
                    descriptions.append(str(description))

                amount = row[amount_column]
                subsidy = row[subsidy_column]
                if amount not in (None, ""):
                    expected_subsidy = min(
                        Decimal(str(amount)) * Decimal("0.15"),
                        Decimal("500"),
                    ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    if Decimal(str(subsidy)) != expected_subsidy:
                        raise RuntimeError(
                            f"{status}工作表存在补贴金额计算错误"
                        )

            if descriptions != sorted(descriptions, reverse=True):
                raise RuntimeError(f"{status}工作表的描述列未按降序排列")
    finally:
        workbook.close()


def process_submitted_files() -> None:
    workbook, file_count, data_row_count = build_workbook()
    save_workbook_atomically(
        workbook,
        OUTPUT_FILE,
        lambda path: validate_output(path, data_row_count),
    )

    print(f"Submitted data complete: merged {file_count} files, {data_row_count} rows")
    print(f"Output file: {OUTPUT_FILE}")


def read_receipt_rows(source: Path) -> list[list[object]]:
    """Read legacy XLS files in pure Python without invoking Excel."""
    source_workbook = xlrd.open_workbook(source)
    try:
        source_sheet = source_workbook.sheet_by_index(0)
        if source_sheet.nrows < 2:
            raise ValueError(f"{source.name} 缺少总标题行或字段标题行")
        source_headers = [
            str(source_sheet.cell_value(1, column_index)).strip()
            for column_index in range(source_sheet.ncols)
        ]
        missing_headers = [
            header for header in RECEIPTS_SOURCE_HEADER
            if header not in source_headers
        ]
        if missing_headers:
            raise ValueError(
                f"{source.name} 缺少必要字段：{tuple(missing_headers)}；"
                f"实际字段 {tuple(source_headers)}"
            )
        source_column_indexes = [
            source_headers.index(header) for header in RECEIPTS_SOURCE_HEADER
        ]

        rows: list[list[object]] = []
        for row_index in range(1, source_sheet.nrows):
            row: list[object] = []
            for column_index in source_column_indexes:
                cell = source_sheet.cell(row_index, column_index)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate.xldate_as_datetime(
                        value,
                        source_workbook.datemode,
                    )
                elif cell.ctype == xlrd.XL_CELL_NUMBER and value == int(value):
                    value = int(value)
                row.append(value)
            if row_index > 1 and str(row[1]).strip() == "合计":
                continue
            rows.append(row)

        actual_header = tuple(
            str(value).strip() if value is not None else ""
            for value in rows[0]
        )
        if actual_header != RECEIPTS_SOURCE_HEADER:
            raise ValueError(
                f"{source.name} 字段标题不一致："
                f"预期 {RECEIPTS_SOURCE_HEADER}，实际 {actual_header}"
            )
        return rows
    finally:
        source_workbook.release_resources()


def receipt_remark(has_original: bool, is_referenced: bool) -> str | None:
    if has_original and is_referenced:
        return RECEIPTS_REMARK_BOTH
    if has_original:
        return RECEIPTS_REMARK_RETURN
    if is_referenced:
        return RECEIPTS_REMARK_ORIGINAL
    return None


def prepare_receipt_data(kept_rows: list[list[object]]):
    records: list[dict[str, object]] = []
    key_rows: dict[str, list[int]] = {}
    original_invoice_numbers: set[str] = set()
    excluded_product_count = 0

    for source_row, row in enumerate(kept_rows[1:], start=3):
        product_name = normalize_receipt_identifier(row[4])
        if RECEIPTS_EXCLUDED_PRODUCT_KEYWORD in product_name:
            excluded_product_count += 1
            continue

        document_number = normalize_document_number(row[0])
        receipt_date = normalize_receipt_date(row[1], source_row=source_row)
        original_invoice_number = normalize_receipt_identifier(row[2])
        match_key = (
            receipt_match_key(receipt_date, document_number)
            if receipt_date is not None and document_number
            else ""
        )
        if match_key:
            key_rows.setdefault(match_key, []).append(source_row)
        if original_invoice_number:
            original_invoice_numbers.add(original_invoice_number)
        records.append(
            {
                "source_row": source_row,
                "document_number": document_number,
                "receipt_date": receipt_date,
                "original_invoice_number": original_invoice_number,
                "summary": row[3],
                "product_name": product_name or None,
                "match_key": match_key,
            }
        )

    issues: list[tuple[str, str, str, str]] = []
    for match_key, source_rows in key_rows.items():
        if len(source_rows) > 1:
            issues.append(
                (
                    "重复匹配键",
                    "、".join(str(row) for row in source_rows),
                    match_key,
                    "多个数据行生成了相同的日期与单据号组合键",
                )
            )

    only_return_count = 0
    only_original_count = 0
    both_count = 0
    unmatched_original_count = 0
    invalid_original_count = 0
    missing_match_key_count = 0
    output_rows: list[list[object]] = []
    for record in records:
        source_row = int(record["source_row"])
        original_invoice_number = str(record["original_invoice_number"])
        match_key = str(record["match_key"])
        has_original = bool(original_invoice_number)
        is_referenced = bool(match_key and match_key in original_invoice_numbers)

        if not match_key:
            missing_match_key_count += 1
            issues.append(
                (
                    "缺少匹配键",
                    str(source_row),
                    "",
                    "日期或单据号为空，无法生成匹配键",
                )
            )
        if has_original and not is_valid_original_invoice_number(
            original_invoice_number
        ):
            invalid_original_count += 1
            issues.append(
                (
                    "原票号格式异常",
                    str(source_row),
                    original_invoice_number,
                    "原票号应为6位日期加单据号",
                )
            )
        if has_original and original_invoice_number not in key_rows:
            unmatched_original_count += 1
            issues.append(
                (
                    "原票号未匹配",
                    str(source_row),
                    original_invoice_number,
                    "未找到日期与单据号组合键相同的原单",
                )
            )

        if has_original and is_referenced:
            both_count += 1
        elif has_original:
            only_return_count += 1
        elif is_referenced:
            only_original_count += 1

        output_rows.append(
            [
                record["document_number"],
                record["receipt_date"],
                original_invoice_number or None,
                record["summary"],
                record["product_name"],
                receipt_remark(has_original, is_referenced),
            ]
        )

    stats = {
        "总数据量": len(records),
        "删除北国商品行数": excluded_product_count,
        "仅退单数量": only_return_count,
        "仅原单数量": only_original_count,
        "退单及原单数量": both_count,
        "备注总数": only_return_count + only_original_count + both_count,
        "未匹配原票号数量": unmatched_original_count,
        "重复匹配键数量": sum(
            1 for source_rows in key_rows.values() if len(source_rows) > 1
        ),
        "原票号格式异常数量": invalid_original_count,
        "缺少匹配键数量": missing_match_key_count,
    }
    duplicate_match_keys = {
        match_key
        for match_key, source_rows in key_rows.items()
        if len(source_rows) > 1
    }
    return output_rows, stats, issues, duplicate_match_keys


def validate_receipts_output(
    path: Path,
    expected_data_rows: int,
    expected_stats: dict[str, int],
    expected_issues: list[tuple[str, str, str, str]],
) -> None:
    workbook = load_workbook(path, read_only=False, data_only=True)
    try:
        if workbook.sheetnames != ["Sheet1", "Processing Report"]:
            raise RuntimeError(
                f"收款单工作表校验失败：实际工作表为 {workbook.sheetnames}"
            )

        sheet = workbook["Sheet1"]
        if sheet.max_row - 1 != expected_data_rows:
            raise RuntimeError(
                f"收款单行数校验失败：预期 {expected_data_rows} 条，"
                f"实际 {sheet.max_row - 1} 条"
            )
        if sheet.max_column != len(RECEIPTS_OUTPUT_HEADER):
            raise RuntimeError(
                f"收款单列数校验失败：预期 {len(RECEIPTS_OUTPUT_HEADER)} 列，"
                f"实际 {sheet.max_column} 列"
            )

        header = tuple(cell.value for cell in sheet[1])
        if header != RECEIPTS_OUTPUT_HEADER:
            raise RuntimeError(
                f"收款单表头校验失败：预期 {RECEIPTS_OUTPUT_HEADER}，"
                f"实际 {header}"
            )

        original_invoice_numbers = {
            normalize_receipt_identifier(sheet.cell(row_number, 3).value)
            for row_number in range(2, sheet.max_row + 1)
            if sheet.cell(row_number, 3).value not in (None, "")
        }
        match_key_counts: dict[str, int] = {}
        for row_number in range(2, sheet.max_row + 1):
            document_number = normalize_receipt_identifier(
                sheet.cell(row_number, 1).value
            )
            date_value = sheet.cell(row_number, 2).value
            if date_value is None or not document_number:
                continue
            match_key = receipt_match_key(
                date_value.date()
                if isinstance(date_value, datetime)
                else date_value,
                document_number,
            )
            match_key_counts[match_key] = match_key_counts.get(match_key, 0) + 1

        for row_number in range(2, sheet.max_row + 1):
            document_number = sheet.cell(row_number, 1).value
            if document_number is not None and (
                not isinstance(document_number, str)
                or document_number.startswith("收款")
            ):
                raise RuntimeError(
                    f"收款单第 {row_number} 行的单据号格式不正确"
                )

            date_cell = sheet.cell(row_number, 2)
            if date_cell.value is not None:
                if not isinstance(date_cell.value, (date, datetime)):
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的日期不是有效日期"
                    )
                if date_cell.number_format != "yyyymmdd":
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的日期格式不正确"
                    )

            original_invoice_number = normalize_receipt_identifier(
                sheet.cell(row_number, 3).value
            )
            match_key = (
                receipt_match_key(
                    date_cell.value.date()
                    if isinstance(date_cell.value, datetime)
                    else date_cell.value,
                    normalize_receipt_identifier(document_number),
                )
                if date_cell.value is not None and document_number is not None
                else ""
            )
            expected_remark = receipt_remark(
                bool(original_invoice_number),
                bool(match_key and match_key in original_invoice_numbers),
            )
            if sheet.cell(row_number, 6).value != expected_remark:
                raise RuntimeError(
                    f"收款单第 {row_number} 行的备注校验失败"
                )

            is_duplicate = bool(match_key and match_key_counts[match_key] > 1)
            for cell in sheet[row_number]:
                fill_color = cell.fill.fgColor.rgb
                is_pink = (
                    cell.fill.fill_type == "solid"
                    and fill_color is not None
                    and fill_color[-6:] == RECEIPTS_DUPLICATE_FILL_COLOR[-6:]
                )
                if is_pink != is_duplicate:
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的重复匹配键标记不正确"
                    )

        report_sheet = workbook["Processing Report"]
        actual_stats = {
            report_sheet.cell(row_number, 1).value: report_sheet.cell(
                row_number,
                2,
            ).value
            for row_number in range(2, 2 + len(expected_stats))
        }
        if actual_stats != expected_stats:
            raise RuntimeError(
                f"处理报告统计校验失败：预期 {expected_stats}，"
                f"实际 {actual_stats}"
            )
        issue_header_row = len(expected_stats) + 3
        issue_header = tuple(
            report_sheet.cell(issue_header_row, column).value
            for column in range(1, len(RECEIPTS_REPORT_HEADER) + 1)
        )
        if issue_header != RECEIPTS_REPORT_HEADER:
            raise RuntimeError("处理报告异常明细表头校验失败")
        actual_issues = [
            tuple(
                (
                    report_sheet.cell(row_number, column).value
                    if report_sheet.cell(row_number, column).value is not None
                    else ""
                )
                for column in range(1, len(RECEIPTS_REPORT_HEADER) + 1)
            )
            for row_number in range(
                issue_header_row + 1,
                issue_header_row + 1 + len(expected_issues),
            )
        ]
        if actual_issues != expected_issues:
            raise RuntimeError("处理报告异常明细校验失败")
    finally:
        workbook.close()


def process_receipts() -> None:
    if not RECEIPTS_SOURCE_FILE.exists():
        raise FileNotFoundError(f"未找到源文件：{RECEIPTS_SOURCE_FILE}")

    kept_rows = read_receipt_rows(RECEIPTS_SOURCE_FILE)
    output_rows, stats, issues, duplicate_match_keys = prepare_receipt_data(
        kept_rows
    )
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(RECEIPTS_OUTPUT_HEADER)
    for row in output_rows:
        sheet.append(row)

    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    normal_font, header_font, header_fill, centered = create_sheet_styles(
        font_name
    )

    measure = width_measurer(measurement_font)
    maximum_widths = [0.0] * sheet.max_column
    duplicate_fill = PatternFill(
        "solid",
        fgColor=RECEIPTS_DUPLICATE_FILL_COLOR,
    )
    for row_number, row in enumerate(sheet.iter_rows(), start=1):
        sheet.row_dimensions[row_number].height = RECEIPTS_ROW_HEIGHT
        row_match_key = ""
        if row_number >= 2:
            document_number = normalize_receipt_identifier(row[0].value)
            receipt_date = row[1].value
            if receipt_date is not None and document_number:
                row_match_key = receipt_match_key(
                    receipt_date.date()
                    if isinstance(receipt_date, datetime)
                    else receipt_date,
                    document_number,
                )
        for cell in row:
            cell.font = header_font if row_number == 1 else normal_font
            cell.alignment = centered
            if row_number == 1:
                cell.fill = header_fill
            elif row_match_key in duplicate_match_keys:
                cell.fill = duplicate_fill

            if row_number >= 2 and cell.column == 1:
                cell.number_format = "@"
            elif row_number >= 2 and cell.column == 2:
                if cell.value not in (None, ""):
                    cell.number_format = "yyyymmdd"

            width = measure(cell.value)
            if width > maximum_widths[cell.column - 1]:
                maximum_widths[cell.column - 1] = width

    sheet.freeze_panes = "A2"
    for column_index, maximum_pixels in enumerate(maximum_widths, start=1):
        column_letter = sheet.cell(1, column_index).column_letter
        sheet.column_dimensions[column_letter].width = min(
            pixels_to_excel_width(maximum_pixels),
            255,
        )

    report_sheet = workbook.create_sheet("Processing Report")
    report_sheet.append(("统计项目", "数量"))
    for label, value in stats.items():
        report_sheet.append((label, value))
    report_sheet.append(())
    report_sheet.append(RECEIPTS_REPORT_HEADER)
    for issue in issues:
        report_sheet.append(issue)
    format_sheet(report_sheet, font_name, measurement_font)

    row_count = len(output_rows)
    save_workbook_atomically(
        workbook,
        RECEIPTS_OUTPUT_FILE,
        lambda path: validate_receipts_output(
            path,
            row_count,
            stats,
            issues,
        ),
    )
    print(f"Receipt statistics complete: {row_count} rows")
    print(
        f"Remarks: {stats['备注总数']}; "
        f"unmatched original invoices: {stats['未匹配原票号数量']}; "
        f"duplicate match keys: {stats['重复匹配键数量']}"
    )
    print(f"Output file: {RECEIPTS_OUTPUT_FILE}")


def read_coupon_rows(source: Path) -> list[list[object]]:
    source_workbook = xlrd.open_workbook(source)
    try:
        source_sheet = source_workbook.sheet_by_index(0)
        if source_sheet.nrows < 3:
            raise ValueError(f"{source.name} 缺少标题行、字段标题行或合计行")
        if source_sheet.ncols < max(COUPON_KEPT_SOURCE_COLUMNS):
            raise ValueError(
                f"{source.name} 列数不足：至少需要 "
                f"{max(COUPON_KEPT_SOURCE_COLUMNS)} 列"
            )
        if str(source_sheet.cell_value(source_sheet.nrows - 1, 0)).strip() != "合计":
            raise ValueError(f"{source.name} 最后一行不是合计行")

        source_header = tuple(
            source_sheet.cell_value(1, column - 1)
            for column in COUPON_KEPT_SOURCE_COLUMNS
        )
        expected_source_header = COUPON_OUTPUT_HEADER[
            :len(COUPON_KEPT_SOURCE_COLUMNS)
        ]
        if source_header != expected_source_header:
            raise ValueError(
                f"{source.name} 保留列字段标题不符合要求："
                f"预期为 {expected_source_header}，实际为 {source_header}。"
                f"如果实际字段包含“2026家电国补（计入收入）”，"
                f"请选择 large appliances 数据类型，或更换为数码销售用券文件。"
            )

        rows: list[list[object]] = [list(COUPON_OUTPUT_HEADER)]
        for row_index in range(2, source_sheet.nrows - 1):
            row: list[object] = []
            for column_index in COUPON_KEPT_SOURCE_COLUMNS:
                cell = source_sheet.cell(row_index, column_index - 1)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate.xldate_as_datetime(
                        value,
                        source_workbook.datemode,
                    )
                elif (
                    cell.ctype == xlrd.XL_CELL_NUMBER
                    and value == int(value)
                ):
                    value = int(value)
                row.append(value)
            if any(value not in (None, "") for value in row):
                document_number = (
                    ""
                    if row[0] is None
                    else str(row[0]).replace("收款", "")
                )
                document_date = normalize_coupon_date(row[1], row_index + 1)
                rows.append(
                    [document_number, document_date, *row[2:], None, None]
                )
        return rows
    finally:
        source_workbook.release_resources()


def load_coupon_remark_lookup(source: Path) -> dict[tuple[str, date], str]:
    if not source.exists():
        raise FileNotFoundError(f"未找到备注匹配文件：{source}")

    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if "Sheet1" not in workbook.sheetnames:
            raise ValueError(f"{source.name} 缺少 Sheet1 工作表")
        sheet = workbook["Sheet1"]
        header = [cell.value for cell in sheet[1]]
        required_headers = ("单据号", "日期", "备注")
        missing_headers = [
            required_header
            for required_header in required_headers
            if required_header not in header
        ]
        if missing_headers:
            raise ValueError(
                f"{source.name} 缺少字段：{'、'.join(missing_headers)}"
            )

        document_index = header.index("单据号")
        date_index = header.index("日期")
        remark_index = header.index("备注")
        lookup: dict[tuple[str, date], str] = {}
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True),
            start=2,
        ):
            document_number = normalize_document_number(row[document_index])
            remark = str(row[remark_index] or "").strip()
            if not document_number or not remark:
                continue
            receipt_date = normalize_coupon_date(
                row[date_index],
                row_number,
            )
            key = (document_number, receipt_date)
            existing_remark = lookup.get(key)
            if existing_remark is not None and existing_remark != remark:
                raise ValueError(
                    f"{source.name} 第 {row_number} 行组合键存在冲突备注："
                    f"{document_number} + {receipt_date:%Y-%m-%d}"
                )
            lookup[key] = remark
        return lookup
    finally:
        workbook.close()


def fill_coupon_remarks(
    rows: list[list[object]],
    remark_lookup: dict[tuple[str, date], str],
) -> tuple[int, Decimal]:
    matched_rows: list[list[object]] = []
    unmatched_rows: list[list[object]] = []
    subsidy_index = COUPON_OUTPUT_HEADER.index("2026数码国补（计入收入）")
    matched_subsidy_total = Decimal("0")
    remark_index = COUPON_OUTPUT_HEADER.index("备注")
    for row in rows[1:]:
        key = (normalize_document_number(row[0]), row[1])
        remark = remark_lookup.get(key, "")
        row[remark_index] = remark
        if remark:
            matched_rows.append(row)
            subsidy = row[subsidy_index]
            if subsidy not in (None, ""):
                try:
                    matched_subsidy_total += Decimal(str(subsidy))
                except InvalidOperation as error:
                    raise ValueError(
                        f"组合键 {key} 的2026数码国补金额无效：{subsidy!r}"
                    ) from error
        else:
            unmatched_rows.append(row)
    rows[1:] = [*unmatched_rows, *matched_rows]
    return len(matched_rows), matched_subsidy_total


def load_uploaded_detail_lookup(source: Path) -> dict[str, str]:
    if not source.exists():
        raise FileNotFoundError(f"未找到已上传匹配文件：{source}")

    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if "Summary" not in workbook.sheetnames:
            raise ValueError(f"{source.name} 缺少“汇总”工作表")
        sheet = workbook["Summary"]
        header = [cell.value for cell in sheet[1]]
        required_headers = ("检索参考号", "状态", "描述")
        missing_headers = [
            required_header
            for required_header in required_headers
            if required_header not in header
        ]
        if missing_headers:
            raise ValueError(
                f"{source.name} 缺少字段：{'、'.join(missing_headers)}"
            )

        reference_index = header.index("检索参考号")
        status_index = header.index("状态")
        description_index = header.index("描述")
        lookup: dict[str, str] = {}
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True),
            start=2,
        ):
            reference = normalize_receipt_identifier(
                row[reference_index]
            ).upper()
            if not reference:
                continue
            status = str(row[status_index] or "").strip()
            description = str(row[description_index] or "").strip()
            detail = f"{status}：{description}"
            existing_detail = lookup.get(reference)
            if existing_detail is not None and existing_detail != detail:
                raise ValueError(
                    f"{source.name} 第 {row_number} 行检索参考号存在冲突："
                    f"{reference}"
                )
            lookup[reference] = detail
        return lookup
    finally:
        workbook.close()


def fill_uploaded_details(
    rows: list[list[object]],
    detail_lookup: dict[str, str],
) -> int:
    summary_index = COUPON_OUTPUT_HEADER.index("明细摘要")
    remark_index = COUPON_OUTPUT_HEADER.index("备注")
    detail_index = COUPON_OUTPUT_HEADER.index("详细情况")
    matched_count = 0
    for row in rows[1:]:
        reference = normalize_receipt_identifier(
            row[summary_index]
        ).upper()
        detail = detail_lookup.get(reference, "")
        row[detail_index] = detail
        if detail:
            row[remark_index] = "已上传"
            matched_count += 1
    return matched_count


def reference_correction_candidates(
    raw_reference: str,
    reference_universe: set[str],
) -> set[str]:
    candidates: set[str] = set()
    upper_reference = raw_reference.upper()
    compact = re.sub(r"\s+", "", upper_reference)
    cleaned = re.sub(r"[^0-9A-Z]", "", upper_reference)

    for token in re.findall(
        r"(?<!\d)(\d{11}[A-Z])(?![A-Z0-9])",
        upper_reference,
    ):
        if token in reference_universe:
            candidates.add(token)
    if cleaned in reference_universe:
        candidates.add(cleaned)
    if re.fullmatch(r"\d{11}", compact):
        candidates.update(
            reference
            for reference in reference_universe
            if reference[:11] == compact
        )
    if len(compact) == 11:
        for reference in reference_universe:
            if len(reference) == 12 and any(
                reference[:index] + reference[index + 1:] == compact
                for index in range(12)
            ):
                candidates.add(reference)
    elif len(compact) == 13:
        for index in range(13):
            candidate = compact[:index] + compact[index + 1:]
            if candidate in reference_universe:
                candidates.add(candidate)
    elif len(compact) == 12:
        for reference in reference_universe:
            if len(reference) == 12 and sum(
                left != right
                for left, right in zip(compact, reference)
            ) == 1:
                candidates.add(reference)
    return candidates


def correct_coupon_references(
    rows: list[list[object]],
    reference_universe: set[str],
) -> tuple[int, int, int, list[tuple[str, str, str, str]]]:
    """Correct references and record every decision for the processing report.

    The universe is built from submitted data only, so an operator has to be
    able to review each applied correction, not just the counts.
    """
    summary_index = COUPON_OUTPUT_HEADER.index("明细摘要")
    existing_counts = Counter(
        normalize_receipt_identifier(row[summary_index]).upper()
        for row in rows[1:]
        if normalize_receipt_identifier(row[summary_index])
    )
    proposed: dict[int, str] = {}
    target_counts: Counter[str] = Counter()
    unresolved_count = 0
    decisions: list[tuple[str, str, str, str]] = []

    for row_index, row in enumerate(rows[1:], start=1):
        raw_reference = normalize_receipt_identifier(
            row[summary_index]
        ).upper()
        if not raw_reference or raw_reference in reference_universe:
            continue
        candidates = reference_correction_candidates(
            raw_reference,
            reference_universe,
        )
        if len(candidates) != 1:
            unresolved_count += 1
            decisions.append(
                (
                    REFERENCE_REPORT_UNRESOLVED,
                    str(row_index + 1),
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
        raw_reference = normalize_receipt_identifier(
            rows[row_index][summary_index]
        ).upper()
        if existing_counts[target] > 0 or target_counts[target] > 1:
            collision_count += 1
            decisions.append(
                (
                    REFERENCE_REPORT_COLLISION,
                    str(row_index + 1),
                    raw_reference,
                    f"目标参考号 {target} 已被其他行占用，未纠正",
                )
            )
            continue
        rows[row_index][summary_index] = target
        corrected_count += 1
        decisions.append(
            (
                REFERENCE_REPORT_CORRECTED,
                str(row_index + 1),
                raw_reference,
                f"已自动纠正为 {target}，请人工复核",
            )
        )

    decisions.sort(key=lambda decision: REFERENCE_REPORT_ORDER[decision[0]])
    return corrected_count, unresolved_count, collision_count, decisions


def fill_unmatched_remarks(
    rows: list[list[object]],
    reference_universe: set[str],
) -> int:
    summary_index = COUPON_OUTPUT_HEADER.index("明细摘要")
    remark_index = COUPON_OUTPUT_HEADER.index("备注")
    unmatched_count = 0
    for row in rows[1:]:
        reference = normalize_receipt_identifier(
            row[summary_index]
        ).upper()
        if reference not in reference_universe:
            row[remark_index] = "未上传"
            unmatched_count += 1
    return unmatched_count


def build_coupon_summary(
    rows: list[list[object]],
    excluded_bottom_rows: int,
) -> list[tuple[object, ...]]:
    remark_index = COUPON_OUTPUT_HEADER.index("备注")
    subsidy_index = COUPON_OUTPUT_HEADER.index("2026数码国补（计入收入）")
    included_rows = (
        rows[1:-excluded_bottom_rows]
        if excluded_bottom_rows > 0
        else rows[1:]
    )
    grouped_counts: Counter[str] = Counter()
    grouped_totals: dict[str, Decimal] = {}
    for row in included_rows:
        remark = str(row[remark_index] or "").strip()
        grouped_counts[remark] += 1
        grouped_totals.setdefault(remark, Decimal("0"))
        subsidy = row[subsidy_index]
        if subsidy not in (None, ""):
            try:
                grouped_totals[remark] += Decimal(str(subsidy))
            except InvalidOperation as error:
                raise ValueError(
                    f"备注 {remark!r} 的2026数码国补金额无效：{subsidy!r}"
                ) from error

    summary_rows = [
        (
            remark,
            grouped_counts[remark],
            float(grouped_totals[remark]),
        )
        for remark in sorted(grouped_counts)
    ]
    summary_rows.append(
        (
            "合计",
            sum(grouped_counts.values()),
            float(sum(grouped_totals.values(), Decimal("0"))),
        )
    )
    return summary_rows


def validate_coupon_output(
    path: Path,
    expected_data_rows: int,
    expected_matched_rows: int,
    remark_lookup: dict[tuple[str, date], str],
    expected_matched_subsidy_total: Decimal,
    detail_lookup: dict[str, str],
    expected_uploaded_rows: int,
    reference_universe: set[str],
    expected_unresolved_rows: int,
    expected_unmatched_rows: int,
    expected_reference_decisions: list[tuple[str, str, str, str]],
    expected_summary_rows: list[tuple[object, ...]],
) -> None:
    workbook = load_workbook(path, data_only=True)
    try:
        if workbook.sheetnames != ["Sheet1", "Summary", REFERENCE_REPORT_SHEET]:
            raise RuntimeError(
                f"销售用券工作表校验失败：实际为 {workbook.sheetnames}"
            )
        report_sheet = workbook[REFERENCE_REPORT_SHEET]
        report_header = tuple(cell.value for cell in report_sheet[1])
        if report_header != REFERENCE_REPORT_HEADER:
            raise RuntimeError(
                f"参考号处理报告字段标题校验失败：实际为 {report_header}"
            )
        actual_decisions = [
            tuple(
                "" if value is None else str(value)
                for value in row
            )
            for row in report_sheet.iter_rows(min_row=2, values_only=True)
        ]
        if actual_decisions != [
            tuple(decision) for decision in expected_reference_decisions
        ]:
            raise RuntimeError(
                f"参考号处理报告内容校验失败：预期 "
                f"{len(expected_reference_decisions)} 条，"
                f"实际 {len(actual_decisions)} 条"
            )
        sheet = workbook["Sheet1"]
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
        actual_unresolved_rows = sum(
            reference not in reference_universe
            for reference in (
                normalize_receipt_identifier(
                    sheet.cell(row_number, summary_column).value
                ).upper()
                for row_number in range(2, sheet.max_row + 1)
            )
            if reference
        )
        if actual_unresolved_rows != expected_unresolved_rows:
            raise RuntimeError(
                f"销售用券未解决参考号数量校验失败：预期 "
                f"{expected_unresolved_rows} 条，实际 "
                f"{actual_unresolved_rows} 条"
            )
        product_name_column = COUPON_OUTPUT_HEADER.index("商品名称") + 1
        if any(
            sheet.cell(row_number, product_name_column).alignment.horizontal
            != "left"
            for row_number in range(2, sheet.max_row + 1)
        ):
            raise RuntimeError("销售用券商品名称列左对齐校验失败")
        matched_start_row = sheet.max_row - expected_matched_rows + 1
        actual_matched_subsidy_total = Decimal("0")
        subsidy_column = COUPON_OUTPUT_HEADER.index(
            "2026数码国补（计入收入）"
        ) + 1
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
                cell.fill.fill_type == "solid"
                and cell.fill.fgColor.rgb
                in {
                    COUPON_MATCH_FILL_COLOR,
                    f"00{COUPON_MATCH_FILL_COLOR}",
                    f"FF{COUPON_MATCH_FILL_COLOR}",
                }
                for cell in sheet[row_number]
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
                "销售用券匹配行的2026数码国补（计入收入）合计不为 0："
                f"{actual_matched_subsidy_total}"
            )

        summary_sheet = workbook["Summary"]
        summary_header = tuple(cell.value for cell in summary_sheet[1])
        if summary_header != COUPON_SUMMARY_HEADER:
            raise RuntimeError(
                f"销售用券汇总字段标题校验失败：实际为 {summary_header}"
            )
        actual_summary_rows = [
            tuple(row)
            for row in summary_sheet.iter_rows(
                min_row=2,
                values_only=True,
            )
        ]
        if actual_summary_rows != expected_summary_rows:
            raise RuntimeError("销售用券备注分组汇总校验失败")
        if not actual_summary_rows or actual_summary_rows[-1][0] != "合计":
            raise RuntimeError("销售用券汇总缺少底部合计行")
        detail_summary_rows = actual_summary_rows[:-1]
        total_summary_row = actual_summary_rows[-1]
        if sum(row[1] for row in detail_summary_rows) != (
            expected_data_rows - expected_matched_rows
        ):
            raise RuntimeError("销售用券汇总包含粉红色数据或数量不完整")
        if total_summary_row[1] != sum(
            row[1] for row in detail_summary_rows
        ):
            raise RuntimeError("销售用券汇总合计数量校验失败")
        if Decimal(str(total_summary_row[2])) != sum(
            (Decimal(str(row[2])) for row in detail_summary_rows),
            Decimal("0"),
        ):
            raise RuntimeError("销售用券汇总合计金额校验失败")
    finally:
        workbook.close()


def process_coupon_sales() -> None:
    if not COUPON_SOURCE_FILE.exists():
        raise FileNotFoundError(f"未找到源文件：{COUPON_SOURCE_FILE}")

    rows = read_coupon_rows(COUPON_SOURCE_FILE)
    remark_lookup = load_coupon_remark_lookup(COUPON_REMARK_SOURCE_FILE)
    matched_count, matched_subsidy_total = fill_coupon_remarks(
        rows,
        remark_lookup,
    )
    detail_lookup = load_uploaded_detail_lookup(COUPON_UPLOADED_SOURCE_FILE)
    # Unsubmitted data is no longer supplied, so submitted data is the only
    # source of valid references.
    reference_universe = set(detail_lookup)
    (
        corrected_count,
        unresolved_count,
        correction_collision_count,
        reference_decisions,
    ) = correct_coupon_references(rows, reference_universe)
    uploaded_count = fill_uploaded_details(rows, detail_lookup)
    unmatched_count = fill_unmatched_remarks(rows, reference_universe)
    summary_rows = build_coupon_summary(rows, matched_count)
    if as_currency(matched_subsidy_total) != Decimal("0.00"):
        raise ValueError(
            "备注匹配行的2026数码国补（计入收入）合计不为 0："
            f"{matched_subsidy_total}"
        )
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    for row in rows:
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
    matched_start_row = sheet.max_row - matched_count + 1
    for row in sheet.iter_rows(
        min_row=matched_start_row,
        max_row=sheet.max_row,
    ):
        for cell in row:
            cell.fill = matched_fill

    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(COUPON_SUMMARY_HEADER)
    for summary_row in summary_rows:
        summary_sheet.append(summary_row)
    format_sheet(summary_sheet, font_name, measurement_font)
    for cell in summary_sheet["C"][1:]:
        cell.number_format = "0.00"

    report_sheet = workbook.create_sheet(REFERENCE_REPORT_SHEET)
    report_sheet.append(REFERENCE_REPORT_HEADER)
    for decision in reference_decisions:
        report_sheet.append(decision)
    format_sheet(report_sheet, font_name, measurement_font, ("说明",))

    data_row_count = max(len(rows) - 1, 0)
    save_workbook_atomically(
        workbook,
        COUPON_OUTPUT_FILE,
        lambda path: validate_coupon_output(
            path,
            data_row_count,
            matched_count,
            remark_lookup,
            matched_subsidy_total,
            detail_lookup,
            uploaded_count,
            reference_universe,
            unresolved_count + correction_collision_count,
            unmatched_count,
            reference_decisions,
            summary_rows,
        ),
    )
    print(f"Subsidy coupon statistics complete: {data_row_count} rows")
    print(f"Remark matches: {matched_count}")
    print(f"Submitted status matches: {uploaded_count}")
    print(f"Rows not found in submitted data (marked 未上传): {unmatched_count}")
    print(
        f"Automatic reference corrections: {corrected_count}; "
        f"no unique candidate: {unresolved_count}; "
        f"duplicate conflicts: {correction_collision_count}"
    )
    print(
        "Total 2026 digital subsidy counted as revenue for matched rows: "
        f"{matched_subsidy_total:.2f}"
    )
    print(f"Output file: {COUPON_OUTPUT_FILE}")


def data_processors() -> tuple[tuple[str, Path, Callable[[], None]], ...]:
    """Report the paths configure_data_dir actually resolved, not assumed names.

    Source directories may carry their original Chinese names, so the label and
    the real path are not interchangeable.
    """
    return (
        ("submitted", INPUT_DIR, process_submitted_files),
        ("receipt_statistics", RECEIPTS_SOURCE_FILE, process_receipts),
        ("subsidy_coupon_statistics", COUPON_SOURCE_FILE, process_coupon_sales),
    )


def process_all() -> None:
    for _, source_path, processor in data_processors():
        print(f"Processing: {source_path}")
        processor()


def choose_data_processor() -> Callable[[], None] | None:
    processors = data_processors()

    print("Select a processing mode:")
    for index, (label, _, _) in enumerate(processors, start=1):
        print(f"  {index}. {label}")
    print("  4. all")
    print("  0. Exit")

    while True:
        try:
            choice = input("Enter a number: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nProcessing cancelled")
            return None

        if choice == "0":
            print("Exited")
            return None
        if choice == "4" or choice.lower() == "all":
            print("Processing all data in order: 1, 2, 3")
            return process_all
        if choice.isdigit():
            selected_index = int(choice) - 1
            if 0 <= selected_index < len(processors):
                _, source_path, processor = processors[selected_index]
                print(f"Processing: {source_path}")
                return processor

        print("Invalid input. Enter a menu number or all.")
