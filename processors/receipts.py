"""收款单统计 (receipt statistics) processing, shared by both projects.

Household appliances and digital both read this same source file and are
processed with the same rules (same-model replacement plus the special
remarks in config/receipt_special_remarks.yaml); there has never been a
digital-specific variant, so this lives at the top level rather than under
either project's package.
"""

from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from python_calamine import CalamineWorkbook
from xlsxwriter import Workbook

from processors.common.config import load_receipt_special_remark_keys
from processors.common.dates import (
    is_valid_original_invoice_number,
    normalize_document_number,
    normalize_receipt_date,
    normalize_receipt_identifier,
    receipt_match_key,
)
from processors.common.excel import (
    FONT_SIZE,
    ROW_HEIGHT,
    load_measurement_font,
    normalize_calamine_value,
    pixels_to_column_pixels,
    resolve_font,
    width_measurer,
    write_xlsx_atomically,
)
from processors.common.paths import find_data_files, resolve_unique_file

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "收款单统计.xlsx"

DATA_DIR: Path
RECEIPTS_SOURCE_FILE: Path | None
RECEIPT_STATISTICS_KEYWORD = "收款单统计"

RECEIPTS_SOURCE_HEADER = ("单据号", "日期", "原票号", "摘要", "商品名称")
RECEIPTS_OUTPUT_HEADER = (*RECEIPTS_SOURCE_HEADER, "备注")
RECEIPTS_REMARK_INDEX = RECEIPTS_OUTPUT_HEADER.index("备注")
RECEIPTS_REMARK_RETURN = "退换货/倒票（退单）"
RECEIPTS_REMARK_ORIGINAL = "退换货/倒票（原单）"
RECEIPTS_REMARK_BOTH = "退换货/倒票（退单及原单）"
RECEIPTS_REMARK_SAME_MODEL_REPLACEMENT = "已做同型号换货处理"
RECEIPTS_REMARK_SPECIAL = r"退换货\倒票"
RECEIPTS_SPECIAL_REMARK_KEYS = load_receipt_special_remark_keys()
RECEIPTS_ROW_HEIGHT = 20
RECEIPTS_REMARK_FILL_COLOR = "FFC7CE"
# The same colour as openpyxl may report it: bare, alpha-prefixed by the reader,
# or alpha-prefixed by XlsxWriter on the way out.
RECEIPTS_REMARK_FILL_SPELLINGS = frozenset(
    {
        RECEIPTS_REMARK_FILL_COLOR,
        f"00{RECEIPTS_REMARK_FILL_COLOR}",
        f"FF{RECEIPTS_REMARK_FILL_COLOR}",
    }
)
RECEIPTS_EXCLUDED_PRODUCT_KEYWORD = "北国"
RECEIPTS_SAME_MODEL_REPLACEMENT_KEYWORD = "同型号换货"

ISSUES_SHEET_NAME = "问题明细"
ISSUES_HEADER = ("问题类型", "行号", "内容", "说明")


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global RECEIPTS_SOURCE_FILE

    DATA_DIR = data_dir
    RECEIPTS_SOURCE_FILE = resolve_unique_file(
        find_data_files(data_dir, RECEIPT_STATISTICS_KEYWORD, (".xlsx",))
    )


def read_receipt_rows(source: Path) -> list[list[object]]:
    """Read the 收款单统计 .xlsx export: title row, header row, data, 合计.

    Returns every source row at its original position, including blank rows
    and the 合计 row. Dropping them here would renumber everything after
    them, and prepare_receipt_data derives each row's Excel row number from
    its position in this list — so an operator chasing a reported problem
    row would be sent to the wrong line. Filtering happens there instead,
    inside the loop that already knows the true row number.

    Reads via python-calamine rather than openpyxl: the source export always
    carries far more columns than the five kept here (60 in a real file), and
    openpyxl's read_only parser still parses every column of every row before
    select() discards the rest, making it the dominant cost of this whole
    processing step. calamine parses the same file roughly 5x faster.
    """
    workbook = CalamineWorkbook.from_path(str(source))
    try:
        sheet = workbook.get_sheet_by_index(0)
        if sheet.start is None:
            # iter_rows() panics inside the Rust extension on a sheet with no
            # cells at all, so a blank export has to be rejected before it.
            raise ValueError(f"{source.name} 缺少总标题行或字段标题行")
        # Raw iter_rows(), not calamine_rows(): normalizing all 60 source
        # columns to keep only 5 of them is what this reader exists to avoid.
        # Column positions are only ever resolved against this same row shape
        # by header name, so calamine's dropped leading blank columns — which
        # calamine_rows pads back — cannot misalign anything here.
        rows_iter = sheet.iter_rows()
        title_row = next(rows_iter, None)
        header_row = next(rows_iter, None)
        if title_row is None or header_row is None:
            raise ValueError(f"{source.name} 缺少总标题行或字段标题行")
        header_row = [normalize_calamine_value(value) for value in header_row]

        source_headers = [
            str(value).strip() if value is not None else ""
            for value in header_row
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

        def select(values: list[object]) -> list[object]:
            return [
                normalize_calamine_value(values[index])
                if index < len(values) else None
                for index in source_column_indexes
            ]

        rows: list[list[object]] = [select(header_row)]
        for source_row in rows_iter:
            rows.append(select(source_row))

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
        workbook.close()


def is_empty_receipt_row(row: list[object]) -> bool:
    """Whether every kept field is blank.

    A row is only blank when all five are: a row carrying just one value is
    real, incomplete data that still belongs in the output and in 问题明细.
    The comparison is against the stripped string so a formatting-only row
    Excel left behind counts as blank, while a numeric 0 does not.
    """
    return all(value is None or str(value).strip() == "" for value in row)


def is_receipt_total_row(row: list[object]) -> bool:
    """Whether this is the export's own 合计 row.

    Only 单据号 and 日期 are examined, and only for an exact match: "合计"
    appearing inside a 摘要 or 商品名称 is ordinary business text, and
    treating it as a total row would silently delete a real sale. Which of
    the two columns carries the word has varied between exports, so both
    are accepted.
    """
    document_number = str(row[0] or "").strip()
    receipt_date = str(row[1] or "").strip()
    return document_number == "合计" or receipt_date == "合计"


def receipt_remark(
    has_original: bool,
    is_referenced: bool,
    is_same_model_replacement: bool = False,
    is_special: bool = False,
) -> str | None:
    if is_special:
        return RECEIPTS_REMARK_SPECIAL
    if has_original and is_referenced:
        return RECEIPTS_REMARK_BOTH
    if is_same_model_replacement:
        return RECEIPTS_REMARK_SAME_MODEL_REPLACEMENT
    if has_original:
        return RECEIPTS_REMARK_RETURN
    if is_referenced:
        return RECEIPTS_REMARK_ORIGINAL
    return None


def receipt_output_sort_key(
    remark: object,
    receipt_date: object,
    document_number: object,
    product_name: object,
) -> tuple[str, bool, date, str, str]:
    """Sort by 备注、日期、单据号、商品名称; missing dates sort last.

    An empty remark sorts before every actual remark, so ordinary sales stay
    together at the top and the highlighted return/exchange rows form the
    trailing section of the output.
    """
    normalized_date = (
        receipt_date.date() if isinstance(receipt_date, datetime) else receipt_date
    )
    has_no_date = not isinstance(normalized_date, date)
    return (
        normalize_receipt_identifier(remark),
        has_no_date,
        date.max if has_no_date else normalized_date,
        normalize_receipt_identifier(document_number),
        normalize_receipt_identifier(product_name),
    )


def prepare_receipt_data(kept_rows: list[list[object]], source_name: str):
    records: list[dict[str, object]] = []
    key_rows: dict[str, list[int]] = {}
    referenced_original_invoice_numbers: set[str] = set()
    same_model_replacement_original_invoice_numbers: set[str] = set()
    excluded_product_count = 0
    blank_row_count = 0
    total_row_count = 0

    # start=3 because kept_rows[0] is the field header (Excel row 2) and the
    # export opens with a title row; every row read stays at its original
    # index, so source_row is the real Excel row number even when rows above
    # it are skipped below.
    for source_row, row in enumerate(kept_rows[1:], start=3):
        if is_empty_receipt_row(row):
            blank_row_count += 1
            continue
        if is_receipt_total_row(row):
            total_row_count += 1
            continue

        product_name = normalize_receipt_identifier(row[4])
        if RECEIPTS_EXCLUDED_PRODUCT_KEYWORD in product_name:
            excluded_product_count += 1
            continue

        document_number = normalize_document_number(row[0])
        receipt_date = normalize_receipt_date(
            row[1], source_row=source_row, source_name=source_name
        )
        original_invoice_number = normalize_receipt_identifier(row[2])
        match_key = (
            receipt_match_key(receipt_date, document_number)
            if receipt_date is not None and document_number
            else ""
        )
        if match_key:
            key_rows.setdefault(match_key, []).append(source_row)
        if original_invoice_number:
            if RECEIPTS_SAME_MODEL_REPLACEMENT_KEYWORD in (
                normalize_receipt_identifier(row[3])
            ):
                same_model_replacement_original_invoice_numbers.add(
                    original_invoice_number
                )
            else:
                referenced_original_invoice_numbers.add(
                    original_invoice_number
                )
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

    # Remarks depend on the complete set of original-invoice references, so
    # compute them only after every source row has been scanned. Sorting then
    # uses the final remark as its primary key.
    for record in records:
        original_invoice_number = str(record["original_invoice_number"])
        match_key = str(record["match_key"])
        has_original = bool(original_invoice_number)
        is_referenced = bool(
            match_key and match_key in referenced_original_invoice_numbers
        )
        is_same_model_replacement = bool(
            match_key
            and match_key in same_model_replacement_original_invoice_numbers
            and not is_referenced
        )
        is_special = match_key in RECEIPTS_SPECIAL_REMARK_KEYS
        record["has_original"] = has_original
        record["is_referenced"] = is_referenced
        record["is_same_model_replacement"] = is_same_model_replacement
        record["is_special"] = is_special
        record["remark"] = receipt_remark(
            has_original,
            is_referenced,
            is_same_model_replacement,
            is_special,
        )

    records.sort(
        key=lambda record: receipt_output_sort_key(
            record["remark"],
            record["receipt_date"],
            record["document_number"],
            record["product_name"],
        )
    )

    # Every record ends up in output_rows in this sorted order, one row per
    # record. Assign output coordinates only now so 问题明细 points at the row
    # an operator will find in the generated Sheet1, not the raw import.
    for position, record in enumerate(records, start=2):
        record["output_row"] = position

    # Same-document 烟机+灶具 (and similar kitchen-suite) sales are entered as
    # multiple line items sharing one 单据号/日期, which is expected here, not
    # a data problem — so it is neither reported in 问题明细 nor highlighted.
    issues: list[tuple[str, str, str, str]] = []

    # 原票号 referencing a sale from a year earlier than anything in this
    # file can never resolve to a match_key here by construction (the file
    # only carries one year's receipts), so it is not a data problem either.
    receipt_years = {
        record["receipt_date"].year
        for record in records
        if record["receipt_date"] is not None
    }
    min_receipt_year = min(receipt_years) if receipt_years else None

    only_return_count = 0
    only_original_count = 0
    both_count = 0
    unmatched_original_count = 0
    invalid_original_count = 0
    missing_match_key_count = 0
    remark_count = 0
    special_remark_count = 0
    output_rows: list[list[object]] = []
    for record in records:
        output_row = int(record["output_row"])
        original_invoice_number = str(record["original_invoice_number"])
        match_key = str(record["match_key"])
        has_original = bool(record["has_original"])
        is_referenced = bool(record["is_referenced"])
        is_same_model_replacement = bool(record["is_same_model_replacement"])
        is_special = bool(record["is_special"])

        if not match_key:
            missing_match_key_count += 1
            issues.append(
                (
                    "缺少匹配键",
                    str(output_row),
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
                    str(output_row),
                    original_invoice_number,
                    "原票号应为6位日期加单据号",
                )
            )
        if has_original and original_invoice_number not in key_rows:
            is_prior_period_reference = (
                min_receipt_year is not None
                and is_valid_original_invoice_number(original_invoice_number)
                and datetime.strptime(
                    original_invoice_number[:6], "%y%m%d"
                ).year
                < min_receipt_year
            )
            if not is_prior_period_reference:
                unmatched_original_count += 1
                issues.append(
                    (
                        "原票号未匹配",
                        str(output_row),
                        original_invoice_number,
                        "未找到日期与单据号组合键相同的原单",
                    )
                )

        if has_original and (is_referenced or is_same_model_replacement):
            both_count += 1
        elif has_original:
            only_return_count += 1
        elif is_referenced or is_same_model_replacement:
            only_original_count += 1

        # Counted from the remark actually written, not by re-adding the
        # three 退单/原单 tallies: a row whose remark comes only from a
        # special match key belongs to none of them, so summing those three
        # silently understated the total.
        remark = record["remark"]
        if remark:
            remark_count += 1
        if is_special:
            special_remark_count += 1

        output_rows.append(
            [
                record["document_number"],
                record["receipt_date"],
                original_invoice_number or None,
                record["summary"],
                record["product_name"],
                remark,
            ]
        )

    applied_special_keys = RECEIPTS_SPECIAL_REMARK_KEYS & key_rows.keys()
    stats = {
        "总数据量": len(records),
        "删除北国商品行数": excluded_product_count,
        "跳过空白行数": blank_row_count,
        "跳过合计行数": total_row_count,
        "仅退单数量": only_return_count,
        "仅原单数量": only_original_count,
        "退单及原单数量": both_count,
        "备注总数": remark_count,
        "特殊备注数量": special_remark_count,
        "未匹配原票号数量": unmatched_original_count,
        "重复匹配键数量": sum(
            1 for source_rows in key_rows.values() if len(source_rows) > 1
        ),
        "原票号格式异常数量": invalid_original_count,
        "缺少匹配键数量": missing_match_key_count,
        "生效特殊匹配键": sorted(applied_special_keys),
        "未找到特殊匹配键": sorted(
            RECEIPTS_SPECIAL_REMARK_KEYS - applied_special_keys
        ),
    }
    return output_rows, stats, issues


def _write_issues_sheet(
    workbook: Workbook,
    issues: list[tuple[str, str, str, str]],
    font_name: str,
    measurement_font,
) -> None:
    sheet = workbook.add_worksheet(ISSUES_SHEET_NAME)
    header_format = workbook.add_format(
        {
            "font_name": font_name,
            "font_size": FONT_SIZE,
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#000000",
            "align": "center",
            "valign": "vcenter",
        }
    )
    centered_format = workbook.add_format(
        {
            "font_name": font_name,
            "font_size": FONT_SIZE,
            "font_color": "#000000",
            "align": "center",
            "valign": "vcenter",
        }
    )
    left_format = workbook.add_format(
        {
            "font_name": font_name,
            "font_size": FONT_SIZE,
            "font_color": "#000000",
            "align": "left",
            "valign": "vcenter",
        }
    )
    left_columns = {
        ISSUES_HEADER.index("内容"),
        ISSUES_HEADER.index("说明"),
    }
    measure = width_measurer(measurement_font)
    maximum_widths = [measure(value) for value in ISSUES_HEADER]

    sheet.set_row(0, ROW_HEIGHT)
    for column, value in enumerate(ISSUES_HEADER):
        sheet.write(0, column, value, header_format)
    for row_number, row in enumerate(issues, start=1):
        sheet.set_row(row_number, ROW_HEIGHT)
        for column, value in enumerate(row):
            cell_format = left_format if column in left_columns else centered_format
            sheet.write(row_number, column, value, cell_format)
            maximum_widths[column] = max(maximum_widths[column], measure(value))

    sheet.freeze_panes(1, 0)
    sheet.autofilter(0, 0, len(issues), len(ISSUES_HEADER) - 1)
    for column, maximum_pixels in enumerate(maximum_widths):
        sheet.set_column_pixels(
            column,
            column,
            pixels_to_column_pixels(maximum_pixels),
        )


def validate_receipts_output(
    path: Path,
    expected_data_rows: int,
    expected_issues: list[tuple[str, str, str, str]],
) -> None:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        expected_sheet_names = (
            ["Sheet1"] if not expected_issues else ["Sheet1", ISSUES_SHEET_NAME]
        )
        if workbook.sheetnames != expected_sheet_names:
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

        rows = sheet.iter_rows(
            min_row=1,
            min_col=1,
            max_col=len(RECEIPTS_OUTPUT_HEADER),
        )
        header_cells = next(rows, ())
        header = tuple(cell.value for cell in header_cells)
        if header != RECEIPTS_OUTPUT_HEADER:
            raise RuntimeError(
                f"收款单表头校验失败：预期 {RECEIPTS_OUTPUT_HEADER}，"
                f"实际 {header}"
            )

        referenced_original_invoice_numbers: set[str] = set()
        same_model_replacement_original_invoice_numbers: set[str] = set()
        snapshots = []
        for row_number, row in enumerate(rows, start=2):
            (
                document_cell,
                date_cell,
                original_cell,
                summary_cell,
                product_cell,
                remark_cell,
            ) = row
            document_number = normalize_receipt_identifier(document_cell.value)
            date_value = date_cell.value
            original_invoice_number = normalize_receipt_identifier(
                original_cell.value
            )
            if original_invoice_number:
                summary = normalize_receipt_identifier(summary_cell.value)
                target = (
                    same_model_replacement_original_invoice_numbers
                    if RECEIPTS_SAME_MODEL_REPLACEMENT_KEYWORD in summary
                    else referenced_original_invoice_numbers
                )
                target.add(original_invoice_number)

            match_key = ""
            if date_value is None or not document_number:
                pass
            else:
                match_key = receipt_match_key(
                    date_value.date()
                    if isinstance(date_value, datetime)
                    else date_value,
                    document_number,
                )

            # openpyxl does not promise a str here: a fill this program wrote
            # yields one, but a theme-coloured cell yields an RGB object that
            # cannot be sliced. Comparing against the three spellings of the
            # colour works for either, and matches how the coupon validators
            # already do it.
            pink_cells = tuple(
                cell.fill.fill_type == "solid"
                and cell.fill.fgColor.rgb in RECEIPTS_REMARK_FILL_SPELLINGS
                for cell in row
            )
            snapshots.append(
                (
                    row_number,
                    document_cell.value,
                    date_value,
                    original_invoice_number,
                    remark_cell.value,
                    date_cell.number_format,
                    match_key,
                    product_cell.value,
                    pink_cells,
                )
            )

        for (
            row_number,
            document_number,
            date_value,
            original_invoice_number,
            actual_remark,
            date_number_format,
            match_key,
            product_name,
            pink_cells,
        ) in snapshots:
            if document_number is not None and (
                not isinstance(document_number, str)
                or document_number.startswith("收款")
            ):
                raise RuntimeError(
                    f"收款单第 {row_number} 行的单据号格式不正确"
                )

            if date_value is not None:
                if not isinstance(date_value, (date, datetime)):
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的日期不是有效日期"
                    )
                if date_number_format != "yyyymmdd":
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的日期格式不正确"
                    )

            expected_remark = receipt_remark(
                bool(original_invoice_number),
                bool(
                    match_key
                    and match_key in referenced_original_invoice_numbers
                ),
                bool(
                    match_key
                    and match_key
                    in same_model_replacement_original_invoice_numbers
                    and match_key
                    not in referenced_original_invoice_numbers
                ),
                match_key in RECEIPTS_SPECIAL_REMARK_KEYS,
            )
            if actual_remark != expected_remark:
                raise RuntimeError(
                    f"收款单第 {row_number} 行的备注校验失败"
                )

            should_be_pink = bool(expected_remark)
            for is_pink in pink_cells:
                if is_pink != should_be_pink:
                    raise RuntimeError(
                        f"收款单第 {row_number} 行的退换货备注填充不正确"
                    )

        actual_sort_keys = [
            receipt_output_sort_key(
                actual_remark,
                date_value,
                document_number,
                product_name,
            )
            for (
                _row_number,
                document_number,
                date_value,
                _original_invoice_number,
                actual_remark,
                _date_number_format,
                _match_key,
                product_name,
                _pink_cells,
            ) in snapshots
        ]
        if actual_sort_keys != sorted(actual_sort_keys):
            raise RuntimeError(
                "收款单排序校验失败：应按备注、日期、单据号、商品名称升序排列"
            )

        if expected_issues:
            issues_sheet = workbook[ISSUES_SHEET_NAME]
            issues_header = tuple(
                cell.value for cell in next(issues_sheet.iter_rows(max_row=1))
            )
            if issues_header != ISSUES_HEADER:
                raise RuntimeError(
                    f"{ISSUES_SHEET_NAME}工作表字段标题校验失败：实际为 {issues_header}"
                )

            def comparable(row: tuple[object, ...]) -> tuple[object, ...]:
                # A cell saved with "" round-trips through openpyxl as None,
                # so blanks on either side must compare equal.
                return tuple("" if value is None else value for value in row)

            actual_issues = [
                comparable(row)
                for row in issues_sheet.iter_rows(min_row=2, values_only=True)
            ]
            expected = [comparable(tuple(issue)) for issue in expected_issues]
            if actual_issues != expected:
                raise RuntimeError(
                    f"{ISSUES_SHEET_NAME}工作表内容校验失败：预期 {len(expected)} 条，"
                    f"实际 {len(actual_issues)} 条"
                )
    finally:
        workbook.close()


def _write_receipts_workbook(
    path: Path,
    output_rows: list[list[object]],
    issues: list[tuple[str, str, str, str]],
) -> None:
    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    with Workbook(
        str(path),
        {
            "constant_memory": True,
            "strings_to_urls": False,
            # 商品名称 and 备注 are free text copied from the source export; a
            # value starting with "=" is data, not a formula to evaluate.
            "strings_to_formulas": False,
        },
    ) as workbook:
        sheet = workbook.add_worksheet("Sheet1")
        header_format = workbook.add_format(
            {
                "font_name": font_name,
                "font_size": FONT_SIZE,
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#000000",
                "align": "center",
                "valign": "vcenter",
            }
        )
        body_formats = {}
        for has_remark in (False, True):
            for number_format_kind in ("general", "text", "date"):
                properties = {
                    "font_name": font_name,
                    "font_size": FONT_SIZE,
                    "font_color": "#000000",
                    "align": "center",
                    "valign": "vcenter",
                }
                if has_remark:
                    properties["bg_color"] = f"#{RECEIPTS_REMARK_FILL_COLOR}"
                if number_format_kind == "text":
                    properties["num_format"] = "@"
                elif number_format_kind == "date":
                    properties["num_format"] = "yyyymmdd"
                body_formats[(has_remark, number_format_kind)] = (
                    workbook.add_format(properties)
                )

        measure = width_measurer(measurement_font)
        maximum_widths = [measure(value) for value in RECEIPTS_OUTPUT_HEADER]
        sheet.set_row(0, RECEIPTS_ROW_HEIGHT)
        for column, value in enumerate(RECEIPTS_OUTPUT_HEADER):
            sheet.write(0, column, value, header_format)

        for row_number, row in enumerate(output_rows, start=1):
            sheet.set_row(row_number, RECEIPTS_ROW_HEIGHT)
            has_remark = row[RECEIPTS_REMARK_INDEX] is not None
            for column, value in enumerate(row):
                number_format_kind = (
                    "text"
                    if column == 0
                    else "date"
                    if column == 1 and value not in (None, "")
                    else "general"
                )
                sheet.write(
                    row_number,
                    column,
                    value,
                    body_formats[(has_remark, number_format_kind)],
                )
                maximum_widths[column] = max(maximum_widths[column], measure(value))

        sheet.freeze_panes(1, 0)
        for column, maximum_pixels in enumerate(maximum_widths):
            sheet.set_column_pixels(
                column,
                column,
                pixels_to_column_pixels(maximum_pixels),
            )

        if issues:
            _write_issues_sheet(
                workbook,
                issues,
                font_name,
                measurement_font,
            )


def process_receipts() -> None:
    if RECEIPTS_SOURCE_FILE is None:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 中找到文件名包含"
            f"“{RECEIPT_STATISTICS_KEYWORD}”的 .XLSX 文件"
        )

    kept_rows = read_receipt_rows(RECEIPTS_SOURCE_FILE)
    output_rows, stats, issues = prepare_receipt_data(
        kept_rows, source_name=RECEIPTS_SOURCE_FILE.name
    )
    row_count = len(output_rows)
    write_xlsx_atomically(
        OUTPUT_FILE,
        lambda path: _write_receipts_workbook(
            path,
            output_rows,
            issues,
        ),
        lambda path: validate_receipts_output(path, row_count, issues),
    )
    print(f"Receipt statistics complete: {row_count} rows")
    print(
        f"Remarks: {stats['备注总数']} "
        f"(special: {stats['特殊备注数量']}); "
        f"unmatched original invoices: {stats['未匹配原票号数量']}; "
        f"duplicate match keys: {stats['重复匹配键数量']}"
    )
    if stats["跳过空白行数"] or stats["跳过合计行数"]:
        print(
            f"Skipped blank rows: {stats['跳过空白行数']}; "
            f"total rows: {stats['跳过合计行数']}"
        )
    print(
        f"Excluded rows containing "
        f"'{RECEIPTS_EXCLUDED_PRODUCT_KEYWORD}': {stats['删除北国商品行数']}"
    )
    print(f"Special receipt remarks applied: {len(stats['生效特殊匹配键'])}")
    # Reported, never fatal: the configured keys are one-off exceptions, and
    # a run covering a different date range legitimately contains none of
    # them. Silence would instead hide a key typo'd into never matching.
    if stats["未找到特殊匹配键"]:
        print(
            "未在本次收款单中找到特殊匹配键："
            + "、".join(stats["未找到特殊匹配键"])
        )
    if issues:
        print(f"Issues logged in '{ISSUES_SHEET_NAME}': {len(issues)}")
    else:
        print(f"No issues; '{ISSUES_SHEET_NAME}' sheet was not created")
    print(f"Output file: {OUTPUT_FILE}")
