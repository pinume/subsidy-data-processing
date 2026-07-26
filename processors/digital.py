
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import partial
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

from processors.common.excel import (
    format_sheet,
    load_measurement_font,
    read_rows,
    resolve_font,
    save_workbook_atomically,
)
from processors.common.dates import (
    normalize_coupon_date,
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.common.paths import (
    find_data_files,
    match_source_file_by_header,
    read_xls_header,
)
from processors.large_appliances import _shared as large_appliances_shared


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
DATA_DIR: Path
SUBMITTED_FILES: tuple[Path, ...]
OUTPUT_FILE = OUTPUT_DIR / "数码_已上传.xlsx"
COUPON_SOURCE_FILE: Path | None
# Receipt statistics are shared with large appliances: same source file,
# same output file, processed once using that project's rules (which
# include special cases digital never had, such as 同型号换货).
COUPON_REMARK_SOURCE_FILE = large_appliances_shared.RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = OUTPUT_FILE

# Files live directly in the flat data directory; the digital submitted
# export is told apart from the large appliances one by filename keyword,
# and the coupon export (whose filename keyword both projects happen to
# share) by header content.
SUBMITTED_FILE_MARKER = "MER_89813014812B06R"
COUPON_STATISTICS_KEYWORD = "销售用券情况统计"
# The coupon export's field header row (row 2) at its last kept column
# (column 26); see COUPON_KEPT_SOURCE_COLUMNS below.
COUPON_SUBSIDY_HEADER = "2026数码国补（计入收入）"


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global SUBMITTED_FILES
    global COUPON_SOURCE_FILE

    DATA_DIR = data_dir
    SUBMITTED_FILES = tuple(
        find_data_files(data_dir, SUBMITTED_FILE_MARKER, (".xlsx",))
    )
    COUPON_SOURCE_FILE = match_source_file_by_header(
        find_data_files(data_dir, COUPON_STATISTICS_KEYWORD, (".xls",)),
        COUPON_SUBSIDY_HEADER,
        read_header=partial(read_xls_header, row=2, column=26),
    )

# Keep the required source columns in their original order.
KEPT_SOURCE_COLUMNS = ("D", "E", "F", "G", "I", "J", "Q", "S", "U", "W", "X")
KEPT_COLUMN_INDEXES = tuple(
    column_index_from_string(column) for column in KEPT_SOURCE_COLUMNS
)
# "补贴金额" is inserted by add_subsidy_column, so only source fields are listed.
REQUIRED_SUBMITTED_HEADERS = ("状态", "描述", "交易金额")
# Digital: 15% of the transaction, capped at 500 per order. Household
# appliances use the same rate but a 1500 cap — see
# processors/large_appliances/submitted.py.
SUBSIDY_RATE = Decimal("0.15")
SUBSIDY_CAP = Decimal("500")
STATUS_ORDER = (
    "核销失败",
    "审核失败",
    "暂存",
    "同步(已上送)",
    "待审核",
    "审核通过",
)
STATUS_PRIORITY = {status: index for index, status in enumerate(STATUS_ORDER)}
COUPON_KEPT_SOURCE_COLUMNS = (3, 4, 6, 8, 15, 18, 26)
COUPON_OUTPUT_HEADER = (
    "单据号",
    "单据日期",
    "商品名称",
    "品牌",
    "财务大类",
    "明细摘要",
    COUPON_SUBSIDY_HEADER,
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
COUPON_MATCH_FILL_COLOR = "FFC7CE"
COUPON_REFERENCE_RE = re.compile(r"\d{11}[A-Z]")
COUPON_SUMMARY_HEADER = (
    "备注",
    "数量",
    f"{COUPON_SUBSIDY_HEADER}合计",
)
DETAILS_SHEET_NAME = "数码-明细总表"


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
            calculated = Decimal(str(amount)) * SUBSIDY_RATE
            subsidy = float(
                min(calculated, SUBSIDY_CAP).quantize(
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
    files = list(SUBMITTED_FILES)
    if not files:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 中找到文件名包含"
            f"“{SUBMITTED_FILE_MARKER}”的 .xlsx 文件"
        )

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
                        Decimal(str(amount)) * SUBSIDY_RATE,
                        SUBSIDY_CAP,
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


def load_uploaded_subsidy_stats(source: Path) -> tuple[int, Decimal]:
    if not source.exists():
        raise FileNotFoundError(f"未找到已上传统计文件：{source}")

    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if "Summary" not in workbook.sheetnames:
            raise ValueError(f"{source.name} 缺少“汇总”工作表")
        sheet = workbook["Summary"]
        header = [cell.value for cell in sheet[1]]
        if "补贴金额" not in header:
            raise ValueError(f"{source.name} 缺少字段：补贴金额")

        subsidy_index = header.index("补贴金额")
        subsidy_count = 0
        subsidy_total = Decimal("0")
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True),
            start=2,
        ):
            subsidy = row[subsidy_index]
            if subsidy in (None, ""):
                continue
            try:
                subsidy_total += Decimal(str(subsidy))
            except InvalidOperation as error:
                raise ValueError(
                    f"{source.name} 第 {row_number} 行补贴金额无效：{subsidy!r}"
                ) from error
            subsidy_count += 1
        return subsidy_count, subsidy_total
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


def reference_decision(
    outcome: str,
    row: list[object],
    raw_reference: str,
    note: str,
) -> tuple[str, str, object, str, str]:
    """Identify a decision by 单据号 + 单据日期, not by row position.

    The detail rows get re-sorted after corrections are applied, so a row
    number recorded here would point at the wrong row in the saved sheet.
    """
    return (
        outcome,
        normalize_document_number(row[COUPON_OUTPUT_HEADER.index("单据号")]),
        row[COUPON_OUTPUT_HEADER.index("单据日期")],
        raw_reference,
        note,
    )


def correct_coupon_references(
    rows: list[list[object]],
    reference_universe: set[str],
) -> tuple[int, int, int, list[tuple[str, str, object, str, str]]]:
    """Correct references and record every decision for the processing report.

    The universe is built from submitted data only, so an operator has to be
    able to review each applied correction, not just the counts.

    A well-formed reference with no candidate at all is simply absent from the
    submitted data — the detail row already carries a 未上传 remark, so it is
    counted but kept out of the report. Only malformed references and genuine
    ambiguities are reported, which is what an operator can actually act on.
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
    decisions: list[tuple[str, str, object, str, str]] = []

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
            if candidates or not COUPON_REFERENCE_RE.fullmatch(raw_reference):
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
        raw_reference = normalize_receipt_identifier(
            row[summary_index]
        ).upper()
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
        row[summary_index] = target
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
    uploaded_count: int,
    uploaded_subsidy_total: Decimal,
) -> list[tuple[object, ...]]:
    subsidy_index = COUPON_OUTPUT_HEADER.index("2026数码国补（计入收入）")
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
    in processors/large_appliances/coupons.py."""

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


def compute_coupon_data() -> CouponComputation:
    if COUPON_SOURCE_FILE is None:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 中找到文件名包含"
            f"“{COUPON_STATISTICS_KEYWORD}”且表头为"
            f"“{COUPON_SUBSIDY_HEADER}”的 .XLS 文件"
        )

    rows = read_coupon_rows(COUPON_SOURCE_FILE)
    remark_lookup = load_coupon_remark_lookup(COUPON_REMARK_SOURCE_FILE)
    matched_count, matched_subsidy_total = fill_coupon_remarks(
        rows,
        remark_lookup,
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
    ) = correct_coupon_references(rows, reference_universe)
    uploaded_match_count = fill_uploaded_details(rows, detail_lookup)
    unmatched_count = fill_unmatched_remarks(rows, reference_universe)
    summary_rows = build_coupon_summary(
        rows,
        uploaded_subsidy_count,
        uploaded_subsidy_total,
    )
    if as_currency(matched_subsidy_total) != Decimal("0.00"):
        raise ValueError(
            "备注匹配行的2026数码国补（计入收入）合计不为 0："
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


