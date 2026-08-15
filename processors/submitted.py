"""已上传数据 (submitted export) processing for household appliances and digital.

Both projects read the same source column layout, apply the same 15%
subsidy rate, and validate the same shape of output; only the subsidy cap,
source marker, and output file differ, so they are captured as two
SubmittedProfile instances sharing one pipeline rather than two
near-identical modules.
"""

from collections import Counter
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl.utils import column_index_from_string
from python_calamine import CalamineWorkbook
from xlsxwriter import Workbook

from processors.common.config import submitted_file_marker
from processors.common.console import ConsoleReporter, format_count
from processors.common.excel import (
    calamine_rows,
    load_measurement_font,
    resolve_font,
    run_with_output_rollback,
    write_formatted_sheet,
    write_xlsx_atomically,
)
from processors.common.paths import find_data_files
from processors.common.references import normalize_reference, validated_reference

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"

# Keep the required source columns in their original order, as they appear in
# a MER_*.xlsx export:
#   D 订单号  E 交易日期  F 交易金额  G 检索参考号  I 状态  J 描述
# Q 详细地址, S tel, U 发票金额, W 图片1 and X S/N码 were kept once but are not
# used by either project's downstream reports, so they are no longer carried
# over. add_subsidy_column depends on the leading four: it reads the amount at
# index 2 and inserts 补贴金额 at index 3.
KEPT_SOURCE_COLUMNS = ("D", "E", "F", "G", "I", "J")


KEPT_COLUMN_INDEXES = tuple(
    column_index_from_string(column) for column in KEPT_SOURCE_COLUMNS
)
# "补贴金额" is inserted by add_subsidy_column, so only source fields are listed.
REQUIRED_SUBMITTED_HEADERS = ("检索参考号", "状态", "描述", "交易金额")
STATUS_ORDER = (
    "核销失败",
    "审核失败",
    "暂存",
    "待同步",
    "同步(已上送)",
    "待审核",
    "审核通过",
)


@dataclass(frozen=True)
class SubmittedProfile:
    output_file: Path
    subsidy_rate: Decimal
    subsidy_cap: Decimal


@dataclass(frozen=True)
class UnknownStatusRecord:
    """One Summary-only row whose status has no sheet of its own.

    Carries where the row came from (source file and row) and what it
    carried (reference and status), so the operator warning can point at
    real rows instead of just counting them. Report-only: the workbook
    keeps using summary_rows/status_rows.
    """

    source_name: str
    source_row: int
    reference: str
    status: str


@dataclass(frozen=True)
class SubmittedReport:
    header: tuple[object, ...]
    summary_rows: list[list[object]]
    status_rows: dict[str, list[list[object]]]
    file_count: int
    data_row_count: int
    # Rows whose status STATUS_ORDER does not name, with their source
    # location, for the console warning only (see UnknownStatusRecord).
    # Those rows reach Summary but no status sheet, so without the warning
    # they are invisible to an operator working from the status tabs.
    unknown_status_records: tuple[UnknownStatusRecord, ...]
    # Invalid exports (no worksheets). Deletion is deferred to
    # process_submitted_files, after every input has been validated, so the
    # corrected() record exists even when a later step fails.
    deleted_invalid_files: tuple[str, ...]


# Household appliances and digital both take 15% of the transaction; the two
# caps were once both written as 500, which silently understated 43% of the
# appliance rows, so they stay pinned separately rather than sharing a
# default.
PROFILES: dict[str, SubmittedProfile] = {
    "家电": SubmittedProfile(
        output_file=OUTPUT_DIR / "家电_已上传.xlsx",
        subsidy_rate=Decimal("0.15"),
        subsidy_cap=Decimal("1500"),
    ),
    "数码": SubmittedProfile(
        output_file=OUTPUT_DIR / "数码_已上传.xlsx",
        subsidy_rate=Decimal("0.15"),
        subsidy_cap=Decimal("500"),
    ),
}
PROFILE_ORDER = ("家电", "数码")
OUTPUT_FILES = tuple(PROFILES[name].output_file for name in PROFILE_ORDER)

DATA_DIR: Path
SUBMITTED_FILE_MARKERS: dict[str, str] = {}
INPUT_FILES: dict[str, tuple[Path, ...]] = {}


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR

    DATA_DIR = data_dir
    for name in PROFILE_ORDER:
        marker = submitted_file_marker(name)
        SUBMITTED_FILE_MARKERS[name] = marker
        INPUT_FILES[name] = tuple(find_data_files(data_dir, marker, (".xlsx",)))


def select_columns(row: list[object]) -> list[object]:
    return [
        row[column - 1] if column <= len(row) else None
        for column in KEPT_COLUMN_INDEXES
    ]


def _trim_trailing_none(row: list[object]) -> list[object]:
    """Drop trailing blank cells so a row's length reflects its real content.

    calamine pads every row to the sheet's overall used width, unlike the
    per-row width the previous reader gave; two source files with different
    formatted-but-empty trailing columns would otherwise make an identical
    header row compare unequal only because of length.
    """
    end = len(row)
    while end > 0 and row[end - 1] is None:
        end -= 1
    return row[:end]


def add_subsidy_column(
    row: list[object],
    *,
    profile_name: str,
    is_header: bool = False,
    source_name: str | None = None,
    source_row: int | None = None,
) -> list[object]:
    profile = PROFILES[profile_name]
    result = list(row)
    if is_header:
        result.insert(3, "补贴金额")
        return result

    amount = result[2]
    if amount in (None, ""):
        subsidy = None
    else:
        try:
            calculated = Decimal(str(amount)) * profile.subsidy_rate
            subsidy = float(
                min(calculated, profile.subsidy_cap).quantize(
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


def build_report(
    profile_name: str,
    deleted_invalid_files: list[str] | None = None,
) -> SubmittedReport:
    """Build one project's submitted report.

    deleted_invalid_files, when given, receives every no-worksheet export
    found during the scan. The caller passes its own list so the names
    survive an exception: whatever fails mid-scan, the garbage files found
    so far can still be removed and recorded.
    """
    if deleted_invalid_files is None:
        deleted_invalid_files = []
    files = list(INPUT_FILES[profile_name])
    if not files:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 中找到文件名包含"
            f"“{SUBMITTED_FILE_MARKERS[profile_name]}”的 .xlsx 文件"
        )

    status_priority = {
        status: index for index, status in enumerate(STATUS_ORDER)
    }

    expected_header: list[object] | None = None
    output_header: list[object] | None = None
    reference_column_index: int | None = None
    status_column_index: int | None = None
    reference_locations: dict[str, tuple[str, int]] = {}
    data_row_count = 0
    data_rows: list[list[object]] = []
    unknown_status_records: list[UnknownStatusRecord] = []
    valid_file_count = 0

    for path in files:
        source_workbook = CalamineWorkbook.from_path(str(path))
        if not source_workbook.sheet_names:
            source_workbook.close()
            # Not deleted here: process_submitted_files removes the files
            # after input validation and records the removal first.
            deleted_invalid_files.append(path.name)
            continue
        try:
            rows = (
                _trim_trailing_none(row)
                for row in calamine_rows(source_workbook.get_sheet_by_index(0))
            )
            title = next(rows, None)
            header = next(rows, None)
            if title is None or header is None:
                raise ValueError(f"{path.name} 缺少标题行或表头行")

            if expected_header is None:
                expected_header = header
                # Drop the source title row and use the selected headers as
                # row 1.
                output_header = add_subsidy_column(
                    select_columns(header),
                    profile_name=profile_name,
                    is_header=True,
                )
                # Reject a wrong export before parsing any rows, so the
                # operator sees the missing fields instead of a downstream
                # value error.
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
                reference_column_index = output_header.index("检索参考号")
                status_column_index = output_header.index("状态")
            elif header != expected_header:
                raise ValueError(f"{path.name} 的表头与第一个文件不一致")

            for source_row, row in enumerate(rows, start=3):
                selected_row = select_columns(row)
                if not any(value not in (None, "") for value in selected_row):
                    continue

                output_row = add_subsidy_column(
                    selected_row,
                    profile_name=profile_name,
                    source_name=path.name,
                    source_row=source_row,
                )
                if reference_column_index is None:
                    raise RuntimeError("未能定位已上传数据必要字段")
                if status_column_index is None:
                    raise RuntimeError("未能定位已上传数据必要字段")

                raw_reference = output_row[reference_column_index]
                if normalize_reference(raw_reference):
                    reference = validated_reference(
                        raw_reference,
                        f"{path.name} 第 {source_row} 行",
                    )
                    existing_location = reference_locations.get(reference)
                    if existing_location is not None:
                        existing_name, existing_row = existing_location
                        raise ValueError(
                            f"检索参考号重复：{reference}；"
                            f"首次出现在 {existing_name} 第 {existing_row} 行，"
                            f"再次出现在 {path.name} 第 {source_row} 行"
                        )
                    reference_locations[reference] = (path.name, source_row)

                status = str(output_row[status_column_index] or "")
                if status not in STATUS_ORDER:
                    unknown_status_records.append(
                        UnknownStatusRecord(
                            source_name=path.name,
                            source_row=source_row,
                            reference=str(output_row[reference_column_index] or ""),
                            status=status,
                        )
                    )

                data_rows.append(output_row)
                data_row_count += 1
            valid_file_count += 1
        finally:
            source_workbook.close()

    if output_header is None:
        raise RuntimeError("未能生成输出表头")

    status_column_index = output_header.index("状态")
    data_rows.sort(
        key=lambda row: status_priority.get(
            str(row[status_column_index]) if row[status_column_index] is not None else "",
            len(STATUS_ORDER),
        )
    )
    description_column_index = output_header.index("描述")
    rows_by_status: dict[str, list[list[object]]] = {
        status: [] for status in STATUS_ORDER
    }
    for row in data_rows:
        status = str(row[status_column_index] or "")
        if status in rows_by_status:
            rows_by_status[status].append(row)
    for status in STATUS_ORDER:
        status_rows = rows_by_status[status]
        status_rows.sort(
            key=lambda row: (
                row[description_column_index] not in (None, ""),
                str(row[description_column_index] or ""),
            ),
            reverse=True,
        )

    return SubmittedReport(
        header=tuple(output_header),
        summary_rows=data_rows,
        status_rows=rows_by_status,
        file_count=valid_file_count,
        data_row_count=data_row_count,
        unknown_status_records=tuple(unknown_status_records),
        deleted_invalid_files=tuple(deleted_invalid_files),
    )


def write_workbook(path: Path, report: SubmittedReport) -> None:
    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    with Workbook(
        str(path),
        {
            "constant_memory": True,
            "strings_to_urls": False,
            # 描述 and 商品名称 are free text copied from the source export; a
            # value starting with "=" is data, not a formula to evaluate.
            "strings_to_formulas": False,
        },
    ) as workbook:
        write_formatted_sheet(
            workbook,
            "Summary",
            report.header,
            report.summary_rows,
            font_name,
            measurement_font,
        )
        for status in STATUS_ORDER:
            write_formatted_sheet(
                workbook,
                status,
                report.header,
                report.status_rows[status],
                font_name,
                measurement_font,
            )


def validate_output(path: Path, expected_data_rows: int, profile_name: str) -> None:
    """Re-read the just-written workbook and check its shape and every row's
    subsidy calculation. Values only (no font/fill checks) — calamine reads
    the same file several times faster than openpyxl here."""
    profile = PROFILES[profile_name]
    workbook = CalamineWorkbook.from_path(str(path))
    try:
        expected_sheet_names = ["Summary", *STATUS_ORDER]
        if workbook.sheet_names != expected_sheet_names:
            raise RuntimeError(
                f"工作表校验失败：预期 {expected_sheet_names}，"
                f"实际 {workbook.sheet_names}"
            )

        sheet_rows = {
            name: list(calamine_rows(workbook.get_sheet_by_name(name)))
            for name in expected_sheet_names
        }
        summary_rows = sheet_rows["Summary"]
        actual_data_rows = max(len(summary_rows) - 1, 0)
        if actual_data_rows != expected_data_rows:
            raise RuntimeError(
                f"输出校验失败：预期 {expected_data_rows} 条，实际 {actual_data_rows} 条"
            )
        header = tuple(summary_rows[0]) if summary_rows else ()
        expected_columns = len(KEPT_SOURCE_COLUMNS) + 1
        if len(header) != expected_columns:
            raise RuntimeError(
                f"输出校验失败：预期 {expected_columns} 列，"
                f"实际 {len(header)} 列"
            )

        status_column = header.index("状态")
        description_column = header.index("描述")
        amount_column = header.index("交易金额")
        subsidy_column = header.index("补贴金额")
        reference_column = header.index("检索参考号")

        seen_references: dict[str, int] = {}
        for row_number, row in enumerate(summary_rows[1:], start=2):
            if normalize_reference(row[reference_column]):
                reference = validated_reference(
                    row[reference_column],
                    f"Summary 第 {row_number} 行",
                    error_type=RuntimeError,
                )
                first_row = seen_references.get(reference)
                if first_row is not None:
                    raise RuntimeError(
                        f"Summary 第 {row_number} 行检索参考号重复：{reference}；"
                        f"首次出现在第 {first_row} 行"
                    )
                seen_references[reference] = row_number

            amount = row[amount_column]
            subsidy = row[subsidy_column]
            if amount in (None, ""):
                if subsidy not in (None, ""):
                    raise RuntimeError(
                        f"Summary 第 {row_number} 行交易金额为空但补贴金额不为空"
                    )
                continue
            expected_subsidy = min(
                Decimal(str(amount)) * profile.subsidy_rate,
                profile.subsidy_cap,
            ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            try:
                actual_subsidy = Decimal(str(subsidy))
            except InvalidOperation as error:
                raise RuntimeError(
                    f"Summary 第 {row_number} 行补贴金额无效：{subsidy!r}"
                ) from error
            if actual_subsidy != expected_subsidy:
                raise RuntimeError(
                    f"Summary 第 {row_number} 行补贴金额计算错误"
                )

        status_total = sum(
            max(len(sheet_rows[status]) - 1, 0) for status in STATUS_ORDER
        )
        known_status_total = sum(
            1
            for row in summary_rows[1:]
            if row[status_column] in STATUS_ORDER
        )
        if status_total != known_status_total:
            raise RuntimeError(
                f"状态工作表校验失败：预期共 {known_status_total} 条，"
                f"实际共 {status_total} 条"
            )

        # The status sheets are written by separate calls from the Summary
        # one, so matching counts do not prove matching content.
        # Comparing them as multisets checks every column of every row —
        # 补贴金额 included — against the Summary rows already recomputed
        # above, which is cheaper than recomputing the subsidy per sheet.
        # Counter rather than set: one 检索参考号 may legitimately carry
        # several identical detail rows, and a status sheet that dropped one
        # of them must not pass.
        expected_status_rows = Counter(
            tuple(row)
            for row in summary_rows[1:]
            if row[status_column] in STATUS_ORDER
        )
        actual_status_rows = Counter(
            tuple(row)
            for status in STATUS_ORDER
            for row in sheet_rows[status][1:]
        )
        if actual_status_rows != expected_status_rows:
            raise RuntimeError(
                "状态工作表校验失败：数据行内容与 Summary 不一致"
            )

        for status in STATUS_ORDER:
            status_rows = sheet_rows[status]
            status_header = tuple(status_rows[0]) if status_rows else ()
            if status_header != header:
                raise RuntimeError(f"{status}工作表的标题行与汇总表不一致")

            descriptions: list[str] = []
            blank_description_found = False
            for row in status_rows[1:]:
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

            if descriptions != sorted(descriptions, reverse=True):
                raise RuntimeError(f"{status}工作表的描述列未按降序排列")
    finally:
        workbook.close()


def _remove_invalid_exports(
    file_names: tuple[str, ...],
    reporter: ConsoleReporter,
) -> None:
    """Remove garbage exports (no worksheets) and record the removal.

    Only files actually removed are recorded, and the record is registered
    in a finally so it survives a mid-loop failure — a file whose deletion
    failed must never appear in the 已删除 list.
    """
    removed: list[str] = []
    try:
        for file_name in file_names:
            (DATA_DIR / file_name).unlink(missing_ok=True)
            removed.append(file_name)
    finally:
        if removed:
            reporter.corrected(
                "已删除无效导出文件（没有工作表）",
                tuple(f"文件：{file_name}" for file_name in removed),
            )


def process_submitted_files(
    profile_name: str,
    reporter: ConsoleReporter,
) -> None:
    invalid_files: list[str] = []
    try:
        report = build_report(profile_name, invalid_files)
    except Exception:
        # Whatever failed, any no-worksheet files found before it are still
        # garbage: remove and record them before the step fails.
        _remove_invalid_exports(tuple(invalid_files), reporter)
        raise
    _remove_invalid_exports(report.deleted_invalid_files, reporter)
    output_file = PROFILES[profile_name].output_file
    write_xlsx_atomically(
        output_file,
        lambda path: write_workbook(path, report),
        lambda path: validate_output(path, report.data_row_count, profile_name),
    )

    metric_parts = [
        f"{format_count(report.file_count)} 个文件",
        f"{format_count(report.data_row_count)} 行",
    ]
    # Non-zero fixed statuses only, in STATUS_ORDER, so the line stays short
    # and the operator can see backlog composition without opening the file.
    for status in STATUS_ORDER:
        count = len(report.status_rows[status])
        if count:
            metric_parts.append(f"{status} {format_count(count)}")
    if report.unknown_status_records:
        status_names = "、".join(
            sorted({record.status or "(空)" for record in report.unknown_status_records})
        )
        metric_parts.append(
            f"{status_names} {format_count(len(report.unknown_status_records))} 行"
            "（保留在 Summary）"
        )
        record_lines = tuple(
            f"源文件 {record.source_name}，源行 {record.source_row}，"
            f"检索参考号 {record.reference or '(空)'}"
            for record in report.unknown_status_records
        )
        reporter.detail(
            f"{profile_name}未知状态数据："
            f"{format_count(len(report.unknown_status_records))} 行",
            record_lines,
        )
    reporter.metric(profile_name, "｜".join(metric_parts))
    reporter.output(output_file)


def process_all(reporter: ConsoleReporter) -> None:
    """Process both projects' submitted data as one all-or-nothing unit.

    An operator has no reason to run one project without the other, so the
    two outputs are rolled back together if either one fails.
    """
    def process_both() -> None:
        for name in PROFILE_ORDER:
            process_submitted_files(name, reporter)

    run_with_output_rollback(OUTPUT_FILES, process_both)
