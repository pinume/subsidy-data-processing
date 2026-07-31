"""已上传数据 (submitted export) processing for household appliances and digital.

Both projects read the same source column layout, apply the same 15%
subsidy rate, and validate the same shape of output; only the subsidy cap,
source marker, and output file differ, so they are captured as two
SubmittedProfile instances sharing one pipeline rather than two
near-identical modules.
"""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string

from processors.common.config import submitted_file_marker
from processors.common.excel import (
    format_sheet,
    load_measurement_font,
    read_rows,
    resolve_font,
    run_with_output_rollback,
    save_workbook_atomically,
)
from processors.common.paths import find_data_files


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"

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


@dataclass(frozen=True)
class SubmittedProfile:
    data_type: str
    output_file: Path
    subsidy_rate: Decimal
    subsidy_cap: Decimal


# Household appliances and digital both take 15% of the transaction; the two
# caps were once both written as 500, which silently understated 43% of the
# appliance rows, so they stay pinned separately rather than sharing a
# default.
PROFILES: dict[str, SubmittedProfile] = {
    "家电": SubmittedProfile(
        data_type="家电",
        output_file=OUTPUT_DIR / "家电_已上传.xlsx",
        subsidy_rate=Decimal("0.15"),
        subsidy_cap=Decimal("1500"),
    ),
    "数码": SubmittedProfile(
        data_type="数码",
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


def build_workbook(profile_name: str) -> tuple[Workbook, int, int]:
    profile = PROFILES[profile_name]
    files = list(INPUT_FILES[profile_name])
    if not files:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 中找到文件名包含"
            f"“{SUBMITTED_FILE_MARKERS[profile_name]}”的 .xlsx 文件"
        )

    status_priority = {
        status: index for index, status in enumerate(STATUS_ORDER)
    }

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
                profile_name=profile_name,
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
                        profile_name=profile_name,
                        source_name=path.name,
                        source_row=source_row,
                    )
                )
                data_row_count += 1

    if output_header is None:
        raise RuntimeError("未能生成输出表头")

    status_column_index = output_header.index("状态")
    data_rows.sort(
        key=lambda row: status_priority.get(
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


def validate_output(path: Path, expected_data_rows: int, profile_name: str) -> None:
    profile = PROFILES[profile_name]
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
                        Decimal(str(amount)) * profile.subsidy_rate,
                        profile.subsidy_cap,
                    ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    if Decimal(str(subsidy)) != expected_subsidy:
                        raise RuntimeError(
                            f"{status}工作表存在补贴金额计算错误"
                        )

            if descriptions != sorted(descriptions, reverse=True):
                raise RuntimeError(f"{status}工作表的描述列未按降序排列")
    finally:
        workbook.close()


def process_submitted_files(profile_name: str) -> None:
    workbook, file_count, data_row_count = build_workbook(profile_name)
    output_file = PROFILES[profile_name].output_file
    save_workbook_atomically(
        workbook,
        output_file,
        lambda path: validate_output(path, data_row_count, profile_name),
    )

    print(f"Submitted data complete: merged {file_count} files, {data_row_count} rows")
    print(f"Output file: {output_file}")


def process_all() -> None:
    """Process both projects' submitted data as one all-or-nothing unit.

    An operator has no reason to run one project without the other, so the
    two outputs are rolled back together if either one fails.
    """
    def process_both() -> None:
        for name in PROFILE_ORDER:
            process_submitted_files(name)

    run_with_output_rollback(OUTPUT_FILES, process_both)
