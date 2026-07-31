"""Locate, read, and classify rows from the merged 销售用券情况统计 export.

家电 and 数码 rows sit in the same sheet, one column of a 国补 pair (columns
26 and 27) populated per row; a single CouponSourceProfile per project
carries the small amount of real per-project variation (which column is
"own" vs "the other side", the output header text, whether brand names get
normalized through config/brand_mapping.yaml) so this module reads the file
exactly once and hands each project its own classified rows, rather than
each project separately locating and re-parsing the same file (the previous
digital.py / large_appliances/_shared.py each ran their own
find_data_files + header match against the same keyword).
"""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

from openpyxl import load_workbook

from processors.common.config import load_brand_mapping
from processors.common.coupons import classify_coupon_row
from processors.common.dates import normalize_coupon_date
from processors.common.paths import find_data_files, resolve_unique_file


BASE_DIR = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = BASE_DIR / "output"

COUPON_STATISTICS_KEYWORD = "销售用券情况统计"
# The coupon export's field header row (row 2) at these columns.
COUPON_FAMILY_SUBSIDY_COLUMN = 26
COUPON_DIGITAL_SUBSIDY_COLUMN = 27
COUPON_FAMILY_SUBSIDY_HEADER = "2026家电国补（计入收入）"
COUPON_DIGITAL_SUBSIDY_HEADER = "2026数码国补（计入收入）"
COUPON_KEPT_SOURCE_COLUMNS_PREFIX = (3, 4, 6, 8, 15, 18)
COUPON_BRAND_REPLACEMENTS = load_brand_mapping()

DATA_DIR: Path
COUPON_SOURCE_FILE: Path | None
COUPON_REFERENCE_SUPPLEMENT_FILE: Path
COUPON_REFERENCE_SUPPLEMENT_KEYWORD = "新建 Microsoft Excel 工作表"


@dataclass(frozen=True)
class CouponSourceProfile:
    name: str
    subsidy_header: str
    other_subsidy_column: int
    normalize_brand: bool

    @property
    def kept_source_columns(self) -> tuple[int, ...]:
        column = (
            COUPON_FAMILY_SUBSIDY_COLUMN
            if self.name == "家电"
            else COUPON_DIGITAL_SUBSIDY_COLUMN
        )
        return (*COUPON_KEPT_SOURCE_COLUMNS_PREFIX, column)

    @property
    def output_header(self) -> tuple[str, ...]:
        return (
            "单据号",
            "单据日期",
            "商品名称",
            "品牌",
            "财务大类",
            "明细摘要",
            self.subsidy_header,
            "备注",
            "详细情况",
        )


APPLIANCE_PROFILE = CouponSourceProfile(
    name="家电",
    subsidy_header=COUPON_FAMILY_SUBSIDY_HEADER,
    other_subsidy_column=COUPON_DIGITAL_SUBSIDY_COLUMN,
    normalize_brand=True,
)
DIGITAL_PROFILE = CouponSourceProfile(
    name="数码",
    subsidy_header=COUPON_DIGITAL_SUBSIDY_HEADER,
    other_subsidy_column=COUPON_FAMILY_SUBSIDY_COLUMN,
    normalize_brand=False,
)


def _read_header_row(path: Path) -> tuple[object, ...]:
    """Read row 2 (the field header row) via one sequential pass.

    Random sheet.cell(row, column) access on a read_only worksheet re-parses
    the sheet XML from the start on every call — fine for a handful of
    lookups, but O(n) per call adds up to O(n^2) once a caller does that once
    per data row on a 10000+ row export. Every reader in this module goes
    through iter_rows() instead, exactly once per full pass.
    """
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        rows_iter = sheet.iter_rows(values_only=True)
        next(rows_iter, None)  # title row
        return next(rows_iter, None) or ()
    finally:
        workbook.close()


def _header_cell(header_row: tuple[object, ...], column: int) -> str:
    if column - 1 >= len(header_row):
        return ""
    value = header_row[column - 1]
    return str(value).strip() if value is not None else ""


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global COUPON_SOURCE_FILE
    global COUPON_REFERENCE_SUPPLEMENT_FILE

    DATA_DIR = data_dir
    candidates = find_data_files(data_dir, COUPON_STATISTICS_KEYWORD, (".xlsx",))
    matches = []
    for candidate in candidates:
        header_row = _read_header_row(candidate)
        if (
            _header_cell(header_row, COUPON_FAMILY_SUBSIDY_COLUMN)
            == COUPON_FAMILY_SUBSIDY_HEADER
            and _header_cell(header_row, COUPON_DIGITAL_SUBSIDY_COLUMN)
            == COUPON_DIGITAL_SUBSIDY_HEADER
        ):
            matches.append(candidate)
    if not matches:
        COUPON_SOURCE_FILE = None
    elif len(matches) > 1:
        raise ValueError(
            f"找到多个表头符合销售用券情况统计导出格式的文件，无法确定使用哪一个：{matches}"
        )
    else:
        COUPON_SOURCE_FILE = matches[0]

    COUPON_REFERENCE_SUPPLEMENT_FILE = resolve_unique_file(
        find_data_files(
            data_dir, COUPON_REFERENCE_SUPPLEMENT_KEYWORD, (".xlsx",)
        )
    ) or data_dir / f"{COUPON_REFERENCE_SUPPLEMENT_KEYWORD}.xlsx"


def read_coupon_source_total(
    source: Path,
    profile: CouponSourceProfile,
    source_workbook=None,
) -> Decimal | None:
    """Read the 国补 value the source file states in its own 合计 row.

    The export writes it as display text (e.g. "￥2,866,331.40"), and it has
    been seen to disagree with the sum of the file's own detail rows. Reading
    it lets the program report that gap instead of silently absorbing it;
    returns None when the cell cannot be parsed.
    """
    owns_workbook = source_workbook is None
    workbook = source_workbook or load_workbook(
        source, read_only=True, data_only=True
    )
    try:
        sheet = workbook.worksheets[0]
        last_row: tuple[object, ...] | None = None
        for values in sheet.iter_rows(values_only=True):
            last_row = values
        if last_row is None:
            return None
        column = max(profile.kept_source_columns)
        cell_value = last_row[column - 1] if column - 1 < len(last_row) else None
        raw = "" if cell_value is None else str(cell_value)
        cleaned = re.sub(r"[^0-9.\-]", "", raw)
        try:
            return Decimal(cleaned) if cleaned else None
        except InvalidOperation:
            return None
    finally:
        if owns_workbook:
            workbook.close()


@dataclass(frozen=True)
class CouponExport:
    appliance_rows: list[list[object]]
    digital_rows: list[list[object]]
    source_total: Decimal | None


def read_coupon_export(
    source: Path,
    source_workbook=None,
) -> CouponExport:
    """Read the merged export once: classify every row for both projects and
    extract the source's own 合计 total in the same iter_rows() pass.

    processors/coupon_report.py needs all three (家电 rows, 数码 rows, the
    source total) every time it runs, so calling read_coupon_rows twice plus
    read_coupon_source_total once — three independent full-sheet reads — is
    real waste on a 10000+ row export even though each individual read is
    already O(n). This does the same classification read_coupon_rows does,
    once per row instead of once per row per profile.
    """
    owns_workbook = source_workbook is None
    workbook = source_workbook or load_workbook(
        source, read_only=True, data_only=True
    )
    try:
        sheet = workbook.worksheets[0]
        all_rows = list(sheet.iter_rows(values_only=True))
        if len(all_rows) < 3:
            raise ValueError(f"{source.name} 缺少标题行、字段标题行或合计行")
        required_columns = max(
            COUPON_FAMILY_SUBSIDY_COLUMN, COUPON_DIGITAL_SUBSIDY_COLUMN
        )
        header_row = all_rows[1]
        if len(header_row) < required_columns:
            raise ValueError(
                f"{source.name} 列数不足：至少需要 {required_columns} 列"
            )
        last_row = all_rows[-1]
        last_row_marker = last_row[0] if last_row else None
        if str(last_row_marker or "").strip() != "合计":
            raise ValueError(f"{source.name} 最后一行不是合计行")

        for profile in (APPLIANCE_PROFILE, DIGITAL_PROFILE):
            source_header = tuple(
                header_row[column - 1] if column - 1 < len(header_row) else None
                for column in profile.kept_source_columns
            )
            expected_source_header = profile.output_header[
                : len(profile.kept_source_columns)
            ]
            if source_header != expected_source_header:
                raise ValueError(
                    f"{source.name} 保留列字段标题不符合要求："
                    f"预期为 {expected_source_header}，实际为 {source_header}。"
                    f"请检查该文件是否为正确的销售用券情况统计导出文件。"
                )

        total_column = max(APPLIANCE_PROFILE.kept_source_columns)
        total_cell_value = (
            last_row[total_column - 1] if total_column - 1 < len(last_row) else None
        )
        total_raw = "" if total_cell_value is None else str(total_cell_value)
        total_cleaned = re.sub(r"[^0-9.\-]", "", total_raw)
        try:
            source_total = Decimal(total_cleaned) if total_cleaned else None
        except InvalidOperation:
            source_total = None

        family_column = COUPON_FAMILY_SUBSIDY_COLUMN - 1
        digital_column = COUPON_DIGITAL_SUBSIDY_COLUMN - 1
        appliance_rows: list[list[object]] = [list(APPLIANCE_PROFILE.output_header)]
        digital_rows: list[list[object]] = [list(DIGITAL_PROFILE.output_header)]
        for row_number, values in enumerate(all_rows[2:-1], start=3):
            appliance_values = [
                values[column - 1] if column - 1 < len(values) else None
                for column in APPLIANCE_PROFILE.kept_source_columns
            ]
            digital_values = [
                values[column - 1] if column - 1 < len(values) else None
                for column in DIGITAL_PROFILE.kept_source_columns
            ]
            if not any(value not in (None, "") for value in appliance_values) and (
                not any(value not in (None, "") for value in digital_values)
            ):
                continue
            appliance_subsidy = (
                values[family_column] if family_column < len(values) else None
            )
            digital_subsidy = (
                values[digital_column] if digital_column < len(values) else None
            )
            classification = classify_coupon_row(
                appliance_subsidy=appliance_subsidy,
                digital_subsidy=digital_subsidy,
                row_number=row_number,
                source_name=source.name,
            )
            if classification == "家电":
                profile, row_values, target = (
                    APPLIANCE_PROFILE,
                    appliance_values,
                    appliance_rows,
                )
            else:
                profile, row_values, target = (
                    DIGITAL_PROFILE,
                    digital_values,
                    digital_rows,
                )
            document_number = (
                ""
                if row_values[0] is None
                else str(row_values[0]).replace("收款", "")
            )
            document_date = normalize_coupon_date(row_values[1], row_number)
            result_row = [document_number, document_date, *row_values[2:], None, None]
            if profile.normalize_brand:
                brand = str(result_row[3] or "").strip()
                result_row[3] = COUPON_BRAND_REPLACEMENTS.get(brand, result_row[3])
            target.append(result_row)
        return CouponExport(appliance_rows, digital_rows, source_total)
    finally:
        if owns_workbook:
            workbook.close()


def read_coupon_rows(
    source: Path,
    profile: CouponSourceProfile,
    source_workbook=None,
) -> list[list[object]]:
    owns_workbook = source_workbook is None
    workbook = source_workbook or load_workbook(
        source, read_only=True, data_only=True
    )
    try:
        sheet = workbook.worksheets[0]
        # One sequential pass, not sheet.cell(row, col) per row: random access
        # on a read_only worksheet re-parses from the start of the sheet XML
        # on every call, which is fine for a header lookup but O(n) per call
        # — and therefore O(n^2) overall — once it runs once per data row on
        # a 10000+ row export.
        all_rows = list(sheet.iter_rows(values_only=True))
        if len(all_rows) < 3:
            raise ValueError(f"{source.name} 缺少标题行、字段标题行或合计行")
        required_columns = max(
            max(profile.kept_source_columns), profile.other_subsidy_column
        )
        header_row = all_rows[1]
        if len(header_row) < required_columns:
            raise ValueError(
                f"{source.name} 列数不足：至少需要 {required_columns} 列"
            )
        last_row = all_rows[-1]
        last_row_marker = last_row[0] if last_row else None
        if str(last_row_marker or "").strip() != "合计":
            raise ValueError(f"{source.name} 最后一行不是合计行")

        source_header = tuple(
            header_row[column - 1] if column - 1 < len(header_row) else None
            for column in profile.kept_source_columns
        )
        expected_source_header = profile.output_header[
            : len(profile.kept_source_columns)
        ]
        if source_header != expected_source_header:
            raise ValueError(
                f"{source.name} 保留列字段标题不符合要求："
                f"预期为 {expected_source_header}，实际为 {source_header}。"
                f"请检查该文件是否为正确的销售用券情况统计导出文件。"
            )

        rows: list[list[object]] = [list(profile.output_header)]
        for row_number, values in enumerate(all_rows[2:-1], start=3):
            row_values = [
                values[column - 1] if column - 1 < len(values) else None
                for column in profile.kept_source_columns
            ]
            if not any(value not in (None, "") for value in row_values):
                continue
            # The merged coupon export carries both projects' rows in one
            # sheet; classify_coupon_row tells 家电 rows apart from 数码
            # ones by which 国补 column is populated, and refuses to guess
            # when both are (source data corruption).
            other_column = profile.other_subsidy_column - 1
            other_subsidy = values[other_column] if other_column < len(values) else None
            own_subsidy = row_values[-1]
            if profile.name == "家电":
                classification = classify_coupon_row(
                    appliance_subsidy=own_subsidy,
                    digital_subsidy=other_subsidy,
                    row_number=row_number,
                    source_name=source.name,
                )
            else:
                classification = classify_coupon_row(
                    appliance_subsidy=other_subsidy,
                    digital_subsidy=own_subsidy,
                    row_number=row_number,
                    source_name=source.name,
                )
            if classification != profile.name:
                continue
            document_number = (
                ""
                if row_values[0] is None
                else str(row_values[0]).replace("收款", "")
            )
            document_date = normalize_coupon_date(row_values[1], row_number)
            result_row = [document_number, document_date, *row_values[2:], None, None]
            if profile.normalize_brand:
                brand = str(result_row[3] or "").strip()
                result_row[3] = COUPON_BRAND_REPLACEMENTS.get(brand, result_row[3])
            rows.append(result_row)
        return rows
    finally:
        if owns_workbook:
            workbook.close()
