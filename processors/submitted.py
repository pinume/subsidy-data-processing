"""已上传数据 (submitted export) processing for household appliances and digital.

Both projects read the same source column layout and share the pipeline in
processors/common/submitted.py; only the subsidy cap, source marker, and
output file differ, so they are captured here as two SubmittedProfile
instances rather than two near-identical modules.
"""

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from processors.common import submitted as common_submitted
from processors.common.config import submitted_file_marker
from processors.common.excel import run_with_output_rollback
from processors.common.paths import find_data_files
from processors.common.submitted import (
    KEPT_SOURCE_COLUMNS,
    KEPT_COLUMN_INDEXES,
    REQUIRED_SUBMITTED_HEADERS,
    STATUS_ORDER,
)


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"


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
    return common_submitted.select_columns(row, KEPT_COLUMN_INDEXES)


def _config(profile_name: str) -> common_submitted.SubmittedConfig:
    profile = PROFILES[profile_name]
    return common_submitted.SubmittedConfig(
        input_files=INPUT_FILES[profile_name],
        data_dir=DATA_DIR,
        source_marker=SUBMITTED_FILE_MARKERS[profile_name],
        output_file=profile.output_file,
        subsidy_rate=profile.subsidy_rate,
        subsidy_cap=profile.subsidy_cap,
        kept_columns=KEPT_SOURCE_COLUMNS,
        required_headers=REQUIRED_SUBMITTED_HEADERS,
        status_order=STATUS_ORDER,
    )


def add_subsidy_column(
    row: list[object],
    *,
    profile_name: str,
    is_header: bool = False,
    source_name: str | None = None,
    source_row: int | None = None,
) -> list[object]:
    profile = PROFILES[profile_name]
    return common_submitted.add_subsidy_column(
        row,
        subsidy_rate=profile.subsidy_rate,
        subsidy_cap=profile.subsidy_cap,
        is_header=is_header,
        source_name=source_name,
        source_row=source_row,
    )


def build_workbook(profile_name: str) -> tuple[Workbook, int, int]:
    return common_submitted.build_workbook(_config(profile_name))


def validate_output(path: Path, expected_data_rows: int, profile_name: str) -> None:
    common_submitted.validate_output(path, expected_data_rows, _config(profile_name))


def process_submitted_files(profile_name: str) -> None:
    common_submitted.process_submitted_files(_config(profile_name))


def process_all() -> None:
    """Process both projects' submitted data as one all-or-nothing unit.

    An operator has no reason to run one project without the other, so the
    two outputs are rolled back together if either one fails.
    """
    def process_both() -> None:
        for name in PROFILE_ORDER:
            process_submitted_files(name)

    run_with_output_rollback(OUTPUT_FILES, process_both)
