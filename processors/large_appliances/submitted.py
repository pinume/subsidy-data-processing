from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from processors.common import submitted as common_submitted
from processors.common.excel import save_workbook_atomically
from processors.common.submitted import (
    KEPT_SOURCE_COLUMNS,
    KEPT_COLUMN_INDEXES,
    REQUIRED_SUBMITTED_HEADERS,
    STATUS_ORDER,
    STATUS_PRIORITY,
    select_columns,
)

from . import _shared
from ._shared import OUTPUT_FILE


# Household appliances: 15% of the transaction, capped at 1500 per order.
# Digital uses the same rate but a 500 cap — see processors/digital.py.
SUBSIDY_RATE = Decimal("0.15")
SUBSIDY_CAP = Decimal("1500")


def add_subsidy_column(
    row: list[object],
    *,
    is_header: bool = False,
    source_name: str | None = None,
    source_row: int | None = None,
) -> list[object]:
    return common_submitted.add_subsidy_column(
        row,
        subsidy_rate=SUBSIDY_RATE,
        subsidy_cap=SUBSIDY_CAP,
        is_header=is_header,
        source_name=source_name,
        source_row=source_row,
    )


def build_workbook() -> tuple[Workbook, int, int]:
    files = list(_shared.INPUT_FILES)
    return common_submitted.build_workbook(
        files,
        data_dir=_shared.DATA_DIR,
        file_marker=_shared.SUBMITTED_FILE_MARKER,
        subsidy_rate=SUBSIDY_RATE,
        subsidy_cap=SUBSIDY_CAP,
    )


def validate_output(path: Path, expected_data_rows: int) -> None:
    common_submitted.validate_output(
        path,
        expected_data_rows,
        subsidy_rate=SUBSIDY_RATE,
        subsidy_cap=SUBSIDY_CAP,
    )


def process_submitted_files() -> None:
    workbook, file_count, data_row_count = build_workbook()
    save_workbook_atomically(
        workbook,
        OUTPUT_FILE,
        lambda path: validate_output(path, data_row_count),
    )

    print(f"Submitted data complete: merged {file_count} files, {data_row_count} rows")
    print(f"Output file: {OUTPUT_FILE}")
