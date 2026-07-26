from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from processors.common import submitted as common_submitted
from processors.common.submitted import (
    KEPT_SOURCE_COLUMNS,
    KEPT_COLUMN_INDEXES,
    REQUIRED_SUBMITTED_HEADERS,
    STATUS_ORDER,
    STATUS_PRIORITY,
)

from . import _shared
from ._shared import OUTPUT_FILE


# Household appliances: 15% of the transaction, capped at 1500 per order.
# Digital uses the same rate but a 500 cap — see processors/digital.py.
SUBSIDY_RATE = Decimal("0.15")
SUBSIDY_CAP = Decimal("1500")


def select_columns(row: list[object]) -> list[object]:
    return common_submitted.select_columns(row, KEPT_COLUMN_INDEXES)


def _config() -> common_submitted.SubmittedConfig:
    return common_submitted.SubmittedConfig(
        input_files=_shared.INPUT_FILES,
        data_dir=_shared.DATA_DIR,
        source_marker=_shared.SUBMITTED_FILE_MARKER,
        output_file=OUTPUT_FILE,
        subsidy_rate=SUBSIDY_RATE,
        subsidy_cap=SUBSIDY_CAP,
        kept_columns=KEPT_SOURCE_COLUMNS,
        required_headers=REQUIRED_SUBMITTED_HEADERS,
        status_order=STATUS_ORDER,
    )


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
    return common_submitted.build_workbook(_config())


def validate_output(path: Path, expected_data_rows: int) -> None:
    common_submitted.validate_output(path, expected_data_rows, _config())


def process_submitted_files() -> None:
    common_submitted.process_submitted_files(_config())
