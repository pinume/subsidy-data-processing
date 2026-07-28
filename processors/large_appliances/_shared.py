from functools import partial
from pathlib import Path

from processors.common.config import submitted_file_marker
from processors.common.paths import (
    find_data_files,
    match_source_file_by_header,
    read_xls_header,
    resolve_unique_file,
)


BASE_DIR = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = BASE_DIR / "output"
DATA_DIR: Path
INPUT_FILES: tuple[Path, ...]
OUTPUT_FILE = OUTPUT_DIR / "家电_已上传.xlsx"
RECEIPTS_SOURCE_FILE: Path | None
# Shared with digital: both projects' receipt statistics come from the same
# source file and are processed with this project's rules (see receipts.py).
RECEIPTS_OUTPUT_FILE = OUTPUT_DIR / "收款单统计.xlsx"
COUPON_SOURCE_FILE: Path | None
COUPON_REFERENCE_SUPPLEMENT_FILE: Path
COUPON_REMARK_SOURCE_FILE = RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = OUTPUT_FILE

DATA_TYPE = "家电"
# Files live directly in the flat data directory; each project tells its own
# files apart by filename keyword. The submitted files for the two projects
# are genuinely separate files, told apart that way; the coupon export is a
# single merged file shared by both projects (家电 and 数码 rows sit in the
# same sheet, one column of the 国补 pair populated per row — see
# COUPON_DIGITAL_SUBSIDY_COLUMN below and read_coupon_rows in coupons.py /
# processors/digital.py), so both projects resolve to the same file, matched
# by each project's own header column.
# The submitted marker is derived from config/merchants.yaml in
# configure_data_dir rather than at import time, so a missing or malformed
# config fails the run with a readable error instead of breaking the import.
SUBMITTED_FILE_MARKER: str
RECEIPT_STATISTICS_KEYWORD = "收款单统计"
COUPON_STATISTICS_KEYWORD = "销售用券情况统计"
COUPON_REFERENCE_SUPPLEMENT_KEYWORD = "新建 Microsoft Excel 工作表"
# The coupon export's field header row (row 2) at its last kept column
# (column 26); see COUPON_KEPT_SOURCE_COLUMNS in coupons.py.
COUPON_SUBSIDY_HEADER = "2026家电国补（计入收入）"
# digital's 国补 column in the same merged file (column 27, right after
# 家电's above); processors/digital.py uses it as its own header/kept
# column, and processors/large_appliances/coupons.py reads it (without
# keeping it) to exclude 数码 rows from 家电 processing.
COUPON_DIGITAL_SUBSIDY_COLUMN = 27


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global INPUT_FILES
    global RECEIPTS_SOURCE_FILE
    global COUPON_SOURCE_FILE
    global COUPON_REFERENCE_SUPPLEMENT_FILE
    global SUBMITTED_FILE_MARKER

    DATA_DIR = data_dir
    SUBMITTED_FILE_MARKER = submitted_file_marker(DATA_TYPE)
    INPUT_FILES = tuple(
        find_data_files(data_dir, SUBMITTED_FILE_MARKER, (".xlsx",))
    )
    RECEIPTS_SOURCE_FILE = resolve_unique_file(
        find_data_files(data_dir, RECEIPT_STATISTICS_KEYWORD, (".xls",))
    )
    COUPON_SOURCE_FILE = match_source_file_by_header(
        find_data_files(data_dir, COUPON_STATISTICS_KEYWORD, (".xls",)),
        COUPON_SUBSIDY_HEADER,
        read_header=partial(read_xls_header, row=2, column=26),
    )
    COUPON_REFERENCE_SUPPLEMENT_FILE = resolve_unique_file(
        find_data_files(
            data_dir, COUPON_REFERENCE_SUPPLEMENT_KEYWORD, (".xlsx",)
        )
    ) or data_dir / f"{COUPON_REFERENCE_SUPPLEMENT_KEYWORD}.xlsx"
