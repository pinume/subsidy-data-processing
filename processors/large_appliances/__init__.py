"""Household appliance data processing: submitted data, receipts, coupons."""

from pathlib import Path

from . import _shared
from ._shared import (
    BASE_DIR,
    OUTPUT_DIR,
    OUTPUT_FILE,
    RECEIPTS_OUTPUT_FILE,
    COUPON_REMARK_SOURCE_FILE,
    COUPON_UPLOADED_SOURCE_FILE,
    SUBMITTED_FILE_MARKER,
)
from .submitted import (
    KEPT_SOURCE_COLUMNS,
    KEPT_COLUMN_INDEXES,
    REQUIRED_SUBMITTED_HEADERS,
    SUBSIDY_RATE,
    SUBSIDY_CAP,
    STATUS_ORDER,
    STATUS_PRIORITY,
    select_columns,
    add_subsidy_column,
    build_workbook,
    validate_output,
    process_submitted_files,
)
from .receipts import (
    RECEIPTS_SOURCE_HEADER,
    RECEIPTS_OUTPUT_HEADER,
    RECEIPTS_REMARK_RETURN,
    RECEIPTS_REMARK_ORIGINAL,
    RECEIPTS_REMARK_BOTH,
    RECEIPTS_REMARK_SAME_MODEL_REPLACEMENT,
    RECEIPTS_REMARK_SPECIAL,
    RECEIPTS_SPECIAL_REMARK_KEYS,
    RECEIPTS_ROW_HEIGHT,
    RECEIPTS_DUPLICATE_FILL_COLOR,
    RECEIPTS_EXCLUDED_PRODUCT_KEYWORD,
    RECEIPTS_SAME_MODEL_REPLACEMENT_KEYWORD,
    read_receipt_rows,
    receipt_remark,
    prepare_receipt_data,
    validate_receipts_output,
    process_receipts,
)
from .coupons import (
    COUPON_KEPT_SOURCE_COLUMNS,
    COUPON_SUBSIDY_HEADER,
    COUPON_EXCLUDED_CATEGORY,
    COUPON_OUTPUT_HEADER,
    COUPON_GROUP_HEADER,
    COUPON_GROUP_COLUMN_INDEXES,
    REFERENCE_REPORT_CORRECTED,
    REFERENCE_REPORT_UNRESOLVED,
    REFERENCE_REPORT_COLLISION,
    REFERENCE_REPORT_ORDER,
    COUPON_MATCH_FILL_COLOR,
    COUPON_BRAND_REPLACEMENTS,
    COUPON_SUMMARY_HEADER,
    COUPON_SUMMARY_PROJECT_LABEL,
    COUPON_REMARK_SORT_PRIORITY,
    COUPON_REFERENCE_SUPPLEMENT_HEADER,
    as_currency,
    read_coupon_rows,
    load_coupon_remark_lookup,
    fill_coupon_remarks,
    load_coupon_reference_supplement,
    coupon_data_rows,
    fill_coupon_reference_supplement,
    load_uploaded_detail_lookup,
    fill_uploaded_details,
    reference_correction_candidates,
    correct_coupon_references,
    fill_unmatched_remarks,
    coupon_text_sort_value,
    coupon_date_sort_value,
    coupon_pink_sort_key,
    coupon_regular_sort_key,
    coupon_group_regular_sort_key,
    select_coupon_group_columns,
    sort_coupon_detail_rows,
    build_coupon_summary,
    coupon_group_sheet_title,
    build_coupon_group_sheets,
    merge_coupon_summary_groups,
    project_summary_blocks,
    merge_coupon_summary_projects,
    merged_coupon_summary_values,
    apply_coupon_summary_borders,
    DETAILS_SHEET_NAME,
    SUMMARY_SHEET_NAME,
    CouponComputation,
    coupon_subsidy_count,
    compute_coupon_data,
    build_summary_and_details_sheets,
    build_group_sheets,
    validate_summary_and_details_sheets,
    validate_group_sheets,
)


DATA_DIR: Path
INPUT_FILES: tuple[Path, ...]
RECEIPTS_SOURCE_FILE: Path | None
COUPON_SOURCE_FILE: Path | None
COUPON_REFERENCE_SUPPLEMENT_FILE: Path


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global INPUT_FILES
    global RECEIPTS_SOURCE_FILE
    global COUPON_SOURCE_FILE
    global COUPON_REFERENCE_SUPPLEMENT_FILE

    _shared.configure_data_dir(data_dir)
    DATA_DIR = _shared.DATA_DIR
    INPUT_FILES = _shared.INPUT_FILES
    RECEIPTS_SOURCE_FILE = _shared.RECEIPTS_SOURCE_FILE
    COUPON_SOURCE_FILE = _shared.COUPON_SOURCE_FILE
    COUPON_REFERENCE_SUPPLEMENT_FILE = _shared.COUPON_REFERENCE_SUPPLEMENT_FILE
