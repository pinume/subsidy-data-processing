"""数码 side of 销售用券情况统计 processing.

数码 has no group sheets, reference-supplement file, or approved-detail
panel, so its summary is a simpler three-column (上传状态, 数量, 合计) table
instead of 家电's five-column one — see processors/coupons/appliance.py's
module docstring for why this stays a separate module rather than a
feature-flagged variant of that one.
"""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation

from processors.common.dates import (
    normalize_document_number,
    normalize_receipt_identifier,
)
from processors.receipts import OUTPUT_FILE as RECEIPTS_OUTPUT_FILE
from processors.submitted import PROFILES as SUBMITTED_PROFILES

from . import matching, sources
from .matching import as_currency
from .sources import load_coupon_remark_lookup, load_uploaded_summary
from .validation import (
    validate_detail_rows_shape,
    validate_document_and_date_values,
    validate_matched_subsidy_total,
    validate_payment_statuses,
    validate_remark_and_detail_values,
    validate_uploaded_and_unmatched_counts,
)

COUPON_SUBSIDY_HEADER = sources.COUPON_DIGITAL_SUBSIDY_HEADER
COUPON_REMARK_SOURCE_FILE = RECEIPTS_OUTPUT_FILE
COUPON_UPLOADED_SOURCE_FILE = SUBMITTED_PROFILES["数码"].output_file
COUPON_OUTPUT_HEADER = sources.DIGITAL_PROFILE.output_header
DETAILS_SHEET_NAME = "数码-明细总表"


def build_coupon_summary(
    rows: list[list[object]],
    uploaded_count: int,
    uploaded_subsidy_total: Decimal,
) -> list[tuple[object, ...]]:
    subsidy_index = matching.SUBSIDY_INDEX
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
                f"第 {row_number} 行的{COUPON_SUBSIDY_HEADER}金额无效：{subsidy!r}"
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
    in processors/coupons/appliance.py."""

    rows: list[list[object]]
    data_row_count: int
    matched_count: int
    matched_subsidy_total: Decimal
    remark_lookup: dict[tuple[str, date], str]
    detail_lookup: dict[str, str]
    reference_universe: set[str]
    payment_references: frozenset[str]
    payment_match_count: int
    uploaded_subsidy_count: int
    uploaded_subsidy_total: Decimal
    uploaded_match_count: int
    unmatched_count: int
    corrected_count: int
    unresolved_count: int
    correction_collision_count: int
    reference_decisions: list[matching.ReferenceDecision]
    summary_rows: list[tuple[object, ...]]


def compute_coupon_data(
    *,
    rows: list[list[object]] | None = None,
    remark_lookup: dict[tuple[str, date], str] | None = None,
    payment_reference_locations: dict[str, str] | None = None,
) -> CouponComputation:
    if sources.COUPON_SOURCE_FILE is None:
        raise FileNotFoundError(
            f"未在 {sources.DATA_DIR} 中找到文件名包含"
            f"“{sources.COUPON_STATISTICS_KEYWORD}”且表头为"
            f"“{COUPON_SUBSIDY_HEADER}”的 .XLSX 文件"
        )

    if remark_lookup is None:
        remark_lookup = load_coupon_remark_lookup(COUPON_REMARK_SOURCE_FILE)
    if rows is None:
        export = sources.read_coupon_export(
            sources.COUPON_SOURCE_FILE,
            remark_lookup=remark_lookup,
        )
        rows = export.digital_rows
    matched_count, matched_subsidy_total = matching.fill_coupon_remarks(
        rows, remark_lookup, COUPON_SUBSIDY_HEADER
    )
    detail_lookup, uploaded_subsidy_count, uploaded_subsidy_total = (
        load_uploaded_summary(COUPON_UPLOADED_SOURCE_FILE)
    )
    # Unsubmitted data is no longer supplied, so submitted data is the only
    # source of valid references.
    reference_universe = set(detail_lookup)
    payment_reference_locations = payment_reference_locations or {}
    sources.validate_payment_reference_subset(
        "数码",
        payment_reference_locations,
        reference_universe,
    )
    payment_references = frozenset(payment_reference_locations)
    (
        corrected_count,
        unresolved_count,
        correction_collision_count,
        reference_decisions,
    # matched_count is passed to every pass below for the same reason 家电
    # passes it: those trailing rows are the 退换货 block, already settled by
    # the receipt remark and pinned pink. Leaving it off let the reference
    # passes overwrite their remarks with 已上传/未上传 and count them into
    # the upload tallies, so the sheet ended up with pink rows labelled
    # 已上传.
    ) = matching.correct_coupon_references(
        rows, reference_universe, matched_count
    )
    uploaded_match_count, unmatched_count = matching.fill_upload_statuses(
        rows,
        detail_lookup,
        matched_count,
    )
    payment_match_count = matching.fill_payment_statuses(
        rows,
        payment_references,
        matched_count,
    )
    summary_rows = build_coupon_summary(
        rows,
        uploaded_subsidy_count,
        uploaded_subsidy_total,
    )
    if as_currency(matched_subsidy_total) != Decimal("0.00"):
        raise ValueError(
            f"备注匹配行的{COUPON_SUBSIDY_HEADER}合计不为 0："
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
        payment_references=payment_references,
        payment_match_count=payment_match_count,
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


def validate_computation(computation: CouponComputation) -> None:
    """Validate business invariants before serializing the workbook."""
    expected_matched_rows = computation.matched_count
    remark_lookup = computation.remark_lookup
    expected_matched_subsidy_total = computation.matched_subsidy_total
    detail_lookup = computation.detail_lookup
    reference_universe = computation.reference_universe
    payment_references = computation.payment_references
    expected_paid_rows = computation.payment_match_count

    rows = computation.rows
    validate_detail_rows_shape(
        rows, COUPON_OUTPUT_HEADER, computation.data_row_count
    )
    validate_payment_statuses(
        rows,
        COUPON_OUTPUT_HEADER,
        payment_references,
        expected_matched_rows,
        expected_paid_rows,
    )
    detail_column, remark_column = validate_uploaded_and_unmatched_counts(
        rows,
        COUPON_OUTPUT_HEADER,
        computation.uploaded_match_count,
        computation.unmatched_count,
    )
    summary_column = COUPON_OUTPUT_HEADER.index("明细摘要")

    matched_start = len(rows) - expected_matched_rows
    actual_matched_subsidy_total = Decimal("0")
    subsidy_column = COUPON_OUTPUT_HEADER.index(COUPON_SUBSIDY_HEADER)
    for row_number, row in enumerate(rows[1:], start=2):
        document = row[0]
        document_date = validate_document_and_date_values(
            document, row[1], row_number
        )
        receipt_remark = remark_lookup.get(
            (
                normalize_document_number(document),
                document_date,
            ),
            "",
        )
        reference = normalize_receipt_identifier(
            row[summary_column]
        ).upper()
        in_matched_partition = (
            expected_matched_rows > 0
            and row_number - 1 >= matched_start
        )
        if in_matched_partition:
            # The 退换货 block keeps the receipt remark and never gets a
            # 详细情况: the reference passes skip it, so expecting an upload
            # status here would be checking for something nothing writes.
            expected_detail = ""
            expected_remark = receipt_remark
        else:
            expected_detail = detail_lookup.get(reference, "")
            if expected_detail:
                expected_remark = "已上传"
            elif reference not in reference_universe:
                expected_remark = "未上传"
            else:
                expected_remark = receipt_remark
        validate_remark_and_detail_values(
            row[remark_column],
            row[detail_column],
            expected_remark,
            expected_detail,
            row_number,
        )
        if in_matched_partition:
            subsidy = row[subsidy_column]
            if subsidy not in (None, ""):
                actual_matched_subsidy_total += Decimal(str(subsidy))

    validate_matched_subsidy_total(
        actual_matched_subsidy_total, expected_matched_subsidy_total
    )
    if as_currency(actual_matched_subsidy_total) != Decimal("0.00"):
        raise RuntimeError(
            f"销售用券匹配行的{COUPON_SUBSIDY_HEADER}合计不为 0："
            f"{actual_matched_subsidy_total}"
        )
