"""Merges the 家电 and 数码 coupon pipelines into one 审核明细.xlsx workbook.

Kept as its own module (rather than living inside either project) because
each project's package needs to refer to the other's output file or menu
entry, and importing across processors.coupons.digital <->
processors.coupons.appliance at module load time would be circular. This
module imports both at the top level; main.py only ever reaches it through a
function-local import inside build_processors(), which runs long after both
packages have finished loading.

Re-exports SUMMARY_SHEET_NAME/SUMMARY_HEADER from
processors.coupons.report_contract so store_report.py can depend on this
module's public surface instead of reaching into
processors.coupons.appliance's internals.
"""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from python_calamine import CalamineWorkbook
from xlsxwriter import Workbook as XlsxWorkbook

from processors.common.config import load_payment_brand_config
from processors.common.excel import (
    calamine_rows,
    load_measurement_font,
    resolve_font,
    write_xlsx_atomically,
)
from processors.coupons import appliance, digital, matching, sources, xlsx_output
from processors.coupons.report_contract import SUMMARY_HEADER, SUMMARY_SHEET_NAME
from processors.coupons.sources import load_coupon_remark_lookup
from processors.payment import OUTPUT_FILE as PAYMENT_FILE
from processors.payment import SUMMARY_HEADERS as PAYMENT_SUMMARY_HEADERS
from processors.payment import SUMMARY_SHEET_NAME as PAYMENT_SUMMARY_SHEET_NAME

# Not unused despite the lack of a local reference: re-exported for
# store_report.py, which imports SUMMARY_SHEET_NAME/SUMMARY_HEADER from this
# module rather than reaching into processors.coupons.report_contract itself.
# Listing them in __all__ (rather than an "as"-aliased import) tells pyflakes
# the same thing without tripping pylint's redundant-alias rule.
__all__ = ["SUMMARY_HEADER", "SUMMARY_SHEET_NAME"]


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_FILE = BASE_DIR / "output" / "审核明细.xlsx"

# Defined in appliance.py because its group-sheet titles must avoid colliding
# with it; this module owns what actually goes into the sheet.
REFERENCE_REPORT_SHEET = appliance.REFERENCE_REPORT_SHEET_NAME
REFERENCE_REPORT_HEADER = (
    "项目",
    "处理结果",
    "单据号",
    "单据日期",
    "原参考号",
    "说明",
)
# Corrections first: they are the rows that were actually rewritten and most
# need a human to confirm. Conflicts next, then the leftovers.
REFERENCE_REPORT_ORDER = matching.REFERENCE_REPORT_ORDER
REPORT_PROJECT_ORDER = ("家电", "数码")
# 财务大类 for digital's 已上传/未上传/合计 block at the foot of 数据汇总.
DIGITAL_SUMMARY_PROJECT_LABEL = "数码"

# Maps coupon-side 财务大类 values to payment-side 财务大类 values for
# matching against 回款明细. Coupon categories come from the source export,
# payment categories from config/payment_brands.yaml — they use different
# coding systems (e.g. "国产彩电" vs "电视").
COUPON_TO_PAYMENT_CATEGORY: dict[str, str] = {
    "国产彩电": "电视",
}

# Maps coupon-side brand names to payment-side brand names.
COUPON_TO_PAYMENT_BRAND: dict[str, str] = {
    "AO史密斯": "A.O.史密斯",
}

# (财务大类, 品牌) pairs where the payment data uses "美的系" instead of
# the individual brand name (see config/payment_brands.yaml midea_group).
_PAYMENT_BRAND_CONFIG = load_payment_brand_config()
MIDEA_GROUP_MAP: dict[tuple[str, str], str] = {
    (category, brand): "美的系"
    for category in _PAYMENT_BRAND_CONFIG.midea_group_categories
    for brand in _PAYMENT_BRAND_CONFIG.midea_group_brands
}

# Payment categories that belong to each project, for aggregating
# project-level summary rows (where 品牌 is None).
HOUSEHOLD_PAYMENT_CATEGORIES = frozenset({"冰箱", "洗衣机", "电视", "空调", "厨卫"})
DIGITAL_PAYMENT_CATEGORIES = frozenset({"手机", "平板", "智能穿戴"})


def load_payment_summary(
    path: Path,
) -> dict[tuple[str, str], tuple[Decimal, int]]:
    """Read 回款明细.xlsx 汇总 and return {(财务大类, 品牌): (补贴金额合计, 补贴金额计数)}."""
    if not path.exists():
        raise FileNotFoundError(f"未找到回款明细文件：{path}")

    workbook = CalamineWorkbook.from_path(str(path))
    try:
        if PAYMENT_SUMMARY_SHEET_NAME not in workbook.sheet_names:
            raise ValueError(
                f"回款明细缺少 {PAYMENT_SUMMARY_SHEET_NAME!r} 工作表"
            )
        rows = list(
            calamine_rows(workbook.get_sheet_by_name(PAYMENT_SUMMARY_SHEET_NAME))
        )
        if not rows:
            raise ValueError("回款明细汇总工作表为空")

        actual_header = tuple(rows[0][: len(PAYMENT_SUMMARY_HEADERS)])
        if actual_header != tuple(PAYMENT_SUMMARY_HEADERS):
            raise ValueError(
                f"回款明细汇总表头不符合预期：{actual_header}"
            )

        result: dict[tuple[str, str], tuple[Decimal, int]] = {}
        current_category = ""
        for row in rows[1:]:
            category_raw = row[0]
            if category_raw not in (None, ""):
                current_category = str(category_raw).strip()
            category = current_category
            brand = str(row[1] or "").strip()
            amount_raw = row[2]
            count_raw = row[3]
            if not category or category == "合计":
                continue
            if not brand:
                continue
            if amount_raw in (None, ""):
                amount = Decimal("0")
            else:
                amount = Decimal(str(amount_raw))
            if count_raw in (None, ""):
                count = 0
            else:
                count = int(Decimal(str(count_raw)).to_integral_value())
            key = (category, brand)
            result[key] = (amount, count)
        return result
    finally:
        workbook.close()


def _resolve_payment_data(
    coupon_category: str,
    coupon_brand: str | None,
    payment_summary: dict[tuple[str, str], tuple[Decimal, int]],
    project_label: str,
) -> tuple[Decimal, int]:
    """Look up (回款金额, 回款数量) for a single row."""
    if coupon_brand is not None:
        # An empty brand is a real brand-level row with no identifiable brand;
        # only None means that the caller is resolving a project total.
        if not coupon_brand:
            return (Decimal("0"), 0)
        key = (coupon_category, coupon_brand)
        if key in payment_summary:
            return payment_summary[key]

        mapped_category = COUPON_TO_PAYMENT_CATEGORY.get(
            coupon_category, coupon_category
        )
        if mapped_category != coupon_category:
            mapped_key = (mapped_category, coupon_brand)
            if mapped_key in payment_summary:
                return payment_summary[mapped_key]

        mapped_brand = COUPON_TO_PAYMENT_BRAND.get(coupon_brand, coupon_brand)
        if mapped_brand != coupon_brand:
            mapped_key = (coupon_category, mapped_brand)
            if mapped_key in payment_summary:
                return payment_summary[mapped_key]
            if mapped_category != coupon_category:
                mapped_key = (mapped_category, mapped_brand)
                if mapped_key in payment_summary:
                    return payment_summary[mapped_key]

        midea_brand = MIDEA_GROUP_MAP.get((coupon_category, coupon_brand))
        if midea_brand is not None:
            midea_key = (coupon_category, midea_brand)
            if midea_key in payment_summary:
                return payment_summary[midea_key]
            if mapped_category != coupon_category:
                midea_key = (mapped_category, midea_brand)
                if midea_key in payment_summary:
                    return payment_summary[midea_key]

        return (Decimal("0"), 0)

    # Project-level row: aggregate all payment data for this project
    if project_label == "家电":
        entries = [
            v
            for (cat, _brand), v in payment_summary.items()
            if cat in HOUSEHOLD_PAYMENT_CATEGORIES
        ]
    elif project_label == "数码":
        entries = [
            v
            for (cat, _brand), v in payment_summary.items()
            if cat in DIGITAL_PAYMENT_CATEGORIES
        ]
    else:
        return (Decimal("0"), 0)

    total_amount = sum((a for a, _ in entries), Decimal("0"))
    total_count = sum(c for _, c in entries)
    return (total_amount, total_count)


def enrich_summary_rows_with_payment(
    summary_rows: list[tuple[object, ...]],
    payment_summary: dict[tuple[str, str], tuple[Decimal, int]],
    project_label: str,
) -> list[tuple[object, ...]]:
    """Append 回款金额 and 回款数量 to every summary row.

    Brand rows and project 合计 rows get payment data; project 已上传/未上传
    rows leave both columns empty.  Brands with payment data but no 未上传
    rows are added with zero counts so their 回款 is still visible.
    """
    if not payment_summary:
        return [(*row, None, None) for row in summary_rows]

    # Collect existing brand keys (only brand-level rows, not project rows).
    # Project rows have status (已上传/未上传/合计) in column 1; brand rows
    # have a brand name there.
    PROJECT_STATUSES = frozenset({"已上传", "未上传", "合计"})
    existing_brand_keys: set[tuple[str, str]] = set()
    brand_rows: list[tuple[object, ...]] = []
    project_rows: list[tuple[object, ...]] = []
    for row in summary_rows:
        col1 = str(row[1] or "").strip()
        if col1 in PROJECT_STATUSES:
            project_rows.append(row)
        else:
            category = str(row[0] or "").strip()
            existing_brand_keys.add((category, col1))
            brand_rows.append(row)

    # Find payment brands missing from the brand rows.
    # Only add missing brands when brand rows already exist (家电);
    # digital has no brand rows, and we don't want to invent them.
    # Only consider categories that already appear in the coupon summary
    # (avoids pulling 数码 brands into 家电, and vice versa).
    extra_brands: list[tuple[str, str]] = []
    if brand_rows:
        coupon_categories = {str(r[0] or "").strip() for r in brand_rows}
        # Also include the mapped payment equivalents of coupon categories
        eligible_payment_categories = set(coupon_categories)
        for coupon_cat in coupon_categories:
            mapped = COUPON_TO_PAYMENT_CATEGORY.get(coupon_cat)
            if mapped:
                eligible_payment_categories.add(mapped)

        # Collect coupon brands that are already represented (including
        # mapped equivalents, to avoid duplicates like 美的 + 美的系).
        represented_brands: set[tuple[str, str]] = set()
        for coupon_cat, coupon_brand in existing_brand_keys:
            represented_brands.add((coupon_cat, coupon_brand))
            # Also mark the payment-side equivalent of this category
            mapped_cat = COUPON_TO_PAYMENT_CATEGORY.get(coupon_cat, coupon_cat)
            if mapped_cat != coupon_cat:
                represented_brands.add((mapped_cat, coupon_brand))
            # If this brand maps to something else in payment, mark that too
            mapped_brand = COUPON_TO_PAYMENT_BRAND.get(coupon_brand)
            if mapped_brand:
                represented_brands.add((coupon_cat, mapped_brand))
                if mapped_cat != coupon_cat:
                    represented_brands.add((mapped_cat, mapped_brand))
            midea = MIDEA_GROUP_MAP.get((coupon_cat, coupon_brand))
            if midea:
                represented_brands.add((coupon_cat, midea))
                if mapped_cat != coupon_cat:
                    represented_brands.add((mapped_cat, midea))

        for (cat, brand), _ in sorted(payment_summary.items()):
            if cat not in eligible_payment_categories:
                continue
            if (cat, brand) in represented_brands:
                continue
            extra_brands.append((cat, brand))

    # Enrich brand rows
    enriched: list[tuple[object, ...]] = []
    for row in brand_rows:
        category = str(row[0] or "").strip()
        brand_str = str(row[1] or "").strip()
        amount, count = _resolve_payment_data(
            category, brand_str, payment_summary, project_label
        )
        display_amount: float | None = float(amount) if amount else None
        display_count: int | None = count if count else None
        enriched.append((*row, display_amount, display_count))

    # Add extra brand rows (no 未上传 data, but have 回款)
    for cat, brand in extra_brands:
        pay_amount, pay_count = payment_summary[(cat, brand)]
        display_amount: float | None = float(pay_amount) if pay_amount else None
        display_count: int | None = pay_count if pay_count else None
        enriched.append((cat, brand, 0, 0.0, display_amount, display_count))

    # Re-sort brand rows: by (category, brand)
    enriched.sort(key=lambda r: (str(r[0] or ""), str(r[1] or "")))

    # Append project rows (已上传/未上传 leave payment empty, 合计 gets totals)
    for row in project_rows:
        category = str(row[0] or "").strip()
        # Status is in column 1 for project rows (former 品牌 column)
        status = str(row[1] or "").strip()
        if status in ("已上传", "未上传"):
            enriched.append((*row, None, None))
        else:
            amount, count = _resolve_payment_data(
                category, None, payment_summary, project_label
            )
            display_amount: float | None = float(amount) if amount else None
            display_count: int | None = count if count else None
            enriched.append((*row, display_amount, display_count))

    return enriched


def merged_reference_decisions(
    appliance_computation: appliance.CouponComputation,
    digital_computation: digital.CouponComputation,
) -> list[tuple[object, ...]]:
    """Label each project's decisions and interleave them into one report.

    Digital's decisions used to be computed and then thrown away, leaving its
    automatic corrections with no audit trail at all.
    """
    decisions = [
        ("家电", *decision) for decision in appliance_computation.reference_decisions
    ] + [("数码", *decision) for decision in digital_computation.reference_decisions]
    decisions.sort(
        key=lambda decision: (
            REFERENCE_REPORT_ORDER[decision[1]],
            REPORT_PROJECT_ORDER.index(decision[0]),
        )
    )
    return decisions


def report_source_total_gap(
    label: str,
    source_total: Decimal | None,
    computed_total: Decimal,
) -> None:
    """Warn when a coupon export's own 合计 row disagrees with its detail rows.

    Both numbers are reported rather than reconciled: the export is a snapshot,
    and a return recorded after the detail rows were written shows up here as a
    gap that closes by itself once the next export includes it. Silently
    trusting either number would hide that.
    """
    if source_total is None or source_total == computed_total:
        return
    print(
        f"[{label}] WARNING: 源文件合计行为 {source_total:,.2f}，"
        f"但其明细行合计为 {computed_total:,.2f}，"
        f"相差 {computed_total - source_total:,.2f}。"
        "报表采用明细行合计以与明细总表保持一致；"
        "该差额通常是导出快照期间产生的退货尚未写入明细，"
        "下次导出补齐后此提示会自动消失"
    )


def digital_extra_summary_rows(
    digital_computation: digital.CouponComputation,
) -> list[tuple[object, ...]]:
    """Recast digital's 3-column summary rows (备注, 数量, 合计) as rows in
    家电's 5-column 数据汇总 table, labeling every row (including digital's
    own "合计" row) with 财务大类="数码" and no 品牌, so the block mirrors the
    家电 one that precedes it (see appliance.COUPON_SUMMARY_PROJECT_LABEL)."""
    return [
        (f"{DIGITAL_SUMMARY_PROJECT_LABEL}合计", remark, count, total)
        for remark, count, total in digital_computation.summary_rows
    ]


def _pad_row_to_width(
    row: tuple[object, ...], width: int, pad_value: object = None
) -> tuple[object, ...]:
    """Ensure a row tuple has exactly `width` elements, padding if needed."""
    if len(row) >= width:
        return row[:width]
    return (*row, *(pad_value for _ in range(width - len(row))))


def write_coupon_workbook(
    path: Path,
    appliance_computation: appliance.CouponComputation,
    digital_computation: digital.CouponComputation,
    extra_summary_rows: list[tuple[object, ...]],
    decisions: list[tuple[object, ...]],
) -> None:
    """Write all 30 sheets in workbook order: 数据汇总, the two 明细总表, the
    group sheets, then the Processing Report."""
    combined_summary_rows = [
        _pad_row_to_width(row, len(appliance.COUPON_SUMMARY_HEADER))
        for row in (
            *appliance_computation.summary_rows,
            *extra_summary_rows,
        )
    ]
    project_blocks = appliance.project_summary_blocks(combined_summary_rows)
    brand_rows_end = (
        project_blocks[0][0] if project_blocks else len(combined_summary_rows)
    )
    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)

    with XlsxWorkbook(
        str(path),
        {
            # No constant_memory: 数据汇总 merges after its rows are written.
            "strings_to_urls": False,
            # 商品名称 and 备注 are free text; a leading "=" is data.
            "strings_to_formulas": False,
        },
    ) as workbook:
        formats = xlsx_output.CouponFormatCache(
            workbook, font_name, appliance.COUPON_MATCH_FILL_COLOR
        )
        xlsx_output.write_summary_sheet(
            workbook,
            appliance.SUMMARY_SHEET_NAME,
            appliance.COUPON_SUMMARY_HEADER,
            combined_summary_rows,
            formats,
            measurement_font,
            group_merges=(
                appliance.coupon_summary_group_merges(
                    combined_summary_rows, brand_rows_end
                )
                if brand_rows_end
                else []
            ),
            project_merges=appliance.coupon_summary_project_merges(project_blocks),
            currency_columns=(
                len(appliance.COUPON_SUMMARY_HEADER) - 3,
                len(appliance.COUPON_SUMMARY_HEADER) - 2,
            ),
        )
        # computation.rows already opens with the header row; the openpyxl
        # writer appended the whole list, header included.
        xlsx_output.write_detail_sheet(
            workbook,
            appliance.DETAILS_SHEET_NAME,
            appliance.COUPON_OUTPUT_HEADER,
            appliance_computation.rows[1:],
            formats,
            measurement_font,
            left_aligned_headers=("商品名称", "详细情况"),
            matched_count=appliance_computation.matched_count,
        )
        xlsx_output.write_detail_sheet(
            workbook,
            digital.DETAILS_SHEET_NAME,
            digital.COUPON_OUTPUT_HEADER,
            digital_computation.rows[1:],
            formats,
            measurement_font,
            left_aligned_headers=("商品名称",),
            matched_count=digital_computation.matched_count,
        )
        for sheet_name, _, _, grouped_rows in appliance_computation.group_sheets:
            xlsx_output.write_group_sheet(
                workbook,
                sheet_name,
                appliance.COUPON_GROUP_HEADER,
                grouped_rows,
                appliance.select_coupon_group_columns,
                formats,
                measurement_font,
                left_aligned_headers=("商品名称", "详细情况"),
            )
        xlsx_output.write_reference_report(
            workbook,
            REFERENCE_REPORT_SHEET,
            REFERENCE_REPORT_HEADER,
            decisions,
            formats,
            measurement_font,
        )


def process_coupon_sales() -> None:
    coupon_source = sources.COUPON_SOURCE_FILE
    if coupon_source is None:
        # compute_coupon_data() raises the same FileNotFoundError once it
        # sees sources.COUPON_SOURCE_FILE is None; raising it here directly
        # names the file operators are missing without pretending a second
        # (unreachable) code path could still succeed after the first raises.
        raise FileNotFoundError(
            f"未在 {sources.DATA_DIR} 中找到文件名包含"
            f"“{sources.COUPON_STATISTICS_KEYWORD}”且表头符合"
            "家电、数码用券导出格式的 .XLSX 文件"
        )

    payment_reference_locations = sources.load_payment_reference_locations(PAYMENT_FILE)
    source_workbook = CalamineWorkbook.from_path(str(coupon_source))
    try:
        export = sources.read_coupon_export(coupon_source, source_workbook)
    finally:
        source_workbook.close()

    remark_lookup = load_coupon_remark_lookup(appliance.COUPON_REMARK_SOURCE_FILE)
    appliance_computation = appliance.compute_coupon_data(
        rows=export.appliance_rows,
        remark_lookup=remark_lookup,
        payment_reference_locations=payment_reference_locations["家电"],
        source_total=export.source_total,
    )
    digital_computation = digital.compute_coupon_data(
        rows=export.digital_rows,
        remark_lookup=remark_lookup,
        payment_reference_locations=payment_reference_locations["数码"],
    )
    extra_summary_rows = digital_extra_summary_rows(digital_computation)

    decisions = merged_reference_decisions(
        appliance_computation,
        digital_computation,
    )
    appliance.validate_computation(
        appliance_computation,
        extra_summary_rows,
    )
    digital.validate_computation(digital_computation)

    # Enrich the four-column summary rows with payment amount and count.
    # Validation above deliberately runs against the pre-enrichment layout.
    payment_summary = load_payment_summary(PAYMENT_FILE)
    appliance_computation.summary_rows = enrich_summary_rows_with_payment(
        appliance_computation.summary_rows, payment_summary, "家电"
    )
    extra_summary_rows = enrich_summary_rows_with_payment(
        extra_summary_rows, payment_summary, "数码"
    )

    write_xlsx_atomically(
        OUTPUT_FILE,
        lambda path: write_coupon_workbook(
            path,
            appliance_computation,
            digital_computation,
            extra_summary_rows,
            decisions,
        ),
        lambda path: validate_merged_coupon_output(
            path,
            appliance_computation,
            digital_computation,
            extra_summary_rows,
            decisions,
        ),
    )

    la = appliance_computation
    print(f"[家电] Subsidy coupon statistics complete: {la.data_row_count} rows")
    print(f"[家电] Receipt remark matches: {la.receipt_remark_count}")
    print(f"[家电] Pink return or exchange rows: {la.matched_count}")
    print(f"[家电] Supplemental reference matches: {la.reference_supplement_count}")
    print(
        "[家电] Ambiguous supplemental reference candidates: "
        f"{la.ambiguous_reference_supplement_count}"
    )
    print(f"[家电] Submitted status matches: {la.uploaded_count}")
    print(f"[家电] Payment status matches: {la.payment_match_count}")
    print(
        f"[家电] Rows not found in submitted data (marked 未上传): {la.unmatched_count}"
    )
    print(
        f"[家电] Automatic reference corrections: {la.corrected_count}; "
        f"no unique candidate: {la.unresolved_count}; "
        f"duplicate conflicts: {la.correction_collision_count}"
    )
    print(
        "[家电] Total 2026 appliance subsidy counted as revenue for "
        f"matched rows: {la.matched_subsidy_total:.2f}"
    )
    if la.zero_subsidy_count:
        print(
            f"[家电] WARNING: {la.zero_subsidy_count} rows have a zero "
            "2026家电国补（计入收入）, which should not occur; they are "
            "counted as 0 — check those source rows"
        )
    report_source_total_gap("家电", la.source_total, la.computed_total)

    dg = digital_computation
    print(f"[数码] Subsidy coupon statistics complete: {dg.data_row_count} rows")
    print(f"[数码] Remark matches: {dg.matched_count}")
    print(f"[数码] Submitted status matches: {dg.uploaded_match_count}")
    print(f"[数码] Payment status matches: {dg.payment_match_count}")
    print(
        "[数码] Uploaded subsidy rows used in Summary: "
        f"{dg.uploaded_subsidy_count}; total: {dg.uploaded_subsidy_total:.2f}"
    )
    print(
        f"[数码] Rows not found in submitted data (marked 未上传): {dg.unmatched_count}"
    )
    print(
        f"[数码] Automatic reference corrections: {dg.corrected_count}; "
        f"no unique candidate: {dg.unresolved_count}; "
        f"duplicate conflicts: {dg.correction_collision_count}"
    )
    print(
        "[数码] Total 2026 digital subsidy counted as revenue for "
        f"matched rows: {dg.matched_subsidy_total:.2f}"
    )
    print(f"Output file: {OUTPUT_FILE}")


def _comparable_output_value(value: object) -> object:
    """Normalize harmless XLSX round trips before comparing written values."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.000000001"))
    if isinstance(value, float):
        return Decimal(str(round(value, 9)))
    return value


def _validate_sheet_rows(
    sheet_name: str,
    actual_rows: list[tuple[object, ...]],
    expected_rows,
) -> None:
    expected_rows = [tuple(row) for row in expected_rows]
    if len(actual_rows) != len(expected_rows):
        raise RuntimeError(
            f"{sheet_name}行数校验失败：预期 {len(expected_rows)} 行，"
            f"实际 {len(actual_rows)} 行"
        )
    for row_number, (actual, expected) in enumerate(
        zip(actual_rows, expected_rows), start=1
    ):
        comparable_actual = tuple(_comparable_output_value(value) for value in actual)
        comparable_expected = tuple(
            _comparable_output_value(value) for value in expected
        )
        if comparable_actual != comparable_expected:
            raise RuntimeError(
                f"{sheet_name}第 {row_number} 行输出校验失败："
                f"实际 {comparable_actual!r}，预期 {comparable_expected!r}"
            )


def _expected_coupon_output(
    appliance_computation: appliance.CouponComputation,
    digital_computation: digital.CouponComputation,
    extra_summary_rows: list[tuple[object, ...]],
    decisions: list[tuple[object, ...]],
) -> tuple[
    dict[str, list[tuple[object, ...]]],
    dict[str, set[tuple[tuple[int, int], tuple[int, int]]]],
]:
    combined_summary_rows = [
        _pad_row_to_width(row, len(appliance.COUPON_SUMMARY_HEADER))
        for row in (
            *appliance_computation.summary_rows,
            *extra_summary_rows,
        )
    ]
    expected_rows = {
        appliance.SUMMARY_SHEET_NAME: [
            appliance.COUPON_SUMMARY_HEADER,
            *appliance.merged_coupon_summary_values(combined_summary_rows),
        ],
        appliance.DETAILS_SHEET_NAME: [
            tuple(row) for row in appliance_computation.rows
        ],
        digital.DETAILS_SHEET_NAME: [tuple(row) for row in digital_computation.rows],
    }
    for sheet_name, _, _, grouped_rows in appliance_computation.group_sheets:
        expected_rows[sheet_name] = [
            appliance.COUPON_GROUP_HEADER,
            *(appliance.select_coupon_group_columns(row) for row, _ in grouped_rows),
        ]
    expected_rows[REFERENCE_REPORT_SHEET] = [
        REFERENCE_REPORT_HEADER,
        *decisions,
    ]

    project_blocks = appliance.project_summary_blocks(combined_summary_rows)
    brand_rows_end = (
        project_blocks[0][0] if project_blocks else len(combined_summary_rows)
    )
    summary_merges = {
        ((first - 1, column - 1), (last - 1, column - 1))
        for first, last, column in (
            appliance.coupon_summary_group_merges(combined_summary_rows, brand_rows_end)
            if brand_rows_end
            else []
        )
    }
    summary_merges.update(
        ((first - 1, first_column - 1), (last - 1, last_column - 1))
        for first, last, first_column, last_column in (
            appliance.coupon_summary_project_merges(project_blocks)
        )
    )
    expected_merges = {name: set() for name in expected_rows}
    expected_merges[appliance.SUMMARY_SHEET_NAME] = summary_merges
    return expected_rows, expected_merges


def validate_merged_coupon_output(
    path: Path,
    appliance_computation: appliance.CouponComputation,
    digital_computation: digital.CouponComputation,
    extra_summary_rows: list[tuple[object, ...]],
    decisions: list[tuple[object, ...]],
) -> None:
    expected_rows, expected_merges = _expected_coupon_output(
        appliance_computation,
        digital_computation,
        extra_summary_rows,
        decisions,
    )
    workbook = CalamineWorkbook.from_path(str(path))
    try:
        expected_sheet_names = list(expected_rows)
        if list(workbook.sheet_names) != expected_sheet_names:
            raise RuntimeError(
                "审核明细工作表校验失败："
                f"预期 {expected_sheet_names}，实际 {workbook.sheet_names}"
            )
        for sheet_name in expected_sheet_names:
            sheet = workbook.get_sheet_by_name(sheet_name)
            actual_rows = [tuple(row) for row in calamine_rows(sheet)]
            _validate_sheet_rows(
                sheet_name,
                actual_rows,
                expected_rows[sheet_name],
            )
            actual_merges = set(sheet.merged_cell_ranges)
            if actual_merges != expected_merges[sheet_name]:
                raise RuntimeError(
                    f"{sheet_name}合并范围校验失败："
                    f"实际 {sorted(actual_merges)!r}，"
                    f"预期 {sorted(expected_merges[sheet_name])!r}"
                )
    finally:
        workbook.close()
