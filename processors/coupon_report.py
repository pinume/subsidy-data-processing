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


def merged_reference_decisions(
    appliance_computation: appliance.CouponComputation,
    digital_computation: digital.CouponComputation,
) -> list[tuple[object, ...]]:
    """Label each project's decisions and interleave them into one report.

    Digital's decisions used to be computed and then thrown away, leaving its
    automatic corrections with no audit trail at all.
    """
    decisions = [
        ("家电", *decision)
        for decision in appliance_computation.reference_decisions
    ] + [
        ("数码", *decision)
        for decision in digital_computation.reference_decisions
    ]
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
        f"[{label}] 警告：源文件合计行为 {source_total:,.2f}，"
        f"但其明细行合计为 {computed_total:,.2f}，"
        f"相差 {computed_total - source_total:,.2f}。"
        "报表采用明细行合计以与明细总表保持一致；"
        "该差额通常是导出快照期间产生的退货尚未写入明细，"
        "下次导出补齐后此提示会自动消失"
    )


def format_subsidy_correction_warning(
    correction: sources.SubsidyCorrection,
    source_name: str,
) -> str:
    """Render one subsidy attribution correction as the operator-facing warning.

    Wording matches what read_coupon_export used to print while reading; the
    reader no longer prints, so a correction is reported exactly once no
    matter how many times the export is read, and tests can assert on the
    formatting directly instead of mocking print. A future Processing Report
    entry consumes the same records.
    """
    return (
        f"警告：{source_name} 第 {correction.row_number} 行单据 "
        f"{correction.document_number} 的财务大类为"
        f"{correction.financial_category!r}，"
        f"已将 {correction.amount} 从“{correction.from_header}”"
        f"调整到“{correction.to_header}”"
    )


def digital_extra_summary_rows(
    digital_computation: digital.CouponComputation,
) -> list[tuple[object, ...]]:
    """Recast digital's 3-column summary rows (上传状态, 数量, 合计) as rows in
    家电's 5-column 数据汇总 table, labeling every row (including digital's
    own "合计" row) with 财务大类="数码" and no 品牌, so the block mirrors the
    家电 one that precedes it (see appliance.COUPON_SUMMARY_PROJECT_LABEL)."""
    return [
        (DIGITAL_SUMMARY_PROJECT_LABEL, None, remark, count, total)
        for remark, count, total in digital_computation.summary_rows
    ]


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
        *appliance_computation.summary_rows,
        *extra_summary_rows,
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
            project_merges=appliance.coupon_summary_project_merges(
                project_blocks
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

    payment_reference_locations = sources.load_payment_reference_locations(
        PAYMENT_FILE
    )
    remark_lookup = load_coupon_remark_lookup(appliance.COUPON_REMARK_SOURCE_FILE)
    source_workbook = CalamineWorkbook.from_path(str(coupon_source))
    try:
        export = sources.read_coupon_export(
            coupon_source,
            source_workbook,
            remark_lookup,
        )
    finally:
        source_workbook.close()

    for correction in export.subsidy_corrections:
        print(format_subsidy_correction_warning(correction, coupon_source.name))

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
    print(f"[家电] 销售用券统计完成：{la.data_row_count} 行")
    print(f"[家电] 退换货（粉色）行数：{la.matched_count}")
    print(f"[家电] 补充参考号匹配：{la.reference_supplement_count}")
    print(
        "[家电] 补充参考号候选不唯一："
        f"{la.ambiguous_reference_supplement_count}"
    )
    print(f"[家电] 已上传状态匹配：{la.uploaded_count}")
    print(f"[家电] 已回款匹配：{la.payment_match_count}")
    print(
        "[家电] 未在已上传数据中找到（标记未上传）："
        f"{la.unmatched_count}"
    )
    print(
        f"[家电] 参考号自动纠正：{la.corrected_count}；"
        f"无唯一候选：{la.unresolved_count}；"
        f"目标冲突：{la.correction_collision_count}"
    )
    print(
        "[家电] 匹配行计入收入的2026家电国补合计："
        f"{la.matched_subsidy_total:.2f}"
    )
    if la.zero_subsidy_count:
        print(
            f"[家电] 警告：{la.zero_subsidy_count} 行的 "
            "2026家电国补（计入收入）为 0，不应出现；"
            "按 0 计入，请检查这些源数据行"
        )
    report_source_total_gap("家电", la.source_total, la.computed_total)

    dg = digital_computation
    print(f"[数码] 销售用券统计完成：{dg.data_row_count} 行")
    print(f"[数码] 备注匹配：{dg.matched_count}")
    print(f"[数码] 已上传状态匹配：{dg.uploaded_match_count}")
    print(f"[数码] 已回款匹配：{dg.payment_match_count}")
    print(
        "[数码] 汇总采用的已上传补贴行："
        f"{dg.uploaded_subsidy_count}；合计：{dg.uploaded_subsidy_total:.2f}"
    )
    print(
        "[数码] 未在已上传数据中找到（标记未上传）："
        f"{dg.unmatched_count}"
    )
    print(
        f"[数码] 参考号自动纠正：{dg.corrected_count}；"
        f"无唯一候选：{dg.unresolved_count}；"
        f"目标冲突：{dg.correction_collision_count}"
    )
    print(
        "[数码] 匹配行计入收入的2026数码国补合计："
        f"{dg.matched_subsidy_total:.2f}"
    )
    print(f"输出文件：{OUTPUT_FILE}")


def _comparable_output_value(value: object) -> object:
    """Normalize harmless XLSX round trips before comparing written values."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, Decimal):
        return round(float(value), 9)
    if isinstance(value, float):
        return round(value, 9)
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
        comparable_actual = tuple(
            _comparable_output_value(value) for value in actual
        )
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
        *appliance_computation.summary_rows,
        *extra_summary_rows,
    ]
    expected_rows = {
        appliance.SUMMARY_SHEET_NAME: [
            appliance.COUPON_SUMMARY_HEADER,
            *appliance.merged_coupon_summary_values(combined_summary_rows),
        ],
        appliance.DETAILS_SHEET_NAME: [
            tuple(row) for row in appliance_computation.rows
        ],
        digital.DETAILS_SHEET_NAME: [
            tuple(row) for row in digital_computation.rows
        ],
    }
    for sheet_name, _, _, grouped_rows in appliance_computation.group_sheets:
        expected_rows[sheet_name] = [
            appliance.COUPON_GROUP_HEADER,
            *(
                appliance.select_coupon_group_columns(row)
                for row, _ in grouped_rows
            ),
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
            appliance.coupon_summary_group_merges(
                combined_summary_rows, brand_rows_end
            )
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
