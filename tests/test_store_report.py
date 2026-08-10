from __future__ import annotations

import io
import unittest
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font

from processors import store_report
from processors.common.excel import run_with_output_rollback

FONT = Font(name="微软雅黑", size=12)


def _fill_columns(sheet) -> None:
    for column_index in range(1, store_report.EXPECTED_COLUMN_COUNT + 1):
        sheet.cell(row=1, column=column_index, value="")


def _build_minimal_template() -> Workbook:
    """A template with just enough structure to pass validate_template."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "益庄"
    for column_index in range(1, store_report.EXPECTED_COLUMN_COUNT + 1):
        sheet.cell(row=1, column=column_index, value="标题" if column_index == 1 else "")
    for coordinate, value in store_report.TEMPLATE_STRUCTURE_CELLS.items():
        sheet[coordinate] = value
    for range_ in store_report.EXPECTED_MERGED_RANGES:
        sheet.merge_cells(range_)
    for offset, number in enumerate(
        range(1, store_report.TOTAL_ROW - 3), start=4
    ):
        sheet[f"B{offset}"] = number
    sheet.row_dimensions[store_report.TEMPLATE_VERSION_ROW].hidden = True
    return workbook


def _upload_detail_row(
    *,
    document: str = "0001",
    category: str = "冰箱",
    brand: str = "海尔",
    reference: str = "12345678901N",
    subsidy: object = 100,
    remark: str = "已上传",
) -> list[object]:
    """One 家电-明细总表 row shaped like the real output (openpyxl 写回后
    calamine 读出的是 datetime，测试里直接给 date 对象)。"""
    header = store_report.coupon_appliance.COUPON_OUTPUT_HEADER
    values = {
        "单据号": document,
        "单据日期": date(2026, 1, 1),
        "商品名称": "测试商品",
        "品牌": brand,
        "财务大类": category,
        "明细摘要": reference,
        header[6]: subsidy,
        "备注": remark,
        "详细情况": "",
        "回款情况": "",
    }
    return [values.get(column) for column in header]


def _write_upload_detail_sheet(workbook, rows) -> None:
    """Append 家电-明细总表 to a fixture workbook."""
    sheet = workbook.create_sheet(store_report.coupon_appliance.DETAILS_SHEET_NAME)
    sheet.append(store_report.coupon_appliance.COUPON_OUTPUT_HEADER)
    for row in rows:
        sheet.append(row)


def _payment_detail_row(
    *,
    raw_category: str = "A02-电冰箱",
    brand: str = "海尔",
    reference: str = "12345678901N",
    subsidy: object = 100,
) -> list[object]:
    """One 回款明细「家电明细」row shaped like the real output."""
    header = (
        store_report.PAYMENT_PROFILES["家电"].detail_headers
        + store_report.PAYMENT_DERIVED_HEADERS
    )
    mapped = store_report.PAYMENT_CATEGORY_MAP.get(raw_category)
    values = {
        "拨付批次": "batch",
        "交易时间": "2026-01-01 10:00:00",
        "交易参考号": reference,
        "商户编号": "89813015722APT1",
        "编码品类": raw_category,
        "商品名称": "测试商品",
        "补贴金额": subsidy,
        "财务大类": mapped,
        "品牌": brand,
    }
    return [values.get(column) for column in header]


def _write_payment_detail_sheet(workbook, rows) -> None:
    """Append 回款明细「家电明细」 to a fixture workbook."""
    sheet = workbook.create_sheet(
        store_report.PAYMENT_PROFILES["家电"].detail_sheet_name
    )
    sheet.append(
        store_report.PAYMENT_PROFILES["家电"].detail_headers
        + store_report.PAYMENT_DERIVED_HEADERS
    )
    for row in rows:
        sheet.append(row)


class ReportRatioTests(unittest.TestCase):
    def test_to_decimal_quantizes_float_accumulation(self) -> None:
        """Excel 浮点尾差在读取时统一量化为两位小数：0.1+0.2 必须等于 0.30，
        而不是 0.30000000000000004 一路带进精确相等校验。"""
        self.assertEqual(
            store_report.to_decimal(0.1) + store_report.to_decimal(0.2),
            Decimal("0.30"),
        )
        self.assertEqual(
            store_report.to_decimal(2280155.70000000000101),
            Decimal("2280155.70"),
        )
        self.assertEqual(
            store_report.to_decimal(779458.0500000000004),
            Decimal("779458.05"),
        )
        self.assertEqual(store_report.to_decimal(None), Decimal("0"))
        self.assertEqual(store_report.to_decimal(""), Decimal("0"))

    def test_to_count_accepts_integers_only(self) -> None:
        """数量不能用金额的两位小数量化：1.004 量化后是 1.00 会漏过
        非整数检查，必须用未量化的 Decimal 原样校验。"""
        self.assertEqual(store_report.to_count(1), 1)
        self.assertEqual(store_report.to_count(1.0), 1)
        self.assertEqual(store_report.to_count("3"), 3)
        for value in (1.004, 1.005, 1.5, "1.004", "1.5"):
            with self.assertRaisesRegex(ValueError, "整数"):
                store_report.to_count(value)
        # NaN 与 Infinity 不是有限数，必须拒绝而不是转成整数。
        for value in (float("nan"), float("inf"), float("-inf")):
            with self.assertRaisesRegex(ValueError, "整数"):
                store_report.to_count(value)
        # 非数值文本（Decimal 抛 InvalidOperation）统一为数量错误，不泄漏底层异常。
        for value in ("abc", "1.2.3", "1,000"):
            with self.assertRaisesRegex(ValueError, "整数"):
                store_report.to_count(value)
        self.assertEqual(store_report.to_count(None), 0)
        self.assertEqual(store_report.to_count(""), 0)

    def test_current_timestamp_uses_shanghai_timezone(self) -> None:
        timestamp = store_report.current_timestamp()

        self.assertEqual(timestamp.tzinfo.key, "Asia/Shanghai")
        self.assertEqual(timestamp.utcoffset(), timedelta(hours=8))

    def test_upload_data_reads_embedded_digital_summary(self) -> None:
        workbook = Workbook()
        household_sheet = workbook.active
        household_sheet.title = store_report.UPLOAD_SHEET_NAME
        household_sheet.append(store_report.UPLOAD_HEADER)
        # 数据汇总的品牌行是明细聚合的校验基准（守恒校验要求逐键一致）。
        household_sheet.append(("冰箱", "海尔", "已上传", 1, 100))
        household_sheet.append((None, None, "未上传", 1, 50))
        household_sheet.append(("家电", None, "已上传", 2, 100))
        household_sheet.append((None, None, "未上传", 1, 50))
        household_sheet.append((None, None, "合计", 3, 150))
        household_sheet.append(("数码", None, "已上传", 4, 200))
        household_sheet.append((None, None, "未上传", 1, 40))
        household_sheet.append((None, None, "合计", 5, 240))
        _write_upload_detail_sheet(
            workbook,
            [
                _upload_detail_row(subsidy=100, remark="已上传"),
                _upload_detail_row(subsidy=50, remark="未上传"),
            ],
        )

        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            workbook.save(upload_file)
            (
                upload_data,
                digital_totals,
                project_metrics,
                corrections,
                reviews,
            ) = store_report.load_upload_data(upload_file)

        self.assertEqual(
            upload_data[("冰箱", "海尔")],
            {"已上传": Decimal("100"), "未上传": Decimal("50")},
        )
        self.assertNotIn(("家电", "海尔"), upload_data)
        self.assertEqual(corrections, [])
        self.assertEqual(reviews, [])
        self.assertEqual(
            digital_totals,
            {"发生额": Decimal("240"), "上传额": Decimal("200")},
        )
        self.assertEqual(
            project_metrics,
            {
                "家电": {
                    "已上传": store_report.CountAmount(2, Decimal("100")),
                    "未上传": store_report.CountAmount(1, Decimal("50")),
                    "合计": store_report.CountAmount(3, Decimal("150")),
                },
                "数码": {
                    "已上传": store_report.CountAmount(4, Decimal("200")),
                    "未上传": store_report.CountAmount(1, Decimal("40")),
                    "合计": store_report.CountAmount(5, Decimal("240")),
                },
            },
        )

    def test_payment_data_sums_known_digital_categories(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        sheet.append(("冰箱", "海尔", 100, 1))
        sheet.append(("合计", None, 100, 1))
        sheet.append((None, None, None, None))
        sheet.append(("手机", "OPPO", 30, 1))
        sheet.append((None, "华为", 50, 1))
        sheet.append(("平板", "华为", 20, 1))
        sheet.append(("合计", None, 100, 3))
        sheet.append(("合计", None, 200, 4))
        _write_payment_detail_sheet(workbook, [_payment_detail_row()])

        with TemporaryDirectory() as directory:
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(payment_file)
            (
                payment_data,
                digital_amount,
                project_metrics,
                payment_index,
                ambiguous_refs,
            ) = store_report.load_payment_data(payment_file)

        self.assertEqual(payment_data, {("冰箱", "海尔"): Decimal("100")})
        self.assertEqual(
            payment_index["12345678901N"],
            store_report.PaymentDetailRecord(
                category="冰箱",
                raw_category="A02-电冰箱",
                brand="海尔",
                subsidy=Decimal("100"),
            ),
        )
        self.assertEqual(ambiguous_refs, {})
        self.assertEqual(digital_amount, Decimal("100"))
        self.assertEqual(
            project_metrics,
            {
                "家电": store_report.CountAmount(1, Decimal("100")),
                "数码": store_report.CountAmount(3, Decimal("100")),
            },
        )

    def test_payment_data_rejects_an_unconfigured_category(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        sheet.append(("电脑", "联想", 100, 1))

        with TemporaryDirectory() as directory:
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(payment_file)
            with self.assertRaisesRegex(ValueError, "电脑"):
                store_report.load_payment_data(payment_file)

    def test_household_row_ratio_uses_occurred_amount(self) -> None:
        sheet = Workbook().active
        rule = store_report.RowRule(8, "冰箱", ("西门子",), "冰箱", ("西门子",))
        upload_data = {("冰箱", "西门子"): {"已上传": Decimal("100"), "未上传": Decimal("50")}}
        payment_data = {("冰箱", "西门子"): Decimal("40")}

        store_report.write_row(
            sheet, rule, upload_data, payment_data,
            {"发生额": Decimal("0"), "上传额": Decimal("0")}, Decimal("0"), FONT, {},
        )

        self.assertAlmostEqual(sheet["L8"].value, 40 / 150)
        self.assertIsNone(sheet["I8"].value)

    def test_fotile_row_includes_refrigerator_payment(self) -> None:
        """方太冰箱行两侧规则一致：审核侧（已纠正为冰箱）与回款侧都归入本行。"""
        sheet = Workbook().active
        rule = next(rule for rule in store_report.ROW_RULES if rule.row == 15)
        upload_data = {
            ("冰箱", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}
        }
        payment_data = {("冰箱", "方太"): Decimal("1500")}

        store_report.write_row(
            sheet, rule, upload_data, payment_data,
            {"发生额": Decimal("0"), "上传额": Decimal("0")}, Decimal("0"), FONT, {},
        )

        self.assertEqual(sheet["D15"].value, 1500)
        self.assertEqual(sheet["F15"].value, 1500)
        self.assertEqual(sheet["J15"].value, 1500)
        self.assertEqual(sheet["L15"].value, 1)

    def test_fotile_kitchen_row_only_takes_kitchen_data(self) -> None:
        """厨卫方太行（29）只收审核/回款两侧真正的厨卫方太；审核侧已纠正为
        冰箱的记录不会占用它，回款侧厨卫也不会流进方太冰箱行。"""
        sheet = Workbook().active
        kitchen_rule = next(
            rule for rule in store_report.ROW_RULES if rule.row == 29
        )
        fridge_rule = next(
            rule for rule in store_report.ROW_RULES if rule.row == 15
        )
        upload_data = {
            ("厨卫", "方太"): {"已上传": Decimal("800"), "未上传": Decimal("0")}
        }
        payment_data = {
            ("厨卫", "方太"): Decimal("800"),
            ("冰箱", "方太"): Decimal("1500"),
        }

        store_report.write_row(
            sheet, kitchen_rule, upload_data, payment_data,
            {"发生额": Decimal("0"), "上传额": Decimal("0")}, Decimal("0"), FONT, {},
        )
        store_report.write_row(
            sheet, fridge_rule, upload_data, payment_data,
            {"发生额": Decimal("0"), "上传额": Decimal("0")}, Decimal("0"), FONT, {},
        )

        self.assertEqual(sheet["D29"].value, 800)
        self.assertEqual(sheet["J29"].value, 800)
        # 方太冰箱行只含回款侧冰箱金额，不含厨卫回款。
        self.assertIsNone(sheet["D15"].value)
        self.assertEqual(sheet["J15"].value, 1500)

    def test_digital_row_fills_amount_and_both_ratios(self) -> None:
        sheet = Workbook().active
        rule = store_report.RowRule(34, None, (), "数码", (), fill_digital=True)
        digital_upload = {"发生额": Decimal("1000"), "上传额": Decimal("800")}

        store_report.write_row(sheet, rule, {}, {}, digital_upload, Decimal("300"), FONT, {})

        self.assertEqual(sheet["E34"].value, 1000)
        self.assertEqual(sheet["G34"].value, 800)
        self.assertEqual(sheet["K34"].value, 300)
        self.assertAlmostEqual(sheet["I34"].value, 0.8)
        self.assertAlmostEqual(sheet["M34"].value, 0.3)
        for coordinate in ("E34", "G34", "K34"):
            self.assertEqual(
                sheet[coordinate].number_format,
                store_report.DATA_NUMBER_FORMAT,
            )
        for coordinate in ("I34", "M34"):
            self.assertEqual(
                sheet[coordinate].number_format,
                store_report.PERCENT_NUMBER_FORMAT,
            )
        self.assertIsNone(sheet["D34"].value)
        self.assertIsNone(sheet["F34"].value)
        self.assertIsNone(sheet["J34"].value)

    def test_update_totals_computes_ratio_of_totals(self) -> None:
        sheet = Workbook().active
        for row in store_report.DETAIL_ROWS:
            sheet[f"D{row}"] = None
            sheet[f"F{row}"] = None
            sheet[f"J{row}"] = None
        sheet["D4"] = 100
        sheet["F4"] = 50
        sheet["J4"] = 20
        sheet["D5"] = 100
        sheet["F5"] = 50
        sheet["J5"] = 30

        store_report.update_totals(sheet, FONT, {})

        self.assertEqual(sheet["D35"].value, 200)
        self.assertEqual(sheet["F35"].value, 100)
        self.assertAlmostEqual(sheet["H35"].value, 0.5)
        self.assertAlmostEqual(sheet["L35"].value, 0.25)
        for coordinate in ("D35", "F35", "J35"):
            self.assertEqual(
                sheet[coordinate].number_format,
                store_report.DATA_NUMBER_FORMAT,
            )
        for coordinate in ("H35", "L35"):
            self.assertEqual(
                sheet[coordinate].number_format,
                store_report.PERCENT_NUMBER_FORMAT,
            )

    def test_table3_subtracts_payment_from_uploaded_metrics(self) -> None:
        sheet = Workbook().active
        upload_metrics = {
            "家电": {
                "已上传": store_report.CountAmount(10, Decimal("1000")),
                "未上传": store_report.CountAmount(3, Decimal("300")),
            },
            "数码": {
                "已上传": store_report.CountAmount(8, Decimal("800")),
                "未上传": store_report.CountAmount(2, Decimal("200")),
            },
        }
        payment_metrics = {
            "家电": store_report.CountAmount(4, Decimal("400")),
            "数码": store_report.CountAmount(3, Decimal("300")),
        }

        store_report.write_table3(sheet, upload_metrics, payment_metrics, FONT, {})

        self.assertEqual(
            [sheet[cell].value for cell in ("D50", "E50", "F50", "G50")],
            [6, 600, 3, 300],
        )
        self.assertEqual(
            [sheet[cell].value for cell in ("D51", "E51", "F51", "G51")],
            [5, 500, 2, 200],
        )
        self.assertEqual(
            [sheet[cell].value for cell in ("D52", "E52", "F52", "G52")],
            [11, 1100, 5, 500],
        )
        for row in (50, 51, 52):
            for column in ("D", "E", "F", "G"):
                self.assertEqual(
                    sheet[f"{column}{row}"].number_format,
                    store_report.DATA_NUMBER_FORMAT,
                )


class SourceHeaderValidationTests(unittest.TestCase):
    def test_load_upload_data_rejects_a_wrong_header(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        # The pre-rename header (备注 / 2026家电国补（计入收入）合计) is rejected:
        # a stale 审核明细.xlsx must fail loudly instead of being misread, and
        # rerunning the 审核明细 mode regenerates it with the new header.
        sheet.append(("财务大类", "品牌", "备注", "数量", "2026家电国补（计入收入）合计"))

        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            workbook.save(upload_file)
            with self.assertRaisesRegex(ValueError, "表头"):
                store_report.load_upload_data(upload_file)

    def test_load_payment_data_rejects_a_wrong_header(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("品类", "品牌", "金额", "笔数"))  # renamed columns

        with TemporaryDirectory() as directory:
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(payment_file)
            with self.assertRaisesRegex(ValueError, "表头"):
                store_report.load_payment_data(payment_file)

    def test_upload_header_literal_is_pinned(self) -> None:
        """The 审核明细 数据汇总 header is a cross-module contract: a rename
        must break this test rather than silently re-parse store data."""
        self.assertEqual(
            store_report.UPLOAD_HEADER,
            ("财务大类", "品牌", "上传状态", "数量", "2026国补金额"),
        )

    def test_load_upload_data_parses_statuses_with_the_literal_header(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        sheet.append(("财务大类", "品牌", "上传状态", "数量", "2026国补金额"))
        sheet.append(("家电", None, "已上传", 2, 100))
        sheet.append((None, None, "未上传", 1, 50))
        sheet.append((None, None, "合计", 3, 150))
        sheet.append(("数码", None, "已上传", 4, 200))
        sheet.append((None, None, "未上传", 1, 40))
        sheet.append((None, None, "合计", 5, 240))
        _write_upload_detail_sheet(workbook, [])

        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            workbook.save(upload_file)
            (
                _amounts,
                digital_totals,
                project_metrics,
                _corrections,
                _reviews,
            ) = store_report.load_upload_data(upload_file)

        self.assertEqual(
            digital_totals,
            {"发生额": Decimal("240"), "上传额": Decimal("200")},
        )
        self.assertEqual(
            project_metrics["家电"]["已上传"],
            store_report.CountAmount(2, Decimal("100")),
        )
        self.assertEqual(
            project_metrics["家电"]["未上传"],
            store_report.CountAmount(1, Decimal("50")),
        )
        self.assertEqual(
            project_metrics["家电"]["合计"],
            store_report.CountAmount(3, Decimal("150")),
        )
        self.assertEqual(
            project_metrics["数码"]["合计"],
            store_report.CountAmount(5, Decimal("240")),
        )


class RuleCoverageTests(unittest.TestCase):
    def test_accepts_when_every_group_is_claimed_by_a_rule(self) -> None:
        upload_data = {("冰箱", "海尔"): {"已上传": Decimal("100"), "未上传": Decimal("0")}}
        payment_data = {("冰箱", "海尔"): Decimal("80")}

        store_report.validate_rule_coverage(upload_data, payment_data)

    def test_accepts_both_fotile_keys(self) -> None:
        """审核侧纠正后的 (冰箱, 方太) 与真正的 (厨卫, 方太) 各有独立行规则。"""
        store_report.validate_rule_coverage(
            {
                ("冰箱", "方太"): {"已上传": Decimal("1500")},
                ("厨卫", "方太"): {"已上传": Decimal("800")},
            },
            {
                ("冰箱", "方太"): Decimal("1500"),
                ("厨卫", "方太"): Decimal("800"),
            },
        )

    def test_ignores_a_zero_amount_unclaimed_group(self) -> None:
        upload_data = {("平板电脑", "未知品牌"): {"已上传": Decimal("0"), "未上传": Decimal("0")}}
        payment_data = {("平板电脑", "未知品牌"): Decimal("0")}

        store_report.validate_rule_coverage(upload_data, payment_data)

    def test_rejects_an_unclaimed_nonzero_upload_group(self) -> None:
        """Simulates a brand newly appearing upstream that ROW_RULES hasn't been updated for."""
        upload_data = {("冰箱", "新品牌"): {"已上传": Decimal("500"), "未上传": Decimal("0")}}
        payment_data: dict[tuple[str, str], Decimal] = {}

        with self.assertRaisesRegex(ValueError, "新品牌"):
            store_report.validate_rule_coverage(upload_data, payment_data)

    def test_rejects_an_unclaimed_nonzero_payment_group(self) -> None:
        upload_data: dict[tuple[str, str], dict[str, Decimal]] = {}
        payment_data = {("冰箱", "新品牌"): Decimal("300")}

        with self.assertRaisesRegex(ValueError, "新品牌"):
            store_report.validate_rule_coverage(upload_data, payment_data)

    def test_real_row_rules_have_no_internal_conflicts(self) -> None:
        """Regression test: RowRule(20, ...) lists both 华为 and 华为（终端）, which
        normalize to the same brand — that must be tolerated as redundancy, not
        flagged as two different rules claiming the same key."""
        store_report.validate_rule_coverage({}, {})


class UploadCategoryCorrectionTests(unittest.TestCase):
    """审核侧品类纠正：以回款明细编码品类为权威，按交易参考号匹配。

    方案要点：A02-电冰箱 的方太 BCD 在审核侧被标成厨卫，经参考号关联到
    回款侧后纠正为冰箱，进入方太冰箱行；真正的厨卫方太与未回款的记录
    保留原品类；参考号重复、品牌/金额不符、编码品类未配置时不自动纠正。
    """

    FOTILE_REFERENCE = "17914133741N"

    def _load(self, detail_rows, payment_rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        sheet.append(store_report.UPLOAD_HEADER)
        # 数据汇总的品牌行必须与明细聚合逐键一致（守恒校验）。
        header = store_report.coupon_appliance.COUPON_OUTPUT_HEADER
        for row in detail_rows:
            category = row[header.index("财务大类")]
            brand = row[header.index("品牌")]
            remark = row[header.index("备注")]
            subsidy = row[header.index(header[6])]
            if category == "数码" or remark not in ("已上传", "未上传"):
                continue
            sheet.append((category, brand, remark, 1, subsidy))
        sheet.append(("家电", None, "已上传", 0, 0))
        sheet.append((None, None, "未上传", 0, 0))
        sheet.append((None, None, "合计", 0, 0))
        sheet.append(("数码", None, "已上传", 0, 0))
        sheet.append((None, None, "未上传", 0, 0))
        sheet.append((None, None, "合计", 0, 0))
        _write_upload_detail_sheet(workbook, detail_rows)
        payment_workbook = Workbook()
        payment_sheet = payment_workbook.active
        payment_sheet.title = store_report.PAYMENT_SHEET_NAME
        payment_sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        _write_payment_detail_sheet(payment_workbook, payment_rows)
        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(upload_file)
            payment_workbook.save(payment_file)
            (
                _payment_data,
                _digital_amount,
                _payment_metrics,
                payment_index,
                ambiguous_refs,
            ) = store_report.load_payment_data(payment_file)
            return store_report.load_upload_data(
                upload_file, payment_index, ambiguous_refs
            )

    def test_corrects_kitchen_fotile_to_refrigerator_by_reference(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    document="ZHLT000259",
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {
                ("冰箱", "方太"): {
                    "已上传": Decimal("1500"),
                    "未上传": Decimal("0"),
                }
            },
        )
        self.assertEqual(len(corrections), 1)
        correction = corrections[0]
        self.assertEqual(
            (
                correction.document_number,
                correction.reference,
                correction.brand,
                correction.original_category,
                correction.corrected_category,
                correction.raw_payment_category,
            ),
            ("ZHLT000259", self.FOTILE_REFERENCE, "方太", "厨卫", "冰箱", "A02-电冰箱"),
        )
        self.assertEqual(reviews, [])

    def test_unmatched_upload_keeps_original_category(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(reviews, [])

    def test_blank_reference_keeps_original_category(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference="",
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(reviews, [])

    def test_brand_mismatch_is_not_corrected_and_flagged(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="海尔",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(len(reviews), 1)
        self.assertIn("品牌不一致", reviews[0][0])

    def test_amount_mismatch_is_not_corrected_and_flagged(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1499.99,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(len(reviews), 1)
        self.assertIn("补贴金额不一致", reviews[0][0])

    def test_system_category_difference_is_not_corrected_silently(self) -> None:
        """国产彩电 ↔ 电视 是两套口径的体系差异，不是错标：不纠正也不提示。"""
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="国产彩电",
                    brand="海信",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=100,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A01-电视机",
                    brand="海信",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=100,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {("国产彩电", "海信"): {"已上传": Decimal("100"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(reviews, [])

    def test_multiple_payment_records_flagged_when_category_differs(self) -> None:
        """同一参考号的正负冲正对（品类一致）→ 无法唯一匹配；审核侧品类不同
        时列入人工核对而不是猜测。"""
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                ),
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=-1500,
                ),
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(len(reviews), 1)
        self.assertIn("多条回款记录", reviews[0][0])

    def test_multiple_distinct_categories_are_flagged_with_candidates(self) -> None:
        """同一参考号对应 A02-电冰箱、A04-空调 两个品类：无法唯一确定，
        必须保留原品类并列出全部候选，绝不自动纠正。"""
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                ),
                _payment_detail_row(
                    raw_category="A04-空调",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                ),
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(len(reviews), 1)
        title, details = reviews[0]
        self.assertIn("参考号对应多个编码品类", title)
        self.assertIn("A02-电冰箱｜方太", details[1])
        self.assertIn("A04-空调｜方太", details[1])

    def test_multiple_payment_records_same_category_are_silent(self) -> None:
        """冲正对但审核侧品类与回款侧相同：无需纠正也无需核对。"""
        _amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                ),
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=-1500,
                ),
            ],
        )

        self.assertEqual(corrections, [])
        self.assertEqual(reviews, [])

    def test_unmapped_payment_category_is_flagged(self) -> None:
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
            [
                _payment_detail_row(
                    raw_category="A99-未知品类",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {("厨卫", "方太"): {"已上传": Decimal("1500"), "未上传": Decimal("0")}},
        )
        self.assertEqual(corrections, [])
        self.assertEqual(len(reviews), 1)
        self.assertIn("编码品类未配置", reviews[0][0])

    def test_category_map_mismatch_raises(self) -> None:
        """编码品类映射结果与明细财务大类不一致是口径漂移，直接报错不猜测。"""
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        _write_payment_detail_sheet(
            workbook,
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )
        detail_header = (
            store_report.PAYMENT_PROFILES["家电"].detail_headers
            + store_report.PAYMENT_DERIVED_HEADERS
        )
        sheet2 = workbook[store_report.PAYMENT_PROFILES["家电"].detail_sheet_name]
        sheet2.cell(row=2, column=detail_header.index("财务大类") + 1, value="厨卫")

        with TemporaryDirectory() as directory:
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(payment_file)
            with self.assertRaisesRegex(ValueError, "不一致"):
                store_report.load_payment_data(payment_file)

    def test_summary_brand_rows_must_match_detail_aggregation(self) -> None:
        """数据汇总品牌行与明细聚合不一致（如明细状态改名、行被跳过）→ 报错。"""
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        sheet.append(store_report.UPLOAD_HEADER)
        sheet.append(("冰箱", "海尔", "已上传", 1, 100))
        sheet.append(("家电", None, "已上传", 0, 0))
        sheet.append((None, None, "未上传", 0, 0))
        sheet.append((None, None, "合计", 0, 0))
        sheet.append(("数码", None, "已上传", 0, 0))
        sheet.append((None, None, "未上传", 0, 0))
        sheet.append((None, None, "合计", 0, 0))
        # 明细里是 150，与数据汇总的 100 不一致。
        _write_upload_detail_sheet(
            workbook,
            [_upload_detail_row(subsidy=150, remark="已上传")],
        )

        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            workbook.save(upload_file)
            with self.assertRaisesRegex(
                ValueError, "品牌金额不一致，共 1 项.*冰箱/海尔/已上传：明细 150.00，汇总 100.00"
            ):
                store_report.load_upload_data(upload_file)

    def test_correction_moves_only_the_matching_row(self) -> None:
        """同一 (原品类, 品牌) 下部分行纠正、部分行保留，金额互不串扰。"""
        amounts, _totals, _metrics, corrections, reviews = self._load(
            [
                _upload_detail_row(
                    document="A1",
                    category="厨卫",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                ),
                _upload_detail_row(
                    document="A2",
                    category="厨卫",
                    brand="方太",
                    reference="",  # 无参考号 → 保留厨卫
                    subsidy=800,
                ),
            ],
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference=self.FOTILE_REFERENCE,
                    subsidy=1500,
                )
            ],
        )

        self.assertEqual(
            amounts,
            {
                ("冰箱", "方太"): {
                    "已上传": Decimal("1500"),
                    "未上传": Decimal("0"),
                },
                ("厨卫", "方太"): {
                    "已上传": Decimal("800"),
                    "未上传": Decimal("0"),
                },
            },
        )
        self.assertEqual(len(corrections), 1)
        self.assertEqual(reviews, [])


class BrandGroupTests(unittest.TestCase):
    def test_haier_group_sums_across_categories(self) -> None:
        rule = next(rule for rule in store_report.BRAND_GROUP_RULES if rule.name == "海尔系")
        upload_data = {
            ("冰箱", "海尔"): {"已上传": Decimal("100"), "未上传": Decimal("20")},
            ("国产彩电", "卡萨帝"): {"已上传": Decimal("50"), "未上传": Decimal("0")},
        }
        payment_data = {
            ("冰箱", "海尔"): Decimal("60"),
            ("电视", "卡萨帝"): Decimal("30"),
        }

        occurred, uploaded, paid = store_report.sum_brand_group(upload_data, payment_data, rule.categories)

        self.assertEqual(occurred, Decimal("170"))
        self.assertEqual(uploaded, Decimal("150"))
        self.assertEqual(paid, Decimal("90"))

    def test_midea_group_sums_across_categories(self) -> None:
        rule = next(rule for rule in store_report.BRAND_GROUP_RULES if rule.name == "美的系")
        upload_data = {
            ("冰箱", "美的"): {"已上传": Decimal("40"), "未上传": Decimal("10")},
            ("空调", "美的"): {"已上传": Decimal("200"), "未上传": Decimal("50")},
        }
        payment_data = {
            ("冰箱", "美的"): Decimal("30"),
            ("空调", "美的"): Decimal("150"),
        }

        occurred, uploaded, paid = store_report.sum_brand_group(upload_data, payment_data, rule.categories)

        self.assertEqual(occurred, Decimal("300"))
        self.assertEqual(uploaded, Decimal("240"))
        self.assertEqual(paid, Decimal("180"))

    def test_upload_category_differs_from_payment_category_for_tv_brands(self) -> None:
        """Regression test: TV rows are 国产彩电 in 审核明细 but 电视 in 回款明细."""
        rule = next(rule for rule in store_report.BRAND_GROUP_RULES if rule.name == "创维")
        upload_data = {("国产彩电", "创维"): {"已上传": Decimal("70"), "未上传": Decimal("10")}}
        payment_data = {("电视", "创维"): Decimal("50")}

        occurred, uploaded, paid = store_report.sum_brand_group(upload_data, payment_data, rule.categories)

        self.assertEqual(occurred, Decimal("80"))
        self.assertEqual(uploaded, Decimal("70"))
        self.assertEqual(paid, Decimal("50"))


class TemplateValidationTests(unittest.TestCase):
    def test_accepts_a_correctly_structured_template(self) -> None:
        store_report.validate_template(_build_minimal_template())

    def test_rejects_a_blank_workbook(self) -> None:
        with self.assertRaises(ValueError):
            store_report.validate_template(Workbook())

    def test_rejects_a_legacy_template_without_version_marker(self) -> None:
        """旧模板（含手工修改版）没有版本标记，必须给出更换提示而不是
        静默产出报表。"""
        workbook = _build_minimal_template()
        workbook.active[store_report.TEMPLATE_VERSION_CELL] = None

        with self.assertRaisesRegex(ValueError, "请更换为新版模板"):
            store_report.validate_template(workbook)

    def test_rejects_a_legacy_template_with_old_row_layout(self) -> None:
        """旧版模板第 15 行是海信电视、没有方太冰箱行：结构校验拒绝。"""
        workbook = _build_minimal_template()
        workbook.active["C15"] = "海信电视"

        with self.assertRaisesRegex(ValueError, "请更换为新版模板"):
            store_report.validate_template(workbook)

    def test_rejects_a_template_with_a_moved_label(self) -> None:
        workbook = _build_minimal_template()
        workbook.active["C39"] = "格力系"  # 海尔系 moved or renamed

        with self.assertRaisesRegex(ValueError, "C39"):
            store_report.validate_template(workbook)

    def test_rejects_a_template_missing_a_key_merge(self) -> None:
        workbook = _build_minimal_template()
        workbook.active.unmerge_cells("A4:A15")

        with self.assertRaisesRegex(ValueError, "合并区域"):
            store_report.validate_template(workbook)

    def test_rejects_a_template_with_broken_sequence_numbers(self) -> None:
        workbook = _build_minimal_template()
        workbook.active["B15"] = 99

        with self.assertRaisesRegex(ValueError, "序号"):
            store_report.validate_template(workbook)

    def test_rejects_a_template_with_visible_version_row(self) -> None:
        workbook = _build_minimal_template()
        workbook.active.row_dimensions[
            store_report.TEMPLATE_VERSION_ROW
        ].hidden = False

        with self.assertRaisesRegex(ValueError, "隐藏"):
            store_report.validate_template(workbook)

    def test_accepts_a_template_with_extra_merges(self) -> None:
        """多余合并允许（模板可能有其他区域合并），关键合并缺失才拒绝。"""
        workbook = _build_minimal_template()
        workbook.active.merge_cells("M52:M53")

        store_report.validate_template(workbook)

    def test_row_rules_cover_every_template_detail_row(self) -> None:
        """行号契约：ROW_RULES 必须恰好覆盖模板明细区 4..TOTAL_ROW，无跳空。"""
        rule_rows = {rule.row for rule in store_report.ROW_RULES}
        self.assertEqual(
            rule_rows,
            set(range(4, store_report.TOTAL_ROW)),
        )
        self.assertEqual(
            sorted(rule_rows),
            list(range(4, store_report.TOTAL_ROW)),
        )

    def test_brand_group_rules_cover_template_rows(self) -> None:
        group_rows = {rule.row for rule in store_report.BRAND_GROUP_RULES}
        self.assertEqual(
            group_rows,
            set(range(39, store_report.BRAND_GROUP_TOTAL_ROW)),
        )

    def test_layout_constants_match_structure_cells(self) -> None:
        """表 1 总计、表 2 合计、表 3 行号与模板关键标签互相锁定。"""
        self.assertEqual(store_report.TOTAL_ROW, 35)
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["A35"], "费用总计")
        self.assertEqual(store_report.BRAND_GROUP_TOTAL_ROW, 46)
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C46"], "合计")
        self.assertEqual(
            store_report.TABLE3_PROJECT_ROWS,
            {"家电": 50, "数码": 51},
        )
        self.assertEqual(store_report.TABLE3_TOTAL_ROW, 52)
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C50"], "家电")
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C51"], "数码")
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C52"], "合计")

    def test_fotile_rows_are_separate_and_both_sided(self) -> None:
        """方太冰箱（15）与方太厨卫（29）是两条独立规则，两侧品类一致，
        不再有跨品类临时规则。"""
        fridge = next(rule for rule in store_report.ROW_RULES if rule.row == 15)
        kitchen = next(rule for rule in store_report.ROW_RULES if rule.row == 29)
        self.assertEqual(
            fridge,
            store_report.RowRule(15, "冰箱", ("方太",), "冰箱", ("方太",)),
        )
        self.assertEqual(
            kitchen,
            store_report.RowRule(29, "厨卫", ("方太",), "厨卫", ("方太",)),
        )
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C15"], "方太冰箱")
        self.assertEqual(store_report.TEMPLATE_STRUCTURE_CELLS["C29"], "方太")
        # 白名单必须覆盖纠正目标：冰箱在审核侧有行规则，电视（对应国产彩电）
        # 不在。
        self.assertIn("冰箱", store_report.UPLOAD_CATEGORIES)
        self.assertNotIn("电视", store_report.UPLOAD_CATEGORIES)

    def test_rejects_extra_sheets(self) -> None:
        workbook = _build_minimal_template()
        workbook.create_sheet("附表")

        with self.assertRaisesRegex(ValueError, "工作表数量"):
            store_report.validate_template(workbook)

    def test_rejects_wrong_column_count(self) -> None:
        workbook = _build_minimal_template()
        workbook.active["N1"] = "多出来的列"

        with self.assertRaisesRegex(ValueError, "列数"):
            store_report.validate_template(workbook)

    def test_rejects_too_few_rows(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.EXPECTED_SHEET_TITLE
        for column_index in range(1, store_report.EXPECTED_COLUMN_COUNT + 1):
            sheet.cell(row=1, column=column_index, value="")
        for coordinate, value in store_report.TEMPLATE_STRUCTURE_CELLS.items():
            if int("".join(filter(str.isdigit, coordinate))) <= 35:
                sheet[coordinate] = value

        with self.assertRaisesRegex(ValueError, "行数"):
            store_report.validate_template(workbook)

    def test_rejects_the_wrong_sheet_title(self) -> None:
        workbook = _build_minimal_template()
        workbook.active.title = "其他门店"

        with self.assertRaisesRegex(ValueError, "工作表名称"):
            store_report.validate_template(workbook)


class TemplateResolutionTests(unittest.TestCase):
    def test_resolves_the_sole_matching_template(self) -> None:
        with TemporaryDirectory() as directory:
            data_dir = Path(directory)
            template = data_dir / "2026年门店国补上传及回款情况表 （益庄店）.xlsx"
            Workbook().save(template)
            store_report.configure_data_dir(data_dir)

            self.assertEqual(store_report.resolve_template_file(), template)

    def test_raises_when_no_template_is_found(self) -> None:
        with TemporaryDirectory() as directory:
            store_report.configure_data_dir(Path(directory))

            with self.assertRaises(FileNotFoundError):
                store_report.resolve_template_file()


class ValidateOutputTests(unittest.TestCase):
    def test_accepts_a_workbook_matching_its_expected_cells_and_totals(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "益庄"
        _fill_columns(sheet)
        sheet["A1"] = "更新时间：2026-07-28 00:00:00"
        sheet["D4"] = 100
        sheet["D35"] = 100
        sheet["D39"] = 100
        sheet["D46"] = 100

        with TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            workbook.save(path)
            store_report.validate_output(path, {"D4": 100.0}, "益庄")

    def test_rejects_a_workbook_whose_total_disagrees_with_its_detail_rows(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "益庄"
        _fill_columns(sheet)
        sheet["A1"] = "更新时间：2026-07-28 00:00:00"
        sheet["D4"] = 100
        sheet["D35"] = 999

        with TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            workbook.save(path)
            with self.assertRaises(ValueError):
                store_report.validate_output(path, {}, "益庄")

    def test_rejects_a_workbook_whose_written_cell_does_not_match_expectation(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "益庄"
        _fill_columns(sheet)
        sheet["A1"] = "更新时间：2026-07-28 00:00:00"
        sheet["D4"] = 100

        with TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            workbook.save(path)
            with self.assertRaisesRegex(ValueError, "D4"):
                store_report.validate_output(path, {"D4": 999.0}, "益庄")

    def test_rejects_a_workbook_missing_the_update_timestamp(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "益庄"
        _fill_columns(sheet)

        with TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            workbook.save(path)
            with self.assertRaisesRegex(ValueError, "更新时间"):
                store_report.validate_output(path, {}, "益庄")

    def test_rejects_a_workbook_with_the_wrong_sheet_title(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "其他门店"
        _fill_columns(sheet)
        sheet["A1"] = "更新时间：2026-07-28 00:00:00"

        with TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            workbook.save(path)
            with self.assertRaisesRegex(ValueError, "工作表名称"):
                store_report.validate_output(path, {}, "益庄")


class ProcessStoreReportIntegrationTests(unittest.TestCase):
    def _write_upload_file(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        sheet.append(store_report.UPLOAD_HEADER)
        # 数据汇总的品牌行是明细聚合的守恒校验基准。
        sheet.append(("冰箱", "海尔", "已上传", 1, 100))
        sheet.append((None, None, "未上传", 1, 20))
        sheet.append(("厨卫", "方太", "已上传", 1, 1500))
        sheet.append(("家电", None, "已上传", 2, 1600))
        sheet.append((None, None, "未上传", 1, 20))
        sheet.append((None, None, "合计", 3, 1620))
        sheet.append(("数码", None, "已上传", 1, 40))
        sheet.append((None, None, "未上传", 1, 10))
        sheet.append((None, None, "合计", 2, 50))
        # 品牌行的权威来源是家电-明细总表（行级纠正）；其中方太 BCD 在
        # 审核侧被标成厨卫，集成链路里靠回款明细的参考号纠正为冰箱。
        _write_upload_detail_sheet(
            workbook,
            [
                _upload_detail_row(
                    document="ZHLT000001",
                    category="冰箱",
                    brand="海尔",
                    reference="12345678901N",
                    subsidy=100,
                    remark="已上传",
                ),
                _upload_detail_row(
                    document="ZHLT000002",
                    category="冰箱",
                    brand="海尔",
                    reference="12345678902N",
                    subsidy=20,
                    remark="未上传",
                ),
                _upload_detail_row(
                    document="ZHLT000259",
                    category="厨卫",
                    brand="方太",
                    reference="17914133741N",
                    subsidy=1500,
                    remark="已上传",
                ),
            ],
        )
        workbook.save(path)

    def _write_payment_file(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        sheet.append(("冰箱", "海尔", 60, 1))
        sheet.append(("冰箱", "方太", 1500, 1))
        sheet.append(("合计", None, 1560, 2))
        sheet.append((None, None, None, None))
        sheet.append(("手机", "OPPO", 15, 1))
        _write_payment_detail_sheet(
            workbook,
            [
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="海尔",
                    reference="12345678901N",
                    subsidy=60,
                ),
                _payment_detail_row(
                    raw_category="A02-电冰箱",
                    brand="方太",
                    reference="17914133741N",
                    subsidy=1500,
                ),
            ],
        )
        workbook.save(path)

    def test_runs_end_to_end_and_produces_a_validated_report(self) -> None:
        with TemporaryDirectory() as directory:
            base = Path(directory)
            data_dir = base / "data"
            data_dir.mkdir()
            upload_file = base / "审核明细.xlsx"
            payment_file = base / "回款明细.xlsx"
            output_file = base / "output.xlsx"
            template_file = data_dir / "2026年门店国补上传及回款情况表 （益庄店）.xlsx"

            self._write_upload_file(upload_file)
            self._write_payment_file(payment_file)
            _build_minimal_template().save(template_file)
            store_report.configure_data_dir(data_dir)

            reporter_stream = io.StringIO()
            reporter = store_report.ConsoleReporter(stream=reporter_stream)
            with (
                patch.object(store_report, "UPLOAD_FILE", upload_file),
                patch.object(store_report, "PAYMENT_FILE", payment_file),
                patch.object(store_report, "OUTPUT_FILE", output_file),
                patch.object(
                    store_report,
                    "resolve_font",
                    return_value=("Maple Mono NF CN", Path("unused.ttf")),
                ),
            ):
                store_report.process_store_report(reporter)
            reporter.finish(success=True, succeeded=1, total=1)
            reporter_text = reporter_stream.getvalue()

            self.assertTrue(output_file.exists())
            result = load_workbook(output_file, data_only=True)
            sheet = result[result.sheetnames[0]]
            self.assertEqual(sheet["D4"].value, 120)
            self.assertEqual(sheet["F4"].value, 100)
            self.assertEqual(sheet["J4"].value, 60)
            # 审核侧厨卫/方太经参考号纠正为冰箱，进入方太冰箱行。
            self.assertEqual(sheet["D15"].value, 1500)
            self.assertEqual(sheet["F15"].value, 1500)
            self.assertEqual(sheet["J15"].value, 1500)
            self.assertEqual(sheet["L15"].value, 1)
            # 厨卫方太行（29）不重复计算冰箱金额。
            self.assertIsNone(sheet["D29"].value)
            self.assertIsNone(sheet["J29"].value)
            self.assertEqual(sheet["E34"].value, 50)
            self.assertEqual(sheet["G34"].value, 40)
            self.assertEqual(sheet["K34"].value, 15)
            self.assertIsNone(sheet["D50"].value)
            self.assertEqual(sheet["E50"].value, 40)
            self.assertEqual(sheet["F50"].value, 1)
            self.assertEqual(sheet["G50"].value, 20)
            self.assertIsNone(sheet["D51"].value)
            self.assertEqual(sheet["E51"].value, 25)
            self.assertEqual(sheet["F51"].value, 1)
            self.assertEqual(sheet["G51"].value, 10)
            self.assertIsNone(sheet["D52"].value)
            self.assertEqual(sheet["E52"].value, 65)
            self.assertEqual(sheet["F52"].value, 2)
            self.assertEqual(sheet["G52"].value, 30)
            self.assertIn("更新时间：", sheet["A1"].value)
            self.assertIn("审核明细品类纠正：1 条", reporter_text)
            # 合并区域内部的 MergedCell 不参与渲染，openpyxl 保存时会丢弃
            # 对其样式的写入（保持模板原样），所以只断言普通单元格的字体。
            output_fonts = {
                cell.font.name
                for row in sheet.iter_rows()
                for cell in row
                if not isinstance(cell, MergedCell)
            }
            self.assertEqual(output_fonts, {"Maple Mono NF CN"})
            leftover = [entry.name for entry in output_file.parent.iterdir() if entry.name.startswith(".")]
            self.assertEqual(leftover, [])


class RollbackTests(unittest.TestCase):
    def test_prior_report_is_restored_when_store_report_fails(self) -> None:
        with TemporaryDirectory() as directory:
            output_file = Path(directory) / "report.xlsx"
            other_output = Path(directory) / "other.xlsx"
            output_file.write_bytes(b"old report")
            other_output.write_bytes(b"old other")

            def fail_after_writing_other() -> None:
                other_output.write_bytes(b"new other")
                output_file.write_bytes(b"corrupted partial report")
                raise ValueError("门店报表处理失败")

            with self.assertRaisesRegex(ValueError, "门店报表处理失败"):
                run_with_output_rollback((other_output, output_file), fail_after_writing_other)

            self.assertEqual(output_file.read_bytes(), b"old report")
            self.assertEqual(other_output.read_bytes(), b"old other")

    def test_no_half_written_output_remains_when_nothing_existed_before(self) -> None:
        with TemporaryDirectory() as directory:
            output_dir = Path(directory)
            output_file = output_dir / "report.xlsx"

            def fail_after_creating_output() -> None:
                output_file.write_bytes(b"half-written report")
                raise ValueError("门店报表处理失败")

            with self.assertRaisesRegex(ValueError, "门店报表处理失败"):
                run_with_output_rollback((output_file,), fail_after_creating_output)

            self.assertFalse(output_file.exists())
            self.assertEqual(list(output_dir.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
