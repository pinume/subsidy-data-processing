from __future__ import annotations

import unittest
from datetime import timedelta
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
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
    return workbook


class ReportRatioTests(unittest.TestCase):
    def test_current_timestamp_uses_shanghai_timezone(self) -> None:
        timestamp = store_report.current_timestamp()

        self.assertEqual(timestamp.tzinfo.key, "Asia/Shanghai")
        self.assertEqual(timestamp.utcoffset(), timedelta(hours=8))

    def test_upload_data_reads_embedded_digital_summary(self) -> None:
        workbook = Workbook()
        household_sheet = workbook.active
        household_sheet.title = store_report.UPLOAD_SHEET_NAME
        household_sheet.append(store_report.UPLOAD_HEADER)
        household_sheet.append(("冰箱", "海尔", "已上传", 1, 100))
        household_sheet.append((None, None, "未上传", 1, 50))
        household_sheet.append(("家电", None, "已上传", 2, 100))
        household_sheet.append((None, None, "未上传", 1, 50))
        household_sheet.append((None, None, "合计", 3, 150))
        household_sheet.append(("数码", None, "已上传", 4, 200))
        household_sheet.append((None, None, "未上传", 1, 40))
        household_sheet.append((None, None, "合计", 5, 240))

        with TemporaryDirectory() as directory:
            upload_file = Path(directory) / "审核明细.xlsx"
            workbook.save(upload_file)
            upload_data, digital_totals = store_report.load_upload_data(upload_file)

        self.assertEqual(
            upload_data[("冰箱", "海尔")],
            {"已上传": Decimal("100"), "未上传": Decimal("50")},
        )
        self.assertNotIn(("家电", "海尔"), upload_data)
        self.assertEqual(
            digital_totals,
            {"发生额": Decimal("240"), "上传额": Decimal("200")},
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

        with TemporaryDirectory() as directory:
            payment_file = Path(directory) / "回款明细.xlsx"
            workbook.save(payment_file)
            payment_data, digital_amount = store_report.load_payment_data(payment_file)

        self.assertEqual(payment_data, {("冰箱", "海尔"): Decimal("100")})
        self.assertEqual(digital_amount, Decimal("100"))

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

    def test_household_row_ratio_uses_uploaded_amount(self) -> None:
        sheet = Workbook().active
        rule = store_report.RowRule(8, "冰箱", ("西门子",), "冰箱", ("西门子",))
        upload_data = {("冰箱", "西门子"): {"已上传": Decimal("100"), "未上传": Decimal("50")}}
        payment_data = {("冰箱", "西门子"): Decimal("40")}

        store_report.write_row(
            sheet, rule, upload_data, payment_data,
            {"发生额": Decimal("0"), "上传额": Decimal("0")}, Decimal("0"), FONT, {},
        )

        self.assertAlmostEqual(sheet["L8"].value, 0.4)
        self.assertIsNone(sheet["I8"].value)

    def test_digital_row_fills_amount_and_both_ratios(self) -> None:
        sheet = Workbook().active
        rule = store_report.RowRule(33, None, (), "数码", (), fill_digital=True)
        digital_upload = {"发生额": Decimal("1000"), "上传额": Decimal("800")}

        store_report.write_row(sheet, rule, {}, {}, digital_upload, Decimal("300"), FONT, {})

        self.assertEqual(sheet["E33"].value, 1000)
        self.assertEqual(sheet["G33"].value, 800)
        self.assertEqual(sheet["K33"].value, 300)
        self.assertAlmostEqual(sheet["I33"].value, 0.8)
        self.assertAlmostEqual(sheet["M33"].value, 0.375)
        self.assertIsNone(sheet["D33"].value)
        self.assertIsNone(sheet["F33"].value)
        self.assertIsNone(sheet["J33"].value)

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

        self.assertEqual(sheet["D34"].value, 200)
        self.assertEqual(sheet["F34"].value, 100)
        self.assertAlmostEqual(sheet["H34"].value, 0.5)
        self.assertAlmostEqual(sheet["L34"].value, 0.5)


class SourceHeaderValidationTests(unittest.TestCase):
    def test_load_upload_data_rejects_a_wrong_header(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.UPLOAD_SHEET_NAME
        sheet.append(("财务大类", "品牌", "备注", "数量", "金额"))  # missing the year-specific label

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


class RuleCoverageTests(unittest.TestCase):
    def test_accepts_when_every_group_is_claimed_by_a_rule(self) -> None:
        upload_data = {("冰箱", "海尔"): {"已上传": Decimal("100"), "未上传": Decimal("0")}}
        payment_data = {("冰箱", "海尔"): Decimal("80")}

        store_report.validate_rule_coverage(upload_data, payment_data)

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
        """Regression test: RowRule(19, ...) lists both 华为 and 华为（终端）, which
        normalize to the same brand — that must be tolerated as redundancy, not
        flagged as two different rules claiming the same key."""
        store_report.validate_rule_coverage({}, {})


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

    def test_rejects_a_template_with_a_moved_label(self) -> None:
        workbook = _build_minimal_template()
        workbook.active["C38"] = "格力系"  # 海尔系 moved or renamed

        with self.assertRaisesRegex(ValueError, "C38"):
            store_report.validate_template(workbook)

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
            if int("".join(filter(str.isdigit, coordinate))) <= 34:
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
        sheet["D34"] = 100
        sheet["D38"] = 100
        sheet["D45"] = 100

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
        sheet["D34"] = 999

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
        sheet.append(("冰箱", "海尔", "已上传", 1, 100))
        sheet.append((None, None, "未上传", 1, 20))
        sheet.append(("数码", None, "已上传", 1, 40))
        sheet.append((None, None, "未上传", 1, 10))
        workbook.save(path)

    def _write_payment_file(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = store_report.PAYMENT_SHEET_NAME
        sheet.append(("财务大类", "品牌", "补贴金额合计", "补贴金额计数"))
        sheet.append(("冰箱", "海尔", 60, 1))
        sheet.append(("合计", None, 60, 1))
        sheet.append((None, None, None, None))
        sheet.append(("手机", "OPPO", 15, 1))
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

            with (
                patch.object(store_report, "UPLOAD_FILE", upload_file),
                patch.object(store_report, "PAYMENT_FILE", payment_file),
                patch.object(store_report, "OUTPUT_FILE", output_file),
            ):
                store_report.process_store_report()

            self.assertTrue(output_file.exists())
            result = load_workbook(output_file, data_only=True)
            sheet = result[result.sheetnames[0]]
            self.assertEqual(sheet["D4"].value, 120)
            self.assertEqual(sheet["F4"].value, 100)
            self.assertEqual(sheet["J4"].value, 60)
            self.assertEqual(sheet["E33"].value, 50)
            self.assertEqual(sheet["G33"].value, 40)
            self.assertEqual(sheet["K33"].value, 15)
            self.assertIn("更新时间：", sheet["A1"].value)
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
