import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from processors import receipts

HEADER = list(receipts.RECEIPTS_SOURCE_HEADER)


class PrepareReceiptDataIssuesTest(unittest.TestCase):
    def test_issues_cover_duplicate_missing_and_invalid_original(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            ["ZH0001", "2026-01-24", "", "", "美的空调"],
            ["ZH0002", None, "", "", "格力空调"],
            ["ZH0003", "2026-01-25", "notvalid", "", "小米手机"],
        ]

        _output_rows, stats, issues, duplicate_match_keys = (
            receipts.prepare_receipt_data(kept_rows)
        )

        self.assertEqual(
            issues,
            [
                (
                    "重复匹配键",
                    "3、4",
                    "260124ZH0001",
                    "多个数据行生成了相同的日期与单据号组合键",
                ),
                (
                    "缺少匹配键",
                    "5",
                    "",
                    "日期或单据号为空，无法生成匹配键",
                ),
                (
                    "原票号格式异常",
                    "6",
                    "notvalid",
                    "原票号应为6位日期加单据号",
                ),
                (
                    "原票号未匹配",
                    "6",
                    "notvalid",
                    "未找到日期与单据号组合键相同的原单",
                ),
            ],
        )
        self.assertEqual(stats["重复匹配键数量"], 1)
        self.assertEqual(stats["缺少匹配键数量"], 1)
        self.assertEqual(stats["原票号格式异常数量"], 1)
        self.assertEqual(duplicate_match_keys, {"260124ZH0001"})

    def test_no_issues_yields_empty_list(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
        ]

        _output_rows, _stats, issues, _duplicates = receipts.prepare_receipt_data(
            kept_rows
        )

        self.assertEqual(issues, [])


class IssuesSheetRoundTripTest(unittest.TestCase):
    def test_issues_sheet_is_written_and_validated(self) -> None:
        issues = [
            ("缺少匹配键", "5", "", "日期或单据号为空，无法生成匹配键"),
        ]
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(receipts.RECEIPTS_OUTPUT_HEADER)

        from processors.common.excel import load_measurement_font, resolve_font

        font_name, font_path = resolve_font()
        measurement_font = load_measurement_font(font_path)
        receipts.build_issues_sheet(workbook, issues, font_name, measurement_font)

        self.assertEqual(workbook.sheetnames, ["Sheet1", receipts.ISSUES_SHEET_NAME])
        issues_sheet = workbook[receipts.ISSUES_SHEET_NAME]
        self.assertEqual(
            tuple(cell.value for cell in issues_sheet[1]),
            receipts.ISSUES_HEADER,
        )
        self.assertEqual(
            [row for row in issues_sheet.iter_rows(min_row=2, values_only=True)],
            issues,
        )

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "收款单统计.xlsx"
            workbook.save(path)
            workbook.close()

            receipts.validate_receipts_output(path, 0, issues)

    def test_validation_rejects_mismatched_issues(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(receipts.RECEIPTS_OUTPUT_HEADER)

        from processors.common.excel import load_measurement_font, resolve_font

        font_name, font_path = resolve_font()
        measurement_font = load_measurement_font(font_path)
        receipts.build_issues_sheet(workbook, [], font_name, measurement_font)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "收款单统计.xlsx"
            workbook.save(path)
            workbook.close()

            with self.assertRaisesRegex(RuntimeError, "问题明细工作表内容校验失败"):
                receipts.validate_receipts_output(
                    path,
                    0,
                    [("缺少匹配键", "5", "", "说明")],
                )


if __name__ == "__main__":
    unittest.main()
