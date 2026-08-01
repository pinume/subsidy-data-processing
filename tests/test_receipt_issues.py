import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from processors import receipts

HEADER = list(receipts.RECEIPTS_SOURCE_HEADER)
SOURCE_NAME = "收款单统计.xlsx"


def prepare(kept_rows):
    return receipts.prepare_receipt_data(kept_rows, source_name=SOURCE_NAME)


class PrepareReceiptDataIssuesTest(unittest.TestCase):
    def test_issues_cover_duplicate_missing_and_invalid_original(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            ["ZH0001", "2026-01-24", "", "", "美的空调"],
            ["ZH0002", None, "", "", "格力空调"],
            ["ZH0003", "2026-01-25", "notvalid", "", "小米手机"],
        ]

        _output_rows, stats, issues, duplicate_match_keys = prepare(kept_rows)

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

        _output_rows, _stats, issues, _duplicates = prepare(kept_rows)

        self.assertEqual(issues, [])

    def test_invalid_date_error_names_source_file_and_real_row(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            ["ZH0002", "2026.1.25", "", "", "美的空调"],
        ]

        with self.assertRaises(ValueError) as caught:
            prepare(kept_rows)

        message = str(caught.exception)
        self.assertIn(SOURCE_NAME, message)
        self.assertIn("第 4 行", message)
        self.assertIn("2026.1.25", message)


class BlankAndTotalRowTest(unittest.TestCase):
    def test_fully_blank_row_is_skipped(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            [None, None, None, None, None],
            ["", "  ", "", "", ""],
        ]

        output_rows, stats, issues, _duplicates = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["总数据量"], 1)
        self.assertEqual(stats["跳过空白行数"], 2)
        self.assertEqual(issues, [])

    def test_row_with_any_value_is_kept(self) -> None:
        kept_rows = [
            HEADER,
            [None, None, None, None, "海尔冰箱"],
        ]

        output_rows, stats, issues, _duplicates = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["跳过空白行数"], 0)
        # Still incomplete data, so it must keep reporting itself.
        self.assertEqual([issue[0] for issue in issues], ["缺少匹配键"])

    def test_numeric_zero_is_not_blank(self) -> None:
        kept_rows = [
            HEADER,
            [0, None, None, None, None],
        ]

        output_rows, stats, _issues, _duplicates = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["跳过空白行数"], 0)

    def test_total_row_is_skipped_in_either_column(self) -> None:
        for total_row in (
            ["合计", "合计", "", "", ""],
            ["合计", None, "", "", ""],
            [None, "合计", "", "", ""],
        ):
            with self.subTest(total_row=total_row):
                kept_rows = [
                    HEADER,
                    ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
                    total_row,
                ]

                output_rows, stats, _issues, _duplicates = prepare(kept_rows)

                self.assertEqual(len(output_rows), 1)
                self.assertEqual(stats["跳过合计行数"], 1)

    def test_合计_inside_text_fields_is_not_a_total_row(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "合计说明", "合计套装"],
        ]

        output_rows, stats, _issues, _duplicates = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["跳过合计行数"], 0)

    def test_skipped_rows_do_not_shift_later_row_numbers(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            [None, None, None, None, None],
            ["合计", "合计", "", "", ""],
            # Excel row 6: two rows above it were skipped, but the reported
            # row number must still point an operator at the real line.
            ["ZH0002", None, "", "", "格力空调"],
        ]

        _output_rows, _stats, issues, _duplicates = prepare(kept_rows)

        self.assertEqual(
            issues,
            [("缺少匹配键", "6", "", "日期或单据号为空，无法生成匹配键")],
        )


class SpecialRemarkStatsTest(unittest.TestCase):
    def _special_key_row(self) -> tuple[str, list[object]]:
        match_key = min(receipts.RECEIPTS_SPECIAL_REMARK_KEYS)
        date_part, document_number = match_key[:6], match_key[6:]
        receipt_date = (
            f"20{date_part[:2]}-{date_part[2:4]}-{date_part[4:6]}"
        )
        return match_key, [document_number, receipt_date, "", "", "测试商品"]

    def test_special_remark_is_counted_in_total(self) -> None:
        match_key, row = self._special_key_row()
        output_rows, stats, _issues, _duplicates = prepare([HEADER, row])

        self.assertEqual(output_rows[0][-1], receipts.RECEIPTS_REMARK_SPECIAL)
        # The row has no 原票号 and is nobody's original, so it belongs to
        # none of the 退单/原单 tallies the total used to be summed from.
        self.assertEqual(stats["仅退单数量"], 0)
        self.assertEqual(stats["仅原单数量"], 0)
        self.assertEqual(stats["退单及原单数量"], 0)
        self.assertEqual(stats["备注总数"], 1)
        self.assertEqual(stats["特殊备注数量"], 1)
        self.assertEqual(stats["生效特殊匹配键"], [match_key])

    def test_special_remark_counted_once_per_row(self) -> None:
        row = self._special_key_row()[1]
        _output_rows, stats, _issues, _duplicates = prepare([HEADER, row, list(row)])

        self.assertEqual(stats["备注总数"], 2)
        self.assertEqual(stats["特殊备注数量"], 2)

    def test_special_key_outranks_other_remark_rules(self) -> None:
        row = self._special_key_row()[1]
        # Give it an 原票号 too: without the special rule this row would be
        # 退换货/倒票（退单）.
        row[2] = "260101ZH9999"
        output_rows, _stats, _issues, _duplicates = prepare([HEADER, row])

        self.assertEqual(output_rows[0][-1], receipts.RECEIPTS_REMARK_SPECIAL)

    def test_missing_special_keys_are_reported_not_fatal(self) -> None:
        _output_rows, stats, _issues, _duplicates = prepare(
            [HEADER, ["ZH0001", "2026-01-24", "", "", "海尔冰箱"]]
        )

        self.assertEqual(stats["生效特殊匹配键"], [])
        self.assertEqual(
            stats["未找到特殊匹配键"],
            sorted(receipts.RECEIPTS_SPECIAL_REMARK_KEYS),
        )

    def test_ordinary_remarks_still_counted(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "260101ZH9999", "", "海尔冰箱"],
        ]

        _output_rows, stats, _issues, _duplicates = prepare(kept_rows)

        self.assertEqual(stats["仅退单数量"], 1)
        self.assertEqual(stats["备注总数"], 1)
        self.assertEqual(stats["特殊备注数量"], 0)


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
