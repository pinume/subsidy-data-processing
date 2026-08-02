import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

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

        _output_rows, stats, issues = prepare(kept_rows)

        # A duplicate match key is expected here (multi-line suite sales
        # share one 单据号/日期) and is never reported as an issue — see
        # test_duplicate_match_keys_are_not_reported_as_issues. Only the
        # genuine problems below reach 问题明细, numbered by their row in the
        # output sheet (row 2 is ZH0001, row 3 the ZH0001 duplicate, row 4
        # ZH0002, row 5 ZH0003).
        self.assertEqual(
            issues,
            [
                (
                    "缺少匹配键",
                    "4",
                    "",
                    "日期或单据号为空，无法生成匹配键",
                ),
                (
                    "原票号格式异常",
                    "5",
                    "notvalid",
                    "原票号应为6位日期加单据号",
                ),
                (
                    "原票号未匹配",
                    "5",
                    "notvalid",
                    "未找到日期与单据号组合键相同的原单",
                ),
            ],
        )
        self.assertEqual(stats["重复匹配键数量"], 1)
        self.assertEqual(stats["缺少匹配键数量"], 1)
        self.assertEqual(stats["原票号格式异常数量"], 1)

    def test_duplicate_match_keys_are_not_reported_as_issues(self) -> None:
        """Same 单据号/日期 with two line items is a normal suite sale (e.g.
        烟机+灶具 sold together), not a data problem — it still counts
        towards 重复匹配键数量 for diagnostics, but is neither highlighted
        nor written to 问题明细."""
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "老板-欧式烟机-K1L"],
            ["ZH0001", "2026-01-24", "", "", "老板-嵌入式灶-9B5-B1"],
        ]

        _output_rows, stats, issues = prepare(kept_rows)

        self.assertEqual(issues, [])
        self.assertEqual(stats["重复匹配键数量"], 1)

    def test_original_invoice_from_before_the_file_is_not_reported(self) -> None:
        """收款单统计 only ever covers one year's receipts, so an 原票号
        pointing at an earlier year can never resolve to a match_key in this
        file by construction — that is not a data problem either."""
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "250101ZH9999", "", "海尔冰箱"],
        ]

        output_rows, stats, issues = prepare(kept_rows)

        self.assertEqual(
            output_rows[0][-1],
            receipts.RECEIPTS_REMARK_RETURN,
        )
        self.assertEqual(issues, [])
        self.assertEqual(stats["未匹配原票号数量"], 0)

    def test_original_invoice_from_the_same_year_is_still_reported(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "260101ZH9999", "", "海尔冰箱"],
        ]

        _output_rows, stats, issues = prepare(kept_rows)

        self.assertEqual(
            issues,
            [
                (
                    "原票号未匹配",
                    "2",
                    "260101ZH9999",
                    "未找到日期与单据号组合键相同的原单",
                )
            ],
        )
        self.assertEqual(stats["未匹配原票号数量"], 1)

    def test_no_issues_yields_empty_list(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
        ]

        _output_rows, _stats, issues = prepare(kept_rows)

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

        output_rows, stats, issues = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["总数据量"], 1)
        self.assertEqual(stats["跳过空白行数"], 2)
        self.assertEqual(issues, [])

    def test_row_with_any_value_is_kept(self) -> None:
        kept_rows = [
            HEADER,
            [None, None, None, None, "海尔冰箱"],
        ]

        output_rows, stats, issues = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["跳过空白行数"], 0)
        # Still incomplete data, so it must keep reporting itself.
        self.assertEqual([issue[0] for issue in issues], ["缺少匹配键"])

    def test_numeric_zero_is_not_blank(self) -> None:
        kept_rows = [
            HEADER,
            [0, None, None, None, None],
        ]

        output_rows, stats, _issues = prepare(kept_rows)

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

                output_rows, stats, _issues = prepare(kept_rows)

                self.assertEqual(len(output_rows), 1)
                self.assertEqual(stats["跳过合计行数"], 1)

    def test_合计_inside_text_fields_is_not_a_total_row(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "合计说明", "合计套装"],
        ]

        output_rows, stats, _issues = prepare(kept_rows)

        self.assertEqual(len(output_rows), 1)
        self.assertEqual(stats["跳过合计行数"], 0)

    def test_skipped_rows_shift_the_reported_output_row(self) -> None:
        kept_rows = [
            HEADER,
            ["ZH0001", "2026-01-24", "", "", "海尔冰箱"],
            [None, None, None, None, None],
            ["合计", "合计", "", "", ""],
            # Source Excel row 6, but the blank and 合计 rows above never
            # reach the output sheet — this is only its second data row,
            # so 问题明细 must report 3 (an operator's coordinate in
            # Sheet1), not 6 (its position in the raw import).
            ["ZH0002", None, "", "", "格力空调"],
        ]

        _output_rows, _stats, issues = prepare(kept_rows)

        self.assertEqual(
            issues,
            [("缺少匹配键", "3", "", "日期或单据号为空，无法生成匹配键")],
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
        output_rows, stats, _issues = prepare([HEADER, row])

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
        _output_rows, stats, _issues = prepare([HEADER, row, list(row)])

        self.assertEqual(stats["备注总数"], 2)
        self.assertEqual(stats["特殊备注数量"], 2)

    def test_special_key_outranks_other_remark_rules(self) -> None:
        row = self._special_key_row()[1]
        # Give it an 原票号 too: without the special rule this row would be
        # 退换货/倒票（退单）.
        row[2] = "260101ZH9999"
        output_rows, _stats, _issues = prepare([HEADER, row])

        self.assertEqual(output_rows[0][-1], receipts.RECEIPTS_REMARK_SPECIAL)

    def test_missing_special_keys_are_reported_not_fatal(self) -> None:
        _output_rows, stats, _issues = prepare(
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

        _output_rows, stats, _issues = prepare(kept_rows)

        self.assertEqual(stats["仅退单数量"], 1)
        self.assertEqual(stats["备注总数"], 1)
        self.assertEqual(stats["特殊备注数量"], 0)


class ReceiptOutputSortTest(unittest.TestCase):
    def test_remark_groups_follow_explicit_business_order(self) -> None:
        remarks = [
            receipts.RECEIPTS_REMARK_SPECIAL,
            receipts.RECEIPTS_REMARK_BOTH,
            receipts.RECEIPTS_REMARK_RETURN,
            receipts.RECEIPTS_REMARK_ORIGINAL,
            receipts.RECEIPTS_REMARK_SAME_MODEL_REPLACEMENT,
            None,
        ]

        self.assertEqual(
            sorted(
                remarks,
                key=lambda remark: receipts.receipt_output_sort_key(
                    remark,
                    date(2026, 1, 1),
                    "ZH0001",
                    "商品",
                ),
            ),
            [
                None,
                receipts.RECEIPTS_REMARK_SAME_MODEL_REPLACEMENT,
                receipts.RECEIPTS_REMARK_ORIGINAL,
                receipts.RECEIPTS_REMARK_RETURN,
                receipts.RECEIPTS_REMARK_BOTH,
                receipts.RECEIPTS_REMARK_SPECIAL,
            ],
        )

    def test_rows_are_sorted_by_remark_date_document_and_product(self) -> None:
        output_rows, _stats, issues = prepare(
            [
                HEADER,
                ["ZH0002", "2026-01-02", "", "", "B商品"],
                ["ZH0004", "2026-01-01", "250101OLD4", "", "D退货"],
                ["ZH0001", "2026-01-01", "", "", "Z商品"],
                ["ZH0001", "2026-01-01", "", "", "A商品"],
                ["ZH0003", "2026-01-01", "250101OLD3", "", "C退货"],
            ]
        )

        self.assertEqual(issues, [])
        self.assertEqual(
            [(row[0], row[4], row[5]) for row in output_rows],
            [
                ("ZH0001", "A商品", None),
                ("ZH0001", "Z商品", None),
                ("ZH0002", "B商品", None),
                ("ZH0003", "C退货", receipts.RECEIPTS_REMARK_RETURN),
                ("ZH0004", "D退货", receipts.RECEIPTS_REMARK_RETURN),
            ],
        )

    def test_issue_row_number_uses_the_sorted_output_position(self) -> None:
        output_rows, _stats, issues = prepare(
            [
                HEADER,
                ["ZH0009", None, "", "", "缺少日期"],
                ["ZH0001", "2026-01-01", "", "", "正常商品"],
            ]
        )

        self.assertEqual([row[0] for row in output_rows], ["ZH0001", "ZH0009"])
        self.assertEqual(
            issues,
            [("缺少匹配键", "3", "", "日期或单据号为空，无法生成匹配键")],
        )

    def test_missing_date_sorts_last_inside_a_remark_group(self) -> None:
        output_rows, _stats, issues = prepare(
            [
                HEADER,
                ["ZH0009", None, "250101OLD9", "", "缺少日期退货"],
                ["ZH0001", "2026-01-01", "250101OLD1", "", "正常日期退货"],
            ]
        )

        self.assertEqual([row[0] for row in output_rows], ["ZH0001", "ZH0009"])
        self.assertEqual(
            [row[5] for row in output_rows],
            [receipts.RECEIPTS_REMARK_RETURN] * 2,
        )
        self.assertEqual(
            issues,
            [("缺少匹配键", "3", "", "日期或单据号为空，无法生成匹配键")],
        )

    def test_validation_rejects_unsorted_output(self) -> None:
        output_rows, _stats, issues = prepare(
            [
                HEADER,
                ["ZH0002", "2026-01-02", "", "", "B商品"],
                ["ZH0001", "2026-01-01", "", "", "A商品"],
            ]
        )
        output_rows.reverse()

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "收款单统计.xlsx"
            receipts._write_receipts_workbook(path, output_rows, issues)

            with self.assertRaisesRegex(RuntimeError, "收款单排序校验失败"):
                receipts.validate_receipts_output(path, len(output_rows), issues)


class ReceiptRemarkFillTest(unittest.TestCase):
    @staticmethod
    def _is_pink(cell) -> bool:
        return (
            cell.fill.patternType == "solid"
            and str(cell.fill.fgColor.rgb)[-6:]
            == receipts.RECEIPTS_REMARK_FILL_COLOR
        )

    def _write_and_read(self, kept_rows):
        output_rows, _stats, issues = prepare(kept_rows)
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "收款单统计.xlsx"
        receipts._write_receipts_workbook(path, output_rows, issues)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        receipts.validate_receipts_output(path, len(output_rows), issues)
        return output_rows, issues, workbook

    def test_suite_sale_duplicate_without_return_is_not_filled(self) -> None:
        output_rows, issues, workbook = self._write_and_read(
            [
                HEADER,
                ["ZH0001", "2026-01-24", "", "", "老板-欧式烟机-K1L"],
                ["ZH0001", "2026-01-24", "", "", "老板-嵌入式灶-9B5-B1"],
            ]
        )

        self.assertEqual([row[-1] for row in output_rows], [None, None])
        self.assertEqual(issues, [])
        self.assertEqual(workbook.sheetnames, ["Sheet1"])
        for row in workbook["Sheet1"].iter_rows(min_row=2):
            self.assertFalse(any(self._is_pink(cell) for cell in row))

    def test_return_referencing_suite_sale_marks_every_related_row(self) -> None:
        output_rows, issues, workbook = self._write_and_read(
            [
                HEADER,
                ["ZH0001", "2026-01-24", "", "", "老板-欧式烟机-K1L"],
                ["ZH0001", "2026-01-24", "", "", "老板-嵌入式灶-9B5-B1"],
                [
                    "ZH0002",
                    "2026-01-25",
                    "260124ZH0001",
                    "",
                    "老板烟灶套装退货",
                ],
            ]
        )

        self.assertEqual(
            [row[-1] for row in output_rows],
            [
                receipts.RECEIPTS_REMARK_ORIGINAL,
                receipts.RECEIPTS_REMARK_ORIGINAL,
                receipts.RECEIPTS_REMARK_RETURN,
            ],
        )
        self.assertEqual(issues, [])
        for row in workbook["Sheet1"].iter_rows(min_row=2):
            self.assertTrue(all(self._is_pink(cell) for cell in row))

    def test_prior_year_original_invoice_is_return_and_is_filled(self) -> None:
        output_rows, issues, workbook = self._write_and_read(
            [
                HEADER,
                [
                    "ZH0001",
                    "2026-01-24",
                    "250101ZH9999",
                    "",
                    "海尔冰箱",
                ],
            ]
        )

        self.assertEqual(
            output_rows[0][-1],
            receipts.RECEIPTS_REMARK_RETURN,
        )
        self.assertEqual(issues, [])
        self.assertEqual(workbook.sheetnames, ["Sheet1"])
        self.assertTrue(
            all(
                self._is_pink(cell)
                for cell in workbook["Sheet1"][2]
            )
        )


class IssuesSheetRoundTripTest(unittest.TestCase):
    def test_issues_sheet_is_written_and_validated(self) -> None:
        issues = [
            ("缺少匹配键", "5", "", "日期或单据号为空，无法生成匹配键"),
        ]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "收款单统计.xlsx"
            receipts._write_receipts_workbook(
                path,
                [],
                issues,
            )

            workbook = load_workbook(path)
            try:
                self.assertEqual(
                    workbook.sheetnames,
                    ["Sheet1", receipts.ISSUES_SHEET_NAME],
                )
                issues_sheet = workbook[receipts.ISSUES_SHEET_NAME]
                self.assertEqual(
                    tuple(cell.value for cell in issues_sheet[1]),
                    receipts.ISSUES_HEADER,
                )
                self.assertEqual(
                    [
                        row
                        for row in issues_sheet.iter_rows(
                            min_row=2,
                            values_only=True,
                        )
                    ],
                    [("缺少匹配键", "5", None, "日期或单据号为空，无法生成匹配键")],
                )
            finally:
                workbook.close()

            receipts.validate_receipts_output(path, 0, issues)

    def test_validation_rejects_mismatched_issues(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "收款单统计.xlsx"
            receipts._write_receipts_workbook(
                path,
                [],
                [("缺少匹配键", "5", "", "错误说明")],
            )

            with self.assertRaisesRegex(RuntimeError, "问题明细工作表内容校验失败"):
                receipts.validate_receipts_output(
                    path,
                    0,
                    [("缺少匹配键", "5", "", "说明")],
                )


if __name__ == "__main__":
    unittest.main()
