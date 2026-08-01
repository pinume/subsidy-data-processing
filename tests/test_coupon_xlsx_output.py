"""The XlsxWriter audit writers must produce what the openpyxl ones did.

Each test states the contract as a fact about the written file, read back with
openpyxl. The rules were not inferred from the writer code — they were read off
the audit output the openpyxl writers actually produce, because several of them
are positional rather than named (数据汇总's currency column is the last one,
and its header differs per project).
"""

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from xlsxwriter import Workbook

from processors.common.excel import load_measurement_font, resolve_font
from processors.coupons.xlsx_output import (
    CouponFormatCache,
    FormatKey,
    write_detail_sheet,
    write_summary_sheet,
)

FILL_COLOR = "FFC7CE"
DETAIL_HEADER = ("单据号", "单据日期", "商品名称", "备注", "详细情况")
SUMMARY_HEADER = ("财务大类", "品牌", "备注", "数量", "国补合计")


class WriterContractTest(unittest.TestCase):
    def setUp(self) -> None:
        self._directory = tempfile.TemporaryDirectory()
        self.path = Path(self._directory.name) / "audit.xlsx"
        self.addCleanup(self._directory.cleanup)
        font_name, font_path = resolve_font()
        self.font_name = font_name
        self.measurement_font = load_measurement_font(font_path)

    def _read(self, name: str):
        return load_workbook(self.path)[name]

    def test_detail_sheet_formats_and_trailing_pink_block(self) -> None:
        rows = [[f"ZH{i:04d}", None, "海尔冰箱", "", ""] for i in range(5)]
        with Workbook(str(self.path)) as workbook:
            formats = CouponFormatCache(workbook, self.font_name, FILL_COLOR)
            write_detail_sheet(
                workbook,
                "明细",
                DETAIL_HEADER,
                rows,
                formats,
                self.measurement_font,
                left_aligned_headers=("商品名称", "详细情况"),
                matched_count=2,
            )
        sheet = self._read("明细")

        self.assertEqual(sheet["A2"].number_format, "@")
        # 单据日期 is empty here, so it keeps the general format — the openpyxl
        # writer guarded the date format on the cell having a value.
        self.assertEqual(sheet["B2"].number_format, "General")
        self.assertEqual(sheet["C2"].alignment.horizontal, "left")
        self.assertEqual(sheet["E2"].alignment.horizontal, "left")
        self.assertEqual(sheet["D2"].alignment.horizontal, "center")

        def pink(row: int) -> bool:
            cell = sheet.cell(row, 1)
            return (
                cell.fill.patternType == "solid"
                and str(cell.fill.fgColor.rgb)[-6:] == FILL_COLOR
            )

        # 5 data rows in sheet rows 2..6; the last 2 are the matched block.
        self.assertEqual([pink(r) for r in range(2, 7)], [False, False, False, True, True])

    def test_detail_sheet_with_no_matches_has_no_pink(self) -> None:
        rows = [["ZH0001", None, "海尔冰箱", "", ""]]
        with Workbook(str(self.path)) as workbook:
            formats = CouponFormatCache(workbook, self.font_name, FILL_COLOR)
            write_detail_sheet(
                workbook,
                "明细",
                DETAIL_HEADER,
                rows,
                formats,
                self.measurement_font,
                left_aligned_headers=(),
                matched_count=0,
            )
        self.assertIsNone(self._read("明细")["A2"].fill.patternType)

    def test_summary_sheet_borders_currency_and_merges(self) -> None:
        rows = [
            ("空调", "格力", "已上传", 1, 1.5),
            ("空调", "格力", "未上传", 2, 2.5),
            ("冰箱", "海尔", "已上传", 3, 3.5),
        ]
        with Workbook(str(self.path)) as workbook:
            formats = CouponFormatCache(workbook, self.font_name, FILL_COLOR)
            write_summary_sheet(
                workbook,
                "数据汇总",
                SUMMARY_HEADER,
                rows,
                formats,
                self.measurement_font,
                group_merges=[(2, 3, 1), (2, 3, 2)],
                project_merges=[],
            )
        sheet = self._read("数据汇总")

        # The currency column is the last one, by position, and unconditional.
        self.assertEqual(sheet["E2"].number_format, "0.00")
        self.assertEqual(sheet["E4"].number_format, "0.00")
        self.assertEqual(sheet["A2"].number_format, "General")
        # Borders cover the header row and every data row.
        for coordinate in ("A1", "E1", "A4", "E4"):
            with self.subTest(coordinate=coordinate):
                self.assertEqual(sheet[coordinate].border.left.style, "thin")
        self.assertEqual(
            {str(r) for r in sheet.merged_cells.ranges}, {"A2:A3", "B2:B3"}
        )
        self.assertEqual(sheet["A2"].value, "空调")

    def test_format_cache_reuses_one_format_per_distinct_key(self) -> None:
        with Workbook(str(self.path)) as workbook:
            formats = CouponFormatCache(workbook, self.font_name, FILL_COLOR)
            first = formats.get(FormatKey(align="left", pink=True))
            second = formats.get(FormatKey(align="left", pink=True))
            third = formats.get(FormatKey(align="left"))
            self.assertIs(first, second)
            self.assertIsNot(first, third)
            workbook.add_worksheet("x")
