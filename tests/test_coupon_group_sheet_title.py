"""Group sheet names are generated from data, and XlsxWriter rejects bad ones.

The 26 group sheets are named 财务大类-品牌, both of which come from the source
export. openpyxl merely warned about a name over 31 characters and wrote the
file anyway; XlsxWriter raises InvalidWorksheetName and takes the whole 审核明细
mode down with it. coupon_group_sheet_title is now the only thing standing
between a long brand name and a failed run, so what it guarantees is pinned
here.
"""

import tempfile
import unittest
from pathlib import Path

from xlsxwriter import Workbook

from processors.coupons.appliance import coupon_group_sheet_title

EXCEL_TITLE_LIMIT = 31


class GroupSheetTitleTest(unittest.TestCase):
    def test_long_names_are_cut_to_excels_limit(self) -> None:
        title = coupon_group_sheet_title("空调", "很长的品牌名称" * 5, set())
        self.assertLessEqual(len(title), EXCEL_TITLE_LIMIT)

    def test_characters_excel_forbids_are_replaced(self) -> None:
        title = coupon_group_sheet_title("冰[箱]", "海:尔*?/\\", set())
        self.assertFalse(set(title) & set("[]:*?/\\"))

    def test_empty_category_or_brand_still_yields_a_name(self) -> None:
        self.assertTrue(coupon_group_sheet_title("", "", set()))

    def test_a_clash_after_truncation_gets_a_suffix(self) -> None:
        """Two brands sharing their first 31 characters must not collide."""
        used: set[str] = set()
        first = coupon_group_sheet_title("空调", "很长的品牌名称" * 5, used)
        second = coupon_group_sheet_title("空调", "很长的品牌名称" * 5, used)

        self.assertNotEqual(first, second)
        self.assertLessEqual(len(second), EXCEL_TITLE_LIMIT)

    def test_a_case_only_clash_gets_a_suffix(self) -> None:
        """Excel and XlsxWriter compare worksheet names case-insensitively."""
        used: set[str] = set()
        first = coupon_group_sheet_title("空调", "TCL", used)
        second = coupon_group_sheet_title("空调", "tcl", used)

        self.assertNotEqual(first.casefold(), second.casefold())

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "case-only-clash.xlsx"
            with Workbook(str(path)) as workbook:
                workbook.add_worksheet(first)
                workbook.add_worksheet(second)

    def test_every_generated_name_is_one_xlsxwriter_accepts(self) -> None:
        """The guarantee that matters: XlsxWriter must not reject the result."""
        used: set[str] = set()
        names = []
        for category, brand in (
            ("空调", "很长的品牌名称" * 5),
            ("空调", "很长的品牌名称" * 5),  # collides with the one above
            ("冰[箱]", "海:尔*?/\\"),
            ("", ""),
            ("厨卫", "A.O.史密斯"),
        ):
            name = coupon_group_sheet_title(category, brand, used)
            names.append(name)

        self.assertEqual(len(set(names)), len(names))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "titles.xlsx"
            with Workbook(str(path)) as workbook:
                for name in names:
                    workbook.add_worksheet(name)
