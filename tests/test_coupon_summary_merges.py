"""The merge ranges must be computable without a worksheet to read back from.

merge_coupon_summary_groups used to walk the built sheet with sheet.cell(),
comparing each row against the one below it. XlsxWriter cannot be read back
that way, so the ranges are now derived from the summary rows themselves and
the openpyxl writer consumes the same result. These tests pin the ranges, so
the two writers cannot drift apart on which cells get merged.
"""

import unittest

from openpyxl import Workbook

from processors.coupons.appliance import (
    coupon_summary_group_merges,
    coupon_summary_project_merges,
    merge_coupon_summary_groups,
    project_summary_blocks,
)

# 财务大类, 品牌, then three columns the merging ignores.
SUMMARY_ROWS = [
    ("空调", "格力", 1, 2, 3.0),
    ("空调", "格力", 1, 2, 3.0),
    ("空调", "美的", 1, 2, 3.0),
    ("冰箱", "海尔", 1, 2, 3.0),
    ("冰箱", "海尔", 1, 2, 3.0),
    ("冰箱", "海尔", 1, 2, 3.0),
    ("家电", None, 1, 2, 3.0),
    ("家电", None, 1, 2, 3.0),
]


class GroupMergeRangesTest(unittest.TestCase):
    def test_runs_longer_than_one_row_become_ranges(self) -> None:
        merges = coupon_summary_group_merges(SUMMARY_ROWS, 6)

        # Sheet rows are 1-based with one header row, so data row 0 is row 2.
        self.assertEqual(
            sorted(merges),
            sorted(
                [
                    (2, 3, 2),  # 格力 spans two rows
                    (5, 7, 2),  # 海尔 spans three
                    (2, 4, 1),  # 空调 spans three
                    (5, 7, 1),  # 冰箱 spans three
                ]
            ),
        )

    def test_a_single_row_run_produces_no_range(self) -> None:
        """美的 occupies one row; merging it would be a one-cell merge."""
        merges = coupon_summary_group_merges(SUMMARY_ROWS, 6)
        self.assertNotIn((4, 4, 2), merges)

    def test_brand_runs_do_not_span_a_category_change(self) -> None:
        rows = [("空调", "海尔", 0, 0, 0), ("冰箱", "海尔", 0, 0, 0)]
        self.assertEqual(coupon_summary_group_merges(rows, 2), [])

    def test_row_count_stops_before_the_project_blocks(self) -> None:
        """The trailing 家电 rows are merged as blocks, not as category runs."""
        merges = coupon_summary_group_merges(SUMMARY_ROWS, 6)
        self.assertTrue(all(last <= 7 for _first, last, _column in merges))


class ProjectMergeRangesTest(unittest.TestCase):
    def test_blocks_become_two_column_ranges(self) -> None:
        blocks = project_summary_blocks(SUMMARY_ROWS)
        self.assertEqual(blocks, [(6, 7)])
        self.assertEqual(coupon_summary_project_merges(blocks), [(8, 9, 1, 2)])


class OpenpyxlWriterUsesTheSameRangesTest(unittest.TestCase):
    def test_written_merges_equal_the_computed_ranges(self) -> None:
        """The openpyxl path must not diverge from what the pure function says."""
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["财务大类", "品牌", "a", "b", "c"])
        for row in SUMMARY_ROWS:
            sheet.append(list(row))

        merge_coupon_summary_groups(sheet, SUMMARY_ROWS, 6)

        expected = {
            f"{chr(ord('A') + column - 1)}{first}"
            f":{chr(ord('A') + column - 1)}{last}"
            for first, last, column in coupon_summary_group_merges(SUMMARY_ROWS, 6)
        }
        self.assertEqual({str(r) for r in sheet.merged_cells.ranges}, expected)
        workbook.close()
