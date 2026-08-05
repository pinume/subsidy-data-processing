"""Row-layout index derivation in processors/coupons/matching.py.

The indexes are computed from the two profiles' output headers at import
time, so a layout drift between the two profiles must fail loudly instead
of silently misreading every row. These tests pin the resolved positions
and prove the divergence detection.
"""

import unittest
from types import SimpleNamespace
from unittest.mock import patch

from processors.coupons import matching, sources


def profile(output_header: tuple[str, ...], name: str = "家电"):
    return SimpleNamespace(name=name, output_header=output_header)


class SharedIndexTests(unittest.TestCase):
    def test_indexes_resolve_the_documented_positions(self) -> None:
        self.assertEqual(matching.DOCUMENT_INDEX, 0)
        self.assertEqual(matching.DATE_INDEX, 1)
        self.assertEqual(matching.SUMMARY_INDEX, 5)
        self.assertEqual(matching.SUBSIDY_INDEX, 6)
        self.assertEqual(matching.REMARK_INDEX, 7)
        self.assertEqual(matching.DETAIL_INDEX, 8)
        self.assertEqual(matching.PAYMENT_STATUS_INDEX, 9)

    def test_shared_index_resolves_each_label_to_its_field(self) -> None:
        for label, expected in (
            ("单据号", 0),
            ("单据日期", 1),
            ("明细摘要", 5),
            ("备注", 7),
            ("详细情况", 8),
            ("回款情况", 9),
        ):
            with self.subTest(label=label):
                self.assertEqual(matching.shared_index(label), expected)

    def test_shared_index_rejects_a_divergent_position(self) -> None:
        divergent = profile(
            (
                "单据号",
                "单据日期",
                "商品名称",
                "品牌",
                "财务大类",
                "备注",
                "明细摘要",
                "X",
                "Y",
                "Z",
            )
        )
        with patch.object(sources, "APPLIANCE_PROFILE", divergent):
            with self.assertRaisesRegex(RuntimeError, "位置不一致"):
                matching.shared_index("明细摘要")

    def test_shared_index_rejects_a_missing_field(self) -> None:
        with patch.object(
            sources,
            "DIGITAL_PROFILE",
            profile(("单据号", "单据日期"), name="数码"),
        ):
            with self.assertRaisesRegex(RuntimeError, "缺失"):
                matching.shared_index("备注")

    def test_shared_profile_field_index_rejects_divergent_subsidy_position(
        self,
    ) -> None:
        divergent = profile(
            (
                "单据号",
                "单据日期",
                "商品名称",
                "品牌",
                "财务大类",
                "明细摘要",
                "备注",
                "2026数码国补（计入收入）",
                "详细情况",
                "回款情况",
            ),
            name="数码",
        )
        with patch.object(sources, "DIGITAL_PROFILE", divergent):
            with self.assertRaisesRegex(RuntimeError, "位置不一致"):
                matching.shared_profile_field_index(
                    sources.COUPON_FAMILY_SUBSIDY_HEADER,
                    sources.COUPON_DIGITAL_SUBSIDY_HEADER,
                )

    def test_shared_profile_field_index_rejects_a_missing_subsidy_field(
        self,
    ) -> None:
        with patch.object(
            sources,
            "DIGITAL_PROFILE",
            profile(
                (
                    "单据号",
                    "单据日期",
                    "商品名称",
                    "品牌",
                    "财务大类",
                    "明细摘要",
                    "备注",
                    "详细情况",
                    "回款情况",
                ),
                name="数码",
            ),
        ):
            with self.assertRaisesRegex(RuntimeError, "缺失"):
                matching.shared_profile_field_index(
                    sources.COUPON_FAMILY_SUBSIDY_HEADER,
                    sources.COUPON_DIGITAL_SUBSIDY_HEADER,
                )


if __name__ == "__main__":
    unittest.main()
