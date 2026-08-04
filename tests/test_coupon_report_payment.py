import tempfile
import unittest
from collections import Counter
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from processors import coupon_report
from processors.coupons import appliance


class CouponPaymentEnrichmentTest(unittest.TestCase):
    def test_unknown_brand_does_not_receive_project_total(self) -> None:
        payment_summary = {
            ("冰箱", "海尔"): (Decimal("100"), 2),
            ("电视", "TCL"): (Decimal("50"), 1),
        }
        rows = [
            ("冰箱", "", 1, 20.0),
            ("家电合计", "已上传", 0, 0.0),
            (None, "未上传", 1, 20.0),
            (None, "合计", 1, 20.0),
        ]

        enriched = coupon_report.enrich_summary_rows_with_payment(
            rows,
            payment_summary,
            "家电",
        )

        self.assertEqual(enriched[0], ("冰箱", "", 1, 20.0, None, None))
        self.assertEqual(enriched[-1][-2:], (150.0, 3))

    def test_all_configured_midea_group_brands_resolve(self) -> None:
        payment_summary = {("冰箱", "美的系"): (Decimal("100"), 2)}

        for brand in ("美的", "小天鹅", "东芝"):
            with self.subTest(brand=brand):
                self.assertEqual(
                    coupon_report._resolve_payment_data(
                        "冰箱",
                        brand,
                        payment_summary,
                        "家电",
                    ),
                    (Decimal("100"), 2),
                )

    def test_midea_group_payment_is_not_added_as_a_duplicate_brand(self) -> None:
        rows = [
            ("冰箱", "东芝", 1, 0.0),
            ("家电合计", "已上传", 0, 0.0),
            (None, "未上传", 1, 0.0),
            (None, "合计", 1, 0.0),
        ]
        payment_summary = {("冰箱", "美的系"): (Decimal("100"), 2)}

        enriched = coupon_report.enrich_summary_rows_with_payment(
            rows,
            payment_summary,
            "家电",
        )

        self.assertEqual(
            [row[:2] for row in enriched],
            [("冰箱", "东芝"), ("家电合计", "已上传"), (None, "未上传"), (None, "合计")],
        )
        self.assertEqual(enriched[0][-2:], (100.0, 2))


class PaymentSummaryLoadingTest(unittest.TestCase):
    def test_missing_summary_sheet_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "回款明细.xlsx"
            workbook = Workbook()
            workbook.save(path)
            workbook.close()

            with self.assertRaisesRegex(ValueError, "缺少.*汇总"):
                coupon_report.load_payment_summary(path)


class ExcludedCategoryValidationTest(unittest.TestCase):
    def test_unmatched_excluded_category_is_not_required_in_brand_summary(self) -> None:
        header = appliance.COUPON_OUTPUT_HEADER
        row = [None] * len(header)
        row[0] = "单据001"
        row[1] = date(2026, 1, 1)
        row[appliance.COUPON_CATEGORY_INDEX] = "数码"
        row[appliance.COUPON_BRAND_INDEX] = "未知品牌"
        row[appliance.COUPON_REMARK_INDEX] = "未上传"
        row[appliance.COUPON_SUBSIDY_INDEX] = 10
        summary_rows, zero_subsidy_count = appliance.build_coupon_summary(
            [list(header), row],
            excluded_bottom_rows=0,
            uploaded_subsidy_count=0,
            uploaded_subsidy_total=Decimal("0"),
        )
        computation = appliance.CouponComputation(
            rows=[list(header), row],
            data_row_count=1,
            matched_count=0,
            matched_subsidy_total=Decimal("0"),
            receipt_remark_count=0,
            remark_lookup={},
            detail_lookup={},
            reference_universe=set(),
            payment_references=frozenset(),
            payment_match_count=0,
            reference_supplement_count=0,
            ambiguous_reference_supplement_count=0,
            reference_supplement_matches=Counter(),
            corrected_count=0,
            unresolved_count=0,
            correction_collision_count=0,
            reference_decisions=[],
            final_unresolved_reference_count=0,
            uploaded_count=0,
            unmatched_count=1,
            uploaded_subsidy_count=0,
            uploaded_subsidy_total=Decimal("0"),
            zero_subsidy_count=zero_subsidy_count,
            source_total=None,
            computed_total=Decimal("0"),
            summary_rows=summary_rows,
            group_sheets=[],
        )

        appliance.validate_computation(computation, [])


if __name__ == "__main__":
    unittest.main()
