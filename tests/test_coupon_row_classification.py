import unittest
from decimal import Decimal

from processors.coupons import sources
from processors.coupons.sources import classify_coupon_row, classify_subsidy_attribution


class CouponRowClassificationTest(unittest.TestCase):
    def test_appliance_subsidy_only_classifies_as_family(self) -> None:
        self.assertEqual(
            classify_coupon_row(
                appliance_subsidy=483.6,
                digital_subsidy=0,
                row_number=3,
                source_name="销售用券情况统计.XLS",
            ),
            "家电",
        )

    def test_digital_subsidy_only_classifies_as_digital(self) -> None:
        self.assertEqual(
            classify_coupon_row(
                appliance_subsidy=0,
                digital_subsidy=299.85,
                row_number=4,
                source_name="销售用券情况统计.XLS",
            ),
            "数码",
        )

    def test_neither_subsidy_populated_defaults_to_family(self) -> None:
        self.assertEqual(
            classify_coupon_row(
                appliance_subsidy=0,
                digital_subsidy=None,
                row_number=5,
                source_name="销售用券情况统计.XLS",
            ),
            "家电",
        )

    def test_both_subsidies_populated_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "第 7 行.*家电国补.*数码国补"):
            classify_coupon_row(
                appliance_subsidy=100,
                digital_subsidy=50,
                row_number=7,
                source_name="销售用券情况统计.XLS",
            )

    def test_known_appliance_categories_override_wrong_digital_column(self) -> None:
        for category in ("冰箱", "厨卫", "国产彩电", "空调", "洗衣机"):
            with self.subTest(category=category):
                self.assertEqual(
                    classify_coupon_row(
                        appliance_subsidy=0,
                        digital_subsidy=100,
                        row_number=8,
                        source_name="销售用券情况统计.XLSX",
                        financial_category=category,
                    ),
                    "家电",
                )

    def test_known_digital_categories_override_wrong_appliance_column(self) -> None:
        for category in ("数码", "新业务类"):
            with self.subTest(category=category):
                self.assertEqual(
                    classify_coupon_row(
                        appliance_subsidy=100,
                        digital_subsidy=0,
                        row_number=9,
                        source_name="销售用券情况统计.XLSX",
                        financial_category=category,
                    ),
                    "数码",
                )

    def test_receipt_remark_preserves_original_subsidy_column(self) -> None:
        self.assertEqual(
            classify_coupon_row(
                appliance_subsidy=0,
                digital_subsidy=100,
                row_number=10,
                source_name="销售用券情况统计.XLSX",
                financial_category="冰箱",
                has_receipt_remark=True,
            ),
            "数码",
        )


class SubsidyAttributionTest(unittest.TestCase):
    """classify_subsidy_attribution decides the move; it reads nothing,
    prints nothing, and records every non-zero move as a correction."""

    def attribution(self, **overrides) -> sources.SubsidyAttribution:
        kwargs = dict(
            appliance_subsidy=0,
            digital_subsidy=0,
            row_number=3,
            source_name="销售用券情况统计.xlsx",
            document_number="001",
        )
        kwargs.update(overrides)
        return classify_subsidy_attribution(**kwargs)

    def test_all_five_appliance_categories_move_digital_amount_into_family(
        self,
    ) -> None:
        for category in ("冰箱", "厨卫", "国产彩电", "空调", "洗衣机"):
            with self.subTest(category=category):
                attribution = self.attribution(
                    digital_subsidy=100,
                    financial_category=category,
                )
                self.assertEqual(attribution.classification, "家电")
                self.assertEqual(attribution.subsidy_value, 100)
                self.assertEqual(
                    attribution.source_total_adjustment,
                    Decimal("100"),
                )
                self.assertIsNotNone(attribution.correction)
                correction = attribution.correction
                self.assertEqual(correction.row_number, 3)
                self.assertEqual(correction.document_number, "001")
                self.assertEqual(correction.financial_category, category)
                self.assertEqual(correction.amount, Decimal("100"))
                self.assertEqual(
                    correction.from_header,
                    sources.COUPON_DIGITAL_SUBSIDY_HEADER,
                )
                self.assertEqual(
                    correction.to_header,
                    sources.COUPON_FAMILY_SUBSIDY_HEADER,
                )

    def test_digital_categories_move_family_amount_into_digital_column(
        self,
    ) -> None:
        for category in ("数码", "新业务类"):
            with self.subTest(category=category):
                attribution = self.attribution(
                    appliance_subsidy=50,
                    financial_category=category,
                )
                self.assertEqual(attribution.classification, "数码")
                self.assertEqual(attribution.subsidy_value, 50)
                self.assertEqual(
                    attribution.source_total_adjustment,
                    Decimal("-50"),
                )
                self.assertIsNotNone(attribution.correction)
                correction = attribution.correction
                self.assertEqual(correction.amount, Decimal("50"))
                self.assertEqual(
                    correction.from_header,
                    sources.COUPON_FAMILY_SUBSIDY_HEADER,
                )
                self.assertEqual(
                    correction.to_header,
                    sources.COUPON_DIGITAL_SUBSIDY_HEADER,
                )

    def test_column_based_rows_need_no_correction(self) -> None:
        for kwargs, expected in (
            (dict(appliance_subsidy=100), "家电"),
            (dict(digital_subsidy=50), "数码"),
            (dict(), "家电"),
        ):
            with self.subTest(kwargs=kwargs):
                attribution = self.attribution(**kwargs)
                self.assertEqual(attribution.classification, expected)
                self.assertEqual(
                    attribution.source_total_adjustment,
                    Decimal("0"),
                )
                self.assertIsNone(attribution.correction)

    def test_receipt_remark_prevents_the_move(self) -> None:
        attribution = self.attribution(
            digital_subsidy=100,
            financial_category="冰箱",
            has_receipt_remark=True,
        )
        self.assertEqual(attribution.classification, "数码")
        self.assertEqual(attribution.subsidy_value, 100)
        self.assertEqual(attribution.source_total_adjustment, Decimal("0"))
        self.assertIsNone(attribution.correction)

    def test_both_columns_populated_still_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "第 3 行.*家电国补.*数码国补"):
            self.attribution(
                appliance_subsidy=100,
                digital_subsidy=50,
                financial_category="冰箱",
            )

    def test_blank_moved_amount_records_nothing(self) -> None:
        attribution = self.attribution(
            appliance_subsidy=None,
            financial_category="新业务类",
        )
        self.assertEqual(attribution.classification, "数码")
        self.assertIsNone(attribution.subsidy_value)
        self.assertEqual(attribution.source_total_adjustment, Decimal("0"))
        self.assertIsNone(attribution.correction)

    def test_zero_moved_amount_records_nothing(self) -> None:
        attribution = self.attribution(
            appliance_subsidy=0,
            financial_category="新业务类",
        )
        self.assertEqual(attribution.classification, "数码")
        self.assertEqual(attribution.source_total_adjustment, Decimal("0"))
        self.assertIsNone(attribution.correction)

    def test_negative_moved_amount_keeps_its_sign(self) -> None:
        attribution = self.attribution(
            appliance_subsidy=-50,
            financial_category="新业务类",
        )
        self.assertEqual(attribution.classification, "数码")
        self.assertEqual(attribution.subsidy_value, -50)
        # Moving a negative amount out of 家电's column raises its total.
        self.assertEqual(
            attribution.source_total_adjustment,
            Decimal("50"),
        )
        self.assertEqual(attribution.correction.amount, Decimal("-50"))

    def test_unparsable_moved_amount_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "第 3 行补贴金额无效"):
            self.attribution(
                appliance_subsidy="abc",
                financial_category="新业务类",
            )


if __name__ == "__main__":
    unittest.main()
