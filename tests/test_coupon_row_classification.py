import unittest

from processors.coupons.sources import classify_coupon_row


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


if __name__ == "__main__":
    unittest.main()
