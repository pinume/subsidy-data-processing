import unittest

from processors import digital, large_appliances


def summary_row(
    processor: object,
    *,
    category: str,
    brand: str,
    remark: str,
    detail: str,
    subsidy: object,
) -> list[object]:
    header = processor.COUPON_OUTPUT_HEADER
    values = {
        "财务大类": category,
        "品牌": brand,
        "备注": remark,
        "详细情况": detail,
        processor.COUPON_OUTPUT_HEADER[6]: subsidy,
    }
    return [values.get(column) for column in header]


class CouponSummaryTest(unittest.TestCase):
    def test_digital_summary_uses_declared_counts_and_amounts(self) -> None:
        rows = [
            list(digital.COUPON_OUTPUT_HEADER),
            summary_row(
                digital,
                category="数码",
                brand="A",
                remark="已上传",
                detail="审核通过",
                subsidy="10.10",
            ),
            summary_row(
                digital,
                category="数码",
                brand="B",
                remark="未上传",
                detail="",
                subsidy=20,
            ),
            summary_row(
                digital,
                category="数码",
                brand="C",
                remark="已上传",
                detail="底部排除行",
                subsidy=999,
            ),
        ]

        result = digital.build_coupon_summary(rows, excluded_bottom_rows=1)

        self.assertEqual(
            result,
            [
                ("已上传", 1, 10.1),
                ("未上传", 1, 20.0),
                ("合计", 2, 30.1),
            ],
        )

    def test_large_appliance_summaries_have_independent_expectations(self) -> None:
        rows = [
            list(large_appliances.COUPON_OUTPUT_HEADER),
            summary_row(
                large_appliances,
                category="冰箱",
                brand="海尔",
                remark="已上传",
                detail="状态：审核通过",
                subsidy="10.11",
            ),
            summary_row(
                large_appliances,
                category="冰箱",
                brand="海尔",
                remark="已上传",
                detail="状态：待审核",
                subsidy="20.20",
            ),
            summary_row(
                large_appliances,
                category="空调",
                brand="格力",
                remark="未上传",
                detail="",
                subsidy=None,
            ),
            summary_row(
                large_appliances,
                category="底部",
                brand="排除",
                remark="已上传",
                detail="审核通过",
                subsidy=999,
            ),
        ]

        summary, approved, remarks = large_appliances.build_coupon_summary(
            rows,
            excluded_bottom_rows=1,
        )

        self.assertEqual(
            summary,
            [
                ("冰箱", "海尔", "已上传", 2, 30.31),
                ("空调", "格力", "未上传", 1, 0.0),
                ("合计", None, None, 3, 30.31),
            ],
        )
        self.assertEqual(
            approved,
            [
                ("冰箱", "海尔", 1, 10.11),
                ("合计", None, 1, 10.11),
            ],
        )
        self.assertEqual(
            remarks,
            [
                ("已上传", 2, 30.31),
                ("未上传", 1, 0.0),
                ("合计", 3, 30.31),
            ],
        )

    def test_invalid_subsidy_is_rejected(self) -> None:
        for processor in (digital, large_appliances):
            with self.subTest(processor=processor.__name__):
                rows = [
                    list(processor.COUPON_OUTPUT_HEADER),
                    summary_row(
                        processor,
                        category="冰箱",
                        brand="海尔",
                        remark="已上传",
                        detail="",
                        subsidy="非数字",
                    ),
                ]

                with self.assertRaisesRegex(ValueError, "国补金额无效"):
                    processor.build_coupon_summary(
                        rows,
                        excluded_bottom_rows=0,
                    )


if __name__ == "__main__":
    unittest.main()
