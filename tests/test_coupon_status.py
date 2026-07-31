import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from processors import digital, large_appliances


PROCESSORS = (digital, large_appliances)


def coupon_row(
    processor: object,
    reference: str,
    *,
    remark: str = "",
    detail: str = "",
) -> list[object]:
    header = processor.COUPON_OUTPUT_HEADER
    values = {
        "明细摘要": reference,
        "备注": remark,
        "详细情况": detail,
    }
    return [values.get(column) for column in header]


class CouponStatusLookupTest(unittest.TestCase):
    def test_uploaded_lookup_normalizes_reference_and_builds_detail(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "uploaded.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet.append(["检索参考号", "状态", "描述"])
            sheet.append([" 12345678901a ", " 已完成 ", "匹配成功 "])
            workbook.save(source)
            workbook.close()

            for processor in PROCESSORS:
                with self.subTest(processor=processor.__name__):
                    self.assertEqual(
                        processor.load_uploaded_detail_lookup(source),
                        {"12345678901A": "已完成：匹配成功"},
                    )

    def test_uploaded_lookup_rejects_malformed_reference(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "uploaded.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet.append(["检索参考号", "状态", "描述"])
            sheet.append(["1234567890A", "已完成", "匹配成功"])
            workbook.save(source)
            workbook.close()

            for processor in PROCESSORS:
                with self.subTest(processor=processor.__name__):
                    with self.assertRaisesRegex(
                        ValueError,
                        "11位数字后跟一个大写字母",
                    ):
                        processor.load_uploaded_detail_lookup(source)

    def test_uploaded_lookup_rejects_conflicting_duplicate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "uploaded.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet.append(["检索参考号", "状态", "描述"])
            sheet.append(["12345678901N", "已完成", "结果一"])
            sheet.append(["12345678901N", "已完成", "结果二"])
            workbook.save(source)
            workbook.close()

            for processor in PROCESSORS:
                with self.subTest(processor=processor.__name__):
                    with self.assertRaisesRegex(
                        ValueError,
                        "检索参考号存在冲突",
                    ):
                        processor.load_uploaded_detail_lookup(source)


class CouponStatusFillTest(unittest.TestCase):
    def test_uploaded_match_sets_detail_and_remark(self) -> None:
        reference = "12345678901N"
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                row = coupon_row(processor, "12345678901n")
                rows = [list(processor.COUPON_OUTPUT_HEADER), row]
                lookup = {reference: "已完成：匹配成功"}

                if processor is large_appliances:
                    count = processor.fill_uploaded_details(
                        rows,
                        lookup,
                        excluded_bottom_rows=0,
                    )
                else:
                    count = processor.fill_uploaded_details(rows, lookup)

                remark_index = processor.COUPON_OUTPUT_HEADER.index("备注")
                detail_index = processor.COUPON_OUTPUT_HEADER.index("详细情况")
                self.assertEqual(count, 1)
                self.assertEqual(row[remark_index], "已上传")
                self.assertEqual(row[detail_index], "已完成：匹配成功")

    def test_reference_outside_submitted_data_is_marked_unsubmitted(self) -> None:
        """Without unsubmitted data, anything not submitted is 未上传."""
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                row = coupon_row(processor, "99999999999Z")
                rows = [list(processor.COUPON_OUTPUT_HEADER), row]

                if processor is large_appliances:
                    count = processor.fill_unmatched_remarks(
                        rows,
                        {"12345678901N"},
                        excluded_bottom_rows=0,
                    )
                else:
                    count = processor.fill_unmatched_remarks(
                        rows,
                        {"12345678901N"},
                    )

                remark_index = processor.COUPON_OUTPUT_HEADER.index("备注")
                self.assertEqual(count, 1)
                self.assertEqual(row[remark_index], "未上传")

    def test_large_appliances_leaves_submitted_and_bottom_rows_alone(self) -> None:
        reference = "12345678901N"
        uploaded_row = coupon_row(
            large_appliances,
            reference,
            remark="已上传",
            detail="已完成：匹配成功",
        )
        bottom_row = coupon_row(large_appliances, "99999999999Z")
        rows = [
            list(large_appliances.COUPON_OUTPUT_HEADER),
            uploaded_row,
            bottom_row,
        ]

        count = large_appliances.fill_unmatched_remarks(
            rows,
            {reference},
            excluded_bottom_rows=1,
        )

        remark_index = large_appliances.COUPON_OUTPUT_HEADER.index("备注")
        self.assertEqual(count, 0)
        self.assertEqual(uploaded_row[remark_index], "已上传")
        self.assertEqual(bottom_row[remark_index], "")


if __name__ == "__main__":
    unittest.main()
