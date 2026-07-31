import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from processors.coupons.sources import load_uploaded_summary
from processors.coupons import appliance, matching
from processors.coupons import digital as coupons_digital


HEADERS = (coupons_digital.COUPON_OUTPUT_HEADER, appliance.COUPON_OUTPUT_HEADER)


def coupon_row(
    header: tuple[str, ...],
    reference: str,
    *,
    remark: str = "",
    detail: str = "",
) -> list[object]:
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
            sheet.append(["检索参考号", "状态", "描述", "补贴金额"])
            sheet.append([" 12345678901a ", " 已完成 ", "匹配成功 ", 1])
            workbook.save(source)
            workbook.close()

            self.assertEqual(
                load_uploaded_summary(source)[0],
                {"12345678901A": "已完成：匹配成功"},
            )

    def test_uploaded_lookup_rejects_malformed_reference(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "uploaded.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet.append(["检索参考号", "状态", "描述", "补贴金额"])
            sheet.append(["1234567890A", "已完成", "匹配成功", 1])
            workbook.save(source)
            workbook.close()

            with self.assertRaisesRegex(
                ValueError,
                "11位数字后跟一个大写字母",
            ):
                load_uploaded_summary(source)

    def test_uploaded_lookup_rejects_conflicting_duplicate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "uploaded.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Summary"
            sheet.append(["检索参考号", "状态", "描述", "补贴金额"])
            sheet.append(["12345678901N", "已完成", "结果一", 1])
            sheet.append(["12345678901N", "已完成", "结果二", 1])
            workbook.save(source)
            workbook.close()

            with self.assertRaisesRegex(
                ValueError,
                "检索参考号存在冲突",
            ):
                load_uploaded_summary(source)


class CouponStatusFillTest(unittest.TestCase):
    def test_uploaded_match_sets_detail_and_remark(self) -> None:
        reference = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                row = coupon_row(header, "12345678901n")
                rows = [list(header), row]
                lookup = {reference: "已完成：匹配成功"}

                count = matching.fill_uploaded_details(rows, lookup)

                remark_index = header.index("备注")
                detail_index = header.index("详细情况")
                self.assertEqual(count, 1)
                self.assertEqual(row[remark_index], "已上传")
                self.assertEqual(row[detail_index], "已完成：匹配成功")

    def test_reference_outside_submitted_data_is_marked_unsubmitted(self) -> None:
        """Without unsubmitted data, anything not submitted is 未上传."""
        for header in HEADERS:
            with self.subTest(header=header):
                row = coupon_row(header, "99999999999Z")
                rows = [list(header), row]

                count = matching.fill_unmatched_remarks(rows, {"12345678901N"})

                remark_index = header.index("备注")
                self.assertEqual(count, 1)
                self.assertEqual(row[remark_index], "未上传")

    def test_excluded_bottom_rows_are_left_alone(self) -> None:
        reference = "12345678901N"
        header = appliance.COUPON_OUTPUT_HEADER
        uploaded_row = coupon_row(
            header,
            reference,
            remark="已上传",
            detail="已完成：匹配成功",
        )
        bottom_row = coupon_row(header, "99999999999Z")
        rows = [list(header), uploaded_row, bottom_row]

        count = matching.fill_unmatched_remarks(
            rows,
            {reference},
            excluded_bottom_rows=1,
        )

        remark_index = header.index("备注")
        self.assertEqual(count, 0)
        self.assertEqual(uploaded_row[remark_index], "已上传")
        self.assertEqual(bottom_row[remark_index], "")


if __name__ == "__main__":
    unittest.main()
