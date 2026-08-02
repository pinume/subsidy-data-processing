import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from processors.coupons import appliance, matching
from processors.coupons import digital as coupons_digital
from processors.coupons.sources import (
    load_payment_reference_locations,
    load_uploaded_summary,
    validate_payment_reference_subset,
)
from processors.coupons.validation import validate_payment_statuses

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
    def test_uploaded_and_payment_matches_are_filled_in_one_pass(self) -> None:
        reference = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                row = coupon_row(header, "12345678901n")
                rows = [list(header), row]
                lookup = {reference: "已完成：匹配成功"}

                counts = matching.fill_reference_statuses(
                    rows,
                    lookup,
                    {reference},
                    {reference},
                )

                remark_index = header.index("备注")
                detail_index = header.index("详细情况")
                payment_index = header.index("回款情况")
                self.assertEqual(counts, (1, 0, 1))
                self.assertEqual(row[remark_index], "已上传")
                self.assertEqual(row[detail_index], "已完成：匹配成功")
                self.assertEqual(row[payment_index], "已回款")

    def test_reference_outside_submitted_data_is_marked_unsubmitted(self) -> None:
        """Without unsubmitted data, anything not submitted is 未上传."""
        for header in HEADERS:
            with self.subTest(header=header):
                row = coupon_row(header, "99999999999Z")
                rows = [list(header), row]

                counts = matching.fill_reference_statuses(
                    rows,
                    {},
                    {"12345678901N"},
                    set(),
                )

                remark_index = header.index("备注")
                payment_index = header.index("回款情况")
                self.assertEqual(counts, (0, 1, 0))
                self.assertEqual(row[remark_index], "未上传")
                self.assertIsNone(row[payment_index])

    def test_payment_reference_is_ignored_until_remark_is_uploaded(self) -> None:
        reference = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                row = coupon_row(header, reference)
                rows = [list(header), row]

                counts = matching.fill_reference_statuses(
                    rows,
                    {},
                    {reference},
                    {reference},
                )

                self.assertEqual(counts, (0, 0, 0))
                self.assertEqual(row[header.index("备注")], "")
                self.assertIsNone(row[header.index("回款情况")])

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

        counts = matching.fill_reference_statuses(
            rows,
            {reference: "已完成：匹配成功"},
            {reference},
            {reference},
            excluded_bottom_rows=1,
        )

        remark_index = header.index("备注")
        payment_index = header.index("回款情况")
        self.assertEqual(counts, (1, 0, 1))
        self.assertEqual(uploaded_row[remark_index], "已上传")
        self.assertEqual(uploaded_row[payment_index], "已回款")
        self.assertEqual(bottom_row[remark_index], "")
        self.assertIsNone(bottom_row[payment_index])

    def test_validation_rejects_an_incorrect_payment_status(self) -> None:
        reference = "12345678901N"
        header = appliance.COUPON_OUTPUT_HEADER
        row = coupon_row(header, reference)
        row[header.index("回款情况")] = "已回款"
        rows = [list(header), row]

        with self.assertRaisesRegex(RuntimeError, "回款情况匹配校验失败"):
            validate_payment_statuses(
                rows,
                header,
                {reference},
                excluded_bottom_rows=0,
                expected_paid_rows=1,
            )


class PaymentReferenceSourceTest(unittest.TestCase):
    def test_loads_appliance_and_digital_references_separately(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "回款明细.xlsx"
            workbook = Workbook()
            appliance_sheet = workbook.active
            appliance_sheet.title = "家电明细"
            appliance_sheet.append(["交易参考号"])
            appliance_sheet.append(["12345678901n"])
            digital_sheet = workbook.create_sheet("数码明细")
            digital_sheet.append(["交易参考号"])
            digital_sheet.append(["12345678902N"])
            workbook.save(source)
            workbook.close()

            result = load_payment_reference_locations(source)

        self.assertEqual(set(result["家电"]), {"12345678901N"})
        self.assertEqual(set(result["数码"]), {"12345678902N"})

    def test_requires_both_detail_sheets_and_reference_headers(self) -> None:
        cases = (
            (False, True, "缺少 数码明细 工作表"),
            (True, False, "家电明细 缺少字段：交易参考号"),
        )
        for include_digital, include_header, expected in cases:
            with self.subTest(expected=expected), tempfile.TemporaryDirectory() as directory:
                source = Path(directory) / "回款明细.xlsx"
                workbook = Workbook()
                appliance_sheet = workbook.active
                appliance_sheet.title = "家电明细"
                appliance_sheet.append(
                    ["交易参考号"] if include_header else ["订单号"]
                )
                if include_digital:
                    digital_sheet = workbook.create_sheet("数码明细")
                    digital_sheet.append(["交易参考号"])
                workbook.save(source)
                workbook.close()

                with self.assertRaisesRegex(ValueError, expected):
                    load_payment_reference_locations(source)

    def test_rejects_payment_reference_outside_submitted_data(self) -> None:
        locations = {"12345678901N": "回款明细.xlsx 的 家电明细 第 2 行"}

        with self.assertRaisesRegex(
            ValueError,
            "家电回款参考号子集校验失败.*第 2 行",
        ):
            validate_payment_reference_subset("家电", locations, set())


if __name__ == "__main__":
    unittest.main()
