"""Lock in the coupon reference matching order.

Required order:
1. match the 明细摘要 reference against submitted data
2. otherwise substitute the reference from the supplement file
3. otherwise correct the reference algorithmically
4. anything still unmatched is marked 未上传

Digital data must never consult the supplement file.
"""
import unittest
from contextlib import ExitStack
from datetime import date
from decimal import Decimal
from inspect import signature
from pathlib import Path
from unittest.mock import patch

from processors.coupons import appliance, matching, sources
from processors.coupons import digital as coupons_digital

SUBMITTED_REFERENCE = "12345678901N"
SUPPLEMENT_REFERENCE = "22222222222N"
CORRECTABLE_REFERENCE = "33333333333N"


def coupon_row(header: tuple[str, ...], reference: str, document: str, day: date):
    row: list[object] = [""] * len(header)
    row[0] = document
    row[1] = day
    row[header.index("明细摘要")] = reference
    return row


class LargeApplianceMatchOrderTest(unittest.TestCase):
    def setUp(self) -> None:
        self.day = date(2026, 1, 24)
        self.universe = {
            SUBMITTED_REFERENCE,
            SUPPLEMENT_REFERENCE,
            CORRECTABLE_REFERENCE,
        }
        self.summary_index = appliance.COUPON_OUTPUT_HEADER.index("明细摘要")

    def test_submitted_match_is_never_rewritten(self) -> None:
        """A reference that already matches submitted data must survive untouched.

        The supplement lists a different reference for the same document, and it
        must not win.
        """
        row = coupon_row(
            appliance.COUPON_OUTPUT_HEADER, SUBMITTED_REFERENCE, "001", self.day
        )
        rows = [list(appliance.COUPON_OUTPUT_HEADER), row]
        supplement = {("001", self.day): frozenset({SUPPLEMENT_REFERENCE})}

        matched, ambiguous, row_ids, _ = (
            appliance.fill_coupon_reference_supplement(
                rows,
                supplement,
                self.universe,
                0,
            )
        )

        self.assertEqual((matched, ambiguous), (0, 0))
        self.assertEqual(row_ids, set())
        self.assertEqual(row[self.summary_index], SUBMITTED_REFERENCE)

    def test_supplement_wins_over_algorithmic_correction(self) -> None:
        """A row fixed by the supplement must not be rewritten by the algorithm."""
        row = coupon_row(
            appliance.COUPON_OUTPUT_HEADER, "99999999999Z", "001", self.day
        )
        rows = [list(appliance.COUPON_OUTPUT_HEADER), row]
        supplement = {("001", self.day): frozenset({SUPPLEMENT_REFERENCE})}

        _, _, protected_row_ids, _ = (
            appliance.fill_coupon_reference_supplement(
                rows,
                supplement,
                self.universe,
                0,
            )
        )
        self.assertEqual(row[self.summary_index], SUPPLEMENT_REFERENCE)

        corrected, unresolved, collisions, decisions = (
            matching.correct_coupon_references(
                rows,
                self.universe,
                0,
                protected_row_ids,
            )
        )

        self.assertEqual((corrected, unresolved, collisions), (0, 0, 0))
        self.assertEqual(decisions, [])
        self.assertEqual(row[self.summary_index], SUPPLEMENT_REFERENCE)

    def test_algorithm_runs_only_when_supplement_has_no_entry(self) -> None:
        row = coupon_row(
            appliance.COUPON_OUTPUT_HEADER, "33333 333333 N", "002", self.day
        )
        rows = [list(appliance.COUPON_OUTPUT_HEADER), row]

        _, _, protected_row_ids, _ = (
            appliance.fill_coupon_reference_supplement(
                rows,
                {},
                self.universe,
                0,
            )
        )
        corrected, _, _, decisions = matching.correct_coupon_references(
            rows,
            self.universe,
            0,
            protected_row_ids,
        )

        self.assertEqual(corrected, 1)
        self.assertEqual(row[self.summary_index], CORRECTABLE_REFERENCE)
        self.assertEqual(
            decisions[0][0],
            matching.REFERENCE_REPORT_CORRECTED,
        )

    def test_reference_matching_nothing_is_marked_unsubmitted(self) -> None:
        row = coupon_row(
            appliance.COUPON_OUTPUT_HEADER, "99999999999Z", "003", self.day
        )
        rows = [list(appliance.COUPON_OUTPUT_HEADER), row]

        counts = matching.fill_reference_statuses(
            rows,
            {},
            self.universe,
            set(),
            0,
        )

        remark_index = appliance.COUPON_OUTPUT_HEADER.index("备注")
        self.assertEqual(counts, (0, 1, 0))
        self.assertEqual(row[remark_index], "未上传")

    def test_final_corrected_reference_drives_both_status_matches(self) -> None:
        target = "12345678901N"
        row = coupon_row(
            appliance.COUPON_OUTPUT_HEADER,
            "12345 678901 N",
            "004",
            self.day,
        )
        rows = [list(appliance.COUPON_OUTPUT_HEADER), row]

        corrected, _, _, _ = matching.correct_coupon_references(
            rows,
            {target},
        )
        counts = matching.fill_reference_statuses(
            rows,
            {target: "审核通过：同意"},
            {target},
            {target},
        )

        self.assertEqual(corrected, 1)
        self.assertEqual(counts, (1, 0, 1))
        self.assertEqual(
            row[appliance.COUPON_OUTPUT_HEADER.index("回款情况")],
            "已回款",
        )


class PinkBlockIsLeftAloneTest(unittest.TestCase):
    """The 退换货 block is settled by the receipt remark and must stay that way.

    Both projects pin those rows at the bottom and colour them pink; the
    reference passes have to skip them, or the remark that made a row pink is
    overwritten with an upload status and the sheet ends up with pink rows
    labelled 已上传.
    """

    def compute(self, module, rows, remark_lookup):
        """Run one project's whole pipeline over rows the test controls.

        Asserting on compute_coupon_data rather than on the matching helpers
        is the point: the helpers always honoured excluded_bottom_rows, and
        the fault was 数码 never passing it to them.
        """
        detail_lookup = {SUBMITTED_REFERENCE: "审核通过：同意"}
        with ExitStack() as stack:
            stack.enter_context(
                patch.object(
                    sources,
                    "COUPON_SOURCE_FILE",
                    Path("销售用券情况统计.xlsx"),
                    create=True,
                )
            )
            stack.enter_context(
                patch.object(
                    module,
                    "load_uploaded_summary",
                    return_value=(detail_lookup, 1, Decimal("0")),
                )
            )
            # 家电 alone reads the optional supplement file; the path is
            # only set by configure_data_dir, which this test does not run.
            if hasattr(module, "load_coupon_reference_supplement"):
                stack.enter_context(
                    patch.object(
                        sources,
                        "COUPON_REFERENCE_SUPPLEMENT_FILE",
                        Path("参考号补充.xlsx"),
                        create=True,
                    )
                )
                stack.enter_context(
                    patch.object(
                        module,
                        "load_coupon_reference_supplement",
                        return_value={},
                    )
                )
            kwargs = {"rows": rows, "remark_lookup": remark_lookup}
            if "source_total" in signature(module.compute_coupon_data).parameters:
                # 家电 would otherwise re-read the export for its 合计 row.
                kwargs["source_total"] = None
            return module.compute_coupon_data(**kwargs)

    def test_the_pink_block_keeps_its_remark(self) -> None:
        for name, module in (
            ("家电", appliance),
            ("数码", coupons_digital),
        ):
            with self.subTest(name):
                header = module.COUPON_OUTPUT_HEADER
                day = date(2026, 1, 24)
                rows = [
                    list(header),
                    coupon_row(header, SUBMITTED_REFERENCE, "001", day),
                    coupon_row(header, SUBMITTED_REFERENCE, "002", day),
                ]
                remark_index = header.index("备注")

                computation = self.compute(
                    module, rows, {("002", day): "退换货/倒票（退单）"}
                )

                self.assertEqual(computation.matched_count, 1)
                result = computation.rows
                # The pink block sits at the bottom; row 1 is the plain sale.
                self.assertEqual(result[1][remark_index], "已上传")
                self.assertEqual(
                    result[-1][remark_index], "退换货/倒票（退单）"
                )


class DigitalSkipsSupplementTest(unittest.TestCase):
    def test_digital_has_no_supplement_step(self) -> None:
        """Digital coupon data must never be matched against the supplement file."""
        for attribute in (
            "load_coupon_reference_supplement",
            "fill_coupon_reference_supplement",
            "COUPON_REFERENCE_SUPPLEMENT_FILE",
        ):
            self.assertFalse(
                hasattr(coupons_digital, attribute),
                f"digital 不应包含参考号异常数据逻辑：{attribute}",
            )


if __name__ == "__main__":
    unittest.main()
