"""Lock in the coupon reference matching order.

Required order:
1. match the 明细摘要 reference against submitted data
2. otherwise substitute the reference from the supplement file
3. otherwise correct the reference algorithmically
4. anything still unmatched is marked 未上传

Digital data must never consult the supplement file.
"""
import unittest
from datetime import date

from processors import digital, large_appliances


SUBMITTED_REFERENCE = "12345678901A"
SUPPLEMENT_REFERENCE = "22222222222B"
CORRECTABLE_REFERENCE = "33333333333C"


def coupon_row(processor, reference: str, document: str, day: date):
    row: list[object] = [""] * len(processor.COUPON_OUTPUT_HEADER)
    row[0] = document
    row[1] = day
    row[processor.COUPON_OUTPUT_HEADER.index("明细摘要")] = reference
    return row


class LargeApplianceMatchOrderTest(unittest.TestCase):
    def setUp(self) -> None:
        self.day = date(2026, 1, 24)
        self.universe = {
            SUBMITTED_REFERENCE,
            SUPPLEMENT_REFERENCE,
            CORRECTABLE_REFERENCE,
        }
        self.summary_index = large_appliances.COUPON_OUTPUT_HEADER.index("明细摘要")

    def test_submitted_match_is_never_rewritten(self) -> None:
        """A reference that already matches submitted data must survive untouched.

        The supplement lists a different reference for the same document, and it
        must not win.
        """
        row = coupon_row(large_appliances, SUBMITTED_REFERENCE, "001", self.day)
        rows = [list(large_appliances.COUPON_OUTPUT_HEADER), row]
        supplement = {("001", self.day): frozenset({SUPPLEMENT_REFERENCE})}

        matched, ambiguous, row_ids, _ = (
            large_appliances.fill_coupon_reference_supplement(
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
        row = coupon_row(large_appliances, "99999999999Z", "001", self.day)
        rows = [list(large_appliances.COUPON_OUTPUT_HEADER), row]
        supplement = {("001", self.day): frozenset({SUPPLEMENT_REFERENCE})}

        _, _, protected_row_ids, _ = (
            large_appliances.fill_coupon_reference_supplement(
                rows,
                supplement,
                self.universe,
                0,
            )
        )
        self.assertEqual(row[self.summary_index], SUPPLEMENT_REFERENCE)

        corrected, unresolved, collisions, decisions = (
            large_appliances.correct_coupon_references(
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
        row = coupon_row(large_appliances, "33333 333333 C", "002", self.day)
        rows = [list(large_appliances.COUPON_OUTPUT_HEADER), row]

        _, _, protected_row_ids, _ = (
            large_appliances.fill_coupon_reference_supplement(
                rows,
                {},
                self.universe,
                0,
            )
        )
        corrected, _, _, decisions = large_appliances.correct_coupon_references(
            rows,
            self.universe,
            0,
            protected_row_ids,
        )

        self.assertEqual(corrected, 1)
        self.assertEqual(row[self.summary_index], CORRECTABLE_REFERENCE)
        self.assertEqual(
            decisions[0][0],
            large_appliances.REFERENCE_REPORT_CORRECTED,
        )

    def test_reference_matching_nothing_is_marked_unsubmitted(self) -> None:
        row = coupon_row(large_appliances, "99999999999Z", "003", self.day)
        rows = [list(large_appliances.COUPON_OUTPUT_HEADER), row]

        count = large_appliances.fill_unmatched_remarks(rows, self.universe, 0)

        remark_index = large_appliances.COUPON_OUTPUT_HEADER.index("备注")
        self.assertEqual(count, 1)
        self.assertEqual(row[remark_index], "未上传")


class DigitalSkipsSupplementTest(unittest.TestCase):
    def test_digital_has_no_supplement_step(self) -> None:
        """Digital coupon data must never be matched against the supplement file."""
        for attribute in (
            "load_coupon_reference_supplement",
            "fill_coupon_reference_supplement",
            "COUPON_REFERENCE_SUPPLEMENT_FILE",
        ):
            self.assertFalse(
                hasattr(digital, attribute),
                f"digital 不应包含参考号异常数据逻辑：{attribute}",
            )


if __name__ == "__main__":
    unittest.main()
