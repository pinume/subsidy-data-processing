import unittest
from datetime import date

from processors.coupons.matching import reference_correction_candidates
from processors.coupons import appliance, matching
from processors.coupons import digital as coupons_digital


HEADERS = (coupons_digital.COUPON_OUTPUT_HEADER, appliance.COUPON_OUTPUT_HEADER)
DOCUMENT_DATE = date(2026, 7, 6)


def coupon_rows(header: tuple[str, ...], *references: str) -> list[list[object]]:
    """Build detail rows carrying the 单据号/单据日期 the report identifies
    decisions by, one document number per row so they stay distinguishable."""
    summary_index = header.index("明细摘要")
    document_index = header.index("单据号")
    date_index = header.index("单据日期")
    rows = [list(header)]
    for offset, reference in enumerate(references):
        row: list[object] = [None] * len(header)
        row[summary_index] = reference
        row[document_index] = f"100{offset}"
        row[date_index] = DOCUMENT_DATE
        rows.append(row)
    return rows


class CouponReferenceCorrectionTest(unittest.TestCase):
    def test_extracts_reference_from_supported_input_variants(self) -> None:
        reference = "12345678901N"
        variants = (
            "12345678901n",
            "12345 678901 N",
            "参考号：12345678901N",
            "12345678901",
            "12345678901NN",
            "12345678901M",
        )
        for raw_reference in variants:
            with self.subTest(raw_reference=raw_reference):
                self.assertEqual(
                    reference_correction_candidates(raw_reference, {reference}),
                    {reference},
                )

    def test_ambiguous_candidate_is_not_selected(self) -> None:
        universe = {"12345678901N", "12345678902N"}
        self.assertEqual(
            reference_correction_candidates("1234567890N", universe),
            universe,
        )

    def test_suffix_letter_comes_from_uploaded_reference_universe(self) -> None:
        universe = {"12345678901A", "12345678902Z"}
        self.assertEqual(
            reference_correction_candidates("12345678901", universe),
            {"12345678901A"},
        )
        self.assertEqual(
            reference_correction_candidates("12345678902", universe),
            {"12345678902Z"},
        )

    def test_unique_candidate_is_corrected(self) -> None:
        target = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, "12345 678901 N")

                corrected, unresolved, collisions, decisions = (
                    matching.correct_coupon_references(rows, {target})
                )

                summary_index = header.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (1, 0, 0))
                self.assertEqual(rows[1][summary_index], target)
                # Every applied correction must be auditable in the report.
                self.assertEqual(len(decisions), 1)
                (
                    result_kind,
                    document_number,
                    document_date,
                    original,
                    explanation,
                ) = decisions[0]
                self.assertEqual(result_kind, matching.REFERENCE_REPORT_CORRECTED)
                # Identified by document, not row position: the detail rows get
                # re-sorted after this runs, so a row number would go stale.
                self.assertEqual(document_number, "1000")
                self.assertEqual(document_date, DOCUMENT_DATE)
                self.assertEqual(original, "12345 678901 N")
                self.assertIn(target, explanation)

    def test_duplicate_target_collision_is_not_corrected(self) -> None:
        target = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(
                    header,
                    "12345 678901 N",
                    "12345678901-N",
                )

                corrected, unresolved, collisions, decisions = (
                    matching.correct_coupon_references(rows, {target})
                )

                summary_index = header.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (0, 0, 2))
                self.assertEqual(rows[1][summary_index], "12345 678901 N")
                self.assertEqual(rows[2][summary_index], "12345678901-N")
                self.assertEqual(
                    [decision[0] for decision in decisions],
                    [matching.REFERENCE_REPORT_COLLISION] * 2,
                )

    def test_well_formed_reference_with_no_candidate_is_left_out_of_report(
        self,
    ) -> None:
        """A well-formed reference that simply was never submitted is noise.

        The detail row already carries a 未上传 remark, so reporting it again
        only buries the corrections an operator actually has to review — on
        real data this was 815 of 898 report lines.
        """
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, "99999999999N")

                corrected, unresolved, collisions, decisions = (
                    matching.correct_coupon_references(rows, {"12345678901N"})
                )

                summary_index = header.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (0, 1, 0))
                self.assertEqual(rows[1][summary_index], "99999999999N")
                self.assertEqual(decisions, [])

    def test_malformed_reference_is_reported(self) -> None:
        """Junk in the reference column is a data-entry problem worth fixing.

        Real exports contained placeholder text such as 预售 here.
        """
        for header in HEADERS:
            for raw_reference in (
                "预售",
                "1234567890",
                "99999999999",
            ):
                with self.subTest(header=header, raw_reference=raw_reference):
                    rows = coupon_rows(header, raw_reference)

                    corrected, unresolved, collisions, decisions = (
                        matching.correct_coupon_references(
                            rows, {"12345678901N"}
                        )
                    )

                    self.assertEqual(
                        (corrected, unresolved, collisions),
                        (0, 1, 0),
                    )
                    self.assertEqual(len(decisions), 1)
                    self.assertEqual(
                        decisions[0][0],
                        matching.REFERENCE_REPORT_UNRESOLVED,
                    )
                    self.assertEqual(decisions[0][3], raw_reference)


if __name__ == "__main__":
    unittest.main()
