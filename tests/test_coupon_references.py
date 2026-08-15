import unittest
from datetime import date

from processors.coupons import appliance, matching
from processors.coupons import digital as coupons_digital
from processors.coupons.matching import reference_correction_candidates

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

    def test_missing_suffix_n_is_filled_from_uploaded_universe(self) -> None:
        """Eleven bare digits resolve to the only legal form: digits + N."""
        universe = {"12345678901N", "12345678902N"}
        self.assertEqual(
            reference_correction_candidates("12345678901", universe),
            {"12345678901N"},
        )
        self.assertEqual(
            reference_correction_candidates("12345678902", universe),
            {"12345678902N"},
        )

    def test_wrong_suffix_letter_can_correct_to_n(self) -> None:
        self.assertEqual(
            reference_correction_candidates(
                "12345678901M", {"12345678901N"}
            ),
            {"12345678901N"},
        )

    def test_reference_glued_to_model_string_takes_digits_before_n(self) -> None:
        """No separator between …N and the product description."""
        target = "16294039444N"
        for raw in (
            "16294039444Np80pro12+512黑",
            "16294039444NP80PRO12+512黑",
            f"备注{target}p80pro",
        ):
            with self.subTest(raw=raw):
                self.assertEqual(
                    reference_correction_candidates(raw, {target}),
                    {target},
                )

    def test_chinese_prefix_before_legal_reference(self) -> None:
        """Chinese remark glued or adjacent to a legal reference."""
        target = "18005652528N"
        for raw in (
            f"回迁新居{target}",
            f"回迁新居 {target}",
        ):
            with self.subTest(raw=raw):
                self.assertEqual(
                    reference_correction_candidates(raw, {target}),
                    {target},
                )

    def test_document_dot_and_wrong_suffix_letter(self) -> None:
        """册号 + 流水号 + 点 + 错后缀 W → universe …N."""
        target = "55978228155N"
        dirty = "册号0002953.55978228155W"
        self.assertEqual(
            reference_correction_candidates(dirty, {target}),
            {target},
        )

    def test_glued_model_unique_candidate_is_written_back(self) -> None:
        target = "16294039444N"
        dirty = "16294039444Np80pro12+512黑"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, dirty)
                decisions = matching.correct_coupon_references(
                    rows, {target}
                )
                summary_index = header.index("明细摘要")
                self.assertEqual(rows[1][summary_index], target)
                # Decisions store the upper-cased original summary text.
                self.assertEqual(len(decisions), 1)
                self.assertEqual(
                    decisions[0][0], matching.REFERENCE_REPORT_CORRECTED
                )
                self.assertEqual(decisions[0][3], dirty.upper())

    def test_chinese_prefix_unique_candidate_is_written_back(self) -> None:
        target = "18005652528N"
        dirty = f"回迁新居{target}"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, dirty)
                decisions = matching.correct_coupon_references(
                    rows, {target}
                )
                summary_index = header.index("明细摘要")
                self.assertEqual(rows[1][summary_index], target)
                self.assertEqual(len(decisions), 1)
                self.assertEqual(
                    decisions[0][0], matching.REFERENCE_REPORT_CORRECTED
                )
                self.assertEqual(decisions[0][3], dirty.upper())

    def test_document_dot_wrong_suffix_is_written_back(self) -> None:
        target = "55978228155N"
        dirty = "册号0002953.55978228155W"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, dirty)
                decisions = matching.correct_coupon_references(
                    rows, {target}
                )
                summary_index = header.index("明细摘要")
                self.assertEqual(rows[1][summary_index], target)
                self.assertEqual(len(decisions), 1)
                self.assertEqual(
                    decisions[0][0], matching.REFERENCE_REPORT_CORRECTED
                )
                self.assertEqual(decisions[0][3], dirty.upper())

    def test_symbol_joined_document_and_valid_reference(self) -> None:
        """Two values in 明细摘要 joined by punctuation or spaces."""
        target = "55992508351N"
        for raw in (
            f"0003099,{target}",
            f"0003099，{target}",
            f"0003099, {target}",
            f"0003099;{target}",
            f"0003099；{target}",
            f"0003099/{target}",
            f"0003099|{target}",
            f"0003099：{target}",
            f"0003099-{target}",
            f"0003099 {target}",
            f"0003099  {target}",
            f"参考号：{target}",
        ):
            with self.subTest(raw=raw):
                self.assertEqual(
                    reference_correction_candidates(raw, {target}),
                    {target},
                )

    def test_symbol_joined_document_and_wrong_suffix_letter(self) -> None:
        """Prefix + 12-char dirty ref: only the reference side is edited."""
        target = "55986838725N"
        for raw in (
            "0002917,55986838725M",
            "0002917;55986838725M",
            "0002917/55986838725M",
            "0002917 55986838725M",
        ):
            with self.subTest(raw=raw):
                self.assertEqual(
                    reference_correction_candidates(raw, {target}),
                    {target},
                )

    def test_symbol_joined_two_universe_references_is_ambiguous(self) -> None:
        left = "12345678901N"
        right = "12345678902N"
        for sep in (",", ";", "/", "|", "、", " "):
            with self.subTest(sep=sep):
                self.assertEqual(
                    reference_correction_candidates(
                        f"{left}{sep}{right}", {left, right}
                    ),
                    {left, right},
                )

    def test_symbol_joined_unique_candidate_is_written_back(self) -> None:
        target = "55986838725N"
        dirty = "0002917 55986838725M"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, dirty)
                decisions = matching.correct_coupon_references(
                    rows, {target}
                )
                summary_index = header.index("明细摘要")
                self.assertEqual(rows[1][summary_index], target)
                self.assertEqual(len(decisions), 1)
                self.assertEqual(
                    decisions[0][0], matching.REFERENCE_REPORT_CORRECTED
                )
                self.assertEqual(decisions[0][3], dirty)

    def test_spaced_single_reference_still_resolves(self) -> None:
        """Internal spaces in one reference still match via the full-cell path."""
        target = "12345678901N"
        self.assertEqual(
            reference_correction_candidates("12345 678901 N", {target}),
            {target},
        )

    def test_universe_rejects_non_n_suffix(self) -> None:
        with self.assertRaisesRegex(ValueError, "大写字母 N"):
            matching.build_reference_correction_index({"12345678901A"})

    def test_unique_candidate_is_corrected(self) -> None:
        target = "12345678901N"
        for header in HEADERS:
            with self.subTest(header=header):
                rows = coupon_rows(header, "12345 678901 N")

                decisions = matching.correct_coupon_references(
                    rows, {target}
                )

                summary_index = header.index("明细摘要")
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
                self.assertEqual(
                    result_kind, matching.REFERENCE_REPORT_CORRECTED
                )
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

                decisions = matching.correct_coupon_references(
                    rows, {target}
                )

                summary_index = header.index("明细摘要")
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

                decisions = matching.correct_coupon_references(
                    rows, {"12345678901N"}
                )

                summary_index = header.index("明细摘要")
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

                    decisions = matching.correct_coupon_references(
                        rows, {"12345678901N"}
                    )

                    self.assertEqual(len(decisions), 1)
                    self.assertEqual(
                        decisions[0][0],
                        matching.REFERENCE_REPORT_UNRESOLVED,
                    )
                    self.assertEqual(decisions[0][3], raw_reference)


if __name__ == "__main__":
    unittest.main()
