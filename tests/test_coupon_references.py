import unittest

from processors import digital, large_appliances


PROCESSORS = (digital, large_appliances)


def coupon_rows(processor: object, *references: str) -> list[list[object]]:
    header = processor.COUPON_OUTPUT_HEADER
    summary_index = header.index("明细摘要")
    rows = [list(header)]
    for reference in references:
        row: list[object] = [None] * len(header)
        row[summary_index] = reference
        rows.append(row)
    return rows


def correct_references(
    processor: object,
    rows: list[list[object]],
    universe: set[str],
) -> tuple[int, int, int]:
    if processor is large_appliances:
        return processor.correct_coupon_references(
            rows,
            universe,
            excluded_bottom_rows=0,
        )
    return processor.correct_coupon_references(rows, universe)


class CouponReferenceCorrectionTest(unittest.TestCase):
    def test_extracts_reference_from_supported_input_variants(self) -> None:
        reference = "12345678901A"
        variants = (
            "12345678901a",
            "12345 678901 A",
            "参考号：12345678901A",
            "12345678901",
            "12345678901AB",
            "12345678901B",
        )
        for processor in PROCESSORS:
            for raw_reference in variants:
                with self.subTest(
                    processor=processor.__name__,
                    raw_reference=raw_reference,
                ):
                    self.assertEqual(
                        processor.reference_correction_candidates(
                            raw_reference,
                            {reference},
                        ),
                        {reference},
                    )

    def test_ambiguous_candidate_is_not_selected(self) -> None:
        universe = {"12345678901A", "12345678901B"}
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                self.assertEqual(
                    processor.reference_correction_candidates(
                        "12345678901",
                        universe,
                    ),
                    universe,
                )

    def test_unique_candidate_is_corrected(self) -> None:
        target = "12345678901A"
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                rows = coupon_rows(processor, "12345 678901 A")

                corrected, unresolved, collisions, decisions = (
                    correct_references(processor, rows, {target})
                )

                summary_index = processor.COUPON_OUTPUT_HEADER.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (1, 0, 0))
                self.assertEqual(rows[1][summary_index], target)
                # Every applied correction must be auditable in the report.
                self.assertEqual(len(decisions), 1)
                result_kind, row_number, original, explanation = decisions[0]
                self.assertEqual(result_kind, processor.REFERENCE_REPORT_CORRECTED)
                self.assertEqual(row_number, "2")
                self.assertEqual(original, "12345 678901 A")
                self.assertIn(target, explanation)

    def test_duplicate_target_collision_is_not_corrected(self) -> None:
        target = "12345678901A"
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                rows = coupon_rows(
                    processor,
                    "12345 678901 A",
                    "12345678901-A",
                )

                corrected, unresolved, collisions, decisions = (
                    correct_references(processor, rows, {target})
                )

                summary_index = processor.COUPON_OUTPUT_HEADER.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (0, 0, 2))
                self.assertEqual(rows[1][summary_index], "12345 678901 A")
                self.assertEqual(rows[2][summary_index], "12345678901-A")
                self.assertEqual(
                    [decision[0] for decision in decisions],
                    [processor.REFERENCE_REPORT_COLLISION] * 2,
                )

    def test_unresolvable_reference_is_reported_and_left_alone(self) -> None:
        """A reference with no match in submitted data must still be visible."""
        for processor in PROCESSORS:
            with self.subTest(processor=processor.__name__):
                rows = coupon_rows(processor, "99999999999Z")

                corrected, unresolved, collisions, decisions = (
                    correct_references(processor, rows, {"12345678901A"})
                )

                summary_index = processor.COUPON_OUTPUT_HEADER.index("明细摘要")
                self.assertEqual((corrected, unresolved, collisions), (0, 1, 0))
                self.assertEqual(rows[1][summary_index], "99999999999Z")
                self.assertEqual(len(decisions), 1)
                self.assertEqual(
                    decisions[0][0],
                    processor.REFERENCE_REPORT_UNRESOLVED,
                )
                self.assertEqual(decisions[0][2], "99999999999Z")


if __name__ == "__main__":
    unittest.main()
