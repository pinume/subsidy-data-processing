import io
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from processors.coupons import appliance
from processors.coupons import sources as coupon_sources
from processors.coupons.appliance import (
    COUPON_OUTPUT_HEADER,
    fill_coupon_reference_supplement,
    load_coupon_reference_supplement,
)


class CouponReferenceSupplementTest(unittest.TestCase):
    @staticmethod
    def coupon_row(
        document_number: str,
        document_date: date,
        reference: str,
    ) -> list[object]:
        values: dict[str, object] = {
            "单据号": document_number,
            "单据日期": document_date,
            "明细摘要": reference,
        }
        return [values.get(header) for header in COUPON_OUTPUT_HEADER]

    def test_missing_optional_file_returns_empty_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "reference_number_supplement.xlsx"
            output = io.StringIO()

            lookup, missing = load_coupon_reference_supplement(source)

            self.assertEqual(lookup, {})
            self.assertTrue(missing)
            # The loader never prints; the caller reports a missing file.
            self.assertEqual(output.getvalue(), "")

    def test_finds_supplement_file_by_keyword_in_flat_directory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            supplement_file = data_dir / "新建 Microsoft Excel 工作表.xlsx"
            supplement_file.touch()

            coupon_sources.configure_data_dir(data_dir)

            self.assertEqual(
                coupon_sources.COUPON_REFERENCE_SUPPLEMENT_FILE,
                supplement_file,
            )

    def test_missing_supplement_file_falls_back_to_a_display_path(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)

            coupon_sources.configure_data_dir(data_dir)

            self.assertEqual(
                coupon_sources.COUPON_REFERENCE_SUPPLEMENT_FILE,
                data_dir / "新建 Microsoft Excel 工作表.xlsx",
            )

    def test_loads_and_deduplicates_valid_references(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "reference_number_supplement.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["参考号", "单据号", "单据日期"])
            sheet.append(["12345678901n", "收款1001", "2026-07-06"])
            sheet.append(["12345678901N", "1001", date(2026, 7, 6)])
            workbook.save(source)
            workbook.close()

            lookup, missing = load_coupon_reference_supplement(source)

            self.assertFalse(missing)
            self.assertEqual(
                lookup,
                {("1001", date(2026, 7, 6)): frozenset({"12345678901N"})},
            )

    def test_unique_reference_replaces_unmatched_value(self) -> None:
        row = self.coupon_row("1001", date(2026, 7, 6), "待补充")
        rows = [list(COUPON_OUTPUT_HEADER), row]
        reference = "12345678901N"

        result = fill_coupon_reference_supplement(
            rows,
            {("1001", date(2026, 7, 6)): frozenset({reference})},
            {reference},
            excluded_bottom_rows=0,
        )

        self.assertEqual(result[0:2], (1, 0))
        self.assertEqual(row[COUPON_OUTPUT_HEADER.index("明细摘要")], reference)
        self.assertEqual(result[2], {id(row)})
        self.assertEqual(result[3][("1001", date(2026, 7, 6), reference)], 1)

    def test_ambiguous_references_do_not_change_row(self) -> None:
        row = self.coupon_row("1001", date(2026, 7, 6), "待补充")
        rows = [list(COUPON_OUTPUT_HEADER), row]
        original = list(row)
        # Unsorted input: the recorded candidates must come out sorted so the
        # console warning is stable across runs.
        references = frozenset({"12345678902N", "12345678901N"})

        result = fill_coupon_reference_supplement(
            rows,
            {("1001", date(2026, 7, 6)): references},
            set(references),
            excluded_bottom_rows=0,
        )

        self.assertEqual(result[0:2], (0, 1))
        self.assertEqual(row[COUPON_OUTPUT_HEADER.index("明细摘要")], "待补充")
        self.assertEqual(result[2], set())
        self.assertEqual(result[3], {})
        # No column touched — upload and payment statuses stay exactly as the
        # row carried them.
        self.assertEqual(row, original)
        self.assertEqual(
            result[4],
            (
                appliance.SupplementReferenceConflict(
                    document_number="1001",
                    document_date=date(2026, 7, 6),
                    current_reference="待补充",
                    candidates=("12345678901N", "12345678902N"),
                ),
            ),
        )

    def test_known_reference_and_excluded_bottom_row_are_unchanged(self) -> None:
        known = "12345678901N"
        known_row = self.coupon_row("1001", date(2026, 7, 6), known)
        bottom_row = self.coupon_row("1002", date(2026, 7, 6), "待补充")
        rows = [list(COUPON_OUTPUT_HEADER), known_row, bottom_row]
        replacement = "12345678902N"

        result = fill_coupon_reference_supplement(
            rows,
            {
                ("1001", date(2026, 7, 6)): frozenset({replacement}),
                ("1002", date(2026, 7, 6)): frozenset({replacement}),
            },
            {known, replacement},
            excluded_bottom_rows=1,
        )

        self.assertEqual(result[0:2], (0, 0))
        self.assertEqual(known_row[COUPON_OUTPUT_HEADER.index("明细摘要")], known)
        self.assertEqual(
            bottom_row[COUPON_OUTPUT_HEADER.index("明细摘要")],
            "待补充",
        )


if __name__ == "__main__":
    unittest.main()
