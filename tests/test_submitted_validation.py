import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string

from processors import digital, large_appliances
from processors.common.excel import (
    load_measurement_font,
    resolve_font,
    text_pixel_width,
    width_measurer,
)


SOURCE_COLUMN_COUNT = 24
# add_subsidy_column reads the amount from the third kept column, which is F.
AMOUNT_COLUMN = "F"
STATUS_COLUMN = "Q"
DESCRIPTION_COLUMN = "U"


def build_header(**overrides: str) -> tuple[str, ...]:
    """Build a source header row, naming fields by their source column letter."""
    header = [
        chr(ord("A") + index) * 2 for index in range(SOURCE_COLUMN_COUNT)
    ]
    names = {
        AMOUNT_COLUMN: "交易金额",
        STATUS_COLUMN: "状态",
        DESCRIPTION_COLUMN: "描述",
    }
    names.update(overrides)
    for letter, name in names.items():
        header[column_index_from_string(letter) - 1] = name
    return tuple(header)


SUBMITTED_HEADER = build_header()


def write_submitted_source(path: Path, header: tuple[str, ...]) -> None:
    """Write a source file shaped like the upload export: title row, then header."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["报表标题"])
    sheet.append(list(header))
    row: list[object] = ["v"] * len(header)
    row[column_index_from_string(AMOUNT_COLUMN) - 1] = 1000
    row[column_index_from_string(STATUS_COLUMN) - 1] = "审核通过"
    row[column_index_from_string(DESCRIPTION_COLUMN) - 1] = "说明"
    sheet.append(row)
    workbook.save(path)
    workbook.close()


class SubsidyCapTest(unittest.TestCase):
    """The two projects share the 15% rate but cap at different amounts.

    Household appliances cap at 1500, digital at 500. They were once both
    written as 500, which silently understated 43% of the appliance rows, so
    the caps are pinned here rather than left to a shared default.
    """

    def subsidy_for(self, module, amount: object) -> object:
        row = module.add_subsidy_column([None, None, amount])
        return row[3]

    def test_rate_is_15_percent_below_either_cap(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                self.assertEqual(self.subsidy_for(module, 1000), 150.00)

    def test_caps_differ_between_projects(self) -> None:
        self.assertEqual(large_appliances.SUBSIDY_CAP, Decimal("1500"))
        self.assertEqual(digital.SUBSIDY_CAP, Decimal("500"))

    def test_appliance_cap_applies_at_1500_not_500(self) -> None:
        # 6000 * 0.15 = 900: above digital's cap, still below the appliance one.
        self.assertEqual(self.subsidy_for(large_appliances, 6000), 900.00)
        self.assertEqual(self.subsidy_for(digital, 6000), 500.00)
        # 20000 * 0.15 = 3000: both capped, but at different values.
        self.assertEqual(self.subsidy_for(large_appliances, 20000), 1500.00)
        self.assertEqual(self.subsidy_for(digital, 20000), 500.00)

    def test_blank_amount_yields_no_subsidy(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                self.assertIsNone(self.subsidy_for(module, None))
                self.assertIsNone(self.subsidy_for(module, ""))


class WidthMeasurerTest(unittest.TestCase):
    """The cached measurer must stay numerically identical to the direct call."""

    def test_matches_text_pixel_width_for_every_value_kind(self) -> None:
        _, font_path = resolve_font()
        font = load_measurement_font(font_path)
        measure = width_measurer(font)

        for value in (
            None,
            "",
            "ABC",
            "商品名称示例",
            0,
            0.0,
            123.456,
            True,
            date(2026, 1, 24),
            "重复值",
            "重复值",
        ):
            with self.subTest(value=value):
                self.assertEqual(measure(value), text_pixel_width(value, font))

    def test_repeated_values_are_measured_once(self) -> None:
        _, font_path = resolve_font()
        font = load_measurement_font(font_path)
        calls = 0
        original = font.getbbox

        def counting_getbbox(text):
            nonlocal calls
            calls += 1
            return original(text)

        font.getbbox = counting_getbbox
        try:
            measure = width_measurer(font)
            for _ in range(50):
                measure("同一个值")
        finally:
            font.getbbox = original

        self.assertEqual(calls, 1)


class SubmittedHeaderValidationTest(unittest.TestCase):
    """A wrong export format must name the file and the missing fields."""

    def run_build(self, module, header: tuple[str, ...]):
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            write_submitted_source(
                data_dir / f"{module.SUBMITTED_FILE_MARKER}_export.xlsx",
                header,
            )
            module.configure_data_dir(data_dir)
            return module.build_workbook()

    def test_missing_required_fields_are_reported_by_name(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                header = build_header(**{STATUS_COLUMN: "占位一", DESCRIPTION_COLUMN: "占位二"})
                with self.assertRaises(ValueError) as raised:
                    self.run_build(module, header)

                message = str(raised.exception)
                self.assertIn("export.xlsx", message)
                self.assertIn("状态", message)
                self.assertIn("描述", message)

    def test_valid_header_is_accepted(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                workbook, file_count, data_rows = self.run_build(
                    module,
                    SUBMITTED_HEADER,
                )
                try:
                    self.assertEqual(file_count, 1)
                    self.assertEqual(data_rows, 1)
                finally:
                    workbook.close()


class ValidatorRejectsBadOutputTest(unittest.TestCase):
    """The atomic save only protects the output if the validators really reject.

    Each case writes a deliberately wrong workbook and asserts it is refused.
    """

    def build_valid_output(self, module, path: Path) -> int:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            write_submitted_source(
                data_dir / f"{module.SUBMITTED_FILE_MARKER}_export.xlsx",
                SUBMITTED_HEADER,
            )
            module.configure_data_dir(data_dir)
            workbook, _, data_rows = module.build_workbook()
            try:
                workbook.save(path)
            finally:
                workbook.close()
            return data_rows

    def test_accepts_the_workbook_it_just_built(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(module, output)
                    module.validate_output(output, data_rows)

    def test_rejects_wrong_row_count(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(module, output)
                    with self.assertRaises(RuntimeError):
                        module.validate_output(output, data_rows + 1)

    def test_rejects_missing_status_worksheet(self) -> None:
        for module in (digital, large_appliances):
            with self.subTest(module=module.__name__):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(module, output)

                    workbook = load_workbook(output)
                    try:
                        del workbook[module.STATUS_ORDER[0]]
                        workbook.save(output)
                    finally:
                        workbook.close()

                    with self.assertRaises(RuntimeError):
                        module.validate_output(output, data_rows)


if __name__ == "__main__":
    unittest.main()
