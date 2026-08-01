import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string

from processors import submitted
from processors.common.config import load_merchants, submitted_file_marker
from processors.common.excel import (
    load_measurement_font,
    measurement_text,
    resolve_font,
    width_measurer,
    widths_are_additive,
)
from processors.submitted import STATUS_ORDER


def text_pixel_width(value: object, font) -> float:
    """Test oracle mirroring width_measurer's uncached, unoptimized path."""
    text = measurement_text(value)
    return font.getlength(text) if text else 0


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

    def subsidy_for(self, profile_name: str, amount: object) -> object:
        row = submitted.add_subsidy_column(
            [None, None, amount], profile_name=profile_name
        )
        return row[3]

    def test_rate_is_15_percent_below_either_cap(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                self.assertEqual(self.subsidy_for(profile_name, 1000), 150.00)

    def test_caps_differ_between_projects(self) -> None:
        self.assertEqual(submitted.PROFILES["家电"].subsidy_cap, Decimal("1500"))
        self.assertEqual(submitted.PROFILES["数码"].subsidy_cap, Decimal("500"))

    def test_appliance_cap_applies_at_1500_not_500(self) -> None:
        # 6000 * 0.15 = 900: above digital's cap, still below the appliance one.
        self.assertEqual(self.subsidy_for("家电", 6000), 900.00)
        self.assertEqual(self.subsidy_for("数码", 6000), 500.00)
        # 20000 * 0.15 = 3000: both capped, but at different values.
        self.assertEqual(self.subsidy_for("家电", 20000), 1500.00)
        self.assertEqual(self.subsidy_for("数码", 20000), 500.00)

    def test_blank_amount_yields_no_subsidy(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                self.assertIsNone(self.subsidy_for(profile_name, None))
                self.assertIsNone(self.subsidy_for(profile_name, ""))


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

    def test_repeated_values_cost_nothing_on_the_installed_font(self) -> None:
        """Whichever path the installed font takes, repeats must be free."""
        _, font_path = resolve_font()
        counter = CallCountingFont(load_measurement_font(font_path))
        measure = width_measurer(counter)

        measure("同一个值")
        after_first = counter.calls
        for _ in range(49):
            measure("同一个值")

        self.assertEqual(counter.calls, after_first)


class CallCountingFont:
    """Wraps a font object and counts getlength calls."""

    def __init__(self, font) -> None:
        self._font = font
        self.calls = 0

    def getlength(self, text: str) -> float:
        self.calls += 1
        return self._font.getlength(text)


class FakeMonospaceFont:
    """Every glyph advances the same width, so widths sum exactly."""

    def __init__(self, width: float = 10.0) -> None:
        self.width = width
        self.calls = 0

    def getlength(self, text: str) -> float:
        self.calls += 1
        return self.width * len(text)


class FakeProportionalFont:
    """Kerns the pairs real proportional fonts kern, so widths do not sum."""

    KERNED_PAIRS = {"AV": -2.0, "To": -1.5, "Ta": -1.0}

    def __init__(self) -> None:
        self.calls = 0

    def getlength(self, text: str) -> float:
        self.calls += 1
        total = sum(7.0 if character.isascii() else 14.0 for character in text)
        for pair, adjustment in self.KERNED_PAIRS.items():
            total += adjustment * (len(text) - len(text.replace(pair, "")) ) / len(pair)
        return total


class WidthMeasurerFontPathsTest(unittest.TestCase):
    """Both measuring strategies, on fonts the test itself controls.

    Using fake fonts rather than whatever is installed keeps the assertions
    about call counts meaningful on any machine.
    """

    SAMPLE = "同一个值"

    def test_additive_font_takes_the_per_character_path(self) -> None:
        self.assertTrue(widths_are_additive(FakeMonospaceFont()))

    def test_kerning_font_is_rejected_by_detection(self) -> None:
        self.assertFalse(widths_are_additive(FakeProportionalFont()))

    def test_additive_font_measures_each_distinct_character_once(self) -> None:
        font = FakeMonospaceFont()
        measure = width_measurer(font)
        font.calls = 0

        measure(self.SAMPLE)
        self.assertLessEqual(font.calls, len(set(self.SAMPLE)))
        after_first = font.calls

        for _ in range(49):
            measure(self.SAMPLE)
        self.assertEqual(font.calls, after_first)

    def test_additive_font_result_equals_the_direct_measurement(self) -> None:
        font = FakeMonospaceFont()
        measure = width_measurer(font)
        for value in ("SN2026ABCD", "商品名称示例", "中文A1，混排", "0123456789"):
            with self.subTest(value=value):
                self.assertEqual(measure(value), font.getlength(value))

    def test_kerning_font_keeps_whole_string_measurement(self) -> None:
        font = FakeProportionalFont()
        measure = width_measurer(font)
        font.calls = 0

        for _ in range(50):
            measure(self.SAMPLE)
        self.assertEqual(font.calls, 1)

        # A kerned string must keep the font's own answer, not a summed one.
        self.assertEqual(measure("AV"), font.getlength("AV"))
        self.assertNotEqual(
            font.getlength("AV"),
            font.getlength("A") + font.getlength("V"),
        )


class SubmittedFileMarkerTest(unittest.TestCase):
    """The marker is derived from the merchant id, never configured twice.

    The export is named MER_<商户编号>_<导出时间>_yjhx.xlsx, so a hardcoded
    marker could drift out of sync with config/merchants.yaml and send a run
    looking for files that no longer exist.
    """

    def test_marker_is_the_merchant_id_with_the_exporter_prefix(self) -> None:
        merchants = load_merchants()
        for data_type in ("家电", "数码"):
            with self.subTest(data_type=data_type):
                self.assertEqual(
                    submitted_file_marker(data_type),
                    f"MER_{merchants[data_type]}",
                )

    def test_each_profile_uses_its_own_data_type(self) -> None:
        self.assertEqual(submitted.PROFILES["数码"].data_type, "数码")
        self.assertEqual(submitted.PROFILES["家电"].data_type, "家电")
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            submitted.configure_data_dir(data_dir)
            self.assertEqual(
                submitted.SUBMITTED_FILE_MARKERS["数码"],
                submitted_file_marker("数码"),
            )
            self.assertEqual(
                submitted.SUBMITTED_FILE_MARKERS["家电"],
                submitted_file_marker("家电"),
            )

    def test_unknown_data_type_is_reported(self) -> None:
        with self.assertRaisesRegex(ValueError, "缺少家具的商户编号"):
            submitted_file_marker("家具")


class SubmittedHeaderValidationTest(unittest.TestCase):
    """A wrong export format must name the file and the missing fields."""

    def run_build(self, profile_name: str, header: tuple[str, ...]):
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            write_submitted_source(
                data_dir / f"{submitted_file_marker(profile_name)}_export.xlsx",
                header,
            )
            submitted.configure_data_dir(data_dir)
            return submitted.build_report(profile_name)

    def test_missing_required_fields_are_reported_by_name(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                header = build_header(**{STATUS_COLUMN: "占位一", DESCRIPTION_COLUMN: "占位二"})
                with self.assertRaises(ValueError) as raised:
                    self.run_build(profile_name, header)

                message = str(raised.exception)
                self.assertIn("export.xlsx", message)
                self.assertIn("状态", message)
                self.assertIn("描述", message)

    def test_valid_header_is_accepted(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                report = self.run_build(
                    profile_name,
                    SUBMITTED_HEADER,
                )
                self.assertEqual(report.file_count, 1)
                self.assertEqual(report.data_row_count, 1)


class ValidatorRejectsBadOutputTest(unittest.TestCase):
    """The atomic save only protects the output if the validators really reject.

    Each case writes a deliberately wrong workbook and asserts it is refused.
    """

    def build_valid_output(self, profile_name: str, path: Path) -> int:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            write_submitted_source(
                data_dir / f"{submitted_file_marker(profile_name)}_export.xlsx",
                SUBMITTED_HEADER,
            )
            submitted.configure_data_dir(data_dir)
            report = submitted.build_report(profile_name)
            submitted.write_workbook(path, report)
            return report.data_row_count

    def test_accepts_the_workbook_it_just_built(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(profile_name, output)
                    submitted.validate_output(output, data_rows, profile_name)

                    workbook = load_workbook(output)
                    try:
                        sheet = workbook["Summary"]
                        self.assertEqual(sheet.freeze_panes, "A2")
                        self.assertEqual(sheet.auto_filter.ref, "A1:L2")
                        self.assertEqual(sheet["A1"].fill.fgColor.rgb[-6:], "000000")
                        self.assertEqual(sheet["D2"].number_format, "0.00")
                    finally:
                        workbook.close()

    def test_rejects_wrong_row_count(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(profile_name, output)
                    with self.assertRaises(RuntimeError):
                        submitted.validate_output(
                            output, data_rows + 1, profile_name
                        )

    def test_rejects_missing_status_worksheet(self) -> None:
        for profile_name in ("家电", "数码"):
            with self.subTest(profile_name=profile_name):
                with tempfile.TemporaryDirectory() as directory:
                    output = Path(directory) / "已上传.xlsx"
                    data_rows = self.build_valid_output(profile_name, output)

                    workbook = load_workbook(output)
                    try:
                        del workbook[STATUS_ORDER[0]]
                        workbook.save(output)
                    finally:
                        workbook.close()

                    with self.assertRaises(RuntimeError):
                        submitted.validate_output(output, data_rows, profile_name)


if __name__ == "__main__":
    unittest.main()
