"""XlsxWriter writers for the audit workbook's four kinds of sheet.

The audit output is 30 sheets but only four shapes: 数据汇总, the two detail
sheets, the 26 group sheets, and the Processing Report. They share a body
style — black header, centered cells, measured column widths, frozen header,
autofilter — and differ in which columns are left-aligned, which carry a text
or date format, where the pink fill goes, and whether cells are merged and
bordered.

This module only writes. Every value, ordering and grouping decision is made
before it is called, and the existing validate_merged_coupon_output() is left
untouched to judge the result: changing the writer and its oracle in the same
step would leave a failure ambiguous.

XlsxWriter's constant_memory mode is deliberately not used. It flushes each
row when the next one starts, which makes merge_range() silently do nothing —
数据汇总 merges 财务大类 and 品牌 runs after its rows are written.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from processors.common.excel import (
    FONT_SIZE,
    ROW_HEIGHT,
    pixels_to_column_pixels,
    width_measurer,
)

TEXT_FORMAT = "@"
DATE_FORMAT = "yyyy-mm-dd"
CURRENCY_FORMAT = "0.00"


@dataclass(frozen=True)
class FormatKey:
    """Every distinct cell appearance the audit workbook produces.

    The sheets between them yield only a dozen or so combinations, so caching
    on this key keeps XlsxWriter's style table as small as the openpyxl one
    was, without naming each combination by hand.
    """

    align: str = "center"
    number_format: str | None = None
    pink: bool = False
    border: bool = False
    header: bool = False


class CouponFormatCache:
    """Hands out one XlsxWriter Format per distinct FormatKey."""

    def __init__(self, workbook, font_name: str, fill_color: str) -> None:
        self._workbook = workbook
        self._font_name = font_name
        self._fill_color = fill_color
        self._formats: dict[FormatKey, object] = {}

    def get(self, key: FormatKey):
        cached = self._formats.get(key)
        if cached is not None:
            return cached

        properties: dict[str, object] = {
            "font_name": self._font_name,
            "font_size": FONT_SIZE,
            "align": key.align,
            "valign": "vcenter",
        }
        if key.header:
            properties.update(
                bold=True, font_color="#FFFFFF", bg_color="#000000"
            )
        else:
            properties["font_color"] = "#000000"
            if key.pink:
                properties["bg_color"] = f"#{self._fill_color}"
        if key.number_format is not None:
            properties["num_format"] = key.number_format
        if key.border:
            properties["border"] = 1
        cell_format = self._workbook.add_format(properties)
        self._formats[key] = cell_format
        return cell_format


def column_number_formats(
    *,
    text_columns: tuple[int, ...] = (),
    date_columns: tuple[int, ...] = (),
    currency_columns: tuple[int, ...] = (),
) -> dict[int, str]:
    """Map a 0-based column index to the number format its cells take.

    Indices, not header names: the openpyxl writers addressed these columns by
    letter (sheet["A"], sheet["E"]), and the names would not identify them
    anyway — the currency column is 2026家电国补（计入收入）合计 on one sheet
    and 2026数码国补（计入收入） on another.
    """
    return {
        **{index: TEXT_FORMAT for index in text_columns},
        **{index: DATE_FORMAT for index in date_columns},
        **{index: CURRENCY_FORMAT for index in currency_columns},
    }


def _write_table(
    workbook,
    sheet_name: str,
    header: tuple[str, ...],
    rows,
    formats: CouponFormatCache,
    measurement_font,
    *,
    left_aligned_headers: tuple[str, ...] = (),
    number_formats: dict[int, str] | None = None,
    pink_rows: frozenset[int] = frozenset(),
    bordered_rows: int = 0,
    autofilter: bool = True,
):
    """Write one sheet and return it, so callers can add merges afterwards.

    pink_rows and bordered_rows are given in 1-based sheet rows, matching how
    the openpyxl writers addressed them.
    """
    number_formats = number_formats or {}
    sheet = workbook.add_worksheet(sheet_name)
    left_columns = {
        index for index, name in enumerate(header) if name in left_aligned_headers
    }
    measure = width_measurer(measurement_font)
    maximum_widths = [measure(value) for value in header]

    sheet.set_row(0, ROW_HEIGHT)
    header_format = formats.get(FormatKey(header=True, border=bordered_rows >= 1))
    for column, value in enumerate(header):
        sheet.write(0, column, value, header_format)

    for row_number, row in enumerate(rows, start=2):
        sheet.set_row(row_number - 1, ROW_HEIGHT)
        is_pink = row_number in pink_rows
        bordered = row_number <= bordered_rows
        for column, value in enumerate(row):
            number_format = number_formats.get(column)
            if number_format == DATE_FORMAT and value in (None, ""):
                # The openpyxl writer guarded this one on the cell having a
                # value; the currency column was stamped unconditionally.
                number_format = None
            if isinstance(value, (date, datetime)) and number_format is None:
                number_format = DATE_FORMAT
            sheet.write(
                row_number - 1,
                column,
                value,
                formats.get(
                    FormatKey(
                        align="left" if column in left_columns else "center",
                        number_format=number_format,
                        pink=is_pink,
                        border=bordered,
                    )
                ),
            )
            if column < len(maximum_widths):
                maximum_widths[column] = max(maximum_widths[column], measure(value))

    sheet.freeze_panes(1, 0)
    if autofilter:
        sheet.autofilter(0, 0, len(rows), len(header) - 1)
    for column, maximum_pixels in enumerate(maximum_widths):
        sheet.set_column_pixels(
            column, column, pixels_to_column_pixels(maximum_pixels)
        )
    return sheet


def write_detail_sheet(
    workbook,
    sheet_name: str,
    header: tuple[str, ...],
    rows,
    formats: CouponFormatCache,
    measurement_font,
    *,
    left_aligned_headers: tuple[str, ...],
    matched_count: int,
) -> None:
    """A 明细总表: the last matched_count rows are pink, the rest plain."""
    first_pink_row = len(rows) + 2 - matched_count
    pink_rows = frozenset(range(first_pink_row, len(rows) + 2))
    _write_table(
        workbook,
        sheet_name,
        header,
        rows,
        formats,
        measurement_font,
        left_aligned_headers=left_aligned_headers,
        number_formats=column_number_formats(
            text_columns=(0,), date_columns=(1,)
        ),
        pink_rows=pink_rows if matched_count else frozenset(),
    )


def write_group_sheet(
    workbook,
    sheet_name: str,
    header: tuple[str, ...],
    grouped_rows,
    select_columns,
    formats: CouponFormatCache,
    measurement_font,
    *,
    left_aligned_headers: tuple[str, ...],
) -> None:
    """One 财务大类-品牌 sheet; each row is pink or not on its own."""
    rows = [select_columns(row) for row, _ in grouped_rows]
    pink_rows = frozenset(
        row_number
        for row_number, (_, is_pink) in enumerate(grouped_rows, start=2)
        if is_pink
    )
    _write_table(
        workbook,
        sheet_name,
        header,
        rows,
        formats,
        measurement_font,
        left_aligned_headers=left_aligned_headers,
        number_formats=column_number_formats(
            text_columns=(0,), date_columns=(1,)
        ),
        pink_rows=pink_rows,
    )


def write_reference_report(
    workbook,
    sheet_name: str,
    header: tuple[str, ...],
    rows,
    formats: CouponFormatCache,
    measurement_font,
) -> None:
    _write_table(
        workbook,
        sheet_name,
        header,
        rows,
        formats,
        measurement_font,
        left_aligned_headers=("说明",),
        number_formats=column_number_formats(
            text_columns=(header.index("单据号"),),
            date_columns=(header.index("单据日期"),),
        ),
    )


def write_summary_sheet(
    workbook,
    sheet_name: str,
    header: tuple[str, ...],
    rows,
    formats: CouponFormatCache,
    measurement_font,
    *,
    group_merges,
    project_merges,
) -> None:
    """数据汇总: bordered throughout, with merged 财务大类/品牌 runs.

    The merges are computed from the rows beforehand (see
    appliance.coupon_summary_group_merges), because XlsxWriter cannot read
    back what it has written the way the openpyxl version did.
    """
    bordered_rows = len(rows) + 1
    sheet = _write_table(
        workbook,
        sheet_name,
        header,
        rows,
        formats,
        measurement_font,
        number_formats=column_number_formats(
            currency_columns=(len(header) - 1,)
        ),
        bordered_rows=bordered_rows,
    )
    merged_format = formats.get(FormatKey(border=True))
    for first_row, last_row, column in group_merges:
        sheet.merge_range(
            first_row - 1,
            column - 1,
            last_row - 1,
            column - 1,
            rows[first_row - 2][column - 1],
            merged_format,
        )
    for first_row, last_row, first_column, last_column in project_merges:
        sheet.merge_range(
            first_row - 1,
            first_column - 1,
            last_row - 1,
            last_column - 1,
            rows[first_row - 2][first_column - 1],
            merged_format,
        )
