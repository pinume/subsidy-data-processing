"""A read-back view of a written workbook, assembled from three readers.

The audit output is validated by re-reading it, and that re-read used to be a
single non-read_only openpyxl load: 2.87s of the mode's 9.6s, because it
materializes a Cell object for all 175k cells across 30 sheets just so the
checks can reach the few hundred they care about.

No one reader replaces it. The work splits by what each library can actually
see:

- values, sheet order and merged ranges — calamine, which parses the whole
  file in Rust in 0.19s.
- number formats, fills, fonts, alignment and borders — openpyxl, the only
  one of the two that exposes styles at all. It stays out of this module,
  because holding a style snapshot for every cell would cost more memory than
  the load it replaces; the validators walk read_only rows and check as they
  go.
- freeze panes, autofilter range and column widths — neither library offers
  these in read_only mode (ReadOnlyWorksheet has no freeze_panes and no
  merged_cells), so they come from one pass over the sheet XML.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from python_calamine import CalamineWorkbook

from processors.common.excel import calamine_rows

# Attribute order is not fixed by the format and the two writers disagree on
# it — openpyxl emits Target before Id, XlsxWriter the reverse — so these match
# the whole tag and the attributes are pulled out by name.
_SHEET_RE = re.compile(r"<sheet ([^>]*?)/?>")
_REL_RE = re.compile(r"<Relationship ([^>]*?)/?>")
_PANE_RE = re.compile(r"<pane ([^>]*)/>")
_AUTOFILTER_RE = re.compile(r'<autoFilter ref="([^"]+)"')
_COL_RE = re.compile(r"<col ([^>]*)/>")
_ATTR_RE = re.compile(r'(\w+)="([^"]*)"')


def column_letter(index: int) -> str:
    """0-based column index to its spreadsheet letter."""
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def a1_range(start: tuple[int, int], end: tuple[int, int]) -> str:
    """calamine's ((row, column), (row, column)) pair as "A1:B5"."""
    first = f"{column_letter(start[1])}{start[0] + 1}"
    last = f"{column_letter(end[1])}{end[0] + 1}"
    return f"{first}:{last}"


@dataclass(frozen=True)
class SheetMetadata:
    """Sheet properties that only the raw XML still carries in a fast read."""

    freeze_panes: str | None
    autofilter_ref: str | None
    column_widths: dict[int, float]


@dataclass(frozen=True)
class WorkbookSnapshot:
    sheet_names: tuple[str, ...]
    values: dict[str, list[tuple[object, ...]]]
    merged_ranges: dict[str, frozenset[str]]
    metadata: dict[str, SheetMetadata]

    def rows(self, sheet_name: str) -> list[tuple[object, ...]]:
        return self.values[sheet_name]


def _attributes(text: str) -> dict[str, str]:
    return dict(_ATTR_RE.findall(text))


def _sheet_xml_paths(archive: ZipFile) -> dict[str, str]:
    """Map each sheet name to its part inside the archive, in workbook order."""
    workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    relationships_xml = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {}
    for tag in _REL_RE.findall(relationships_xml):
        attributes = _attributes(tag)
        if "Id" in attributes and "Target" in attributes:
            targets[attributes["Id"]] = attributes["Target"]

    paths: dict[str, str] = {}
    for tag in _SHEET_RE.findall(workbook_xml):
        attributes = _attributes(tag)
        name = attributes.get("name")
        relationship_id = attributes.get("r:id") or attributes.get("id")
        if name is None or relationship_id is None:
            continue
        target = targets[relationship_id].replace("\\", "/").lstrip("/")
        paths[name] = target if target.startswith("xl/") else f"xl/{target}"
    return paths


def _read_metadata(sheet_xml: str) -> SheetMetadata:
    freeze_panes = None
    # Only the header of the XML holds <pane> and <cols>; <autoFilter> sits
    # after the row data, so the whole part has to be searched for it.
    header = sheet_xml[: sheet_xml.find("<sheetData")]
    pane = _PANE_RE.search(header)
    if pane is not None:
        attributes = _attributes(pane.group(1))
        if attributes.get("state") == "frozen":
            freeze_panes = attributes.get("topLeftCell")

    autofilter = _AUTOFILTER_RE.search(sheet_xml)
    widths: dict[int, float] = {}
    for column in _COL_RE.findall(header):
        attributes = _attributes(column)
        width = attributes.get("width")
        if width is None:
            continue
        for index in range(int(attributes["min"]), int(attributes["max"]) + 1):
            widths[index] = float(width)
    return SheetMetadata(
        freeze_panes=freeze_panes,
        autofilter_ref=autofilter.group(1) if autofilter else None,
        column_widths=widths,
    )


def read_workbook_snapshot(path: Path) -> WorkbookSnapshot:
    """Values, merged ranges and sheet metadata, without building any cells."""
    workbook = CalamineWorkbook.from_path(str(path))
    try:
        sheet_names = tuple(workbook.sheet_names)
        values: dict[str, list[tuple[object, ...]]] = {}
        merged: dict[str, frozenset[str]] = {}
        for name in sheet_names:
            sheet = workbook.get_sheet_by_name(name)
            values[name] = [tuple(row) for row in calamine_rows(sheet)]
            merged[name] = frozenset(
                a1_range(start, end) for start, end in sheet.merged_cell_ranges
            )
    finally:
        workbook.close()

    with ZipFile(path) as archive:
        paths = _sheet_xml_paths(archive)
        metadata = {
            name: _read_metadata(archive.read(paths[name]).decode("utf-8"))
            for name in sheet_names
        }

    return WorkbookSnapshot(
        sheet_names=sheet_names,
        values=values,
        merged_ranges=merged,
        metadata=metadata,
    )
