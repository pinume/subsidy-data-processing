import os
import re
import shutil
import stat
import time
from collections.abc import Callable, Iterator
from copy import copy
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree.ElementTree import iterparse
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import ImageFont


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TAG = f"{{{MAIN_NS}}}"
CELL_REF_RE = re.compile(r"([A-Z]+)")
FONT_NAME = "Maple Mono NF CN"
FALLBACK_FONT_NAME = "微软雅黑"
FONT_SIZE = 11
ROW_HEIGHT = 18
FONT_PATH = Path(r"C:\Windows\Fonts\MapleMono-NF-CN-Regular.ttf")
FALLBACK_FONT_PATH = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_CANDIDATES = (
    (
        FONT_NAME,
        (
            FONT_PATH,
            Path.home()
            / ".local/share/fonts/MapleMono-NF-CN/MapleMono-NF-CN-Regular.ttf",
        ),
        (
            "MapleMono-NF-CN-Regular.ttf",
            "MapleMonoNormal-NF-CN-Regular.ttf",
        ),
    ),
    (
        FALLBACK_FONT_NAME,
        (
            FALLBACK_FONT_PATH,
            Path("/usr/local/share/fonts/github-fonts/MSYH.TTC"),
            Path("/usr/local/share/fonts/github-fonts/MSYHL.TTC"),
        ),
        (
            "msyh.ttc",
            "msyh.ttf",
            "MSYH.TTC",
            "MSYHL.TTC",
        ),
    ),
    (
        "Noto Sans CJK SC",
        (
            Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
            Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        ),
        (
            "NotoSansCJK-Regular.ttc",
            "NotoSansSC-Regular.otf",
            "NotoSansCJKsc-Regular.otf",
        ),
    ),
    (
        "Source Han Sans SC",
        (),
        (
            "SourceHanSansSC-Regular.otf",
            "SourceHanSansCN-Regular.otf",
        ),
    ),
    (
        "PingFang SC",
        (Path("/System/Library/Fonts/PingFang.ttc"),),
        ("PingFang.ttc",),
    ),
    (
        "WenQuanYi Micro Hei",
        (),
        (
            "wqy-microhei.ttc",
            "WenQuanYi Micro Hei.ttf",
        ),
    ),
)
FONT_SEARCH_ROOTS = (
    Path.home() / ".local/share/fonts",
    Path.home() / ".fonts",
    Path.home() / "Library/Fonts",
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts"),
    Path(r"C:\Windows\Fonts"),
)


def _shared_strings(archive: ZipFile) -> list[str]:
    try:
        stream = archive.open("xl/sharedStrings.xml")
    except KeyError:
        return []

    values: list[str] = []
    with stream:
        for _, element in iterparse(stream, events=("end",)):
            if element.tag == TAG + "si":
                values.append(
                    "".join(node.text or "" for node in element.iter(TAG + "t"))
                )
                element.clear()
    return values


def _first_sheet_path(archive: ZipFile) -> str:
    from xml.etree.ElementTree import fromstring

    workbook = fromstring(archive.read("xl/workbook.xml"))
    sheet = workbook.find(f"{TAG}sheets/{TAG}sheet")
    if sheet is None:
        raise ValueError("The workbook contains no worksheets")

    relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
    relationships = fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    for relationship in relationships:
        if (
            relationship.tag == f"{{{package_rel_ns}}}Relationship"
            and relationship.attrib.get("Id") == relationship_id
        ):
            target = relationship.attrib["Target"].replace("\\", "/").lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise ValueError("Unable to locate the first worksheet")


def _cell_value(cell, shared_strings: list[str]):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(TAG + "t"))

    value_node = cell.find(TAG + "v")
    if value_node is None or value_node.text is None:
        return None
    value = value_node.text

    if cell_type == "s":
        return shared_strings[int(value)]
    if cell_type == "b":
        return value == "1"
    if cell_type in {"str", "e"}:
        return value
    return value


def read_rows(path: Path) -> Iterator[list[object]]:
    """Stream rows from the first sheet, ignoring an incorrect dimension value."""
    with ZipFile(path) as archive:
        shared_strings = _shared_strings(archive)
        sheet_path = _first_sheet_path(archive)
        with archive.open(sheet_path) as stream:
            for _, row in iterparse(stream, events=("end",)):
                if row.tag != TAG + "row":
                    continue

                cells: dict[int, object] = {}
                max_column = 0
                for cell in row.findall(TAG + "c"):
                    reference = cell.attrib.get("r", "")
                    match = CELL_REF_RE.match(reference)
                    if not match:
                        continue
                    column = column_index_from_string(match.group(1))
                    cells[column] = _cell_value(cell, shared_strings)
                    max_column = max(max_column, column)

                yield [cells.get(column) for column in range(1, max_column + 1)]
                row.clear()


def _find_font_file(file_names: tuple[str, ...]) -> Path | None:
    for root in FONT_SEARCH_ROOTS:
        if not root.exists():
            continue

        for file_name in file_names:
            direct_path = root / file_name
            if direct_path.exists():
                return direct_path

        for file_name in file_names:
            try:
                return next(root.rglob(file_name))
            except (OSError, StopIteration):
                continue
    return None


def resolve_font() -> tuple[str, Path]:
    configured_path = os.environ.get("UPLOAD_DATA_FONT_PATH")
    if configured_path:
        font_path = Path(configured_path).expanduser()
        if not font_path.exists():
            raise FileNotFoundError(f"Configured font file does not exist: {font_path}")
        return os.environ.get("UPLOAD_DATA_FONT_NAME", font_path.stem), font_path

    for font_name, paths, file_names in FONT_CANDIDATES:
        for path in paths:
            if path.exists():
                return font_name, path

        font_path = _find_font_file(file_names)
        if font_path is not None:
            return font_name, font_path

    raise FileNotFoundError(
        "No supported font file was found. Install Maple Mono NF CN, Microsoft "
        "YaHei, Noto Sans CJK SC, or PingFang SC, or set UPLOAD_DATA_FONT_PATH."
    )


def load_measurement_font(font_path: Path):
    return ImageFont.truetype(str(font_path), size=15)


def measurement_text(value: object) -> str:
    if value is None:
        return ""
    return value.strftime("%Y%m%d") if isinstance(value, (date, datetime)) else str(value)


def text_pixel_width(value: object, font) -> float:
    text = measurement_text(value)
    if not text:
        return 0
    # getlength (the rendered advance width) rather than getbbox (the ink
    # extent) — about 3x faster per call, and it is what Excel's own
    # character-width metric approximates anyway. On a proportional font the
    # two differ by a fraction of a pixel per string, well inside the ±10%
    # slack pixels_to_excel_width already adds; on the monospace font this
    # project prefers (Maple Mono NF CN) they are identical.
    return font.getlength(text)


# Strings whose width would stop being the sum of their characters' widths the
# moment a font kerned, ligated, or shaped them: classic kerning pairs, an "ffi"
# ligature candidate, and the character classes this program actually writes
# (digits and uppercase in SN/IMEI/invoice numbers, CJK punctuation, runs of
# Chinese). A font passes only if every one of them measures additively.
WIDTH_ADDITIVITY_PROBES = (
    "AV",
    "To",
    "Ta",
    "ffi",
    "0123456789",
    "SN2026ABCD",
    "，。、；：",
    "连续中文字符",
    "中文A1，混排",
)


def widths_are_additive(value_font) -> bool:
    """Whether a string's width always equals the sum of its characters'.

    True for the monospace font this project prefers, false for proportional
    fonts such as 微软雅黑, which kern. Probing the font itself rather than
    trusting its name keeps any unrecognised font on the exact, slower path.
    """
    for probe in WIDTH_ADDITIVITY_PROBES:
        summed = sum(value_font.getlength(character) for character in probe)
        if value_font.getlength(probe) != summed:
            return False
    return True


def width_measurer(value_font) -> Callable[[object], float]:
    """Measure text width, caching by rendered text.

    Sheets repeat dates, statuses, and remarks heavily, but SN/IMEI/invoice
    columns are effectively all-distinct, so caching by whole string cannot
    bound the number of font calls. When the font measures additively (see
    widths_are_additive) a string's width can be summed from its characters
    instead, which replaces that unbounded set of distinct strings with the
    bounded set of distinct characters. Fonts that kern keep the whole-string
    measurement, so their output is unchanged.
    """
    cache: dict[str, float] = {}

    if not widths_are_additive(value_font):

        def measure(value: object) -> float:
            text = measurement_text(value)
            if not text:
                return 0
            width = cache.get(text)
            if width is None:
                width = value_font.getlength(text)
                cache[text] = width
            return width

        return measure

    character_widths: dict[str, float] = {}

    def measure(value: object) -> float:
        text = measurement_text(value)
        if not text:
            return 0
        width = cache.get(text)
        if width is None:
            width = 0.0
            for character in text:
                character_width = character_widths.get(character)
                if character_width is None:
                    character_width = value_font.getlength(character)
                    character_widths[character] = character_width
                width += character_width
            cache[text] = width
        return width

    return measure


def pixels_to_excel_width(pixels: float) -> float:
    return min(round(((pixels * 1.1) + 16) / 7, 2), 255)


# Reusing an already-computed cell style has no public openpyxl API, so these
# three helpers reach into Cell._style. They are the only place in the project
# that does; StyleReuseTest pins their behaviour so that an openpyxl upgrade
# which changes the attribute fails loudly instead of silently corrupting
# styles or quietly losing the speedup.
def style_snapshot(cell) -> tuple | None:
    """An immutable key describing the cell's style right now.

    Deliberately not the StyleArray itself: it is mutable and the assignments
    that follow keep writing to it, so a live reference used as a dict key
    would change value underneath the dict.
    """
    style = cell._style
    return tuple(style) if style is not None else None


def capture_style(cell):
    """Detach a finished style so later cells can reuse it."""
    return copy(cell._style)


def reuse_style(cell, style) -> None:
    """Give the cell its own copy of a previously computed style.

    A copy, never the shared instance. StyleArray is mutable, so cells sharing
    one instance are not merely equal but linked: code that later sets a fill
    or a number format on any one of them would change all of them.
    """
    cell._style = copy(style)


def create_sheet_styles(font_name: str):
    normal_font = Font(name=font_name, size=FONT_SIZE, color="000000")
    header_font = Font(name=font_name, size=FONT_SIZE, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="000000")
    centered = Alignment(horizontal="center", vertical="center")
    return normal_font, header_font, header_fill, centered


def format_sheet(
    sheet,
    font_name: str,
    measurement_font,
    left_aligned_headers: tuple[str, ...] = ("描述",),
) -> None:
    column_count = sheet.max_column
    normal_font, header_font, header_fill, centered = create_sheet_styles(font_name)
    left_aligned = Alignment(horizontal="left", vertical="center")
    left_aligned_columns: set[int] = set()
    subsidy_column = None
    for cell in sheet[1]:
        if cell.value in left_aligned_headers:
            left_aligned_columns.add(cell.column)
        if cell.value == "补贴金额":
            subsidy_column = cell.column
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = centered

    measure = width_measurer(measurement_font)
    maximum_pixel_widths = [measure(cell.value) for cell in sheet[1]]
    # Setting font/alignment/number_format re-registers each value in the
    # workbook's style tables, and that lookup hashes the whole style object —
    # by far the most expensive thing this function does on a large sheet. The
    # body cells only ever produce a handful of distinct results, so each
    # distinct one is computed once, by the ordinary assignments below, and
    # afterwards reused. Two separate hazards shape how that is done:
    #
    # Styles set before this function runs (payment's 汇总 already carries
    # #,##0.00) must survive, because the per-cell assignments would have left
    # them alone. That is why the cell's incoming style is part of the key
    # rather than something to overwrite.
    #
    # Styles set after this function returns (the coupon pipelines' pink fill)
    # must not leak between cells. That is why reuse_style hands out a copy: a
    # shared StyleArray would make one later fill assignment repaint every cell
    # that shares it.
    styles: dict[tuple, object] = {}

    for row_number, row in enumerate(
        sheet.iter_rows(min_row=2, max_row=sheet.max_row), start=2
    ):
        sheet.row_dimensions[row_number].height = ROW_HEIGHT
        for cell in row:
            is_left_aligned = cell.column in left_aligned_columns
            is_subsidy = cell.column == subsidy_column and cell.value is not None
            key = (style_snapshot(cell), is_left_aligned, is_subsidy)
            style = styles.get(key)
            if style is None:
                cell.font = normal_font
                cell.alignment = left_aligned if is_left_aligned else centered
                if is_subsidy:
                    cell.number_format = "0.00"
                styles[key] = capture_style(cell)
            else:
                reuse_style(cell, style)
            width = measure(cell.value)
            if width > maximum_pixel_widths[cell.column - 1]:
                maximum_pixel_widths[cell.column - 1] = width

    sheet.row_dimensions[1].height = ROW_HEIGHT
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(column_count)}{sheet.max_row}"
    for column, maximum_pixels in enumerate(maximum_pixel_widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = (
            pixels_to_excel_width(maximum_pixels)
        )


STALE_TEMPORARY_FILE_AGE_SECONDS = 180


def remove_stale_temporary_files(
    output_dir: Path,
    minimum_age_seconds: float = STALE_TEMPORARY_FILE_AGE_SECONDS,
) -> list[str]:
    """Delete leftover intermediate files from an interrupted earlier run.

    Everything this program writes into the output directory as an
    intermediate is dot-prefixed: save_workbook_atomically's ".<名字>-<随机>"
    temporary file. It is removed on a normal run; a crash or a killed
    process leaves it behind, where it accumulates invisibly and still holds
    business data. Excel's own lock files are named "~$...", so they are
    never touched.

    Only files older than minimum_age_seconds are removed. This program was
    never designed for two instances to run against the same output
    directory at once, but without an age check this cleanup would make that
    actively unsafe instead of merely unsupported: a second instance
    starting up would delete the first instance's temporary file while it is
    still being written, and the first instance's own save would then fail
    trying to rename a file that no longer exists. The age check does not
    make concurrent runs supported — it only keeps a second instance's
    startup from corrupting a first instance's in-flight save. The default
    (180s) is comfortably longer than any single sheet this program writes
    has been observed to take to save.
    """
    if not output_dir.is_dir():
        return []

    now = time.time()
    removed: list[str] = []
    for path in sorted(output_dir.iterdir()):
        if not path.is_file() or not path.name.startswith("."):
            continue
        try:
            age_seconds = now - path.stat().st_mtime
        except OSError:
            continue
        if age_seconds < minimum_age_seconds:
            continue
        try:
            path.chmod(path.stat().st_mode | stat.S_IWRITE)
            path.unlink()
        except OSError as error:
            print(f"Could not remove leftover file {path.name}: {error}")
            continue
        removed.append(path.name)
    return removed


def save_workbook_atomically(
    workbook: Workbook,
    output_path: Path,
    validator: Callable[[Path], None],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with NamedTemporaryFile(
            prefix=f".{output_path.stem}-",
            suffix=output_path.suffix,
            dir=output_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        workbook.save(temporary_path)
        validator(temporary_path)
        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def run_with_output_rollback(
    output_paths: tuple[Path, ...],
    operation: Callable[[], object],
) -> object:
    """Restore every managed output if a multi-file operation fails.

    Individual workbooks are already written atomically. This adds the missing
    transaction boundary around commands that intentionally produce several
    files: an error in a later step must not leave earlier outputs from the new
    run beside older outputs from the previous run.
    """
    unique_paths = tuple(dict.fromkeys(output_paths))
    backups: dict[Path, Path | None] = {}
    temporary_backups: set[Path] = set()
    try:
        for output_path in unique_paths:
            if not output_path.exists():
                backups[output_path] = None
                continue
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with NamedTemporaryFile(
                prefix=f".{output_path.stem}-rollback-",
                suffix=output_path.suffix,
                dir=output_path.parent,
                delete=False,
            ) as backup_file:
                backup_path = Path(backup_file.name)
            temporary_backups.add(backup_path)
            shutil.copy2(output_path, backup_path)
            backups[output_path] = backup_path

        return operation()
    except BaseException as operation_error:
        rollback_errors: list[str] = []
        for output_path, backup_path in backups.items():
            try:
                if backup_path is None:
                    output_path.unlink(missing_ok=True)
                else:
                    os.replace(backup_path, output_path)
                    backups[output_path] = None
                    temporary_backups.discard(backup_path)
            except OSError as rollback_error:
                rollback_errors.append(f"{output_path.name}: {rollback_error}")
        if rollback_errors:
            raise RuntimeError(
                "处理失败且输出回滚不完整："
                + "；".join(rollback_errors)
            ) from operation_error
        raise
    finally:
        for backup_path in temporary_backups:
            backup_path.unlink(missing_ok=True)
