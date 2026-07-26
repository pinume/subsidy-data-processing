from collections.abc import Callable
from pathlib import Path

import xlrd


def resolve_data_dir() -> Path | None:
    current_data_dir = (Path.cwd() / "data").resolve()
    if current_data_dir.is_dir():
        print(f"Found data directory: {current_data_dir}")
        return current_data_dir

    if current_data_dir.exists():
        raise NotADirectoryError(
            f"Cannot create data directory; the path is a file: {current_data_dir}"
        )

    current_data_dir.mkdir()
    print(f"Created data directory: {current_data_dir}")
    print(
        "Add the required files directly into this directory (no subfolders); "
        "each project identifies its files by filename keywords. "
        "See README.md for the exact naming rules."
    )
    print("This run has ended.")
    return None


def find_data_files(
    data_dir: Path,
    keyword: str,
    suffixes: tuple[str, ...],
) -> list[Path]:
    """List files directly under data_dir matching keyword and suffix.

    Excel leaves ~$-prefixed lock files next to open workbooks, so those are
    always skipped.
    """
    lowered_suffixes = {suffix.lower() for suffix in suffixes}
    return sorted(
        path
        for path in data_dir.iterdir()
        if path.is_file()
        and not path.name.startswith("~$")
        and keyword in path.name
        and path.suffix.lower() in lowered_suffixes
    )


def resolve_unique_file(candidates: list[Path]) -> Path | None:
    """Return the sole candidate, or None; refuse to silently pick one of many.

    A stale duplicate left in the data directory is an operator mistake, not
    a choice the program should make on its own.
    """
    if not candidates:
        return None
    if len(candidates) > 1:
        raise ValueError(
            f"找到多个符合条件的文件，无法确定使用哪一个：{candidates}"
        )
    return candidates[0]


def read_xls_header(path: Path, *, row: int, column: int) -> str:
    """Read a single header cell from a legacy .xls workbook (1-based row/column)."""
    workbook = xlrd.open_workbook(path)
    try:
        sheet = workbook.sheet_by_index(0)
        if sheet.nrows < row or sheet.ncols < column:
            return ""
        return str(sheet.cell_value(row - 1, column - 1)).strip()
    finally:
        workbook.release_resources()


def match_source_file_by_header(
    candidates: list[Path],
    expected_header: str,
    *,
    read_header: Callable[[Path], str],
) -> Path | None:
    """Pick the candidate whose header cell equals expected_header.

    Two files can share the same filename keyword (e.g. the digital and large
    appliances coupon exports both contain "销售用券情况统计"), so content is
    what actually distinguishes them.
    """
    matches = [
        candidate
        for candidate in candidates
        if read_header(candidate) == expected_header
    ]
    if not matches:
        return None
    if len(matches) > 1:
        raise ValueError(
            f"找到多个表头为“{expected_header}”的文件，无法确定使用哪一个：{matches}"
        )
    return matches[0]
