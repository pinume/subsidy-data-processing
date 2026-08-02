from pathlib import Path


def resolve_data_dir() -> Path | None:
    current_data_dir = (Path.cwd() / "data").resolve()
    if current_data_dir.is_dir():
        return current_data_dir

    if current_data_dir.exists():
        raise NotADirectoryError(
            f"Cannot create data directory; the path is a file: {current_data_dir}"
        )

    # Returning None ends the run: main() has nothing to offer until the
    # operator has put the exports into the directory just created.
    current_data_dir.mkdir()
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
