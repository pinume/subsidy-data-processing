import io
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from processors.common.paths import (
    find_data_files,
    match_source_file_by_header,
    resolve_data_dir,
    resolve_unique_file,
)


class DataDirectoryTest(unittest.TestCase):
    def test_creates_empty_flat_data_directory(self) -> None:
        original_directory = Path.cwd()
        try:
            with tempfile.TemporaryDirectory() as directory:
                os.chdir(directory)
                with redirect_stdout(io.StringIO()):
                    result = resolve_data_dir()
                os.chdir(original_directory)

                data_dir = Path(directory) / "data"
                self.assertIsNone(result)
                self.assertTrue(data_dir.is_dir())
                self.assertEqual(list(data_dir.iterdir()), [])
        finally:
            os.chdir(original_directory)


class FindDataFilesTest(unittest.TestCase):
    def test_filters_by_keyword_and_suffix_and_skips_temp_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            matching = data_dir / "收款单统计.XLS"
            wrong_suffix = data_dir / "收款单统计.xlsx"
            wrong_keyword = data_dir / "销售用券情况统计.XLS"
            lock_file = data_dir / "~$收款单统计.XLS"
            for path in (matching, wrong_suffix, wrong_keyword, lock_file):
                path.touch()

            self.assertEqual(
                find_data_files(data_dir, "收款单统计", (".xls",)),
                [matching],
            )

    def test_is_case_insensitive_on_suffix_and_sorts_results(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            second = data_dir / "MER_2_export.xlsx"
            first = data_dir / "MER_1_export.xlsx"
            second.touch()
            first.touch()

            self.assertEqual(
                find_data_files(data_dir, "MER_", (".XLSX",)),
                [first, second],
            )


class ResolveUniqueFileTest(unittest.TestCase):
    def test_returns_none_when_no_candidates(self) -> None:
        self.assertIsNone(resolve_unique_file([]))

    def test_returns_the_sole_candidate(self) -> None:
        candidate = Path("only.xlsx")
        self.assertEqual(resolve_unique_file([candidate]), candidate)

    def test_raises_on_multiple_candidates(self) -> None:
        with self.assertRaises(ValueError):
            resolve_unique_file([Path("a.xlsx"), Path("b.xlsx")])


class MatchSourceFileByHeaderTest(unittest.TestCase):
    def test_returns_the_candidate_whose_header_matches(self) -> None:
        digital = Path("digital.XLS")
        large_appliances = Path("large_appliances.XLS")
        headers = {
            digital: "2026数码国补（计入收入）",
            large_appliances: "2026家电国补（计入收入）",
        }

        result = match_source_file_by_header(
            [digital, large_appliances],
            "2026家电国补（计入收入）",
            read_header=headers.__getitem__,
        )

        self.assertEqual(result, large_appliances)

    def test_returns_none_when_no_header_matches(self) -> None:
        candidate = Path("digital.XLS")

        result = match_source_file_by_header(
            [candidate],
            "2026家电国补（计入收入）",
            read_header=lambda _path: "2026数码国补（计入收入）",
        )

        self.assertIsNone(result)

    def test_raises_when_multiple_candidates_share_the_header(self) -> None:
        candidates = [Path("first.XLS"), Path("second.XLS")]

        with self.assertRaises(ValueError):
            match_source_file_by_header(
                candidates,
                "2026家电国补（计入收入）",
                read_header=lambda _path: "2026家电国补（计入收入）",
            )


if __name__ == "__main__":
    unittest.main()
