import io
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from processors.common.paths import (
    DATA_SUBDIRECTORIES,
    resolve_data_dir,
    resolve_existing_data_file,
)


class DataDirectoryTest(unittest.TestCase):
    def test_creates_current_standard_directory_structure(self) -> None:
        original_directory = Path.cwd()
        try:
            with tempfile.TemporaryDirectory() as directory:
                os.chdir(directory)
                with patch("builtins.input", return_value="2"):
                    with redirect_stdout(io.StringIO()):
                        result = resolve_data_dir()
                os.chdir(original_directory)

                data_dir = Path(directory) / "data"
                self.assertIsNone(result)
                self.assertTrue(data_dir.is_dir())
                self.assertEqual(
                    {path.name for path in data_dir.iterdir()},
                    set(DATA_SUBDIRECTORIES),
                )
                self.assertTrue(
                    (data_dir / "reference_number_supplement").is_dir()
                )
                self.assertFalse((data_dir / "invoice").exists())
        finally:
            os.chdir(original_directory)

    def test_resolves_first_existing_data_file_candidate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            chinese_file = data_dir / "收款单统计" / "收款单统计.XLS"
            chinese_file.parent.mkdir()
            chinese_file.touch()

            self.assertEqual(
                resolve_existing_data_file(
                    data_dir,
                    (
                        Path("receipt_statistics") / "receipt_statistics.XLS",
                        Path("收款单统计") / "收款单统计.XLS",
                    ),
                ),
                chinese_file,
            )

    def test_prefers_standard_data_file_candidate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            standard_file = data_dir / "receipt_statistics" / "receipt_statistics.XLS"
            chinese_file = data_dir / "收款单统计" / "收款单统计.XLS"
            standard_file.parent.mkdir()
            chinese_file.parent.mkdir()
            standard_file.touch()
            chinese_file.touch()

            self.assertEqual(
                resolve_existing_data_file(
                    data_dir,
                    (
                        Path("receipt_statistics") / "receipt_statistics.XLS",
                        Path("收款单统计") / "收款单统计.XLS",
                    ),
                ),
                standard_file,
            )


if __name__ == "__main__":
    unittest.main()
