import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import main as app_main


class CombinedOutputRollbackTest(unittest.TestCase):
    def test_submitted_outputs_are_restored_when_second_project_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            appliance_output = output_dir / "家电_已上传.xlsx"
            digital_output = output_dir / "数码_已上传.xlsx"
            appliance_output.write_bytes(b"old appliance")
            digital_output.write_bytes(b"old digital")

            def write_appliance() -> None:
                appliance_output.write_bytes(b"new appliance")

            def fail_digital() -> None:
                digital_output.write_bytes(b"new digital")
                raise ValueError("digital failed")

            with (
                patch.object(
                    app_main.large_appliances,
                    "OUTPUT_FILE",
                    appliance_output,
                ),
                patch.object(
                    app_main.digital,
                    "OUTPUT_FILE",
                    digital_output,
                ),
                patch.object(
                    app_main.large_appliances,
                    "process_submitted_files",
                    write_appliance,
                ),
                patch.object(
                    app_main.digital,
                    "process_submitted_files",
                    fail_digital,
                ),
            ):
                with self.assertRaisesRegex(ValueError, "digital failed"):
                    app_main.process_submitted_files()

            self.assertEqual(appliance_output.read_bytes(), b"old appliance")
            self.assertEqual(digital_output.read_bytes(), b"old digital")


class MainErrorHandlingTest(unittest.TestCase):
    def test_configuration_failure_uses_normal_error_report(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            with (
                patch.object(app_main, "resolve_data_dir", return_value=data_dir),
                patch.object(
                    app_main.digital,
                    "configure_data_dir",
                    side_effect=ValueError("bad config"),
                ),
                patch.object(app_main, "report_failure") as report_failure,
            ):
                self.assertEqual(app_main.main(), 1)

            report_failure.assert_called_once()
            self.assertEqual(str(report_failure.call_args.args[0]), "bad config")


if __name__ == "__main__":
    unittest.main()
