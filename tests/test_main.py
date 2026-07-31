import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import main as app_main
from processors import coupon_report


class CombinedOutputRollbackTest(unittest.TestCase):
    def test_submitted_outputs_are_restored_when_second_project_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            appliance_output = output_dir / "家电_已上传.xlsx"
            digital_output = output_dir / "数码_已上传.xlsx"
            appliance_output.write_bytes(b"old appliance")
            digital_output.write_bytes(b"old digital")

            def fake_process(profile_name: str) -> None:
                if profile_name == "家电":
                    appliance_output.write_bytes(b"new appliance")
                else:
                    digital_output.write_bytes(b"new digital")
                    raise ValueError("digital failed")

            with (
                patch.object(
                    app_main.submitted,
                    "OUTPUT_FILES",
                    (appliance_output, digital_output),
                ),
                patch.object(
                    app_main.submitted,
                    "process_submitted_files",
                    fake_process,
                ),
            ):
                with self.assertRaisesRegex(ValueError, "digital failed"):
                    app_main.submitted.process_all()

            self.assertEqual(appliance_output.read_bytes(), b"old appliance")
            self.assertEqual(digital_output.read_bytes(), b"old digital")


class AllOutputFilesTest(unittest.TestCase):
    def test_includes_the_store_report_output(self) -> None:
        self.assertIn(app_main.store_report.OUTPUT_FILE, app_main.all_output_files())


class ProcessorOrderTest(unittest.TestCase):
    def test_store_report_runs_after_its_two_upstream_processing_modes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            app_main.submitted.configure_data_dir(data_dir)
            app_main.receipts.configure_data_dir(data_dir)
            app_main.digital.configure_data_dir(data_dir)
            app_main.large_appliances.configure_data_dir(data_dir)
            app_main.payment.configure_data_dir(data_dir)
            app_main.store_report.configure_data_dir(data_dir)

            labels = [label for label, _, _ in app_main.build_processors()]
            store_report_index = labels.index("门店国补上传及回款情况表")

            self.assertLess(labels.index("审核明细（销售用券情况统计）"), store_report_index)
            self.assertLess(labels.index("回款明细（家电+数码）"), store_report_index)


class AllModeRollbackTest(unittest.TestCase):
    def test_every_real_output_target_is_restored_when_store_report_fails_last(self) -> None:
        """Exercises the actual all_output_files() rollback set — not just the two
        files a given step happens to touch — the way `all` mode really runs."""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            paths = {
                "large_appliances": output_dir / "家电_已上传.xlsx",
                "digital": output_dir / "数码_已上传.xlsx",
                "receipts": output_dir / "收款单统计.xlsx",
                "coupon_report": output_dir / "审核明细.xlsx",
                "payment": output_dir / "回款明细.xlsx",
                "store_report": output_dir / "门店报表.xlsx",
            }
            for key, path in paths.items():
                path.write_bytes(f"old {key}".encode())

            def write_ok(key: str) -> None:
                paths[key].write_bytes(f"new {key}".encode())

            def fail_store_report() -> None:
                paths["store_report"].write_bytes(b"corrupted partial store report")
                raise ValueError("store report failed")

            processors = (
                ("已上传数据（家电+数码）", paths["large_appliances"], lambda: write_ok("large_appliances")),
                ("审核明细（销售用券情况统计）", paths["coupon_report"], lambda: write_ok("coupon_report")),
                ("回款明细（家电+数码）", paths["payment"], lambda: write_ok("payment")),
                ("门店国补上传及回款情况表", paths["store_report"], fail_store_report),
            )

            with (
                patch.object(
                    app_main.submitted,
                    "OUTPUT_FILES",
                    (paths["large_appliances"], paths["digital"]),
                ),
                patch.object(app_main.receipts, "OUTPUT_FILE", paths["receipts"]),
                patch.object(coupon_report, "OUTPUT_FILE", paths["coupon_report"]),
                patch.object(app_main.payment, "OUTPUT_FILE", paths["payment"]),
                patch.object(app_main.store_report, "OUTPUT_FILE", paths["store_report"]),
            ):
                with self.assertRaisesRegex(ValueError, "store report failed"):
                    app_main.process_all(processors)

            for key, path in paths.items():
                self.assertEqual(path.read_bytes(), f"old {key}".encode())


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
