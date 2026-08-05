import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import main as app_main
from processors import coupon_report
from processors.common.console import ConsoleReporter
from processors.common.excel import OutputCleanupError, StaleFileCleanup


class CombinedOutputRollbackTest(unittest.TestCase):
    def test_submitted_outputs_are_restored_when_second_project_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            appliance_output = output_dir / "家电_已上传.xlsx"
            digital_output = output_dir / "数码_已上传.xlsx"
            appliance_output.write_bytes(b"old appliance")
            digital_output.write_bytes(b"old digital")

            def fake_process(profile_name: str, reporter) -> None:
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
                    app_main.submitted.process_all(
                        ConsoleReporter(stream=io.StringIO())
                    )

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
            app_main.coupon_sources.configure_data_dir(data_dir)
            app_main.payment.configure_data_dir(data_dir)
            app_main.store_report.configure_data_dir(data_dir)

            labels = [label for label, _, _, _ in app_main.build_processors()]
            coupon_report_index = labels.index("审核明细（销售用券情况统计）")
            store_report_index = labels.index("门店国补上传及回款情况表")

            self.assertLess(labels.index("回款明细（家电+数码）"), coupon_report_index)
            self.assertLess(coupon_report_index, store_report_index)


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

            def write_ok(key: str):
                return lambda reporter: paths[key].write_bytes(
                    f"new {key}".encode()
                )

            def fail_store_report(reporter) -> None:
                paths["store_report"].write_bytes(b"corrupted partial store report")
                raise ValueError("store report failed")

            processors = (
                ("已上传数据（家电+数码）", "已上传数据", paths["large_appliances"], write_ok("large_appliances")),
                ("审核明细（销售用券情况统计）", "审核明细", paths["coupon_report"], write_ok("coupon_report")),
                ("回款明细（家电+数码）", "回款明细", paths["payment"], write_ok("payment")),
                ("门店国补上传及回款情况表", "门店报表", paths["store_report"], fail_store_report),
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
                    app_main.process_all(
                        processors,
                        ConsoleReporter(stream=io.StringIO()),
                    )

            for key, path in paths.items():
                self.assertEqual(path.read_bytes(), f"old {key}".encode())


class MainErrorHandlingTest(unittest.TestCase):
    def test_configuration_failure_uses_normal_error_report(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory)
            with (
                patch.object(app_main, "resolve_data_dir", return_value=data_dir),
                patch.object(
                    app_main.coupon_sources,
                    "configure_data_dir",
                    side_effect=ValueError("bad config"),
                ),
                patch.object(
                    app_main.ConsoleReporter, "error"
                ) as error,
            ):
                self.assertEqual(app_main.main(), 1)

            error.assert_called_once()
            self.assertEqual(str(error.call_args.args[1]), "bad config")


class TransactionLifecycleTest(unittest.TestCase):
    """process_all's console lifecycle: success only after commit, failures
    and cancellations flush concerns, and a post-commit cleanup failure is
    never reported as a rollback."""

    def _processors(self, *behaviors):
        return tuple(
            (f"模式{index}", f"步骤{index}", None, behavior)
            for index, behavior in enumerate(behaviors, start=1)
        )

    def test_batch_success_finishes_after_commit(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        processors = self._processors(
            lambda r: None,
            lambda r: None,
        )
        with patch.object(
            app_main,
            "run_with_output_rollback",
            side_effect=lambda paths, operation: operation(),
        ):
            app_main.process_all(processors, reporter)
        text = reporter.stream.getvalue()
        self.assertIn("处理完成：2/2 步骤成功", text)
        self.assertIn("已提交输出", text)
        self.assertEqual(reporter.error_stream.getvalue(), "")

    def test_batch_failure_reports_rollback_and_no_output_list(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        processors = self._processors(
            lambda r: None,
            lambda r: (_ for _ in ()).throw(ValueError("坏了")),
        )
        with patch.object(
            app_main,
            "run_with_output_rollback",
            side_effect=lambda paths, operation: operation(),
        ):
            with self.assertRaisesRegex(ValueError, "坏了"):
                app_main.process_all(processors, reporter)
        self.assertIn("处理失败：1/2 步骤成功", reporter.error_stream.getvalue())
        self.assertIn(
            "本次输出已回滚，原有输出文件保持不变",
            reporter.error_stream.getvalue(),
        )
        self.assertNotIn("已提交输出", reporter.stream.getvalue())
        self.assertNotIn("处理完成", reporter.stream.getvalue())

    def test_batch_cancel_reports_cancelled_summary(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        processors = self._processors(lambda r: (_ for _ in ()).throw(KeyboardInterrupt()))
        with patch.object(
            app_main,
            "run_with_output_rollback",
            side_effect=lambda paths, operation: operation(),
        ):
            with self.assertRaises(KeyboardInterrupt):
                app_main.process_all(processors, reporter)
        self.assertIn("处理已取消：0/1 步骤已完成", reporter.error_stream.getvalue())
        self.assertIn(
            "本次输出已回滚，原有输出文件保持不变",
            reporter.error_stream.getvalue(),
        )

    def test_commit_cleanup_failure_is_not_reported_as_rollback(self) -> None:
        """Every step committed, then backup cleanup raised: the summary must
        say the outputs were committed, not rolled back."""
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        processors = self._processors(lambda r: None)

        def cleanup_fails(paths, operation):
            operation()  # steps all succeed
            raise OutputCleanupError("输出已提交，备份清理失败")

        with patch.object(
            app_main,
            "run_with_output_rollback",
            side_effect=cleanup_fails,
        ):
            with self.assertRaisesRegex(OutputCleanupError, "备份清理失败"):
                app_main.process_all(processors, reporter)
        text = reporter.error_stream.getvalue()
        self.assertIn("输出已提交，未回滚", text)
        self.assertIn("备份清理失败", text)
        self.assertNotIn("本次输出已回滚", text)
        # No success summary on the main stream either.
        self.assertNotIn("已提交输出", reporter.stream.getvalue())
        self.assertNotIn("处理完成：", reporter.stream.getvalue())

    def test_nested_cleanup_failure_in_all_mode_reports_rollback(self) -> None:
        """Mode 1 has its own inner rollback transaction. Its post-commit
        cleanup failure reaches process_all as OutputCleanupError, but the
        outer transaction rolled every output back — the summary must say
        rolled back, with no second [失败] and no committed claim."""
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )

        def inner_cleanup_fails(reporter) -> None:
            raise OutputCleanupError("内层清理失败")

        processors = self._processors(inner_cleanup_fails)
        with patch.object(
            app_main,
            "run_with_output_rollback",
            side_effect=lambda paths, operation: operation(),
        ):
            with self.assertRaisesRegex(OutputCleanupError, "内层清理失败"):
                app_main.process_all(processors, reporter)
        text = reporter.error_stream.getvalue()
        # Only the step-level [失败] (with the rollback remedy) prints.
        self.assertEqual(text.count("[失败]"), 1)
        self.assertIn("处理失败：0/1 步骤成功", text)
        self.assertIn("本次输出已回滚，原有输出文件保持不变", text)
        self.assertNotIn("输出已提交，未回滚", text)
        self.assertNotIn("已提交输出", reporter.stream.getvalue())

    def test_single_mode_cleanup_failure_does_not_claim_rollback(self) -> None:
        """Mode 1's own rollback transaction can hit a post-commit cleanup
        failure; the single-mode handler must not claim the files were
        preserved or rolled back."""

        class CleanupFailRun:
            is_all = False
            step_label = "已上传数据"

            def run(self, reporter) -> None:
                raise OutputCleanupError("输出已提交，备份清理失败")

        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        with (
            patch.object(app_main, "ConsoleReporter", return_value=reporter),
            patch.object(app_main, "resolve_data_dir", return_value=Path("data")),
            patch.object(
                app_main,
                "choose_data_processor",
                return_value=CleanupFailRun(),
            ),
            patch.object(
                app_main,
                "remove_stale_temporary_files",
                return_value=StaleFileCleanup((), ()),
            ),
        ):
            self.assertEqual(app_main.main(), 1)
        text = reporter.error_stream.getvalue()
        self.assertIn("输出已提交，未回滚", text)
        self.assertIn("备份清理失败", text)
        self.assertNotIn("本次输出已回滚", text)
        self.assertNotIn("现有输出文件保持不变", text)


if __name__ == "__main__":
    unittest.main()
