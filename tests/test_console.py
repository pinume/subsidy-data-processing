"""ConsoleReporter contract: injected streams, explicit state, formatting.

The reporter is the only thing allowed to print in the pipeline, so its
formatting is pinned here: thousands separators, relative paths, stage
lines, warning blocks, and the final success/failure summaries. Two
instances never share state, so a broken test cannot leak warnings into a
later one.
"""

import io
import os
import unittest
from decimal import Decimal
from pathlib import Path

from processors.common.console import (
    ConsoleReporter,
    display_path,
    format_amount,
    format_count,
)


class FormattingTests(unittest.TestCase):
    def test_format_count_adds_thousands_separators(self) -> None:
        self.assertEqual(format_count(17682), "17,682")
        self.assertEqual(format_count(0), "0")
        self.assertEqual(format_count(5511), "5,511")

    def test_format_amount_keeps_two_decimals_and_separators(self) -> None:
        self.assertEqual(format_amount(314.85), "314.85")
        self.assertEqual(format_amount(Decimal("2867621.40")), "2,867,621.40")
        self.assertEqual(format_amount(0), "0.00")

    def test_display_path_relativizes_under_the_working_directory(self) -> None:
        path = Path.cwd() / "output" / "审核明细.xlsx"
        self.assertEqual(display_path(path), os.path.join("output", "审核明细.xlsx"))

    def test_display_path_keeps_relative_inputs_relative(self) -> None:
        self.assertEqual(display_path("output/审核明细.xlsx"), "output/审核明细.xlsx")

    def test_display_path_keeps_outside_paths_absolute(self) -> None:
        outside = Path("/tmp/not-in-project/x.xlsx")
        self.assertEqual(display_path(outside), str(outside))


class ConsoleReporterTests(unittest.TestCase):
    def test_step_start_prints_stage_line_with_blank_line(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_start(1, 5, "已上传数据")
        self.assertEqual(reporter.stream.getvalue(), "[1/5] 处理已上传数据\n\n")

    def test_metric_prints_label_and_value(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.metric("家电", "8 个文件｜4,125 行")
        self.assertEqual(reporter.stream.getvalue(), "家电：8 个文件｜4,125 行\n")

    def test_detail_is_silent_by_default_and_never_counts(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.detail("按业务规则剔除“北国”商品：2 行", ("源第 3 行",))
        self.assertEqual(reporter.stream.getvalue(), "")
        self.assertEqual(reporter.warning_count, 0)

    def test_detail_shows_only_in_verbose_mode(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), verbose=True)
        reporter.detail("数码待同步数据：2 行", ("源文件 a.xlsx，源行 3", "其余 0 行"))
        self.assertEqual(
            reporter.stream.getvalue(),
            "[明细] 数码待同步数据：2 行\n"
            "       源文件 a.xlsx，源行 3\n"
            "       其余 0 行\n"
            "\n",
        )
        self.assertEqual(reporter.warning_count, 0)

    def test_warning_counts_and_prints_block(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.warning("测试警告", ("明细一", "明细二"))
        self.assertEqual(reporter.warning_count, 1)
        self.assertEqual(
            reporter.stream.getvalue(),
            "[警告] 测试警告\n       明细一\n       明细二\n",
        )

    def test_output_prints_relative_path(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.output(Path.cwd() / "output" / "回款明细.xlsx")
        self.assertEqual(
            reporter.stream.getvalue(),
            f"输出  {os.path.join('output', '回款明细.xlsx')}\n",
        )

    def test_step_success_mentions_the_report(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_success("已上传数据")
        self.assertEqual(
            reporter.stream.getvalue(),
            "[成功] 已生成已上传数据报表\n",
        )

    def test_step_success_does_not_double_the_report_word(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_success("门店报表")
        self.assertEqual(
            reporter.stream.getvalue(),
            "[成功] 已生成门店报表\n",
        )

    def test_run_success_single_mode_has_no_transaction_line(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.warning("一条警告")
        reporter.run_success(1)
        self.assertIn("[成功] 全部 1 个步骤已完成", reporter.stream.getvalue())
        self.assertIn("警告 1 项｜失败 0 项", reporter.stream.getvalue())
        self.assertNotIn("输出事务已提交", reporter.stream.getvalue())

    def test_run_success_all_mode_mentions_the_transaction(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.run_success(5, transaction=True)
        self.assertIn("全部 5 个步骤已完成", reporter.stream.getvalue())
        self.assertIn("输出事务已提交｜警告 0 项｜失败 0 项", reporter.stream.getvalue())

    def test_run_start_announces_the_rollback_rule(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.run_start(5)
        self.assertEqual(
            reporter.stream.getvalue(),
            "全部模式：任一步失败将回滚本次所有输出\n",
        )

    def test_error_goes_to_the_error_stream_and_counts(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        reporter.error("审核明细", ValueError("坏了"), "本次输出已回滚，原文件保持不变")
        self.assertEqual(reporter.failure_count, 1)
        self.assertEqual(reporter.stream.getvalue(), "")
        self.assertEqual(
            reporter.error_stream.getvalue(),
            "[失败] 审核明细处理失败\n"
            "       原因：坏了\n"
            "       处理：本次输出已回滚，原文件保持不变\n",
        )

    def test_error_without_a_step_label(self) -> None:
        reporter = ConsoleReporter(error_stream=io.StringIO())
        reporter.error(None, ValueError("配置错误"), "请检查配置")
        self.assertIn("[失败] 处理失败\n", reporter.error_stream.getvalue())

    def test_instances_do_not_share_state(self) -> None:
        first = ConsoleReporter(stream=io.StringIO())
        second = ConsoleReporter(stream=io.StringIO())
        first.warning("一")
        second.warning("二")
        second.warning("三")
        self.assertEqual(first.warning_count, 1)
        self.assertEqual(second.warning_count, 2)
        self.assertNotIn("二", first.stream.getvalue())


if __name__ == "__main__":
    unittest.main()
