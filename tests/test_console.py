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
    display_width,
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

    def test_warning_counts_and_prints_a_table(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.warning("测试警告", ("单据：ZG2J000016", "金额：314.85"))
        self.assertEqual(reporter.warning_count, 1)
        self.assertEqual(
            reporter.stream.getvalue(),
            "[警告] 测试警告\n"
            "┌──────┬────────────┐\n"
            "│ 单据 │ ZG2J000016 │\n"
            "│ 金额 │ 314.85     │\n"
            "└──────┴────────────┘\n",
        )

    def test_warning_table_aligns_cjk_keys_by_display_width(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.warning("对齐", ("当前参考号：为空", "单：1"))
        lines = reporter.stream.getvalue().splitlines()
        # 当前参考号 occupies 10 columns (5 CJK chars), 单 only 2; the value
        # column must start at the same display column on both rows (char
        # indexes differ because CJK chars count as two columns).
        first_value = display_width(lines[2][: lines[2].index("为空")])
        second_value = display_width(lines[3][: lines[3].index("1")])
        self.assertEqual(first_value, second_value)

    def test_warning_bare_note_becomes_an_empty_key_row(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.warning("说明", ("其余 2 行未展开",))
        lines = reporter.stream.getvalue().splitlines()
        self.assertIn("其余 2 行未展开", lines[2])
        self.assertTrue(lines[2].startswith("│ "))

    def test_warning_renders_red_when_color_is_enabled(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=True)
        reporter.warning("测试警告", ("单据：1",))
        text = reporter.stream.getvalue()
        self.assertIn("\033[31m[警告] 测试警告\033[0m", text)
        self.assertIn("\033[31m│", text)

    def test_success_lines_render_green_when_color_is_enabled(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=True)
        reporter.step_success("审核明细")
        reporter.output(Path.cwd() / "output" / "审核明细.xlsx")
        text = reporter.stream.getvalue()
        self.assertIn("\033[32m[成功] 已生成审核明细报表\033[0m", text)
        self.assertIn("\033[32m输出  ", text)

    def test_no_color_when_explicitly_disabled(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=False)
        reporter.warning("测试警告", ("单据：1",))
        reporter.step_success("审核明细")
        self.assertNotIn("\033[", reporter.stream.getvalue())

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
