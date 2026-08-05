"""ConsoleReporter contract: injected streams, explicit state, formatting.

The reporter is the only thing allowed to print in the pipeline, so its
formatting is pinned here: thousands separators, relative paths, stage
lines, deferred attention blocks, truncation, and the final
success/failure summaries. Two instances never share state, so a broken
test cannot leak attention items into a later one.
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
        self.assertEqual(reporter.stream.getvalue(), "[1/5] 已上传数据\n\n")

    def test_metric_prints_label_and_value(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.metric("家电", "8 个文件｜4,125 行")
        self.assertEqual(reporter.stream.getvalue(), "家电：8 个文件｜4,125 行\n")

    def test_metric_is_indented_inside_a_step(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_start(1, 1, "审核明细")
        reporter.metric("家电", "7,862 行")
        self.assertEqual(reporter.stream.getvalue(), "[1/1] 审核明细\n\n  家电：7,862 行\n")

    def test_detail_is_silent_by_default_and_never_counts(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.detail("按业务规则剔除“北国”商品：2 行", ("源第 3 行",))
        self.assertEqual(reporter.stream.getvalue(), "")
        self.assertEqual(reporter.corrected_count, 0)
        self.assertEqual(reporter.review_count, 0)

    def test_detail_shows_only_in_verbose_mode(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), verbose=True)
        reporter.detail("数码未知状态数据：2 行", ("源文件 a.xlsx，源行 3", "其余 0 行"))
        self.assertEqual(
            reporter.stream.getvalue(),
            "[明细] 数码未知状态数据：2 行\n"
            "       源文件 a.xlsx，源行 3\n"
            "       其余 0 行\n"
            "\n",
        )
        self.assertEqual(reporter.corrected_count, 0)

    def test_corrected_defers_until_finish_and_counts(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.corrected("补贴归属已自动调整", ("单据：ZG2J000016",))
        self.assertEqual(reporter.stream.getvalue(), "")  # nothing yet
        self.assertEqual(reporter.corrected_count, 1)
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("[已修正 1] 补贴归属已自动调整", text)
        self.assertIn("单据：ZG2J000016", text)

    def test_review_required_defers_until_finish_and_counts(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.review_required("补充参考号候选不唯一", ("候选：A、B",))
        self.assertEqual(reporter.stream.getvalue(), "")
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("[待核对 1] 补充参考号候选不唯一", text)
        self.assertIn("候选：A、B", text)
        self.assertEqual(reporter.review_count, 1)

    def test_attention_blocks_render_aligned_without_box_characters(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.review_required(
            "补充参考号候选不唯一",
            ("当前参考号：为空", "单：1"),
        )
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        start = lines.index("[待核对 1] 补充参考号候选不唯一")
        # 当前参考号 occupies 10 display columns (5 CJK chars), 单 only 2;
        # both values must start at the same display column.
        first_value = display_width(lines[start + 1][: lines[start + 1].index("为空")])
        second_value = display_width(lines[start + 2][: lines[start + 2].index("1")])
        self.assertEqual(first_value, second_value)
        self.assertNotIn("┌", reporter.stream.getvalue())
        self.assertNotIn("│", reporter.stream.getvalue())

    def test_success_finish_prints_summary_and_committed_outputs(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.corrected("补贴归属已自动调整", ("单据：ZG2J000016",))
        reporter.output(Path.cwd() / "output" / "审核明细.xlsx")
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("处理完成：1/1 步骤成功", text)
        self.assertIn("数据修正：1 项｜待人工核对：0 项", text)
        self.assertIn("需关注内容", text)
        self.assertIn("已提交输出：1 个文件", text)
        self.assertIn(f"  {os.path.join('output', '审核明细.xlsx')}", text)

    def test_success_summary_skips_concerns_when_empty(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.finish(success=True, succeeded=5, total=5)
        text = reporter.stream.getvalue()
        self.assertIn("处理完成：5/5 步骤成功", text)
        self.assertIn("数据修正：0 项｜待人工核对：0 项", text)
        self.assertNotIn("需关注内容", text)

    def test_failure_finish_flushes_concerns_and_never_lists_outputs(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        reporter.corrected("补贴归属已自动调整", ("单据：ZG2J000016",))
        reporter.output(Path.cwd() / "output" / "审核明细.xlsx")
        reporter.finish(success=False, succeeded=2, total=5)
        # Concerns are about data, not files: they flush even after rollback.
        self.assertIn("需关注内容", reporter.stream.getvalue())
        self.assertIn("单据：ZG2J000016", reporter.stream.getvalue())
        # A rolled-back run must not pretend the outputs were committed.
        self.assertNotIn("已提交输出", reporter.stream.getvalue())
        self.assertNotIn("output", reporter.stream.getvalue())
        self.assertIn("处理失败：2/5 步骤成功", reporter.error_stream.getvalue())
        self.assertIn(
            "本次输出已回滚，原有输出文件保持不变",
            reporter.error_stream.getvalue(),
        )

    def test_cancelled_finish_uses_cancel_wording(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        reporter.finish(success=False, succeeded=3, total=5, cancelled=True)
        self.assertIn("处理已取消：3/5 步骤已完成", reporter.error_stream.getvalue())

    def test_committed_finish_never_claims_rollback(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        reporter.finish(success=False, succeeded=5, total=5, rolled_back=False)
        self.assertIn("输出已提交，未回滚", reporter.error_stream.getvalue())
        self.assertNotIn("本次输出已回滚", reporter.error_stream.getvalue())
        self.assertNotIn("已提交输出", reporter.stream.getvalue())

    def test_attention_rows_are_truncated_per_item(self) -> None:
        details = tuple(f"单据：{index}" for index in range(12))
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.review_required("一条", details)
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("单据：0", text)
        self.assertIn("单据：9", text)
        self.assertNotIn("单据：10", text)
        self.assertIn("其余 2 行未展开（UPLOAD_DATA_VERBOSE=1 查看全部）", text)

    def test_verbose_mode_shows_every_attention_row(self) -> None:
        details = tuple(f"单据：{index}" for index in range(12))
        reporter = ConsoleReporter(stream=io.StringIO(), verbose=True)
        reporter.review_required("一条", details)
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("单据：11", text)
        self.assertNotIn("未展开", text)

    def test_attention_items_are_truncated_per_title_group(self) -> None:
        """11 items of one title must not crowd out a different title."""
        reporter = ConsoleReporter(stream=io.StringIO())
        for index in range(11):
            reporter.corrected("补贴归属已自动调整", (f"单据：A{index}",))
        reporter.corrected("已删除无效导出文件（没有工作表）", ("文件：b.xlsx",))
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        # The noisy group is capped at 10 items...
        self.assertIn("[已修正 1] 补贴归属已自动调整", text)
        self.assertIn("A9", text)
        self.assertNotIn("A10", text)
        self.assertIn("其余 1 项未展开（UPLOAD_DATA_VERBOSE=1 查看全部）", text)
        # ...and the different title still shows in full (single item keeps
        # its key-value rows).
        self.assertIn("[已修正 2] 已删除无效导出文件（没有工作表）", text)
        self.assertIn("文件：b.xlsx", text)

    def test_multi_item_group_renders_header_and_data_rows(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.corrected("补贴归属已自动调整", ("单据：A", "金额：1"))
        reporter.corrected("补贴归属已自动调整", ("单据：B", "金额：2"))
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        header = "[已修正 1] 补贴归属已自动调整"
        self.assertEqual(lines.count(header), 1)
        start = lines.index(header)
        # One header row naming every column, then one data row per item.
        self.assertIn("单据", lines[start + 1])
        self.assertIn("金额", lines[start + 1])
        row_a, row_b = lines[start + 2], lines[start + 3]
        self.assertIn("A", row_a)
        self.assertIn("1", row_a)
        self.assertIn("B", row_b)
        self.assertIn("2", row_b)
        # Values align under their columns.
        self.assertEqual(
            display_width(row_a[: row_a.index("1")]),
            display_width(row_b[: row_b.index("2")]),
        )

    def test_repeated_field_rows_render_as_a_table(self) -> None:
        """A list of file names inside one call repeats the same field; it
        renders as a header row plus one value per row."""
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.corrected(
            "已删除无效导出文件（没有工作表）",
            ("文件：a.xlsx", "文件：b.xlsx"),
        )
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        start = lines.index("[已修正 1] 已删除无效导出文件（没有工作表）")
        self.assertIn("文件", lines[start + 1])  # the column header
        self.assertEqual(lines[start + 2], "  a.xlsx")
        self.assertEqual(lines[start + 3], "  b.xlsx")
        self.assertNotIn("文件：a.xlsx", reporter.stream.getvalue())

    def test_single_item_group_keeps_key_value_rows(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.review_required("补充参考号候选不唯一", ("单据：A", "金额：1"))
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        start = lines.index("[待核对 1] 补充参考号候选不唯一")
        self.assertIn("单据：A", lines[start + 1])
        self.assertIn("金额：1", lines[start + 2])

    def test_verbose_mode_shows_every_attention_item(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), verbose=True)
        for index in range(12):
            reporter.corrected("已删除无效导出文件", (f"文件：{index}.xlsx",))
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("11.xlsx", text)
        self.assertNotIn("未展开", text)

    def test_step_success_counts_outputs_and_attention_of_that_step(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_start(1, 5, "已上传数据")
        reporter.output(Path("output/a.xlsx"))
        reporter.output(Path("output/b.xlsx"))
        reporter.corrected("一")
        reporter.step_success()
        text = reporter.stream.getvalue()
        self.assertIn("[完成] 已准备 2 个输出｜记录 1 项处理事项", text)

    def test_step_success_omits_attention_when_none(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.step_start(1, 5, "门店报表")
        reporter.output(Path("output/store.xlsx"))
        reporter.step_success()
        text = reporter.stream.getvalue()
        self.assertIn("[完成] 已准备 1 个输出", text)
        self.assertNotIn("记录", text)

    def test_run_start_announces_the_rollback_rule(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO())
        reporter.run_start()
        self.assertEqual(
            reporter.stream.getvalue(),
            "开始全部处理\n"
            "任一步骤失败，本次输出将回滚，原有输出文件保持不变\n",
        )

    def test_corrected_and_section_header_stay_uncolored_when_color_is_on(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=True)
        reporter.corrected("补贴归属已自动调整", ("单据：1",))
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        section = next(line for line in lines if "需关注内容" in line)
        header = next(line for line in lines if "[已修正 1]" in line)
        detail = next(line for line in lines if "单据：1" in line)
        # Auto-fixed items need no action: they must not read as failures.
        self.assertNotIn("\033[", section)
        self.assertNotIn("\033[", header)
        self.assertNotIn("\033[", detail)

    def test_review_required_renders_yellow_when_color_is_on(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=True)
        reporter.review_required("补充参考号候选不唯一", ("单据：1",))
        reporter.finish(success=True, succeeded=1, total=1)
        lines = reporter.stream.getvalue().splitlines()
        header = next(line for line in lines if "[待核对 1]" in line)
        detail = next(line for line in lines if "单据：1" in line)
        self.assertIn("\033[33m[待核对 1]", header)
        self.assertIn("\033[33m", detail)

    def test_success_lines_render_green_when_color_is_enabled(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=True)
        reporter.step_start(1, 1, "已上传数据")
        reporter.output(Path.cwd() / "output" / "a.xlsx")
        reporter.step_success()
        reporter.finish(success=True, succeeded=1, total=1)
        text = reporter.stream.getvalue()
        self.assertIn("\033[32m[完成] 已准备 1 个输出\033[0m", text)
        self.assertIn("\033[32m处理完成：1/1 步骤成功\033[0m", text)

    def test_no_color_when_explicitly_disabled(self) -> None:
        reporter = ConsoleReporter(stream=io.StringIO(), color=False)
        reporter.review_required("测试", ("单据：1",))
        reporter.step_start(1, 1, "已上传数据")
        reporter.output(Path.cwd() / "output" / "a.xlsx")
        reporter.step_success()
        reporter.finish(success=True, succeeded=1, total=1)
        self.assertNotIn("\033[", reporter.stream.getvalue())

    def test_error_goes_to_the_error_stream(self) -> None:
        reporter = ConsoleReporter(
            stream=io.StringIO(),
            error_stream=io.StringIO(),
        )
        reporter.error(
            "审核明细",
            ValueError("金额列不存在"),
            "本次输出已回滚，原文件保持不变",
        )
        self.assertEqual(reporter.stream.getvalue(), "")
        self.assertEqual(
            reporter.error_stream.getvalue(),
            "[失败] 审核明细处理失败：金额列不存在\n"
            "       处理：本次输出已回滚，原文件保持不变\n",
        )

    def test_error_without_a_step_label(self) -> None:
        reporter = ConsoleReporter(error_stream=io.StringIO())
        reporter.error(None, ValueError("配置错误"), "请检查配置")
        self.assertIn("[失败] 处理失败：配置错误\n", reporter.error_stream.getvalue())

    def test_instances_do_not_share_state(self) -> None:
        first = ConsoleReporter(stream=io.StringIO())
        second = ConsoleReporter(stream=io.StringIO())
        first.corrected("一")
        second.corrected("二")
        second.review_required("三")
        self.assertEqual(first.corrected_count, 1)
        self.assertEqual(second.corrected_count, 1)
        self.assertEqual(second.review_count, 1)
        self.assertNotIn("一", second.stream.getvalue())


if __name__ == "__main__":
    unittest.main()
