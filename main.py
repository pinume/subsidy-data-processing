import os
import sys
import traceback
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from processors import payment, receipts, store_report, submitted
from processors.common.console import ConsoleReporter, format_count
from processors.common.excel import (
    OutputCleanupError,
    remove_stale_temporary_files,
    run_with_output_rollback,
)
from processors.common.paths import resolve_data_dir
from processors.coupons import sources as coupon_sources


def all_output_files() -> tuple[Path, ...]:
    from processors.coupon_report import OUTPUT_FILE as coupon_output_file

    return (
        *submitted.OUTPUT_FILES,
        receipts.OUTPUT_FILE,
        coupon_output_file,
        payment.OUTPUT_FILE,
        store_report.OUTPUT_FILE,
    )


def build_processors() -> tuple[
    tuple[str, str, Path, Callable[[ConsoleReporter], None]],
    ...,
]:
    """List every processing mode across both projects.

    Each entry is (menu label, step label, source path, processor). The step
    label is the short name shown on the [i/N] stage lines; the menu label
    keeps the full descriptive text. Receipt statistics and subsidy coupon
    statistics are shared, produced once regardless of which entry triggers
    them. The coupon_report import is deferred so it only runs after both
    projects have finished loading (see processors/coupon_report.py for why).
    """
    from processors.coupon_report import process_coupon_sales as process_coupon_report

    return (
        (
            "已上传数据（家电+数码）",
            "已上传数据",
            submitted.DATA_DIR,
            submitted.process_all,
        ),
        (
            "收款单统计",
            "收款单统计",
            receipts.RECEIPTS_SOURCE_FILE or receipts.DATA_DIR,
            receipts.process_receipts,
        ),
        (
            "回款明细（家电+数码）",
            "回款明细",
            payment.DATA_DIR,
            payment.process_payment_files,
        ),
        (
            "审核明细（销售用券情况统计）",
            "审核明细",
            coupon_sources.COUPON_SOURCE_FILE or coupon_sources.DATA_DIR,
            process_coupon_report,
        ),
        (
            "门店国补上传及回款情况表",
            "门店报表",
            store_report.DATA_DIR,
            store_report.process_store_report,
        ),
    )


@dataclass(frozen=True)
class ProcessorSelection:
    """What the operator picked: the runner plus how to stage and report it."""

    run: Callable[[ConsoleReporter], None]
    step_label: str
    is_all: bool


def process_all(
    processors: tuple[tuple[str, str, Path, Callable[[ConsoleReporter], None]], ...],
    reporter: ConsoleReporter,
) -> None:
    total = len(processors)
    succeeded = 0

    def process_everything() -> None:
        nonlocal succeeded
        reporter.run_start()
        for index, (_menu_label, step_label, _source_path, processor) in enumerate(
            processors,
            start=1,
        ):
            reporter.step_start(index, total, step_label)
            try:
                processor(reporter)
            except KeyboardInterrupt:
                raise
            except BaseException as error:
                reporter.error(
                    step_label,
                    error,
                    "本次输出已回滚，原文件保持不变",
                )
                if os.environ.get("UPLOAD_DATA_DEBUG"):
                    traceback.print_exc()
                raise
            reporter.step_success()
            succeeded = index

    try:
        run_with_output_rollback(all_output_files(), process_everything)
    except KeyboardInterrupt:
        # On cancel after every step committed, the exception came from
        # post-commit backup cleanup, so no rollback happened.
        reporter.finish(
            success=False,
            succeeded=succeeded,
            total=total,
            cancelled=True,
            rolled_back=succeeded < total,
        )
        raise
    except OutputCleanupError as error:
        if succeeded == total:
            # Every step committed; only post-commit backup cleanup failed,
            # so the outputs were NOT rolled back.
            reporter.error(
                None,
                error,
                "输出已提交，备份清理失败，请检查 output 目录中的残留文件",
            )
            reporter.finish(
                success=False,
                succeeded=succeeded,
                total=total,
                rolled_back=False,
            )
        else:
            # A nested transaction (e.g. mode 1's own rollback) raised it;
            # the outer transaction rolled every output back and the
            # step-level [失败] already reported that.
            reporter.finish(success=False, succeeded=succeeded, total=total)
        raise
    except BaseException:
        reporter.finish(success=False, succeeded=succeeded, total=total)
        raise
    # The transaction committed: only now may the success summary claim it.
    reporter.finish(success=True, succeeded=total, total=total)


def choose_data_processor() -> ProcessorSelection | None:
    processors = build_processors()
    all_choice = len(processors) + 1

    print("请选择处理模式：")
    for index, (menu_label, _, _, _) in enumerate(processors, start=1):
        print(f"  {index}. {menu_label}")
    print(f"  {all_choice}. 全部处理")
    print("  0. 退出")

    while True:
        try:
            choice = input("输入编号后回车：").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n处理已取消")
            return None

        if choice == "0":
            print("已退出")
            return None
        if choice == str(all_choice) or choice.lower() == "all":
            print(f"按顺序处理全部数据：1-{len(processors)}")
            return ProcessorSelection(
                run=lambda reporter: process_all(processors, reporter),
                step_label="全部模式",
                is_all=True,
            )
        if choice.isdigit():
            selected_index = int(choice) - 1
            if 0 <= selected_index < len(processors):
                _, step_label, _, processor = processors[selected_index]
                return ProcessorSelection(
                    run=processor,
                    step_label=step_label,
                    is_all=False,
                )

        print("输入无效，请输入菜单编号或 all。")


def main() -> int:
    reporter = ConsoleReporter(
        verbose=os.environ.get("UPLOAD_DATA_VERBOSE") == "1"
    )
    try:
        data_dir = resolve_data_dir()
        if data_dir is None:
            return 0

        # Both projects share this data directory, and several processing modes
        # are no longer project-specific, so both need to be configured up
        # front rather than only the one the operator picks.
        submitted.configure_data_dir(data_dir)
        receipts.configure_data_dir(data_dir)
        coupon_sources.configure_data_dir(data_dir)
        payment.configure_data_dir(data_dir)
        store_report.configure_data_dir(data_dir)

        # Every pipeline writes into the same output directory and cleans up
        # after itself; anything dot-prefixed still sitting there is from a run
        # that was interrupted before it could.
        cleanup = remove_stale_temporary_files(submitted.OUTPUT_DIR)
        if cleanup.removed:
            reporter.metric(
                "已清理残留临时文件",
                f"{format_count(len(cleanup.removed))} 个",
            )
        for name, reason in cleanup.failed:
            reporter.review_required(
                "无法删除残留临时文件",
                (f"文件：{name}", f"原因：{reason}"),
            )

        selection = choose_data_processor()
        if selection is None:
            return 0
    except KeyboardInterrupt:
        print("\n处理已取消", file=sys.stderr)
        return 130
    except Exception as error:
        reporter.error(
            None,
            error,
            "现有输出文件保持不变，请检查配置或源文件后重试",
        )
        if os.environ.get("UPLOAD_DATA_DEBUG"):
            traceback.print_exc()
        return 1

    try:
        if selection.is_all:
            selection.run(reporter)
            return 0
        reporter.step_start(1, 1, selection.step_label)
        try:
            selection.run(reporter)
        except KeyboardInterrupt:
            reporter.finish(
                success=False,
                succeeded=0,
                total=1,
                cancelled=True,
            )
            return 130
        except OutputCleanupError as error:
            # The step's outputs were committed; only backup cleanup failed,
            # so the remedy must not claim the files were preserved/rolled
            # back.
            reporter.error(
                selection.step_label,
                error,
                "输出已提交，备份清理失败，请检查 output 目录中的残留文件",
            )
            reporter.finish(
                success=False,
                succeeded=0,
                total=1,
                rolled_back=False,
            )
            return 1
        except Exception as error:
            reporter.error(
                selection.step_label,
                error,
                "现有输出文件保持不变，请检查源文件后重试",
            )
            reporter.finish(success=False, succeeded=0, total=1)
            return 1
        reporter.step_success()
        reporter.finish(success=True, succeeded=1, total=1)
        return 0
    except KeyboardInterrupt:
        # All mode reported its own cancelled summary inside process_all.
        return 130
    except Exception:
        # All mode reported the failing step and the rollback inside
        # process_all; only the traceback is still useful here.
        if os.environ.get("UPLOAD_DATA_DEBUG"):
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
