import os
import sys
import traceback
from collections.abc import Callable
from pathlib import Path

from processors import payment, receipts, store_report, submitted
from processors.common.excel import (
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


def build_processors() -> tuple[tuple[str, Path, Callable[[], None]], ...]:
    """List every processing mode across both projects.

    Receipt statistics and subsidy coupon statistics are shared, produced
    once regardless of which entry triggers them. The coupon_report import
    is deferred so it only runs after both projects have finished loading
    (see processors/coupon_report.py for why).
    """
    from processors.coupon_report import process_coupon_sales as process_coupon_report

    return (
        (
            "已上传数据（家电+数码）",
            submitted.DATA_DIR,
            submitted.process_all,
        ),
        (
            "收款单统计",
            receipts.RECEIPTS_SOURCE_FILE or receipts.DATA_DIR,
            receipts.process_receipts,
        ),
        (
            "回款明细（家电+数码）",
            payment.DATA_DIR,
            payment.process_payment_files,
        ),
        (
            "审核明细（销售用券情况统计）",
            coupon_sources.COUPON_SOURCE_FILE or coupon_sources.DATA_DIR,
            process_coupon_report,
        ),
        (
            "门店国补上传及回款情况表",
            store_report.DATA_DIR,
            store_report.process_store_report,
        ),
    )


def process_all(processors: tuple[tuple[str, Path, Callable[[], None]], ...]) -> None:
    def process_everything() -> None:
        print(
            "Batch mode: step success messages are provisional; "
            "a later failure rolls every output back."
        )
        for _, source_path, processor in processors:
            print(f"处理中：{source_path}")
            processor()
        print("全部处理模式已完成；输出已统一提交。")

    run_with_output_rollback(all_output_files(), process_everything)


def choose_data_processor() -> Callable[[], None] | None:
    processors = build_processors()
    all_choice = len(processors) + 1

    print("请选择处理模式：")
    for index, (label, _, _) in enumerate(processors, start=1):
        print(f"  {index}. {label}")
    print(f"  {all_choice}. all")
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
            return lambda: process_all(processors)
        if choice.isdigit():
            selected_index = int(choice) - 1
            if 0 <= selected_index < len(processors):
                _, source_path, processor = processors[selected_index]
                print(f"处理中：{source_path}")
                return processor

        print("输入无效，请输入菜单编号或 all。")


def report_failure(error: BaseException) -> None:
    """Show operators the cause, not a Python stack trace."""
    print(f"\n处理失败：{error}", file=sys.stderr)
    print(
        "现有输出文件保持不变。请检查上方指出的源文件后重新运行。",
        file=sys.stderr,
    )
    if os.environ.get("UPLOAD_DATA_DEBUG"):
        traceback.print_exc()
    else:
        print(
            "设置 UPLOAD_DATA_DEBUG=1 可查看完整堆栈。",
            file=sys.stderr,
        )


def main() -> int:
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
        removed = remove_stale_temporary_files(submitted.OUTPUT_DIR)
        if removed:
            print(f"已清理 {len(removed)} 个残留临时文件：{'、'.join(removed)}")

        processor = choose_data_processor()
        if processor is None:
            return 0
        processor()
    except KeyboardInterrupt:
        print("\n处理已取消", file=sys.stderr)
        return 130
    except Exception as error:
        report_failure(error)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
