import os
import sys
import traceback
from collections.abc import Callable
from pathlib import Path

from processors import digital, large_appliances, payment
from processors.common.excel import remove_stale_temporary_files
from processors.common.paths import resolve_data_dir


def process_submitted_files() -> None:
    """Process both projects' submitted data together.

    Each project reads a different set of source files (told apart by
    filename marker) and writes its own output, but an operator has no
    reason to run one without the other, so a single menu entry covers both.
    """
    large_appliances.process_submitted_files()
    digital.process_submitted_files()


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
            large_appliances.DATA_DIR,
            process_submitted_files,
        ),
        (
            "收款单统计",
            large_appliances.RECEIPTS_SOURCE_FILE or large_appliances.DATA_DIR,
            large_appliances.process_receipts,
        ),
        (
            "审核明细（销售用券情况统计）",
            large_appliances.COUPON_SOURCE_FILE or large_appliances.DATA_DIR,
            process_coupon_report,
        ),
        (
            "回款明细（家电+数码）",
            payment.DATA_DIR,
            payment.process_payment_files,
        ),
    )


def process_all(processors: tuple[tuple[str, Path, Callable[[], None]], ...]) -> None:
    for _, source_path, processor in processors:
        print(f"Processing: {source_path}")
        processor()


def choose_data_processor() -> Callable[[], None] | None:
    processors = build_processors()
    all_choice = len(processors) + 1

    print("Select a processing mode:")
    for index, (label, _, _) in enumerate(processors, start=1):
        print(f"  {index}. {label}")
    print(f"  {all_choice}. all")
    print("  0. Exit")

    while True:
        try:
            choice = input("Enter a number: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nProcessing cancelled")
            return None

        if choice == "0":
            print("Exited")
            return None
        if choice == str(all_choice) or choice.lower() == "all":
            print(f"Processing all data in order: 1-{len(processors)}")
            return lambda: process_all(processors)
        if choice.isdigit():
            selected_index = int(choice) - 1
            if 0 <= selected_index < len(processors):
                _, source_path, processor = processors[selected_index]
                print(f"Processing: {source_path}")
                return processor

        print("Invalid input. Enter a menu number or all.")


def report_failure(error: BaseException) -> None:
    """Show operators the cause, not a Python stack trace."""
    print(f"\nProcessing failed: {error}", file=sys.stderr)
    print(
        "Existing output files were left unchanged. "
        "Check the source files named above, then run the program again.",
        file=sys.stderr,
    )
    if os.environ.get("UPLOAD_DATA_DEBUG"):
        traceback.print_exc()
    else:
        print(
            "Set UPLOAD_DATA_DEBUG=1 for the full traceback.",
            file=sys.stderr,
        )


def main() -> int:
    data_dir = resolve_data_dir()
    if data_dir is None:
        return 0

    # Both projects share this data directory, and several processing modes
    # are no longer project-specific, so both need to be configured up
    # front rather than only the one the operator picks.
    digital.configure_data_dir(data_dir)
    large_appliances.configure_data_dir(data_dir)
    payment.configure_data_dir(data_dir)

    # Every pipeline writes into the same output directory and cleans up after
    # itself; anything dot-prefixed still sitting there is from a run that was
    # interrupted before it could.
    removed = remove_stale_temporary_files(large_appliances.OUTPUT_DIR)
    if removed:
        print(f"Removed {len(removed)} leftover file(s): {'、'.join(removed)}")

    processor = choose_data_processor()
    if processor is None:
        return 0

    try:
        processor()
    except KeyboardInterrupt:
        print("\nProcessing cancelled", file=sys.stderr)
        return 130
    except Exception as error:
        report_failure(error)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
