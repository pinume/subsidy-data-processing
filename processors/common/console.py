"""One console reporter shared by every processing mode.

Presentation only: all business results, warnings and failures flow through
a ConsoleReporter instance so the terminal reads uniformly — stage lines,
compact metrics, warning blocks, and a final success/failure summary. No
Excel content depends on this module.

The reporter is passed explicitly (never a module-level global) so the all
mode can share one instance and accumulate warning counts, single modes can
use their own, and tests can inject StringIO streams. Each processor's
top-level entry takes a reporter; deep functions stay print-free and return
structured records instead (see SubmittedReport.unknown_status_records,
ExcludedProductRecord, SubsidyCorrection, SupplementReferenceConflict).
"""

from __future__ import annotations

import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

INDENT = "       "


def format_count(value: object) -> str:
    """Thousands-separated integer for counts (17682 -> 17,682)."""
    return f"{int(value):,}"


def format_amount(value: object) -> str:
    """Thousands-separated two-decimal amount (314.85 -> 314.85)."""
    try:
        amount = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return str(value)
    return f"{amount:,.2f}"


def display_path(path: object) -> str:
    """Show a path relative to the working directory when possible.

    output/审核明细.xlsx instead of the absolute path; anything outside the
    working directory stays absolute rather than being misleading.
    """
    try:
        return str(Path(path).resolve().relative_to(Path.cwd().resolve()))
    except (OSError, ValueError):
        return str(path)


class ConsoleReporter:
    """Writes the operator-facing console transcript to injected streams.

    warning_count/failure_count are accumulated on the instance so the final
    summary can state them; two instances never share state.
    """

    def __init__(self, stream=None, error_stream=None) -> None:
        self.stream = stream if stream is not None else sys.stdout
        self.error_stream = (
            error_stream if error_stream is not None else sys.stderr
        )
        self.warning_count = 0
        self.failure_count = 0

    def step_start(self, index: int, total: int, label: str) -> None:
        print(f"[{index}/{total}] {label}", file=self.stream)

    def metric(self, label: str, value: object = "") -> None:
        print(f"{label}  {value}", file=self.stream)

    def warning(self, title: str, details: tuple[str, ...] = ()) -> None:
        self.warning_count += 1
        print(f"[警告] {title}", file=self.stream)
        for detail in details:
            print(f"{INDENT}{detail}", file=self.stream)

    def output(self, path: object) -> None:
        print(f"输出  {display_path(path)}", file=self.stream)

    def step_success(self, label: str) -> None:
        print(f"[成功] {label}完成", file=self.stream)

    def run_start(self, total: int) -> None:
        print("全部模式：任一步失败将回滚本次所有输出", file=self.stream)

    def run_success(self, total: int, *, transaction: bool = False) -> None:
        print(f"[成功] 全部 {total} 个步骤已完成", file=self.stream)
        prefix = "输出事务已提交｜" if transaction else ""
        print(
            f"{INDENT}{prefix}警告 {self.warning_count} 项｜"
            f"失败 {self.failure_count} 项",
            file=self.stream,
        )

    def failure(
        self,
        step_label: str | None,
        error: BaseException,
        remedy: str,
    ) -> None:
        self.failure_count += 1
        title = f"{step_label}处理失败" if step_label else "处理失败"
        print(f"[失败] {title}", file=self.error_stream)
        print(f"{INDENT}原因：{error}", file=self.error_stream)
        print(f"{INDENT}处理：{remedy}", file=self.error_stream)
