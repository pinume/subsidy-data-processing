"""One console reporter shared by every processing mode.

Presentation only: all business results, warnings and failures flow through
a ConsoleReporter instance so the terminal reads uniformly — stage lines,
compact metrics, warning blocks (rendered as tables), and a final
success/failure summary. No Excel content depends on this module.

Coloring is ANSI-only (no third-party dependency) and only active when the
stream is a terminal, so redirected output and StringIO tests stay plain.
Warnings render red, success/positive output green.

The reporter is passed explicitly (never a module-level global) so the all
mode can share one instance and accumulate warning counts, single modes can
use their own, and tests can inject StringIO streams. Each processor's
top-level entry takes a reporter; deep functions stay print-free and return
structured records instead (see SubmittedReport.unknown_status_records,
ExcludedProductRecord, SubsidyCorrection, SupplementReferenceConflict).
"""

from __future__ import annotations

import os
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

INDENT = "       "

_RED = "\033[31m"
_GREEN = "\033[32m"
_RESET = "\033[0m"

# Characters at or above this code point display two columns wide (CJK and
# full-width forms); everything below is one column.
_FULLWIDTH_FLOOR = 0x2E80


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


def display_width(text: str) -> int:
    """Terminal columns a string occupies: CJK/full-width chars count 2."""
    return sum(2 if ord(char) >= _FULLWIDTH_FLOOR else 1 for char in text)


def _pad(text: str, width: int) -> str:
    return text + " " * (width - display_width(text))


def _split_key_value(detail: str) -> tuple[str, str]:
    """Split one warning detail at its first full-width colon.

    Details are written as "键：值"; a line without a colon (a bare note
    such as 其余 2 行未展开) becomes a row with an empty key.
    """
    if "：" in detail:
        key, value = detail.split("：", 1)
        return key, value
    return "", detail


def _table_lines(details: tuple[str, ...]) -> list[str]:
    """Render warning details as a two-column box table.

    Column widths are computed from display widths so CJK keys and values
    stay aligned.
    """
    rows = [_split_key_value(detail) for detail in details]
    key_width = max((display_width(key) for key, _ in rows), default=0)
    value_width = max((display_width(value) for _, value in rows), default=0)
    border = "─" * (key_width + 2)
    lines = [f"┌{border}┬{'─' * (value_width + 2)}┐"]
    for key, value in rows:
        lines.append(
            f"│ {_pad(key, key_width)} │ {_pad(value, value_width)} │"
        )
    lines.append(f"└{border}┴{'─' * (value_width + 2)}┘")
    return lines


def _color_enabled(stream, override: bool | None) -> bool:
    """Whether ANSI color should be emitted for this stream.

    Defaults to on only for terminals; an explicit override wins. On
    Windows the console is switched into VT processing mode so the ANSI
    codes render (a no-op elsewhere and harmless when redirected).
    """
    if override is not None:
        return override
    if not (hasattr(stream, "isatty") and stream.isatty()):
        return False
    if os.name == "nt":
        os.system("")
    return True


class ConsoleReporter:
    """Writes the operator-facing console transcript to injected streams.

    warning_count/failure_count are accumulated on the instance so the final
    summary can state them; two instances never share state.

    The output channels are deliberately distinct:
    - metric(): normal statistics, always shown (待同步 count, 北国 count)
    - detail(): traceability lines shown only when verbose is on; never a
      warning, never counted
    - warning(): things a human should look at, rendered as a red table;
      counted in the summary
    - error(): a failed step, written red to the error stream
    """

    def __init__(
        self,
        stream=None,
        error_stream=None,
        *,
        verbose: bool = False,
        color: bool | None = None,
    ) -> None:
        self.stream = stream if stream is not None else sys.stdout
        self.error_stream = (
            error_stream if error_stream is not None else sys.stderr
        )
        self.verbose = verbose
        self.color = _color_enabled(self.stream, override=color)
        self.warning_count = 0
        self.failure_count = 0

    def _red(self, text: str) -> str:
        return f"{_RED}{text}{_RESET}" if self.color else text

    def _green(self, text: str) -> str:
        return f"{_GREEN}{text}{_RESET}" if self.color else text

    def step_start(self, index: int, total: int, label: str) -> None:
        # Trailing blank line keeps the stage block from gluing onto the
        # metrics that follow.
        print(f"[{index}/{total}] 处理{label}\n", file=self.stream)

    def metric(self, label: str, value: object = "") -> None:
        print(f"{label}：{value}", file=self.stream)

    def detail(self, title: str, lines: tuple[str, ...] = ()) -> None:
        """Traceability block, shown only when verbose is on.

        Not a warning: normal processing can carry source locations without
        alarming anyone or inflating the summary's warning count.
        """
        if not self.verbose:
            return
        print(self._green(f"[明细] {title}"), file=self.stream)
        for line in lines:
            print(self._green(f"{INDENT}{line}"), file=self.stream)
        # Blank line so the next block never glues onto the last detail row.
        print(file=self.stream)

    def warning(self, title: str, details: tuple[str, ...] = ()) -> None:
        self.warning_count += 1
        print(self._red(f"[警告] {title}"), file=self.stream)
        for line in _table_lines(details):
            print(self._red(line), file=self.stream)

    def output(self, path: object) -> None:
        print(self._green(f"输出  {display_path(path)}"), file=self.stream)

    def step_success(self, label: str) -> None:
        suffix = "" if label.endswith("报表") else "报表"
        print(self._green(f"[成功] 已生成{label}{suffix}"), file=self.stream)

    def run_start(self, total: int) -> None:
        print("全部模式：任一步失败将回滚本次所有输出", file=self.stream)

    def run_success(self, total: int, *, transaction: bool = False) -> None:
        print(self._green(f"[成功] 全部 {total} 个步骤已完成"), file=self.stream)
        prefix = "输出事务已提交｜" if transaction else ""
        print(
            self._green(
                f"{INDENT}{prefix}警告 {self.warning_count} 项｜"
                f"失败 {self.failure_count} 项"
            ),
            file=self.stream,
        )

    def error(
        self,
        step_label: str | None,
        error: BaseException,
        remedy: str,
    ) -> None:
        self.failure_count += 1
        title = f"{step_label}处理失败" if step_label else "处理失败"
        red = self._red
        print(red(f"[失败] {title}"), file=self.error_stream)
        print(red(f"{INDENT}原因：{error}"), file=self.error_stream)
        print(red(f"{INDENT}处理：{remedy}"), file=self.error_stream)
