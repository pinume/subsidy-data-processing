"""One console reporter shared by every processing mode.

Presentation only: all business results, attention items and failures flow
through a ConsoleReporter instance so the terminal reads uniformly — stage
lines, compact metrics, verbose-only traceability, deferred attention blocks
(auto-fixed corrections vs items needing a human), and a final
success/failure summary. No Excel content depends on this module.

Two attention channels exist because they demand different responses:
- corrected(): the program already handled it (subsidy re-attribution, an
  invalid export deleted); shown once at the end in default color.
- review_required(): a human must look at it (ambiguous supplement match, an
  unparsable source total); shown once at the end in yellow.

Both are collected during the run and flushed by finish(), so the operator
sees live progress first and reads the full list once the run settles.
Unless verbose is on, same-title items merge into one block, each title
group shows at most MAX_ATTENTION_ITEMS items and each item at most
MAX_ATTENTION_DETAILS rows — so one noisy title can never crowd out a
different one.

Coloring is ANSI-only (no third-party dependency) and only active when the
stream is a terminal, so redirected output and StringIO tests stay plain.
Success reads green, review_required yellow, failures red on the error
stream; corrected() stays uncolored — it needs no action.

The reporter is passed explicitly (never a module-level global) so the all
mode can share one instance and accumulate counts, single modes can use
their own, and tests can inject StringIO streams. Each processor's
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
_YELLOW = "\033[33m"
_RESET = "\033[0m"

MAX_ATTENTION_ITEMS = 10
MAX_ATTENTION_DETAILS = 10
TRUNCATION_HINT = "（UPLOAD_DATA_VERBOSE=1 查看全部）"

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
    """Split one detail line at its first full-width colon.

    Details are written as "键：值"; a line without a colon (a bare note)
    becomes a row with an empty key and spans the full width.
    """
    if "：" in detail:
        key, value = detail.split("：", 1)
        return key, value
    return "", detail


def _group_by_title(
    items: list[tuple[str, tuple[str, ...]]],
) -> list[tuple[str, list[tuple[str, ...]]]]:
    """Merge same-title items into one block, first occurrence order kept.

    Truncation is per title group, so a noisy title can never crowd out a
    different one (10 corrections must not hide one deleted-file record).
    """
    groups: list[tuple[str, list[tuple[str, ...]]]] = []
    index_by_title: dict[str, int] = {}
    for title, details in items:
        if title not in index_by_title:
            index_by_title[title] = len(groups)
            groups.append((title, []))
        groups[index_by_title[title]][1].append(details)
    return groups


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

    corrected_count/review_count accumulate on the instance so the summary
    can state them; two instances never share state.

    The output channels are deliberately distinct:
    - metric(): normal statistics, always shown (待同步 count, 北国 count)
    - detail(): traceability lines shown only when verbose is on; never
      counted
    - corrected(): an anomaly the program already fixed; collected, shown
      once at the end in default color, counted in the summary
    - review_required(): something a human must confirm; collected, shown
      once at the end in yellow, counted in the summary
    - error(): a failed step, written red to the error stream, immediate
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
        self.corrected_count = 0
        self.review_count = 0
        self._corrected: list[tuple[str, tuple[str, ...]]] = []
        self._review: list[tuple[str, tuple[str, ...]]] = []
        self._outputs: list[Path] = []
        self._in_step = False
        self._step_output_mark = 0
        self._step_corrected_mark = 0
        self._step_review_mark = 0

    def _red(self, text: str) -> str:
        return f"{_RED}{text}{_RESET}" if self.color else text

    def _green(self, text: str) -> str:
        return f"{_GREEN}{text}{_RESET}" if self.color else text

    def _yellow(self, text: str) -> str:
        return f"{_YELLOW}{text}{_RESET}" if self.color else text

    def step_start(self, index: int, total: int, label: str) -> None:
        # Trailing blank line keeps the stage block from gluing onto the
        # metrics that follow.
        print(f"[{index}/{total}] {label}\n", file=self.stream)
        self._in_step = True
        self._step_output_mark = len(self._outputs)
        self._step_corrected_mark = len(self._corrected)
        self._step_review_mark = len(self._review)

    def metric(self, label: str, value: object = "") -> None:
        # Indented under the running step so metrics read as that step's
        # results; direct calls outside a step stay flush-left.
        prefix = "  " if self._in_step else ""
        print(f"{prefix}{label}：{value}", file=self.stream)

    def detail(self, title: str, lines: tuple[str, ...] = ()) -> None:
        """Traceability block, shown only when verbose is on.

        Not an attention item: normal processing can carry source locations
        without alarming anyone or inflating the summary counts.
        """
        if not self.verbose:
            return
        print(self._green(f"[明细] {title}"), file=self.stream)
        for line in lines:
            print(self._green(f"{INDENT}{line}"), file=self.stream)
        # Blank line so the next block never glues onto the last detail row.
        print(file=self.stream)

    def corrected(self, title: str, details: tuple[str, ...] = ()) -> None:
        """An anomaly the program already handled; shown once at the end."""
        self.corrected_count += 1
        self._corrected.append((title, tuple(details)))

    def review_required(self, title: str, details: tuple[str, ...] = ()) -> None:
        """Something a human must confirm; shown once at the end, yellow."""
        self.review_count += 1
        self._review.append((title, tuple(details)))

    def output(self, path: object) -> None:
        # Collected, not printed: paths are only shown once the transaction
        # commits (finish success); a rolled-back run prints none.
        self._outputs.append(Path(path))

    def step_success(self) -> None:
        outputs = len(self._outputs) - self._step_output_mark
        corrected = len(self._corrected) - self._step_corrected_mark
        review = len(self._review) - self._step_review_mark
        parts = [f"已准备 {outputs} 个输出"]
        attention = corrected + review
        if attention:
            parts.append(f"记录 {attention} 项处理事项")
        print(self._green(f"[完成] {'｜'.join(parts)}"), file=self.stream)
        # Blank line so the next stage header never glues onto this line.
        print(file=self.stream)

    def run_start(self) -> None:
        print("开始全部处理", file=self.stream)
        print("任一步骤失败，本次输出将回滚，原有输出文件保持不变", file=self.stream)

    def finish(
        self,
        *,
        success: bool,
        succeeded: int,
        total: int,
        cancelled: bool = False,
        rolled_back: bool = True,
    ) -> None:
        """End the run: flush deferred attention, then the summary.

        Success shows the committed output list; failure/cancellation shows
        the rolled-back summary on the error stream and never a file list.
        rolled_back=False states the transaction committed (a post-commit
        cleanup failure must not be misreported as a rollback).
        """
        if success:
            self._finish_success(succeeded, total)
        else:
            self._flush_concerns()
            self._finish_failure(succeeded, total, cancelled, rolled_back)

    def error(
        self,
        step_label: str | None,
        error: BaseException,
        remedy: str,
    ) -> None:
        title = f"{step_label}处理失败" if step_label else "处理失败"
        red = self._red
        print(red(f"[失败] {title}：{error}"), file=self.error_stream)
        print(red(f"{INDENT}处理：{remedy}"), file=self.error_stream)

    # -- internals ---------------------------------------------------------

    def _finish_success(self, succeeded: int, total: int) -> None:
        green = self._green
        print(green("─" * 40), file=self.stream)
        print(green(f"处理完成：{succeeded}/{total} 步骤成功"), file=self.stream)
        print(
            green(
                f"数据修正：{self.corrected_count} 项｜"
                f"待人工核对：{self.review_count} 项"
            ),
            file=self.stream,
        )
        print(file=self.stream)
        self._flush_concerns()
        print(green(f"已提交输出：{len(self._outputs)} 个文件"), file=self.stream)
        for path in self._outputs:
            print(f"  {display_path(path)}", file=self.stream)

    def _finish_failure(
        self,
        succeeded: int,
        total: int,
        cancelled: bool,
        rolled_back: bool,
    ) -> None:
        if not rolled_back:
            # The transaction committed; only post-commit cleanup failed.
            print(
                self._red("输出已提交，未回滚"),
                file=self.error_stream,
            )
            return
        color = self._yellow if cancelled else self._red
        if cancelled:
            print(
                color(f"处理已取消：{succeeded}/{total} 步骤已完成"),
                file=self.error_stream,
            )
        else:
            print(
                color(f"处理失败：{succeeded}/{total} 步骤成功"),
                file=self.error_stream,
            )
        print(
            self._red("本次输出已回滚，原有输出文件保持不变"),
            file=self.error_stream,
        )

    def _flush_concerns(self) -> None:
        if not (self._corrected or self._review):
            return
        print("需关注内容", file=self.stream)
        print(file=self.stream)
        self._flush_type(self._corrected, "已修正", None)
        self._flush_type(self._review, "待核对", self._yellow)

    def _flush_type(
        self,
        items: list[tuple[str, tuple[str, ...]]],
        label: str,
        colorize,
    ) -> None:
        for group_index, (title, group_items) in enumerate(
            _group_by_title(items),
            start=1,
        ):
            shown = group_items
            truncated_items = 0
            if not self.verbose and len(shown) > MAX_ATTENTION_ITEMS:
                shown = shown[:MAX_ATTENTION_ITEMS]
                truncated_items = len(group_items) - MAX_ATTENTION_ITEMS
            header = f"[{label} {group_index}] {title}"
            print(colorize(header) if colorize else header, file=self.stream)
            for line in self._merged_lines(shown):
                print(colorize(line) if colorize else line, file=self.stream)
            if truncated_items:
                print(
                    f"其余 {truncated_items} 项未展开{TRUNCATION_HINT}",
                    file=self.stream,
                )
            print(file=self.stream)

    def _merged_lines(self, items: list[tuple[str, ...]]) -> list[str]:
        """Render a title group's rows.

        Multiple items share one header row plus a data row per item; a
        single item whose rows repeat one field (a list of file names) also
        renders as a table. Everything else renders as aligned key-value
        rows. All values align by display width so CJK text lines up.
        """
        split_rows = [
            [_split_key_value(line) for line in details]
            for details in items
        ]
        if self._is_table(split_rows):
            if len(split_rows) == 1:
                # One item repeating one field (a list of file names): each
                # detail line becomes its own data row.
                split_rows = [[pair] for pair in split_rows[0]]
            return self._render_table(split_rows)
        return self._render_key_value_rows(split_rows)

    def _is_table(self, split_rows) -> bool:
        """Rectangular rows: several items with one shared key sequence, or
        a single item repeating one field at least twice. Bare notes and
        over-long rows fall back to key-value rendering."""
        if len(split_rows) == 1:
            row = split_rows[0]
            keys = [key for key, _ in row]
            return (
                len(keys) >= 2
                and all(keys)
                and len(set(keys)) == 1
                and len(row) <= MAX_ATTENTION_DETAILS
            )
        keys = [key for key, _ in split_rows[0]]
        return bool(keys) and all(keys) and all(
            [key for key, _ in row] == keys
            and len(row) <= MAX_ATTENTION_DETAILS
            for row in split_rows
        )

    def _render_table(self, split_rows) -> list[str]:
        keys = [key for key, _ in split_rows[0]]
        widths = [
            max(
                [display_width(keys[column]), *(
                    display_width(row[column][1])
                    for row in split_rows
                )]
            )
            for column in range(len(keys))
        ]
        lines = [
            "  " + " ".join(_pad(key, widths[column]) for column, key in enumerate(keys))
        ]
        for row in split_rows:
            cells = [
                _pad(row[column][1], widths[column])
                for column in range(len(keys))
            ]
            lines.append("  " + " ".join(cells))
        return lines

    def _render_key_value_rows(self, split_rows) -> list[str]:
        """Concatenated key-value rows with one alignment pass; each item's
        rows are capped unless verbose is on."""
        flat: list[tuple[str, str]] = []
        for row in split_rows:
            shown = row
            if not self.verbose and len(shown) > MAX_ATTENTION_DETAILS:
                shown = shown[:MAX_ATTENTION_DETAILS]
                shown.append(
                    ("", f"其余 {len(row) - MAX_ATTENTION_DETAILS} 行未展开"
                     f"{TRUNCATION_HINT}")
                )
            flat.extend(shown)
        key_width = max((display_width(key) for key, _ in flat), default=0)
        lines = []
        for key, value in flat:
            if key:
                lines.append(f"  {_pad(key, key_width)}：{value}")
            else:
                lines.append(f"  {value}")
        return lines
