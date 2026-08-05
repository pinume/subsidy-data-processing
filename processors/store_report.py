"""Store subsidy upload/payment status report (门店国补上传及回款情况表).

Unlike the other processors, this one does not read raw exports from the
data directory. It reads two of this project's own outputs — 审核明细.xlsx
(from the 审核明细 processing mode) and 回款明细.xlsx (from the 回款明细
processing mode) — aggregates them by brand and category, and fills a blank
report template that the operator keeps in the data directory.

Current scope is intentionally narrow: one store (益庄店), one year (2026),
one fixed template. ROW_RULES/BRAND_GROUP_RULES and the output filename are
hardcoded rather than configurable — see the module docstring history for why
this is deferred until a second store is actually needed.
"""

from __future__ import annotations

import math
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.workbook import Workbook
from python_calamine import CalamineWorkbook

from processors.common.console import ConsoleReporter
from processors.common.excel import (
    calamine_rows,
    resolve_font,
    save_workbook_atomically,
)
from processors.common.paths import find_data_files, resolve_unique_file
from processors.coupon_report import OUTPUT_FILE as UPLOAD_FILE
from processors.coupon_report import SUBSIDY_YEAR
from processors.coupon_report import SUMMARY_HEADER as UPLOAD_HEADER
from processors.coupon_report import SUMMARY_SHEET_NAME as UPLOAD_SHEET_NAME
from processors.payment import OUTPUT_FILE as PAYMENT_FILE
from processors.payment import SUMMARY_HEADERS as PAYMENT_HEADER
from processors.payment import SUMMARY_SHEET_NAME as PAYMENT_SHEET_NAME

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = (
    OUTPUT_DIR / f"{SUBSIDY_YEAR}年门店国补上传及回款情况表（益庄店）.xlsx"
)
TEMPLATE_FILE_KEYWORD = "门店国补上传及回款情况表"
DATA_DIR: Path

SHANGHAI_TZ = ZoneInfo("Asia/Shanghai")
FILLED_ALIGNMENT = Alignment(horizontal="left", vertical="center")
DATA_NUMBER_FORMAT = "General"
PERCENT_NUMBER_FORMAT = "0.00%"
CURRENCY_COLUMN_WIDTH = 16.93
CURRENCY_COLUMNS = ("D", "E", "F", "G", "J", "K")
TOTAL_ROW = 34
DETAIL_ROWS = range(4, TOTAL_ROW)
BRAND_GROUP_TOTAL_ROW = 45
BRAND_GROUP_DETAIL_ROWS = range(38, BRAND_GROUP_TOTAL_ROW)
TABLE3_PROJECT_ROWS = {"家电": 49, "数码": 50}
TABLE3_TOTAL_ROW = 51
EXPECTED_SHEET_COUNT = 1
EXPECTED_COLUMN_COUNT = 13  # A..M
EXPECTED_SHEET_TITLE = "益庄"
MIN_TEMPLATE_ROW_COUNT = TABLE3_TOTAL_ROW

# A handful of fixed labels that only exist in the right blank template.
# Checked before writing (wrong/stale template must not silently produce a
# wrong report) and again after saving (writing must not have corrupted them).
TEMPLATE_STRUCTURE_CELLS: dict[str, str] = {
    "A2": "品类",
    "C4": "海尔/卡萨帝冰箱",
    "A34": "费用总计",
    "C37": "品牌",
    "C38": "海尔系",
    "C45": "合计",
    "C47": "表3",
    "D47": "审核中",
    "F47": "未上传",
    "D48": "数量",
    "E48": "26年国补上传额",
    "F48": "数量",
    "G48": "26年国补上传额",
    "C49": "家电",
    "C50": "数码",
    "C51": "合计",
}


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    DATA_DIR = data_dir


def resolve_template_file() -> Path:
    candidates = find_data_files(DATA_DIR, TEMPLATE_FILE_KEYWORD, (".xlsx",))
    template_file = resolve_unique_file(candidates)
    if template_file is None:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 找到包含“{TEMPLATE_FILE_KEYWORD}”的空白报表模板"
        )
    return template_file


def validate_template(workbook: Workbook) -> None:
    """Reject a template that doesn't match the expected layout by content, not just filename.

    Checked before any cell is written, so a wrong/stale template fails fast
    instead of producing a wrong report that only the (more expensive)
    post-save validate_output would catch.
    """
    if len(workbook.sheetnames) != EXPECTED_SHEET_COUNT:
        raise ValueError(
            f"空白模板工作表数量应为 {EXPECTED_SHEET_COUNT}，实际为 {len(workbook.sheetnames)}"
        )
    sheet = workbook[workbook.sheetnames[0]]
    if sheet.title != EXPECTED_SHEET_TITLE:
        raise ValueError(f"空白模板工作表名称应为 {EXPECTED_SHEET_TITLE!r}，实际为 {sheet.title!r}")
    if sheet.max_column != EXPECTED_COLUMN_COUNT:
        raise ValueError(f"空白模板列数应为 {EXPECTED_COLUMN_COUNT}，实际为 {sheet.max_column}")
    if sheet.max_row < MIN_TEMPLATE_ROW_COUNT:
        raise ValueError(f"空白模板行数至少应为 {MIN_TEMPLATE_ROW_COUNT}，实际为 {sheet.max_row}")

    mismatches = [
        f"{coordinate}（应为 {expected!r}，实际为 {sheet[coordinate].value!r}）"
        for coordinate, expected in TEMPLATE_STRUCTURE_CELLS.items()
        if sheet[coordinate].value != expected
    ]
    if mismatches:
        raise ValueError(
            "空白模板结构校验失败，可能选错了文件或模板已被改动：" + "；".join(mismatches)
        )


@dataclass(frozen=True)
class RowRule:
    row: int
    upload_category: str | None
    upload_brands: tuple[str, ...]
    payment_category: str | None
    payment_brands: tuple[str, ...]
    fill_digital: bool = False


ROW_RULES = (
    RowRule(4, "冰箱", ("海尔", "卡萨帝"), "冰箱", ("海尔", "卡萨帝")),
    RowRule(5, "洗衣机", ("海尔", "卡萨帝"), "洗衣机", ("海尔", "卡萨帝")),
    RowRule(6, "冰箱", ("美的", "COLMO", "东芝"), "冰箱", ("美的系", "COLMO", "东芝JX")),
    RowRule(7, "洗衣机", ("美的", "小天鹅", "COLMO"), "洗衣机", ("美的系", "小天鹅", "COLMO")),
    RowRule(8, "冰箱", ("西门子",), "冰箱", ("西门子",)),
    RowRule(9, "洗衣机", ("西门子",), "洗衣机", ("西门子",)),
    RowRule(10, "冰箱", ("博世",), "冰箱", ("博世",)),
    RowRule(11, "洗衣机", ("博世",), "洗衣机", ("博世",)),
    RowRule(12, "冰箱", ("美菱",), "冰箱", ("美菱",)),
    RowRule(13, "洗衣机", ("美菱",), "洗衣机", ("美菱",)),
    RowRule(14, "洗衣机", ("小鸭",), "洗衣机", ("小鸭",)),
    RowRule(15, "国产彩电", ("海信",), "电视", ("海信",)),
    RowRule(16, "国产彩电", ("创维",), "电视", ("创维",)),
    RowRule(17, "国产彩电", ("TCL",), "电视", ("TCL",)),
    RowRule(18, "国产彩电", ("海尔", "卡萨帝"), "电视", ("海尔", "卡萨帝")),
    RowRule(19, "国产彩电", ("华为", "华为（终端）"), "电视", ("华为", "华为（终端）")),
    RowRule(20, "空调", ("格力",), "空调", ("格力",)),
    RowRule(21, "空调", ("美的",), "空调", ("美的",)),
    RowRule(22, "空调", ("海尔", "卡萨帝"), "空调", ("海尔", "卡萨帝")),
    RowRule(23, "空调", ("海信",), "空调", ("海信",)),
    RowRule(24, "空调", ("奥克斯",), "空调", ("奥克斯",)),
    RowRule(25, "空调", ("科龙",), "空调", ("科龙",)),
    RowRule(26, "空调", ("TCL",), "空调", ("TCL",)),
    RowRule(27, "厨卫", ("老板",), "厨卫", ("老板",)),
    RowRule(28, "厨卫", ("方太",), "厨卫", ("方太",)),
    RowRule(29, "厨卫", ("AO史密斯", "A.O.史密斯"), "厨卫", ("AO史密斯", "A.O.史密斯")),
    RowRule(30, "厨卫", ("海尔", "卡萨帝"), "厨卫", ("海尔", "卡萨帝")),
    RowRule(31, "厨卫", ("美的", "COLMO"), "厨卫", ("美的系", "美的", "COLMO")),
    RowRule(32, "厨卫", ("万家乐",), "厨卫", ("万家乐",)),
    RowRule(33, None, (), "数码", (), fill_digital=True),
)

# Every payment-file 财务大类 this report knows how to place. A category
# outside both sets raises rather than being silently folded into 数码 — see
# load_payment_data.
HOUSEHOLD_PAYMENT_CATEGORIES = frozenset({"冰箱", "洗衣机", "电视", "空调", "厨卫"})
DIGITAL_PAYMENT_CATEGORIES = frozenset({"手机", "平板", "智能穿戴"})


@dataclass(frozen=True)
class BrandGroupCategory:
    upload_category: str
    payment_category: str
    brands: tuple[str, ...]


@dataclass(frozen=True)
class BrandGroupRule:
    row: int
    name: str
    categories: tuple[BrandGroupCategory, ...]


@dataclass(frozen=True)
class CountAmount:
    count: int
    amount: Decimal


BRAND_GROUP_RULES = (
    BrandGroupRule(
        38,
        "海尔系",
        (
            BrandGroupCategory("冰箱", "冰箱", ("海尔", "卡萨帝")),
            BrandGroupCategory("洗衣机", "洗衣机", ("海尔", "卡萨帝")),
            BrandGroupCategory("国产彩电", "电视", ("海尔", "卡萨帝")),
            BrandGroupCategory("空调", "空调", ("海尔", "卡萨帝")),
            BrandGroupCategory("厨卫", "厨卫", ("海尔", "卡萨帝")),
        ),
    ),
    BrandGroupRule(
        39,
        "美的系",
        (
            BrandGroupCategory("冰箱", "冰箱", ("美的", "COLMO", "东芝")),
            BrandGroupCategory("洗衣机", "洗衣机", ("美的", "小天鹅", "COLMO")),
            BrandGroupCategory("空调", "空调", ("美的",)),
            BrandGroupCategory("厨卫", "厨卫", ("美的", "COLMO")),
        ),
    ),
    BrandGroupRule(40, "格力", (BrandGroupCategory("空调", "空调", ("格力",)),)),
    BrandGroupRule(
        41,
        "博西",
        (
            BrandGroupCategory("冰箱", "冰箱", ("西门子", "博世")),
            BrandGroupCategory("洗衣机", "洗衣机", ("西门子", "博世")),
        ),
    ),
    BrandGroupRule(
        42,
        "海信系",
        (
            BrandGroupCategory("国产彩电", "电视", ("海信",)),
            BrandGroupCategory("空调", "空调", ("海信", "科龙")),
        ),
    ),
    BrandGroupRule(43, "创维", (BrandGroupCategory("国产彩电", "电视", ("创维",)),)),
    BrandGroupRule(
        44,
        "TCL",
        (
            BrandGroupCategory("国产彩电", "电视", ("TCL",)),
            BrandGroupCategory("空调", "空调", ("TCL",)),
        ),
    ),
)


def to_decimal(value: object) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value))


def to_count(value: object) -> int:
    count = to_decimal(value)
    if count != count.to_integral_value():
        raise ValueError(f"数量应为整数，实际为 {value!r}")
    return int(count)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    replacements = {
        "A.O.史密斯": "AO史密斯",
        "美的系": "美的",
        "东芝JX": "东芝",
        "华为（终端）": "华为",
    }
    return replacements.get(text, text)


def safe_ratio(numerator: Decimal, denominator: Decimal) -> float | None:
    if denominator == 0:
        return None
    ratio = numerator / denominator
    if ratio == 0:
        return None
    return float(ratio)


def decimal_to_cell_value(value: Decimal) -> float | None:
    if value == 0:
        return None
    return float(value)


def apply_filled_style(sheet, coords: tuple[str, ...], font: Font) -> None:
    for coord in coords:
        sheet[coord].alignment = FILLED_ALIGNMENT
        sheet[coord].font = font


def apply_data_number_format(sheet, coords: tuple[str, ...]) -> None:
    for coord in coords:
        sheet[coord].number_format = DATA_NUMBER_FORMAT


def apply_report_font(sheet, font_name: str) -> None:
    for row in sheet.iter_rows():
        for cell in row:
            font = copy(cell.font)
            font.name = font_name
            cell.font = font


def widen_currency_columns(sheet) -> None:
    for column in CURRENCY_COLUMNS:
        sheet.column_dimensions[column].width = CURRENCY_COLUMN_WIDTH


def _open_source_workbook(path: Path, business_name: str, processing_mode: str):
    """Open one of this program's own outputs with calamine, not openpyxl.

    Both inputs are read for one small summary sheet each, but openpyxl parses
    the whole shared string table up front no matter which sheet is wanted and
    no matter that read_only was asked for. Since these files are written by
    XlsxWriter, which shares strings rather than inlining them the way openpyxl
    did, that table now holds every string of the thousands of detail rows
    neither loader looks at — 0.17s to reach a 48-row sheet. calamine parses it
    in Rust and the cost disappears.
    """
    if not path.exists():
        raise FileNotFoundError(
            f"未找到{business_name}：{path}，请先运行“{processing_mode}”处理模式生成该文件"
        )
    return CalamineWorkbook.from_path(str(path))


def _sheet_rows_or_raise(
    workbook, sheet_name: str, source_name: str, business_name: str
) -> list[list[object]]:
    if sheet_name not in workbook.sheet_names:
        raise ValueError(
            f"{source_name} 缺少工作表 {sheet_name!r}，文件可能不是本程序生成的{business_name}"
        )
    return list(calamine_rows(workbook.get_sheet_by_name(sheet_name)))


def _validate_header(
    rows: list[list[object]],
    expected_header: tuple[str, ...],
    source_name: str,
    business_name: str,
) -> None:
    """Check the first row by field name, not just position, so a column reorder or a
    year-label change upstream fails clearly instead of silently misreading data."""
    actual_header = tuple(rows[0][: len(expected_header)]) if rows else ()
    if actual_header != tuple(expected_header):
        raise ValueError(
            f"{source_name} 的{business_name}表头不符合预期：预期 {tuple(expected_header)}，实际 {actual_header}"
        )


def load_upload_data(
    upload_file: Path = UPLOAD_FILE,
) -> tuple[
    dict[tuple[str, str], dict[str, Decimal]],
    dict[str, Decimal],
    dict[str, dict[str, CountAmount]],
]:
    workbook = _open_source_workbook(upload_file, "审核明细", "审核明细（销售用券情况统计）")
    try:
        rows = _sheet_rows_or_raise(workbook, UPLOAD_SHEET_NAME, upload_file.name, "审核明细")
        _validate_header(rows, UPLOAD_HEADER, upload_file.name, "审核明细")
        header_width = len(UPLOAD_HEADER)

        amounts: dict[tuple[str, str], dict[str, Decimal]] = {}
        current_category = ""
        current_brand = ""
        digital_uploaded = Decimal("0")
        digital_not_uploaded = Decimal("0")
        digital_total: Decimal | None = None
        project_metrics: dict[str, dict[str, CountAmount]] = {
            "家电": {},
            "数码": {},
        }

        for row in (r[:header_width] for r in rows[1:]):
            category_raw, brand_raw, status_raw, count_raw, amount_raw = row
            if category_raw:
                current_category = normalize_text(category_raw)
                if not brand_raw:
                    current_brand = ""
            if brand_raw:
                current_brand = normalize_text(brand_raw)
            if current_category in {"财务大类", "合计"}:
                current_brand = ""
                continue

            status = normalize_text(status_raw)
            amount = to_decimal(amount_raw)

            if current_category in project_metrics and not current_brand:
                if status in {"已上传", "未上传", "合计"}:
                    project_metrics[current_category][status] = CountAmount(
                        to_count(count_raw),
                        amount,
                    )
                if current_category == "数码":
                    if status == "已上传":
                        digital_uploaded += amount
                    elif status == "未上传":
                        digital_not_uploaded += amount
                    elif status == "合计":
                        digital_total = amount
                continue

            if not current_brand:
                continue

            brand = current_brand
            key = (current_category, brand)
            bucket = amounts.setdefault(
                key,
                {"已上传": Decimal("0"), "未上传": Decimal("0")},
            )
            bucket[status] = bucket.get(status, Decimal("0")) + amount
    finally:
        workbook.close()

    digital_totals = {
        "发生额": digital_total or digital_uploaded + digital_not_uploaded,
        "上传额": digital_uploaded,
    }

    for project in TABLE3_PROJECT_ROWS:
        missing_statuses = {
            status
            for status in ("已上传", "未上传")
            if status not in project_metrics[project]
        }
        if missing_statuses:
            raise ValueError(
                f"{upload_file.name} 的数据汇总缺少{project}项目汇总状态："
                f"{'、'.join(sorted(missing_statuses))}"
            )

    return amounts, digital_totals, project_metrics


def load_payment_data(
    payment_file: Path = PAYMENT_FILE,
) -> tuple[
    dict[tuple[str, str], Decimal],
    Decimal,
    dict[str, CountAmount],
]:
    workbook = _open_source_workbook(payment_file, "回款明细", "回款明细（家电+数码）")
    try:
        rows = _sheet_rows_or_raise(workbook, PAYMENT_SHEET_NAME, payment_file.name, "回款明细")
        _validate_header(rows, PAYMENT_HEADER, payment_file.name, "回款明细")
        header_width = len(PAYMENT_HEADER)

        amounts: dict[tuple[str, str], Decimal] = {}
        current_category = ""
        digital_amount = Decimal("0")
        project_metrics = {
            "家电": CountAmount(0, Decimal("0")),
            "数码": CountAmount(0, Decimal("0")),
        }

        for row in (r[:header_width] for r in rows[1:]):
            category_raw, brand_raw, amount_raw, count_raw = row
            if category_raw:
                current_category = normalize_text(category_raw)
            category = current_category

            if not category or category == "合计":
                continue
            if not brand_raw:
                continue

            brand = normalize_text(brand_raw)
            amount = to_decimal(amount_raw)
            count = to_count(count_raw)

            if category in HOUSEHOLD_PAYMENT_CATEGORIES:
                key = (category, brand)
                amounts[key] = amounts.get(key, Decimal("0")) + amount
                current = project_metrics["家电"]
                project_metrics["家电"] = CountAmount(
                    current.count + count,
                    current.amount + amount,
                )
            elif category in DIGITAL_PAYMENT_CATEGORIES:
                digital_amount += amount
                current = project_metrics["数码"]
                project_metrics["数码"] = CountAmount(
                    current.count + count,
                    current.amount + amount,
                )
            else:
                raise ValueError(
                    f"门店报表尚未配置回款品类：{category!r}（品牌 {brand!r}）；"
                    "需先确认业务口径（单独一行/并入数码/不纳入报表），"
                    "再把该品类加入 HOUSEHOLD_PAYMENT_CATEGORIES 或 DIGITAL_PAYMENT_CATEGORIES"
                )
    finally:
        workbook.close()

    return amounts, digital_amount, project_metrics


def sum_upload_amount(
    upload_data: dict[tuple[str, str], dict[str, Decimal]],
    category: str | None,
    brands: tuple[str, ...],
) -> tuple[Decimal, Decimal]:
    if not category:
        return Decimal("0"), Decimal("0")

    occurred = Decimal("0")
    uploaded = Decimal("0")
    normalized_brands = {normalize_text(brand) for brand in brands}

    for brand in normalized_brands:
        values = upload_data.get((normalize_text(category), brand))
        if not values:
            continue
        uploaded += values.get("已上传", Decimal("0"))
        occurred += values.get("已上传", Decimal("0")) + values.get("未上传", Decimal("0"))

    return occurred, uploaded


def sum_payment_amount(
    payment_data: dict[tuple[str, str], Decimal],
    category: str | None,
    brands: tuple[str, ...],
) -> Decimal:
    if not category:
        return Decimal("0")

    total = Decimal("0")
    normalized_brands = {normalize_text(brand) for brand in brands}

    for brand in normalized_brands:
        total += payment_data.get((normalize_text(category), brand), Decimal("0"))

    return total


def _rule_claims(
    category_attr: str, brands_attr: str, business_name: str
) -> dict[tuple[str, str], int]:
    """Map each (品类, 品牌) ROW_RULES claims to the row that claims it.

    A rule listing brand aliases that normalize to the same value (e.g. both
    "华为" and "华为（终端）") claims that key twice, which is redundant but
    harmless — sum_upload_amount/sum_payment_amount already dedupe brands via
    a set. Only a *different* row claiming an already-claimed key is a real
    ROW_RULES config conflict, not a data problem.
    """
    claims: dict[tuple[str, str], int] = {}
    for rule in ROW_RULES:
        category = getattr(rule, category_attr)
        if not category:
            continue
        normalized_category = normalize_text(category)
        for brand in getattr(rule, brands_attr):
            key = (normalized_category, normalize_text(brand))
            if key in claims and claims[key] != rule.row:
                raise ValueError(
                    f"ROW_RULES 配置冲突：{business_name}品类品牌 {key} 同时被第 {claims[key]} 行"
                    f"和第 {rule.row} 行规则匹配"
                )
            claims[key] = rule.row
    return claims


def _upload_rule_claims() -> dict[tuple[str, str], int]:
    return _rule_claims("upload_category", "upload_brands", "审核明细")


def _payment_rule_claims() -> dict[tuple[str, str], int]:
    return _rule_claims("payment_category", "payment_brands", "回款明细")


def validate_rule_coverage(
    upload_data: dict[tuple[str, str], dict[str, Decimal]],
    payment_data: dict[tuple[str, str], Decimal],
) -> None:
    """Every non-zero (品类, 品牌) group in the source data must hit exactly one ROW_RULES row.

    Guards against a new brand appearing upstream (or a category rename) that
    ROW_RULES hasn't been updated for: without this, that brand's money is
    simply never summed into any row, yet every existing total-vs-detail check
    still balances, because both sides of that check are already missing the
    same money.
    """
    upload_claims = _upload_rule_claims()
    payment_claims = _payment_rule_claims()

    unmatched: list[str] = [
        f"审核明细未配置规则：{category}/{brand}，发生额 {occurred}"
        for (category, brand), values in upload_data.items()
        for occurred in (values.get("已上传", Decimal("0")) + values.get("未上传", Decimal("0")),)
        if (category, brand) not in upload_claims and occurred != 0
    ]
    unmatched += [
        f"回款明细未配置规则：{category}/{brand}，回款额 {amount}"
        for (category, brand), amount in payment_data.items()
        if (category, brand) not in payment_claims and amount != 0
    ]
    if unmatched:
        raise ValueError(
            "门店报表规则覆盖率校验失败，以下品类/品牌未命中 ROW_RULES，"
            "请确认是否为新增品牌并同步更新规则：" + "；".join(unmatched)
        )


def write_metrics_row(
    sheet,
    row: int,
    occurred_col: str,
    uploaded_col: str,
    paid_col: str,
    upload_ratio_col: str,
    payment_ratio_col: str,
    blank_cols: tuple[str, ...],
    occurred: Decimal,
    uploaded: Decimal,
    paid: Decimal,
    font: Font,
    expected_cells: dict[str, object],
) -> None:
    for col in blank_cols:
        sheet[f"{col}{row}"] = None
        expected_cells[f"{col}{row}"] = None

    occurred_value = decimal_to_cell_value(occurred)
    uploaded_value = decimal_to_cell_value(uploaded)
    paid_value = decimal_to_cell_value(paid)
    upload_ratio_value = safe_ratio(uploaded, occurred)
    payment_ratio_value = safe_ratio(paid, occurred)

    sheet[f"{occurred_col}{row}"] = occurred_value
    sheet[f"{uploaded_col}{row}"] = uploaded_value
    sheet[f"{paid_col}{row}"] = paid_value
    sheet[f"{upload_ratio_col}{row}"] = upload_ratio_value
    sheet[f"{payment_ratio_col}{row}"] = payment_ratio_value

    expected_cells[f"{occurred_col}{row}"] = occurred_value
    expected_cells[f"{uploaded_col}{row}"] = uploaded_value
    expected_cells[f"{paid_col}{row}"] = paid_value
    expected_cells[f"{upload_ratio_col}{row}"] = upload_ratio_value
    expected_cells[f"{payment_ratio_col}{row}"] = payment_ratio_value

    apply_filled_style(
        sheet,
        (
            f"{occurred_col}{row}",
            f"{uploaded_col}{row}",
            f"{upload_ratio_col}{row}",
            f"{paid_col}{row}",
            f"{payment_ratio_col}{row}",
        ),
        font,
    )
    apply_data_number_format(
        sheet,
        (
            f"{occurred_col}{row}",
            f"{uploaded_col}{row}",
            f"{paid_col}{row}",
        ),
    )
    sheet[f"{upload_ratio_col}{row}"].number_format = PERCENT_NUMBER_FORMAT
    sheet[f"{payment_ratio_col}{row}"].number_format = PERCENT_NUMBER_FORMAT


def update_totals_row(
    sheet,
    amount_columns: tuple[str, ...],
    ratio_columns: dict[str, tuple[str, str]],
    source_rows: range,
    total_row: int,
    font: Font,
    expected_cells: dict[str, object],
) -> None:
    totals = {
        column: sum(
            (to_decimal(sheet[f"{column}{row}"].value) for row in source_rows),
            Decimal("0"),
        )
        for column in amount_columns
    }

    for column, total in totals.items():
        value = decimal_to_cell_value(total)
        sheet[f"{column}{total_row}"] = value
        expected_cells[f"{column}{total_row}"] = value

    for ratio_column, (numerator_column, denominator_column) in ratio_columns.items():
        value = safe_ratio(totals[numerator_column], totals[denominator_column])
        sheet[f"{ratio_column}{total_row}"] = value
        expected_cells[f"{ratio_column}{total_row}"] = value

    apply_filled_style(
        sheet,
        tuple(f"{column}{total_row}" for column in (*amount_columns, *ratio_columns)),
        font,
    )
    apply_data_number_format(
        sheet,
        tuple(f"{column}{total_row}" for column in amount_columns),
    )
    for ratio_column in ratio_columns:
        sheet[f"{ratio_column}{total_row}"].number_format = PERCENT_NUMBER_FORMAT


def write_row(
    sheet,
    row_rule: RowRule,
    upload_data: dict[tuple[str, str], dict[str, Decimal]],
    payment_data: dict[tuple[str, str], Decimal],
    digital_upload: dict[str, Decimal],
    digital_payment: Decimal,
    font: Font,
    expected_cells: dict[str, object],
) -> None:
    if row_rule.fill_digital:
        occurred = digital_upload["发生额"]
        uploaded = digital_upload["上传额"]
        paid = digital_payment
        write_metrics_row(
            sheet, row_rule.row, "E", "G", "K", "I", "M", ("D", "F", "H", "J", "L"),
            occurred, uploaded, paid, font, expected_cells,
        )
        return

    occurred, uploaded = sum_upload_amount(upload_data, row_rule.upload_category, row_rule.upload_brands)
    paid = sum_payment_amount(payment_data, row_rule.payment_category, row_rule.payment_brands)
    write_metrics_row(
        sheet, row_rule.row, "D", "F", "J", "H", "L", ("E", "G", "I", "K", "M"),
        occurred, uploaded, paid, font, expected_cells,
    )


def update_totals(sheet, font: Font, expected_cells: dict[str, object]) -> None:
    update_totals_row(
        sheet,
        amount_columns=("D", "E", "F", "G", "J", "K"),
        ratio_columns={"H": ("F", "D"), "I": ("G", "E"), "L": ("J", "D"), "M": ("K", "E")},
        source_rows=DETAIL_ROWS,
        total_row=TOTAL_ROW,
        font=font,
        expected_cells=expected_cells,
    )


def sum_brand_group(
    upload_data: dict[tuple[str, str], dict[str, Decimal]],
    payment_data: dict[tuple[str, str], Decimal],
    categories: tuple[BrandGroupCategory, ...],
) -> tuple[Decimal, Decimal, Decimal]:
    occurred = Decimal("0")
    uploaded = Decimal("0")
    paid = Decimal("0")

    for category in categories:
        category_occurred, category_uploaded = sum_upload_amount(
            upload_data, category.upload_category, category.brands
        )
        occurred += category_occurred
        uploaded += category_uploaded
        paid += sum_payment_amount(payment_data, category.payment_category, category.brands)

    return occurred, uploaded, paid


def write_brand_group_row(
    sheet,
    brand_group_rule: BrandGroupRule,
    upload_data: dict[tuple[str, str], dict[str, Decimal]],
    payment_data: dict[tuple[str, str], Decimal],
    font: Font,
    expected_cells: dict[str, object],
) -> None:
    occurred, uploaded, paid = sum_brand_group(upload_data, payment_data, brand_group_rule.categories)
    write_metrics_row(
        sheet, brand_group_rule.row, "D", "E", "F", "G", "H", (),
        occurred, uploaded, paid, font, expected_cells,
    )


def update_brand_group_totals(sheet, font: Font, expected_cells: dict[str, object]) -> None:
    update_totals_row(
        sheet,
        amount_columns=("D", "E", "F"),
        ratio_columns={"G": ("E", "D"), "H": ("F", "D")},
        source_rows=BRAND_GROUP_DETAIL_ROWS,
        total_row=BRAND_GROUP_TOTAL_ROW,
        font=font,
        expected_cells=expected_cells,
    )


def write_table3(
    sheet,
    upload_metrics: dict[str, dict[str, CountAmount]],
    payment_metrics: dict[str, CountAmount],
    font: Font,
    expected_cells: dict[str, object],
) -> None:
    pending_total = CountAmount(0, Decimal("0"))
    not_uploaded_total = CountAmount(0, Decimal("0"))
    for project, row in TABLE3_PROJECT_ROWS.items():
        uploaded = upload_metrics[project]["已上传"]
        not_uploaded = upload_metrics[project]["未上传"]
        paid = payment_metrics[project]
        pending = CountAmount(
            uploaded.count - paid.count,
            uploaded.amount - paid.amount,
        )
        pending_total = CountAmount(
            pending_total.count + pending.count,
            pending_total.amount + pending.amount,
        )
        not_uploaded_total = CountAmount(
            not_uploaded_total.count + not_uploaded.count,
            not_uploaded_total.amount + not_uploaded.amount,
        )
        values = {
            f"D{row}": pending.count or None,
            f"E{row}": decimal_to_cell_value(pending.amount),
            f"F{row}": not_uploaded.count or None,
            f"G{row}": decimal_to_cell_value(not_uploaded.amount),
        }
        for coordinate, value in values.items():
            sheet[coordinate] = value
            sheet[coordinate].font = font
            expected_cells[coordinate] = value
        apply_data_number_format(sheet, tuple(values))

    total_values = {
        f"D{TABLE3_TOTAL_ROW}": pending_total.count or None,
        f"E{TABLE3_TOTAL_ROW}": decimal_to_cell_value(pending_total.amount),
        f"F{TABLE3_TOTAL_ROW}": not_uploaded_total.count or None,
        f"G{TABLE3_TOTAL_ROW}": decimal_to_cell_value(not_uploaded_total.amount),
    }
    for coordinate, value in total_values.items():
        sheet[coordinate] = value
        sheet[coordinate].font = font
        expected_cells[coordinate] = value
    apply_data_number_format(sheet, tuple(total_values))


def current_timestamp() -> datetime:
    return datetime.now(SHANGHAI_TZ)


def update_header(sheet, timestamp: datetime, expected_cells: dict[str, object]) -> None:
    formatted_timestamp = timestamp.strftime("%Y-%m-%d %H:%M:%S")
    value = (
        f"        {SUBSIDY_YEAR}年（益庄店 ）门店国补上传及回款情况表"
        f"        更新时间：{formatted_timestamp}"
    )
    sheet["A1"] = value
    expected_cells["A1"] = value


def _values_match(expected: object, actual: object) -> bool:
    if expected is None or actual is None:
        return expected == actual
    if isinstance(expected, float) or isinstance(actual, float):
        try:
            return math.isclose(float(expected), float(actual), rel_tol=1e-9, abs_tol=1e-9)
        except (TypeError, ValueError):
            return False
    return expected == actual


def _validate_totals_match_details(sheet, path_name: str) -> None:
    for column in ("D", "E", "F", "G", "J", "K"):
        recomputed = sum(
            (to_decimal(sheet[f"{column}{row}"].value) for row in DETAIL_ROWS),
            Decimal("0"),
        )
        actual = to_decimal(sheet[f"{column}{TOTAL_ROW}"].value)
        if recomputed != actual:
            raise ValueError(
                f"{path_name} 第 {TOTAL_ROW} 行合计校验失败："
                f"{column}{TOTAL_ROW} 应为 {recomputed}，实际为 {actual}"
            )
    for column in ("D", "E", "F"):
        recomputed = sum(
            (to_decimal(sheet[f"{column}{row}"].value) for row in BRAND_GROUP_DETAIL_ROWS),
            Decimal("0"),
        )
        actual = to_decimal(sheet[f"{column}{BRAND_GROUP_TOTAL_ROW}"].value)
        if recomputed != actual:
            raise ValueError(
                f"{path_name} 第 {BRAND_GROUP_TOTAL_ROW} 行合计校验失败："
                f"{column}{BRAND_GROUP_TOTAL_ROW} 应为 {recomputed}，实际为 {actual}"
            )


def _validate_ratios_match_totals(sheet, path_name: str) -> None:
    """Independently recompute each ratio cell from the totals actually saved in
    the file (not from the in-memory value written), catching a numerator/
    denominator mixup that expected-cell comparison alone would miss, since
    that comparison reuses the same safe_ratio() call that wrote the cell."""
    checks = (
        (TOTAL_ROW, {"H": ("F", "D"), "I": ("G", "E"), "L": ("J", "D"), "M": ("K", "E")}),
        (BRAND_GROUP_TOTAL_ROW, {"G": ("E", "D"), "H": ("F", "D")}),
    )
    for total_row, ratio_columns in checks:
        for ratio_column, (numerator_column, denominator_column) in ratio_columns.items():
            numerator = to_decimal(sheet[f"{numerator_column}{total_row}"].value)
            denominator = to_decimal(sheet[f"{denominator_column}{total_row}"].value)
            expected_ratio = safe_ratio(numerator, denominator)
            actual_ratio = sheet[f"{ratio_column}{total_row}"].value
            if not _values_match(expected_ratio, actual_ratio):
                raise ValueError(
                    f"{path_name} {ratio_column}{total_row} 与其对应金额之比不符："
                    f"应为 {expected_ratio}，实际为 {actual_ratio}"
                )


def validate_output(
    path: Path,
    expected_cells: dict[str, object],
    expected_sheet_title: str,
) -> None:
    """Re-read the saved workbook and check it matches what process_store_report intended:
    sheet identity, every written cell exactly, and totals/ratios independently recomputed."""
    workbook = load_workbook(path, data_only=True)
    try:
        if len(workbook.sheetnames) != EXPECTED_SHEET_COUNT:
            raise ValueError(
                f"{path.name} 工作表数量应为 {EXPECTED_SHEET_COUNT}，实际为 {len(workbook.sheetnames)}"
            )
        sheet = workbook[workbook.sheetnames[0]]
        if sheet.title != expected_sheet_title:
            raise ValueError(f"{path.name} 工作表名称应为 {expected_sheet_title!r}，实际为 {sheet.title!r}")
        if sheet.max_column != EXPECTED_COLUMN_COUNT:
            raise ValueError(
                f"{path.name} 列数应为 {EXPECTED_COLUMN_COUNT}，实际为 {sheet.max_column}"
            )
        if "更新时间：" not in str(sheet["A1"].value or ""):
            raise ValueError(f"{path.name} A1 未写入更新时间")

        mismatches = [
            f"{coordinate}（应为 {expected!r}，实际为 {sheet[coordinate].value!r}）"
            for coordinate, expected in expected_cells.items()
            if not _values_match(expected, sheet[coordinate].value)
        ]
        if mismatches:
            raise ValueError(
                f"{path.name} 明细单元格校验失败，共 {len(mismatches)} 处不一致，"
                f"示例：{'；'.join(mismatches[:5])}"
            )

        _validate_totals_match_details(sheet, path.name)
        _validate_ratios_match_totals(sheet, path.name)
    finally:
        workbook.close()


def process_store_report(reporter: ConsoleReporter) -> None:
    timestamp = current_timestamp()
    font_name, _ = resolve_font()
    upload_data, digital_upload, upload_metrics = load_upload_data(UPLOAD_FILE)
    payment_data, digital_payment, payment_metrics = load_payment_data(PAYMENT_FILE)
    validate_rule_coverage(upload_data, payment_data)

    template_file = resolve_template_file()
    workbook = load_workbook(template_file)
    try:
        validate_template(workbook)
        sheet = workbook[workbook.sheetnames[0]]
        font = Font(name=font_name, size=12)

        expected_cells: dict[str, object] = dict(TEMPLATE_STRUCTURE_CELLS)

        for row_rule in ROW_RULES:
            write_row(
                sheet, row_rule, upload_data, payment_data, digital_upload, digital_payment, font, expected_cells
            )

        update_totals(sheet, font, expected_cells)

        for brand_group_rule in BRAND_GROUP_RULES:
            write_brand_group_row(sheet, brand_group_rule, upload_data, payment_data, font, expected_cells)

        update_brand_group_totals(sheet, font, expected_cells)
        write_table3(
            sheet,
            upload_metrics,
            payment_metrics,
            font,
            expected_cells,
        )
        widen_currency_columns(sheet)
        update_header(sheet, timestamp, expected_cells)
        apply_report_font(sheet, font_name)
    except BaseException:
        # save_workbook_atomically (below) takes ownership of closing the
        # workbook once writing succeeds; anything raised before that point —
        # a bad template, a write bug — must close it here instead, or the
        # template file stays locked open (fatal on Windows) until GC.
        workbook.close()
        raise

    save_workbook_atomically(
        workbook,
        OUTPUT_FILE,
        lambda path: validate_output(path, expected_cells, sheet.title),
    )
    reporter.output(OUTPUT_FILE)
