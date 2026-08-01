"""Payment collection (回款明细) processing for household appliances and digital.

Both subsidy programs export a "补贴明细" workbook per store. This module
normalizes those exports into one workbook holding a detail sheet per data
type plus a shared 汇总 sheet. The data type is detected per source file, so
the two programs never have to be selected by hand.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from processors.common.config import load_payment_brand_config, merchant_id
from processors.common.excel import (
    FONT_SIZE,
    format_sheet,
    load_measurement_font,
    resolve_font,
    save_workbook_atomically,
)
from processors.common.paths import find_data_files

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
DATA_DIR: Path
SOURCE_FILES: tuple[Path, ...]
OUTPUT_FILE = OUTPUT_DIR / "回款明细.xlsx"

# Both programs name their export "...补贴明细"; which program a file belongs
# to is decided per file by detect_profile, not by the search keyword.
SOURCE_FILE_KEYWORD = "补贴明细"
SUPPORTED_SUFFIXES = (".xlsx",)

HEADER_ALIASES = {
    "交易完成时间": "交易时间",
    "销方名称": "销售企业",
    "销售企业名称": "销售企业",
    "核销商编": "商户编号",
    "经销商编号": "商户编号",
    "应收销售金额": "销售金额",
    "参考号": "交易参考号",
    "订单号": "交易订单号",
    "补贴销售金额": "补贴金额",
    "品类": "编码品类",
    "商品明细": "商品名称",
    "发票号码": "发票号",
}
APPLIANCE_DETAIL_HEADERS = (
    "拨付批次",
    "交易时间",
    "交易参考号",
    "交易订单号",
    "销售企业",
    "商户编号",
    "销售金额",
    "其他支付",
    "实收销售金额",
    "补贴金额",
    "补贴比例",
    "SN码",
    "所在地区",
    "商品编码",
    "能耗等级",
    "编码品类",
    "商品名称",
    "发票金额",
    "发票号",
    "ID",
)
DIGITAL_DETAIL_HEADERS = (
    "拨付批次",
    "交易时间",
    "交易参考号",
    "交易订单号",
    "销售企业",
    "商户编号",
    "销售金额",
    "其他支付",
    "实收销售金额",
    "补贴金额",
    "编码品类",
    "商品名称",
    "SN码",
    "IMEI1码",
    "IMEI2码",
    "商品编码",
    "所在地区",
    "发票号",
    "备注",
)
DERIVED_HEADERS = ("财务大类", "品牌")
# Category maps, brand keyword lists, brand normalization, the 美的系 group,
# and model->brand aliases all live in config/payment_brands.yaml — see that
# file for what each section means and why (list order is match priority for
# brand_keywords, and category order there drives sort/summary order).
_BRAND_CONFIG = load_payment_brand_config()
APPLIANCE_CATEGORY_MAP = _BRAND_CONFIG.appliance_categories
DIGITAL_CATEGORY_MAP = _BRAND_CONFIG.digital_categories
APPLIANCE_BRAND_KEYWORDS = _BRAND_CONFIG.appliance_brand_keywords
DIGITAL_BRAND_KEYWORDS = _BRAND_CONFIG.digital_brand_keywords
APPLIANCE_BRAND_NORMALIZATION_MAP = _BRAND_CONFIG.appliance_brand_normalization
MIDEA_GROUP_CATEGORIES = _BRAND_CONFIG.midea_group_categories
MIDEA_GROUP_BRANDS = _BRAND_CONFIG.midea_group_brands
APPLIANCE_BRAND_MODEL_ALIASES = _BRAND_CONFIG.appliance_brand_model_aliases
MODEL_TOKEN_PATTERN = re.compile(
    r"(?=[A-Z0-9._/+\-]*\d)[A-Z0-9]+(?:[._/+\-][A-Z0-9]+)*",
    re.IGNORECASE,
)
SUMMARY_SHEET_NAME = "汇总"
SUMMARY_HEADERS = ["财务大类", "品牌", "补贴金额合计", "补贴金额计数"]
DETAIL_SORT_HEADERS = ("财务大类", "品牌", "交易时间", "商品名称")


@dataclass(frozen=True)
class ProcessingProfile:
    name: str
    detail_sheet_name: str
    detail_headers: tuple[str, ...]
    optional_headers: frozenset[str]
    category_map: dict[str, str]
    brand_keywords: tuple[tuple[str, tuple[str, ...]], ...]
    brand_normalization_map: dict[str, str]
    brand_model_aliases: dict[str, str]


@dataclass(frozen=True)
class PaymentOutputExpectation:
    profile: ProcessingProfile
    row_count: int
    merchant_id: str


PROFILES = {
    "家电": ProcessingProfile(
        name="家电",
        detail_sheet_name="家电明细",
        detail_headers=APPLIANCE_DETAIL_HEADERS,
        optional_headers=frozenset({"发票金额"}),
        category_map=APPLIANCE_CATEGORY_MAP,
        brand_keywords=APPLIANCE_BRAND_KEYWORDS,
        brand_normalization_map=APPLIANCE_BRAND_NORMALIZATION_MAP,
        brand_model_aliases=APPLIANCE_BRAND_MODEL_ALIASES,
    ),
    "数码": ProcessingProfile(
        name="数码",
        detail_sheet_name="数码明细",
        detail_headers=DIGITAL_DETAIL_HEADERS,
        optional_headers=frozenset(),
        category_map=DIGITAL_CATEGORY_MAP,
        brand_keywords=DIGITAL_BRAND_KEYWORDS,
        brand_normalization_map={},
        brand_model_aliases={},
    ),
}
# 汇总 keeps this order, so 家电 always precedes 数码 no matter which source
# file the data directory happens to list first.
PROFILE_ORDER = ("家电", "数码")
FILENAME_KEYWORDS = (
    ("数码补贴明细", "数码"),
    ("以旧换新补贴明细", "家电"),
)


def configure_data_dir(data_dir: Path) -> None:
    global DATA_DIR
    global SOURCE_FILES

    DATA_DIR = data_dir
    SOURCE_FILES = tuple(
        find_data_files(data_dir, SOURCE_FILE_KEYWORD, SUPPORTED_SUFFIXES)
    )


def _iter_actual_rows_with_numbers(worksheet):
    """Yield non-empty rows without trusting the source's declared dimensions."""
    if hasattr(worksheet, "reset_dimensions"):
        worksheet.reset_dimensions()
    for row_number, row in enumerate(worksheet.iter_rows(min_col=1, values_only=True), 1):
        values = tuple(row)
        if any(value not in (None, "") for value in values):
            yield row_number, values


def _iter_actual_rows(worksheet):
    for _, values in _iter_actual_rows_with_numbers(worksheet):
        yield values


def _cell_value(row: tuple, index: int):
    return row[index] if index < len(row) else None


def _normalize_header_names(row: tuple) -> list[str]:
    headers = []
    for value in row:
        normalized = value.strip() if isinstance(value, str) else value
        headers.append(
            HEADER_ALIASES.get(normalized, normalized)
            if normalized is not None
            else ""
        )
    return headers


def _is_header_row(row: tuple) -> bool:
    return any(
        isinstance(value, str) and value.strip() == "拨付批次" for value in row
    )


def _missing_required_headers(
    headers: list[str], profile: ProcessingProfile
) -> list[str]:
    present = set(headers)
    return [
        name
        for name in profile.detail_headers
        if name not in profile.optional_headers and name not in present
    ]


def _detect_profile_from_filename(source_name: str) -> ProcessingProfile | None:
    matches = {
        profile_name
        for keyword, profile_name in FILENAME_KEYWORDS
        if keyword in source_name
    }
    if len(matches) == 1:
        return PROFILES[next(iter(matches))]
    return None


def detect_profile(path: Path, source_name: str | None = None) -> ProcessingProfile:
    """Pick the processing profile from the file name, else the header row."""
    if source_name:
        profile = _detect_profile_from_filename(source_name)
        if profile is not None:
            print(f"已按文件名识别数据类型：{profile.name}")
            return profile
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book.worksheets:
            if sheet.title == SUMMARY_SHEET_NAME:
                continue
            for row in _iter_actual_rows(sheet):
                if not _is_header_row(row):
                    continue
                headers = _normalize_header_names(row)
                matches = {
                    name: profile
                    for name, profile in PROFILES.items()
                    if not _missing_required_headers(headers, profile)
                }
                if len(matches) == 1:
                    profile = next(iter(matches.values()))
                    print(f"已按表头识别数据类型：{profile.name}")
                    return profile
                if matches:
                    raise ValueError(
                        "无法确定数据类型，表头同时符合："
                        f"{sorted(matches)}，请手动拆分原始数据"
                    )
                details = "；".join(
                    f"{profile.name}缺少 {_missing_required_headers(headers, profile)}"
                    for profile in PROFILES.values()
                )
                raise ValueError(f"无法识别数据类型：{details}")
    finally:
        book.close()
    raise ValueError(f"工作簿 {path.name!r} 中没有找到明细表头")


def _get_source_positions(
    row: tuple, sheet_name: str, profile: ProcessingProfile
) -> dict[str, int]:
    """Normalize and validate one source header row."""
    headers = _normalize_header_names(row)
    header_counts = Counter(name for name in headers if name)
    duplicates = sorted(name for name, count in header_counts.items() if count > 1)
    if duplicates:
        raise ValueError(f"工作表 {sheet_name!r} 存在重复表头：{duplicates}")

    positions = {name: index for index, name in enumerate(headers) if name}
    missing = _missing_required_headers(headers, profile)
    if missing:
        raise ValueError(f"工作表 {sheet_name!r} 缺少字段：{missing}")
    return positions


def _collect_normalized_detail(
    source,
    *,
    profile: ProcessingProfile,
    merchant_id: str,
    formula_workbook=None,
) -> tuple[list[list[object]], int]:
    """Normalize one source sheet without creating output Cell objects yet."""
    source_positions: dict[str, int] = {}
    normalized_rows: list[list[object]] = []
    unidentified_brands = 0
    merchant_index = profile.detail_headers.index("商户编号")
    category_index = profile.detail_headers.index("编码品类")
    product_index = profile.detail_headers.index("商品名称")
    subsidy_index_in_headers = profile.detail_headers.index("补贴金额")
    subsidy_header = "补贴金额"
    formula_rows_iterator = None

    def _find_formula_value(row_number: int, column_index: int):
        # Advances the formula-view iterator forward only as far as needed;
        # opens the (expensive) formula workbook lazily on first use.
        nonlocal formula_rows_iterator
        if formula_rows_iterator is None:
            formula_sheet = formula_workbook()[source.title]
            formula_rows_iterator = _iter_actual_rows_with_numbers(formula_sheet)
        for formula_row_number, formula_row in formula_rows_iterator:
            if formula_row_number == row_number:
                return _cell_value(formula_row, column_index)
            if formula_row_number > row_number:
                break
        return None

    for row_number, row in _iter_actual_rows_with_numbers(source):
        if not source_positions and _is_header_row(row):
            source_positions = _get_source_positions(row, source.title, profile)
            continue

        if not source_positions:
            continue

        normalized = [
            _cell_value(row, source_positions[name]) if name in source_positions else None
            for name in profile.detail_headers
        ]
        row_merchant_id = normalized[merchant_index]
        if str(row_merchant_id).strip() == merchant_id:
            subsidy_value = normalized[subsidy_index_in_headers]
            if (
                subsidy_value is None
                and formula_workbook is not None
                and subsidy_header in source_positions
            ):
                formula = _find_formula_value(row_number, source_positions[subsidy_header])
                if isinstance(formula, str) and formula.startswith("="):
                    raise ValueError(
                        f"工作表 {source.title!r} 第 {row_number} 行字段 "
                        f"{subsidy_header!r} 是公式但没有缓存计算结果；"
                        "请先用 Excel/WPS 打开并保存，或将公式转换为数值"
                    )
            encoded_category = normalized[category_index]
            financial_category = profile.category_map.get(encoded_category)
            if financial_category is None:
                raise ValueError(
                    f"工作表 {source.title!r} 第 {row_number} 行存在未配置的编码品类："
                    f"{encoded_category!r}"
                )
            product_name = normalized[product_index]
            brand = _extract_brand(product_name, profile)
            brand = _normalize_financial_brand(brand, financial_category)
            if brand is None:
                unidentified_brands += 1
            normalized_rows.append(normalized + [financial_category, brand])

    if not source_positions:
        raise ValueError(f"工作表 {source.title!r} 未找到明细表头")
    return normalized_rows, unidentified_brands


def _extract_brand(product_name, profile: ProcessingProfile) -> str | None:
    if product_name in (None, ""):
        return None
    normalized_name = str(product_name).strip().casefold()
    for brand, keywords in profile.brand_keywords:
        if any(keyword.casefold() in normalized_name for keyword in keywords):
            return profile.brand_normalization_map.get(brand, brand)
    for model, brand in profile.brand_model_aliases.items():
        if model.casefold() in normalized_name:
            return profile.brand_normalization_map.get(brand, brand)
    return None


def _normalize_financial_brand(
    brand: str | None, financial_category: str
) -> str | None:
    if (
        financial_category in MIDEA_GROUP_CATEGORIES
        and brand in MIDEA_GROUP_BRANDS
    ):
        return "美的系"
    return brand


def _extract_model_tokens(product_name) -> set[str]:
    if product_name in (None, ""):
        return set()
    return {
        token.upper().strip("._/+-")
        for token in MODEL_TOKEN_PATTERN.findall(str(product_name).upper())
        if len(token.strip("._/+-")) >= 6
    }


def _infer_missing_brands(
    rows: list[list[object]], headers: tuple[str, ...]
) -> int:
    """Infer brands in plain records before they become worksheet cells."""
    product_column = headers.index("商品名称")
    brand_column = headers.index("品牌")
    token_brands: dict[str, Counter] = defaultdict(Counter)

    for row in rows:
        brand = row[brand_column]
        if brand in (None, ""):
            continue
        for token in _extract_model_tokens(row[product_column]):
            token_brands[token][brand] += 1

    unique_token_brand = {
        token: next(iter(counts))
        for token, counts in token_brands.items()
        if len(counts) == 1
    }
    inferred = 0
    for row in rows:
        if row[brand_column] not in (None, ""):
            continue
        candidates = {
            unique_token_brand[token]
            for token in _extract_model_tokens(row[product_column])
            if token in unique_token_brand
        }
        if len(candidates) == 1:
            row[brand_column] = candidates.pop()
            inferred += 1
    return inferred


def _category_order(category_map: dict[str, str]) -> dict[str, int]:
    return {
        category: index
        for index, category in enumerate(dict.fromkeys(category_map.values()))
    }


def _sort_scalar(value) -> tuple[int, str]:
    """Return a type-stable scalar sort key for worksheet values."""
    if value in (None, ""):
        return (1, "")
    if hasattr(value, "isoformat") and not isinstance(value, str):
        return (0, value.isoformat())
    return (0, str(value).strip())


def _sort_detail_rows(
    rows: list[list[object]],
    headers: tuple[str, ...],
    category_map: dict[str, str],
) -> int:
    """Sort plain detail records before writing the output worksheet."""
    data_row_count = len(rows)
    if data_row_count <= 1:
        return data_row_count

    header_positions = {header: index for index, header in enumerate(headers)}
    missing = [header for header in DETAIL_SORT_HEADERS if header not in header_positions]
    if missing:
        raise ValueError(f"明细缺少排序字段：{missing}")

    category_column = header_positions["财务大类"]
    brand_column = header_positions["品牌"]
    transaction_time_column = header_positions["交易时间"]
    product_column = header_positions["商品名称"]
    category_order = _category_order(category_map)

    rows.sort(
        key=lambda item: (
            category_order.get(
                str(item[category_column]),
                len(category_order),
            ),
            _sort_scalar(item[brand_column]),
            _sort_scalar(item[transaction_time_column]),
            _sort_scalar(item[product_column]),
        ),
    )
    return data_row_count


def _sum_detail_groups(detail_sheets) -> dict[tuple[str, str], list]:
    groups: dict[tuple[str, str], list] = {}
    for detail_sheet in detail_sheets:
        rows = detail_sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if header is None:
            raise ValueError(f"{detail_sheet.title}缺少表头")
        header_positions = {
            value: index
            for index, value in enumerate(header)
            if value not in (None, "")
        }
        category_column = header_positions["财务大类"]
        brand_column = header_positions["品牌"]
        subsidy_column = header_positions["补贴金额"]

        for row_number, row in enumerate(rows, start=2):
            category = row[category_column]
            brand = row[brand_column]
            subsidy = row[subsidy_column]
            if category in (None, ""):
                raise ValueError(
                    f"{detail_sheet.title}第 {row_number} 行缺少财务大类"
                )
            key = (str(category), "" if brand in (None, "") else str(brand))
            if key not in groups:
                groups[key] = [Decimal("0"), 0]
            if subsidy not in (None, ""):
                try:
                    subsidy_amount = Decimal(str(subsidy))
                except (InvalidOperation, ValueError) as error:
                    raise ValueError(
                        f"{detail_sheet.title}第 {row_number} 行补贴金额不是有效数值"
                    ) from error
                groups[key][0] += subsidy_amount
                groups[key][1] += -1 if subsidy_amount < 0 else 1
    return groups


def _append_summary_row(summary_sheet, category, brand, amount: Decimal, count: int) -> None:
    rounded_amount = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    summary_sheet.append([category, brand, float(rounded_amount), count])
    summary_sheet.cell(summary_sheet.max_row, 3).number_format = "#,##0.00"
    summary_sheet.cell(summary_sheet.max_row, 4).number_format = "0"


def _build_summary_sheet(
    summary_sheet, sections: list[tuple[str, list, dict[str, str]]]
) -> tuple[int, list[int]]:
    """Write one subtotaled block per data type into 汇总; sections keep their
    given order so 家电 and 数码 never interleave, and are separated by one blank
    row instead of a labeled title row."""
    summary_sheet.append(SUMMARY_HEADERS)
    total_groups = 0
    bold_rows = []
    grand_amount = Decimal("0")
    grand_count = 0
    written_sections = 0

    for _label, detail_sheets, category_map in sections:
        groups = _sum_detail_groups(detail_sheets)
        if not groups:
            continue
        if written_sections:
            summary_sheet.append([None, None, None, None])
        written_sections += 1

        category_order = _category_order(category_map)
        for (category, brand), (amount, count) in sorted(
            groups.items(),
            key=lambda item: (
                category_order.get(item[0][0], len(category_order)),
                item[0][1],
            ),
        ):
            _append_summary_row(summary_sheet, category, brand, amount, count)
        total_groups += len(groups)

        section_amount = sum((values[0] for values in groups.values()), Decimal("0"))
        section_count = sum(values[1] for values in groups.values())
        _append_summary_row(summary_sheet, "合计", None, section_amount, section_count)
        bold_rows.append(summary_sheet.max_row)
        grand_amount += section_amount
        grand_count += section_count

    _append_summary_row(summary_sheet, "合计", None, grand_amount, grand_count)
    bold_rows.append(summary_sheet.max_row)
    return total_groups, bold_rows


def _merge_summary_categories(summary_sheet) -> int:
    category_column = 1
    first_data_row = 2
    last_data_row = summary_sheet.max_row - 1  # Exclude the bottom total row.
    if last_data_row < first_data_row:
        return 0

    merged_groups = 0
    group_start = first_data_row
    current_value = summary_sheet.cell(group_start, category_column).value
    for row in range(first_data_row + 1, last_data_row + 2):
        next_value = (
            summary_sheet.cell(row, category_column).value
            if row <= last_data_row
            else object()
        )
        if next_value == current_value:
            continue
        group_end = row - 1
        if group_end > group_start:
            summary_sheet.merge_cells(
                start_row=group_start,
                start_column=category_column,
                end_row=group_end,
                end_column=category_column,
            )
            merged_groups += 1
        summary_sheet.cell(group_start, category_column).alignment = Alignment(
            horizontal="center", vertical="center"
        )
        group_start = row
        current_value = next_value
    return merged_groups


def _process_sources(
    sources: list[Path],
    profile: ProcessingProfile,
    merchant_id: str,
    target_book: Workbook,
):
    """Merge every source file of one data type into a single detail sheet."""
    merged_sheet = target_book.create_sheet(profile.detail_sheet_name)
    merged_rows = 0
    merged_header_written = False
    normalized_rows: list[list[object]] = []

    unidentified_brands = 0
    for path in sources:
        # Use values cached by Excel for formula cells; openpyxl does not
        # calculate formulas.
        source_book = load_workbook(path, read_only=True, data_only=True)
        formula_book: object | None = None

        def get_formula_book(path=path):
            # Opened only if a subsidy cell needs checking for an uncached
            # formula; a second full workbook parse is otherwise skipped.
            nonlocal formula_book
            if formula_book is None:
                formula_book = load_workbook(path, read_only=True, data_only=False)
            return formula_book

        try:
            for source_sheet in source_book.worksheets:
                if source_sheet.title == SUMMARY_SHEET_NAME:
                    continue
                sheet_rows, missing_brands = _collect_normalized_detail(
                    source_sheet,
                    profile=profile,
                    merchant_id=merchant_id,
                    formula_workbook=get_formula_book,
                )
                normalized_rows.extend(sheet_rows)
                merged_rows += len(sheet_rows)
                unidentified_brands += missing_brands
                merged_header_written = True
        finally:
            source_book.close()
            if formula_book is not None:
                formula_book.close()

    if not merged_header_written:
        raise ValueError(f"{profile.name}的原始数据中没有可整合的明细 Sheet")
    if merged_rows == 0:
        source_names = "、".join(path.name for path in sources)
        raise ValueError(
            f"{profile.name}来源文件中未找到商户 {merchant_id} 的数据："
            f"{source_names}；请检查 config/merchants.yaml"
        )
    output_headers = profile.detail_headers + DERIVED_HEADERS
    inferred_brands = _infer_missing_brands(normalized_rows, output_headers)
    unidentified_brands -= inferred_brands
    sorted_detail_rows = _sort_detail_rows(
        normalized_rows,
        output_headers,
        profile.category_map,
    )
    merged_sheet.append(output_headers)
    for row in normalized_rows:
        merged_sheet.append(row)
    print(
        f"已处理{profile.name}明细：Sheet {merged_sheet.title}，"
        f"商户 {merchant_id} 共 {merged_rows} 条，"
        f"品牌推断 {inferred_brands} 条、未识别 {unidentified_brands} 条，"
        f"明细排序 {sorted_detail_rows} 条"
    )
    return merged_sheet


def _classify_sources() -> dict[str, list[Path]]:
    """Group the data directory's subsidy exports by data type."""
    if not SOURCE_FILES:
        raise FileNotFoundError(
            f"未在 {DATA_DIR} 找到包含“{SOURCE_FILE_KEYWORD}”的回款原始数据文件"
        )

    classified: dict[str, list[Path]] = {}
    for source in SOURCE_FILES:
        profile = detect_profile(source, source.name)
        classified.setdefault(profile.name, []).append(source)
    return classified


def build_workbook() -> tuple[Workbook, list[tuple[ProcessingProfile, object]], int]:
    """Build the payment workbook in memory and return it with its 汇总 stats."""
    target_book = Workbook()
    target_book.remove(target_book.active)
    classified = _classify_sources()
    sections: list[tuple[ProcessingProfile, object]] = []
    for profile_name in PROFILE_ORDER:
        sources = classified.get(profile_name)
        if not sources:
            continue
        sections.append(
            (
                PROFILES[profile_name],
                _process_sources(
                    sources,
                    PROFILES[profile_name],
                    merchant_id(profile_name),
                    target_book,
                ),
            )
        )

    if not sections:
        raise ValueError("没有任何明细数据可输出")

    summary_sheet = target_book.create_sheet(SUMMARY_SHEET_NAME, 0)
    summary_groups, bold_rows = _build_summary_sheet(
        summary_sheet,
        [
            (profile.name, [detail_sheet], profile.category_map)
            for profile, detail_sheet in sections
        ],
    )

    font_name, font_path = resolve_font()
    measurement_font = load_measurement_font(font_path)
    for sheet in target_book.worksheets:
        format_sheet(sheet, font_name, measurement_font)
    # 汇总 is a report with merged category cells and a blank separator row,
    # not a filterable table.
    summary_sheet.auto_filter.ref = None
    _merge_summary_categories(summary_sheet)
    for row in bold_rows:
        for cell in summary_sheet[row]:
            cell.font = Font(name=font_name, size=FONT_SIZE, bold=True)

    print(f"汇总分 {len(sections)} 类共 {summary_groups} 组，字体 {font_name}")
    return target_book, sections, summary_groups


def _summary_snapshot(worksheet) -> list[tuple[object, ...]]:
    return [
        tuple(row)
        for row in worksheet.iter_rows(
            min_row=1,
            max_col=len(SUMMARY_HEADERS),
            values_only=True,
        )
    ]


def validate_output(
    path: Path,
    expectations: dict[str, PaymentOutputExpectation],
    expected_summary: list[tuple[object, ...]],
) -> None:
    """Re-read the saved workbook and cross-check details against the summary."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        expected_sheet_names = [SUMMARY_SHEET_NAME, *expectations]
        if workbook.sheetnames != expected_sheet_names:
            raise ValueError(
                f"{path.name} 工作表应为 {expected_sheet_names}，"
                f"实际为 {workbook.sheetnames}"
            )

        detail_sheets = []
        for sheet_name, expectation in expectations.items():
            sheet = workbook[sheet_name]
            expected_header = expectation.profile.detail_headers + DERIVED_HEADERS
            actual_header = tuple(cell.value for cell in sheet[1])
            if actual_header != expected_header:
                raise ValueError(
                    f"{path.name} 的“{sheet_name}”表头不符合要求："
                    f"预期 {expected_header}，实际 {actual_header}"
                )

            written = max(sheet.max_row - 1, 0)
            if written != expectation.row_count:
                raise ValueError(
                    f"{path.name} 的“{sheet_name}”应有 "
                    f"{expectation.row_count} 行数据，"
                    f"实际 {written} 行"
                )

            merchant_column = expected_header.index("商户编号")
            wrong_merchants = [
                row_number
                for row_number, row in enumerate(
                    sheet.iter_rows(min_row=2, values_only=True),
                    start=2,
                )
                if str(row[merchant_column]).strip() != expectation.merchant_id
            ]
            if wrong_merchants:
                raise ValueError(
                    f"{path.name} 的“{sheet_name}”存在非目标商户数据，"
                    f"首个异常行：{wrong_merchants[0]}"
                )
            detail_sheets.append(sheet)

        summary_sheet = workbook[SUMMARY_SHEET_NAME]
        actual_summary = _summary_snapshot(summary_sheet)
        if actual_summary != expected_summary:
            raise ValueError(f"{path.name} 的“{SUMMARY_SHEET_NAME}”内容校验失败")

        detail_groups = _sum_detail_groups(detail_sheets)
        detail_total = sum(
            (values[0] for values in detail_groups.values()),
            Decimal("0"),
        ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        detail_count = sum(values[1] for values in detail_groups.values())
        grand_total = actual_summary[-1]
        if (
            grand_total[0] != "合计"
            or Decimal(str(grand_total[2])).quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )
            != detail_total
            or grand_total[3] != detail_count
        ):
            raise ValueError(
                f"{path.name} 的汇总总计与明细不一致："
                f"明细为 {detail_total} / {detail_count}"
            )
    finally:
        workbook.close()


def process_payment_files() -> None:
    target_book, sections, _ = build_workbook()
    expectations = {
        detail_sheet.title: PaymentOutputExpectation(
            profile=profile,
            row_count=max(detail_sheet.max_row - 1, 0),
            merchant_id=merchant_id(profile.name),
        )
        for profile, detail_sheet in sections
    }
    expected_summary = _summary_snapshot(target_book[SUMMARY_SHEET_NAME])
    save_workbook_atomically(
        target_book,
        OUTPUT_FILE,
        lambda path: validate_output(path, expectations, expected_summary),
    )
    print(f"处理完成：{OUTPUT_FILE}")
