from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import yaml

BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / "config"
BRAND_MAPPING_FILE = CONFIG_DIR / "brand_mapping.yaml"
MERCHANTS_FILE = CONFIG_DIR / "merchants.yaml"
PAYMENT_BRANDS_FILE = CONFIG_DIR / "payment_brands.yaml"


def _load_yaml_mapping(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as file:
        loaded = yaml.safe_load(file)
    if loaded is None:
        return {}
    if not isinstance(loaded, dict):
        raise ValueError(
            f"{path.name} 顶层应为映射，实际为 {type(loaded).__name__}"
        )
    return loaded


def _string_mapping(value: object, location: str) -> dict[str, str]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError(
            f"{location} 应为映射，实际为 {type(value).__name__}"
        )
    result: dict[str, str] = {}
    for key, item in value.items():
        if not isinstance(key, str) or not key.strip():
            raise ValueError(f"{location} 的键必须为非空字符串：{key!r}")
        if not isinstance(item, str) or not item.strip():
            raise ValueError(
                f"{location}.{key} 的值必须为非空字符串：{item!r}"
            )
        result[key.strip()] = item.strip()
    return result


def _string_list(value: object, location: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if not isinstance(value, list):
        raise ValueError(
            f"{location} 应为列表，实际为 {type(value).__name__}"
        )
    result: list[str] = []
    for index, item in enumerate(value):
        if not isinstance(item, str) or not item.strip():
            raise ValueError(
                f"{location}[{index}] 必须为非空字符串：{item!r}"
            )
        result.append(item.strip())
    return tuple(result)


@lru_cache(maxsize=1)
def load_brand_mapping() -> dict[str, str]:
    if not BRAND_MAPPING_FILE.exists():
        return {}

    config = _load_yaml_mapping(BRAND_MAPPING_FILE)
    return _string_mapping(
        config.get("brand_mapping"),
        f"{BRAND_MAPPING_FILE.name} 的 brand_mapping",
    )


@lru_cache(maxsize=1)
def load_merchants() -> dict[str, str]:
    """Map each data type (家电 / 数码) to its merchant id.

    The two subsidy programs number the same store differently, so the id is
    per data type rather than per store. Both pipelines read it from here:
    回款明细 filters source rows by it, and 已上传数据 locates its export files
    by it (see submitted_file_marker).
    """
    if not MERCHANTS_FILE.exists():
        raise FileNotFoundError(f"未找到商户编号配置文件：{MERCHANTS_FILE}")

    with MERCHANTS_FILE.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}

    merchants = config.get("merchants") or {}
    if not isinstance(merchants, dict):
        raise ValueError(
            f"{MERCHANTS_FILE.name} 的 merchants 应为「数据类型: 商户编号」映射"
        )
    return {str(key).strip(): str(value).strip() for key, value in merchants.items()}


def merchant_id(data_type: str) -> str:
    merchants = load_merchants()
    merchant = merchants.get(data_type)
    if not merchant:
        raise ValueError(f"{MERCHANTS_FILE.name} 缺少{data_type}的商户编号")
    return merchant


def submitted_file_marker(data_type: str) -> str:
    """Filename marker of one data type's 已上传 export.

    The export is named MER_<商户编号>_<导出时间>_yjhx.xlsx, so the marker is
    derived from the merchant id rather than configured separately — the
    MER_ prefix is the exporter's naming rule, not a per-store setting.
    """
    return f"MER_{merchant_id(data_type)}"


BrandKeywords = tuple[tuple[str, tuple[str, ...]], ...]


@dataclass(frozen=True)
class PaymentBrandConfig:
    """The 回款明细 category/brand data that processors/payment.py used to
    hardcode. Kept here in one dataclass, rather than as separate top-level
    constants in payment.py, so payment.py only ever reads it through
    load_payment_brand_config() and adding a brand/model entry is a
    config-file edit, not a code change."""

    appliance_categories: dict[str, str] = field(default_factory=dict)
    digital_categories: dict[str, str] = field(default_factory=dict)
    appliance_brand_keywords: BrandKeywords = ()
    digital_brand_keywords: BrandKeywords = ()
    appliance_brand_normalization: dict[str, str] = field(default_factory=dict)
    midea_group_categories: frozenset[str] = frozenset()
    midea_group_brands: frozenset[str] = frozenset()
    appliance_brand_model_aliases: dict[str, str] = field(default_factory=dict)


def _brand_keywords(value: object, location: str) -> BrandKeywords:
    if value is None:
        return ()
    if not isinstance(value, list):
        raise ValueError(
            f"{location} 应为列表，实际为 {type(value).__name__}"
        )
    result: list[tuple[str, tuple[str, ...]]] = []
    for index, entry in enumerate(value):
        entry_location = f"{location}[{index}]"
        if not isinstance(entry, dict):
            raise ValueError(
                f"{entry_location} 应为映射，实际为 {type(entry).__name__}"
            )
        brand = entry.get("brand")
        if not isinstance(brand, str) or not brand.strip():
            raise ValueError(f"{entry_location}.brand 必须为非空字符串")
        keywords = _string_list(
            entry.get("keywords"),
            f"{entry_location}.keywords",
        )
        result.append((brand.strip(), keywords))
    return tuple(result)


@lru_cache(maxsize=1)
def load_payment_brand_config() -> PaymentBrandConfig:
    """Missing or empty file yields an all-empty config (same convention as
    load_brand_mapping) rather than raising: this module is imported before
    an operator has necessarily set anything up, and payment.py's own row
    processing already raises a clear, row-specific error the first time it
    hits a category or brand it can't resolve."""
    if not PAYMENT_BRANDS_FILE.exists():
        return PaymentBrandConfig()

    config = _load_yaml_mapping(PAYMENT_BRANDS_FILE)
    categories = config.get("categories")
    brand_keywords = config.get("brand_keywords")
    brand_normalization = config.get("brand_normalization")
    midea_group = config.get("midea_group")
    brand_model_aliases = config.get("brand_model_aliases")
    for location, value in (
        ("categories", categories),
        ("brand_keywords", brand_keywords),
        ("brand_normalization", brand_normalization),
        ("midea_group", midea_group),
        ("brand_model_aliases", brand_model_aliases),
    ):
        if value is not None and not isinstance(value, dict):
            raise ValueError(
                f"{PAYMENT_BRANDS_FILE.name} 的 {location} 应为映射，"
                f"实际为 {type(value).__name__}"
            )
    categories = categories or {}
    brand_keywords = brand_keywords or {}
    brand_normalization = brand_normalization or {}
    midea_group = midea_group or {}
    brand_model_aliases = brand_model_aliases or {}

    return PaymentBrandConfig(
        appliance_categories=_string_mapping(
            categories.get("appliance"),
            f"{PAYMENT_BRANDS_FILE.name} 的 categories.appliance",
        ),
        digital_categories=_string_mapping(
            categories.get("digital"),
            f"{PAYMENT_BRANDS_FILE.name} 的 categories.digital",
        ),
        appliance_brand_keywords=_brand_keywords(
            brand_keywords.get("appliance"),
            f"{PAYMENT_BRANDS_FILE.name} 的 brand_keywords.appliance",
        ),
        digital_brand_keywords=_brand_keywords(
            brand_keywords.get("digital"),
            f"{PAYMENT_BRANDS_FILE.name} 的 brand_keywords.digital",
        ),
        appliance_brand_normalization=_string_mapping(
            brand_normalization.get("appliance"),
            f"{PAYMENT_BRANDS_FILE.name} 的 brand_normalization.appliance",
        ),
        midea_group_categories=frozenset(
            _string_list(
                midea_group.get("categories"),
                f"{PAYMENT_BRANDS_FILE.name} 的 midea_group.categories",
            )
        ),
        midea_group_brands=frozenset(
            _string_list(
                midea_group.get("brands"),
                f"{PAYMENT_BRANDS_FILE.name} 的 midea_group.brands",
            )
        ),
        appliance_brand_model_aliases=_string_mapping(
            brand_model_aliases.get("appliance"),
            f"{PAYMENT_BRANDS_FILE.name} 的 brand_model_aliases.appliance",
        ),
    )
