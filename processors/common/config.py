from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import yaml


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / "config"
BRAND_MAPPING_FILE = CONFIG_DIR / "brand_mapping.yaml"
RECEIPT_SPECIAL_REMARKS_FILE = CONFIG_DIR / "receipt_special_remarks.yaml"
MERCHANTS_FILE = CONFIG_DIR / "merchants.yaml"
PAYMENT_BRANDS_FILE = CONFIG_DIR / "payment_brands.yaml"


@lru_cache(maxsize=1)
def load_brand_mapping() -> dict[str, str]:
    if not BRAND_MAPPING_FILE.exists():
        return {}

    with BRAND_MAPPING_FILE.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}

    return config.get("brand_mapping", {})


@lru_cache(maxsize=1)
def load_receipt_special_remark_keys() -> frozenset[str]:
    if not RECEIPT_SPECIAL_REMARKS_FILE.exists():
        return frozenset()

    with RECEIPT_SPECIAL_REMARKS_FILE.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}

    return frozenset(config.get("match_keys", []))


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


def _brand_keywords(entries: list[dict]) -> BrandKeywords:
    return tuple((entry["brand"], tuple(entry["keywords"])) for entry in entries)


@lru_cache(maxsize=1)
def load_payment_brand_config() -> PaymentBrandConfig:
    """Missing or empty file yields an all-empty config (same convention as
    load_brand_mapping) rather than raising: this module is imported before
    an operator has necessarily set anything up, and payment.py's own row
    processing already raises a clear, row-specific error the first time it
    hits a category or brand it can't resolve."""
    if not PAYMENT_BRANDS_FILE.exists():
        return PaymentBrandConfig()

    with PAYMENT_BRANDS_FILE.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}

    categories = config.get("categories") or {}
    brand_keywords = config.get("brand_keywords") or {}
    brand_normalization = config.get("brand_normalization") or {}
    midea_group = config.get("midea_group") or {}
    brand_model_aliases = config.get("brand_model_aliases") or {}

    return PaymentBrandConfig(
        appliance_categories=categories.get("appliance") or {},
        digital_categories=categories.get("digital") or {},
        appliance_brand_keywords=_brand_keywords(brand_keywords.get("appliance") or []),
        digital_brand_keywords=_brand_keywords(brand_keywords.get("digital") or []),
        appliance_brand_normalization=brand_normalization.get("appliance") or {},
        midea_group_categories=frozenset(midea_group.get("categories") or ()),
        midea_group_brands=frozenset(midea_group.get("brands") or ()),
        appliance_brand_model_aliases=brand_model_aliases.get("appliance") or {},
    )
