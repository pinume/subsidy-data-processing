from functools import lru_cache
from pathlib import Path

import yaml


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / "config"
BRAND_MAPPING_FILE = CONFIG_DIR / "brand_mapping.yaml"


@lru_cache(maxsize=1)
def load_brand_mapping() -> dict[str, str]:
    if not BRAND_MAPPING_FILE.exists():
        return {}

    with BRAND_MAPPING_FILE.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}

    return config.get("brand_mapping", {})
