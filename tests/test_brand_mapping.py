import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from processors.common import config
from processors.common.config import (
    load_brand_mapping,
    load_payment_brand_config,
)
from processors.coupons.sources import COUPON_BRAND_REPLACEMENTS


def load_from(contents: str, *, payment: bool = False):
    loader = load_payment_brand_config if payment else load_brand_mapping
    path_attribute = "PAYMENT_BRANDS_FILE" if payment else "BRAND_MAPPING_FILE"
    filename = "payment_brands.yaml" if payment else "brand_mapping.yaml"
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / filename
        path.write_text(contents, encoding="utf-8")
        loader.cache_clear()
        try:
            with patch.object(config, path_attribute, path):
                return loader()
        finally:
            loader.cache_clear()


class BrandMappingTest(unittest.TestCase):
    def test_expected_brand_replacements_are_loaded_from_config(self):
        mapping = load_brand_mapping()

        self.assertEqual(mapping["卡萨帝"], "海尔")
        self.assertEqual(mapping["统帅"], "海尔")
        self.assertEqual(mapping["COLMO"], "美的")
        self.assertEqual(mapping["华凌"], "美的")

    def test_coupon_sources_uses_configured_brand_mapping(self):
        # Content, not identity: load_from() clears the lru_cache in its
        # finally block, so any other test that ran before this one makes
        # load_brand_mapping() return a fresh dict that can never be the
        # import-time object COUPON_BRAND_REPLACEMENTS was bound to.
        # assertIs here would fail or pass depending on test order.
        self.assertEqual(COUPON_BRAND_REPLACEMENTS, load_brand_mapping())

    def test_values_are_nonempty_strings_and_are_stripped(self) -> None:
        self.assertEqual(
            load_from('brand_mapping:\n  " 海尔 ": " 海尔系 "\n'),
            {"海尔": "海尔系"},
        )


class BrandConfigValidationTest(unittest.TestCase):
    def test_payment_brand_keywords_preserve_boundary_spaces(self) -> None:
        loaded = load_from(
            'brand_keywords:\n  appliance:\n    - brand: LG\n      keywords: ["LG "]\n',
            payment=True,
        )

        self.assertEqual(loaded.appliance_brand_keywords, (("LG", ("LG ",)),))

    def test_brand_mapping_rejects_invalid_shapes(self) -> None:
        cases = (
            ('- "海尔"\n', "顶层应为映射"),
            ("brand_mapping: []\n", "brand_mapping 应为映射"),
            ("brand_mapping:\n  海尔: []\n", "海尔 的值必须为非空字符串"),
        )
        for contents, expected in cases:
            with self.subTest(contents=contents):
                with self.assertRaisesRegex(ValueError, expected) as caught:
                    load_from(contents)
                self.assertIn("brand_mapping.yaml", str(caught.exception))

    def test_payment_brand_config_rejects_invalid_shapes(self) -> None:
        cases = (
            ('- "categories"\n', "顶层应为映射"),
            (
                "categories:\n  appliance: []\n",
                "categories.appliance 应为映射",
            ),
            (
                "brand_keywords:\n  appliance:\n    - keywords: [海尔]\n",
                r"brand_keywords\.appliance\[0\]\.brand",
            ),
            (
                "brand_keywords:\n  appliance:\n"
                "    - brand: 海尔\n      keywords: 海尔\n",
                r"brand_keywords\.appliance\[0\]\.keywords 应为列表",
            ),
            (
                "midea_group:\n  categories: {}\n",
                "midea_group.categories 应为列表",
            ),
            (
                "brand_model_aliases:\n  appliance:\n    MODEL: []\n",
                "brand_model_aliases.appliance.MODEL 的值必须为非空字符串",
            ),
        )
        for contents, expected in cases:
            with self.subTest(contents=contents):
                with self.assertRaisesRegex(ValueError, expected) as caught:
                    load_from(contents, payment=True)
                self.assertIn("payment_brands.yaml", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
