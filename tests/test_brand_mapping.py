import unittest

from processors.common.config import load_brand_mapping
from processors.large_appliances import COUPON_BRAND_REPLACEMENTS


class BrandMappingTest(unittest.TestCase):
    def test_expected_brand_replacements_are_loaded_from_config(self):
        mapping = load_brand_mapping()

        self.assertEqual(mapping["卡萨帝"], "海尔")
        self.assertEqual(mapping["统帅"], "海尔")
        self.assertEqual(mapping["COLMO"], "美的")
        self.assertEqual(mapping["华凌"], "美的")

    def test_large_appliances_uses_configured_brand_mapping(self):
        self.assertIs(COUPON_BRAND_REPLACEMENTS, load_brand_mapping())


if __name__ == "__main__":
    unittest.main()
