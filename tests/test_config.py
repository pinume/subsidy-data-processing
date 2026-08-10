import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from processors.common import config


class MerchantConfigValidationTest(unittest.TestCase):
    def _load(self, content: str):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "merchants.yaml"
            path.write_text(content, encoding="utf-8")
            with patch.object(config, "MERCHANTS_FILE", path):
                config.load_merchants.cache_clear()
                try:
                    return config.load_merchants()
                finally:
                    config.load_merchants.cache_clear()

    def test_rejects_null_appliance_merchant_id(self) -> None:
        with self.assertRaisesRegex(ValueError, r"merchants\.家电.*非空字符串"):
            self._load(
                "merchants:\n"
                "  家电:\n"
                "  数码: DIGITAL-ID\n"
            )

    def test_rejects_non_string_digital_merchant_id(self) -> None:
        with self.assertRaisesRegex(ValueError, r"merchants\.数码.*非空字符串"):
            self._load(
                "merchants:\n"
                "  家电: APPLIANCE-ID\n"
                "  数码: 12345\n"
            )

    def test_rejects_missing_required_merchant_id(self) -> None:
        with self.assertRaisesRegex(ValueError, "缺少数码的商户编号"):
            self._load(
                "merchants:\n"
                "  家电: APPLIANCE-ID\n"
            )

    def test_strips_valid_merchant_ids(self) -> None:
        self.assertEqual(
            self._load(
                "merchants:\n"
                "  家电: ' APPLIANCE-ID '\n"
                "  数码: ' DIGITAL-ID '\n"
            ),
            {"家电": "APPLIANCE-ID", "数码": "DIGITAL-ID"},
        )


if __name__ == "__main__":
    unittest.main()
