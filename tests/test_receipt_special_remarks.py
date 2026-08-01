import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from processors.common import config
from processors.common.config import load_receipt_special_remark_keys
from processors.receipts import RECEIPTS_SPECIAL_REMARK_KEYS


def load_from(contents: str) -> frozenset[str]:
    """Load match keys from a temporary config file.

    load_receipt_special_remark_keys is lru_cached, so its cache is cleared
    around each call — otherwise the first test's result would be handed to
    every later one regardless of the file it points at.
    """
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "receipt_special_remarks.yaml"
        path.write_text(contents, encoding="utf-8")
        load_receipt_special_remark_keys.cache_clear()
        try:
            with patch.object(config, "RECEIPT_SPECIAL_REMARKS_FILE", path):
                return load_receipt_special_remark_keys()
        finally:
            load_receipt_special_remark_keys.cache_clear()


class ReceiptSpecialRemarksTest(unittest.TestCase):
    def test_expected_match_keys_are_loaded_from_config(self):
        keys = load_receipt_special_remark_keys()

        self.assertIn("2605050233000077", keys)
        self.assertIn("2605030233000049", keys)
        self.assertIn("260426ZH3X000025", keys)

    def test_receipts_uses_configured_special_remark_keys(self):
        self.assertIs(
            RECEIPTS_SPECIAL_REMARK_KEYS,
            load_receipt_special_remark_keys(),
        )


class SpecialRemarkConfigValidationTest(unittest.TestCase):
    def test_valid_keys_are_loaded_and_stripped(self) -> None:
        keys = load_from(
            'match_keys:\n  - "260426ZH3X000025"\n  - "  2605030233000049  "\n'
        )

        self.assertEqual(keys, {"260426ZH3X000025", "2605030233000049"})

    def test_missing_file_yields_empty_set(self) -> None:
        load_receipt_special_remark_keys.cache_clear()
        try:
            with patch.object(
                config,
                "RECEIPT_SPECIAL_REMARKS_FILE",
                Path("/nonexistent/receipt_special_remarks.yaml"),
            ):
                self.assertEqual(load_receipt_special_remark_keys(), frozenset())
        finally:
            load_receipt_special_remark_keys.cache_clear()

    def test_absent_match_keys_yields_empty_set(self) -> None:
        self.assertEqual(load_from("# no match_keys here\n"), frozenset())

    def test_top_level_list_is_rejected(self) -> None:
        # The keys written without the "match_keys:" header — a plain YAML
        # list, which has no .get() to read a key out of.
        with self.assertRaisesRegex(ValueError, "顶层应为映射"):
            load_from('- "2605030233000049"\n')

    def test_top_level_scalar_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "顶层应为映射"):
            load_from('"2605030233000049"\n')

    def test_non_list_match_keys_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "应为列表"):
            load_from('match_keys: "2605030233000049"\n')

    def test_non_string_key_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "必须为非空字符串"):
            load_from("match_keys:\n  - 2605030233000049\n")

    def test_blank_key_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "必须为非空字符串"):
            load_from('match_keys:\n  - "   "\n')

    def test_invalid_date_prefix_is_rejected(self) -> None:
        # 260532 is not a real date, so this key could never match a row.
        with self.assertRaisesRegex(ValueError, "格式无效"):
            load_from('match_keys:\n  - "260532ZH0001"\n')

    def test_key_without_document_number_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "格式无效"):
            load_from('match_keys:\n  - "260503"\n')

    def test_error_names_the_config_file_and_value(self) -> None:
        with self.assertRaises(ValueError) as caught:
            load_from('match_keys:\n  - "260532ZH0001"\n')

        message = str(caught.exception)
        self.assertIn("receipt_special_remarks.yaml", message)
        self.assertIn("260532ZH0001", message)


if __name__ == "__main__":
    unittest.main()
