import unittest

from processors.common.config import load_receipt_special_remark_keys
from processors.large_appliances import RECEIPTS_SPECIAL_REMARK_KEYS


class ReceiptSpecialRemarksTest(unittest.TestCase):
    def test_expected_match_keys_are_loaded_from_config(self):
        keys = load_receipt_special_remark_keys()

        self.assertIn("2605050233000077", keys)
        self.assertIn("2605030233000049", keys)
        self.assertIn("260426ZH3X000025", keys)

    def test_large_appliances_uses_configured_special_remark_keys(self):
        self.assertIs(
            RECEIPTS_SPECIAL_REMARK_KEYS,
            load_receipt_special_remark_keys(),
        )


if __name__ == "__main__":
    unittest.main()
