import re
from datetime import date, datetime


def normalize_receipt_date(value, source_row: int):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for date_format in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return datetime.strptime(text, date_format).date()
            except ValueError:
                continue
    raise ValueError(
        f"Invalid date at row {source_row} in receipt_statistics.XLS: {value!r}"
    )


def normalize_coupon_date(value: object, source_row: int) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for date_format in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    raise ValueError(
        f"Invalid document date at row {source_row} in "
        f"subsidy_coupon_statistics.XLS: {value!r}"
    )


def normalize_receipt_identifier(value: object) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_document_number(value: object) -> str:
    return normalize_receipt_identifier(value).removeprefix("收款")


def receipt_match_key(receipt_date: date, document_number: str) -> str:
    return f"{receipt_date:%y%m%d}{document_number}"


def is_valid_original_invoice_number(value: str) -> bool:
    if not re.fullmatch(r"\d{6}.+", value):
        return False
    try:
        datetime.strptime(value[:6], "%y%m%d")
    except ValueError:
        return False
    return True
