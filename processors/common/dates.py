import re
from datetime import date, datetime

RECEIPT_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d")


def normalize_receipt_date(value, source_row: int, source_name: str):
    """Parse one 收款单 date cell, or stop the run naming where it came from.

    source_name is the actual input file's name rather than a fixed string:
    an operator reading the error needs to know which file to open, and the
    message used to hardcode "receipt_statistics.xlsx", a name no file in
    this pipeline has ever actually had.
    """
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for date_format in RECEIPT_DATE_FORMATS:
            try:
                return datetime.strptime(text, date_format).date()
            except ValueError:
                continue
    raise ValueError(
        f"{source_name} 第 {source_row} 行日期格式无效：{value!r}；"
        "支持 YYYY-MM-DD、YYYY/MM/DD、YYYYMMDD 或 Excel 日期。"
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
        f"subsidy_coupon_statistics.xlsx: {value!r}"
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
