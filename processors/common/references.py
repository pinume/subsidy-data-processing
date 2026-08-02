"""检索参考号 handling, shared by the submitted export and the coupon sources.

The same reference appears in the upload export, in the generated 已上传
workbook, and in the appliance reference supplement; each of those used to
normalize and format-check it on its own, so the wording of the error and even
the shape of the check had started to drift apart. They all go through here
instead.
"""

import re

from processors.common.dates import normalize_receipt_identifier

REFERENCE_RE = re.compile(r"\d{11}[A-Z]")
REFERENCE_FORMAT_HINT = "正确格式应为11位数字后跟一个大写字母"


def normalize_reference(value: object) -> str:
    """The one spelling every comparison uses: trimmed, then upper-cased.

    Callers must compare normalized values on both sides — the source export
    is not consistent about case, so a raw comparison silently misses matches.
    """
    return normalize_receipt_identifier(value).upper()


def validated_reference(
    value: object,
    location: str,
    *,
    error_type: type[Exception] = ValueError,
) -> str:
    """Normalize and format-check one reference, reporting where it came from.

    error_type exists because the same check runs on both sides of the atomic
    save: a bad input file is a ValueError, while a workbook this program just
    wrote failing its own read-back is a RuntimeError.
    """
    reference = normalize_reference(value)
    if not REFERENCE_RE.fullmatch(reference):
        raise error_type(
            f"{location}检索参考号格式无效：{value!r}；{REFERENCE_FORMAT_HINT}"
        )
    return reference
