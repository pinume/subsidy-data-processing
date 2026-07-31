"""Stable surface for 审核明细.xlsx that other modules depend on.

store_report.py needs the summary sheet's name and header to aggregate
审核明细.xlsx; it depends on this module (re-exported by
processors/coupon_report.py) rather than reaching into
processors.coupons.appliance directly, so a change to appliance.py's
internals can't silently break store_report.py.
"""

SUMMARY_SHEET_NAME = "数据汇总"
SUMMARY_HEADER = (
    "财务大类",
    "品牌",
    "备注",
    "数量",
    "2026家电国补（计入收入）合计",
)
