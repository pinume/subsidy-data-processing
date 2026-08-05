"""Stable surface for 审核明细.xlsx that other modules depend on.

store_report.py needs the summary sheet's name and header to aggregate
审核明细.xlsx; it depends on this module (re-exported by
processors/coupon_report.py) rather than reaching into
processors.coupons.appliance directly, so a change to appliance.py's
internals can't silently break store_report.py.
"""

# 补贴年度。跨年度运行（如进入 2027 年）时改这一处，并按 README
# 「年度迁移清单」逐项核对——外部导出表头与列号、门店空白模板等不
# 从这里派生，仍需人工同步。
SUBSIDY_YEAR = 2026

SUMMARY_SHEET_NAME = "数据汇总"
SUMMARY_SUBSIDY_HEADER = f"{SUBSIDY_YEAR}国补金额"
SUMMARY_HEADER = (
    "财务大类",
    "品牌",
    "上传状态",
    "数量",
    SUMMARY_SUBSIDY_HEADER,
)
