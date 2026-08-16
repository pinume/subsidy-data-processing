"""Stable surface for 审核明细.xlsx that other modules depend on.

store_report.py needs the summary sheet's name and header to aggregate
审核明细.xlsx; it depends on this module (re-exported by
processors/coupon_report.py) rather than reaching into
processors.coupons.appliance directly, so a change to appliance.py's
internals can't silently break store_report.py.
"""

# 补贴年度。跨年度运行（如进入 2027 年）时改这一处，并逐项核对：
# 1. 外部用券导出表头与列号：processors/coupons/sources.py 的
#    COUPON_FAMILY_SUBSIDY_HEADER / COUPON_DIGITAL_SUBSIDY_HEADER 与
#    COUPON_FAMILY_SUBSIDY_COLUMN / COUPON_DIGITAL_SUBSIDY_COLUMN。
# 2. 审核明细汇总表头：由 SUBSIDY_YEAR 派生（SUMMARY_SUBSIDY_HEADER），核对生成表头。
# 3. 门店报表：data/ 目录中的空白模板文件名、processors/store_report.py 的
#    结构校验常量（TEMPLATE_STRUCTURE_CELLS 年度标签）、输出文件名由 SUBSIDY_YEAR 派生。
# 4. 金额错误文案：家电/数码明细的补贴金额校验文案引用补贴列名常量，随列名自动更新。
# 5. 测试夹具与预期表头：tests/ 中硬编码的年度字段与表头断言。
# 6. 完整回归：用真实新年度文件运行 uv run python main.py --all，核对各输出与门店报表。
SUBSIDY_YEAR = 2026

SUMMARY_SHEET_NAME = "数据汇总"
SUMMARY_SUBSIDY_HEADER = f"{SUBSIDY_YEAR}国补金额"
SUMMARY_PAYMENT_COUNT_HEADER = "回款数量"
SUMMARY_PAYMENT_AMOUNT_HEADER = "回款金额"
SUMMARY_CORE_HEADER = (
    "财务大类",
    "品牌",
    "上传状态",
    "数量",
    SUMMARY_SUBSIDY_HEADER,
)
SUMMARY_HEADER = (
    *SUMMARY_CORE_HEADER,
    "退回",
    SUMMARY_PAYMENT_COUNT_HEADER,
    SUMMARY_PAYMENT_AMOUNT_HEADER,
)
