"""Coupon (销售用券情况统计) processing, shared source layer plus two profiles.

processors/coupons/sources.py locates and reads the merged source file and
classifies rows into 家电/数码; processors/coupons/matching.py holds the
matching logic identical in shape between the two projects; appliance.py and
digital.py each build their own report (家电 is a real superset — reference
supplement file, category/brand group sheets, a six-column summary — so it
keeps its own module rather than being forced through 数码's shape or vice
versa). processors/coupons/report_contract.py is the stable surface other
modules (store_report.py) depend on instead of reaching into appliance.py's
internals.
"""
