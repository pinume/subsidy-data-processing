# Subsidy Data Processing

统一处理家电、数码国补 Excel 数据，支持已上传数据、收款单、审核明细、回款明细，
以及门店国补上传及回款情况表。

## 快速开始

安装 Git 和 [uv](https://docs.astral.sh/uv/) 后运行：

```powershell
git clone https://github.com/pinume/subsidy-data-processing.git
Set-Location subsidy-data-processing
uv sync --locked
uv run python main.py
```

程序首次运行会创建 `data` 目录。把原始文件直接放入该目录，不要放进子目录，
然后重新运行并按菜单选择处理模式。结果统一写入 `output` 目录。

## 配置

- `config/merchants.yaml`：家电、数码各一份商户编号。回款明细按它筛选源数据
  行，已上传数据按它定位 `MER_<商户编号>_*.xlsx` 文件。**首次使用必须填写。**
- `config/brand_mapping.yaml`：品牌归一化映射（统帅→海尔、COLMO→美的等）。
- `config/payment_brands.yaml`：回款明细的品类映射、品牌关键词、美的系归并等，
  改品牌配置无需动代码。

## 年度迁移清单

补贴年度变更（例如进入 2027 年）时，先修改
`processors/coupons/report_contract.py` 的 `SUBSIDY_YEAR`，再逐项核对：

1. **外部用券导出表头与列号**：`processors/coupons/sources.py` 的
   `COUPON_FAMILY_SUBSIDY_HEADER` / `COUPON_DIGITAL_SUBSIDY_HEADER`（补贴列名）
   与 `COUPON_FAMILY_SUBSIDY_COLUMN` / `COUPON_DIGITAL_SUBSIDY_COLUMN`（列号）。
   这是源系统的导出契约，不随 `SUBSIDY_YEAR` 自动变化，必须人工同步。
2. **审核明细汇总表头**：由 `SUBSIDY_YEAR` 派生（`SUMMARY_SUBSIDY_HEADER`），
   确认生成结果无误即可。
3. **门店报表**：`data` 目录里的空白模板文件名与结构（表 3 的"26年国补上传额"
   等标签写在模板内，不在代码里）；`processors/store_report.py` 的输出文件名与
   报表标题由 `SUBSIDY_YEAR` 派生，确认无误即可。
4. **金额错误文案**：家电/数码明细的补贴金额校验文案引用补贴列名常量，随列名
   自动更新，无需单独修改。
5. **测试夹具与预期表头**：`tests/` 中硬编码的年度字段与表头。
6. **完整回归**：用真实的新年度文件完整运行一次全部模式（菜单 `6` 或 `all`），
   核对各输出与门店报表。

## 更新

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
