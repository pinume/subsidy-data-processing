# Subsidy Data Processing

## Quick Start

Install Git and [uv](https://docs.astral.sh/uv/), then run:

```powershell
git clone https://github.com/pinume/subsidy-data-processing.git
Set-Location subsidy-data-processing
uv sync --locked
uv run python main.py
```

## 商户编号配置

两类补贴的编号体系不同，同一门店在两张表里的编号也不同，需分别填写
`config/merchants.yaml`：

```yaml
merchants:
  家电: 89813015722APT1
  数码: 89813014812B06R
```

两条流水线共用这一份配置，换门店只改这里：

* 回款明细：按编号筛选源数据行
* 已上传数据：源文件名形如 `MER_<商户编号>_<导出时间>_yjhx.xlsx`，程序用
  `MER_` + 编号定位文件，前缀是导出方的命名规则，不单独配置

## 品类与品牌配置（仅回款明细使用）

回款明细的编码品类映射、品牌关键词、品牌归并规则、型号别名都在
`config/payment_brands.yaml`，不写在代码里；补充新品牌或型号别名只改这个文件。
文件内注释说明了各部分的含义，其中品牌关键词列表的顺序即匹配优先级、品类映射
的顺序决定汇总和明细里财务大类的排序，改动时需保留顺序。

## 回款明细（菜单第 4 项）

原始数据放在 `data` 目录下，文件名需包含 `补贴明细`；数据类型（家电 / 数码）
自动识别，先按文件名关键词（`以旧换新补贴明细` / `数码补贴明细`），识别不出时
读取明细表头按字段判定。支持 `.xlsx`、`.xls`、`.xlsm`、`.csv`、`.tsv`。

按商户编号筛选数据行，编号写在 `config/merchants.yaml`（见下）。

结果写入 `output\回款明细.xlsx`：

| Sheet  | 内容                                     |
| ------ | -------------------------------------- |
| `汇总`   | 家电、数码各一个区块，中间空一行分隔，块末为加粗 `合计`，最末行为总合计 |
| `家电明细` | 家电整合明细                                 |
| `数码明细` | 数码整合明细                                 |

明细分表是因为两类列结构不同（家电有 `补贴比例`、`能耗等级`，数码有 `IMEI1码`、
`IMEI2码`、`备注`），合表会产生大量空列。明细排序为：财务大类 → 品牌 → 交易时间
→ 商品名称；财务大类为洗衣机或冰箱时，品牌美的、小天鹅、东芝统一归并为美的系。

说明：

* 原始数据不会被覆盖；`.xlsx` 源文件直接只读读取，不生成工作副本，`.xls`/
  `.xlsm`/`.csv`/`.tsv` 才会生成工作副本，处理完自动删除
* 程序启动时会清理 `output` 目录里超过 3 分钟的残留临时文件（上一次运行被
  中断或崩溃留下的）；不要同时运行两份本程序处理同一个 `output` 目录
* 公式单元格读取 Excel 已保存的计算结果；补贴金额是公式但没有缓存结果时，程序
  会停止并提示先用 Excel/WPS 打开保存，或将公式转换为数值
* 输出含商户编号、SN/IMEI、发票号、金额等敏感数据，对账完成后请按组织要求处理

## Update

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
