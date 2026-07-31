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

## 商户编号

在 `config/merchants.yaml` 中分别填写家电和数码商户编号：

```yaml
merchants:
  家电: 你的家电商户编号
  数码: 你的数码商户编号
```

商户编号同时用于定位已上传数据文件和筛选回款明细。如果来源文件存在但没有匹配
到配置的商户编号，程序会停止并保留原输出，不会生成空报表。

## 输入与输出

| 菜单 | 主要输入文件特征 | 输出文件 |
| --- | --- | --- |
| 已上传数据（家电+数码） | `MER_<商户编号>_*.xlsx` | `家电_已上传.xlsx`、`数码_已上传.xlsx` |
| 收款单统计 | 文件名包含 `收款单统计` 的 `.xlsx` | `收款单统计.xlsx` |
| 审核明细 | 家电、数码合并在同一个 `销售用券情况统计*.xlsx` 内（按行区分：每行仅 2026家电国补或2026数码国补其中一列有值） | `审核明细.xlsx` |
| 回款明细（家电+数码） | 文件名包含 `补贴明细` 的 `.xlsx` | `回款明细.xlsx` |
| 门店国补上传及回款情况表 | `data` 目录中文件名包含“门店国补上传及回款情况表”的空白模板 | `2026年门店国补上传及回款情况表（益庄店）.xlsx` |

回款明细只接受 `.xlsx`，只读打开。程序会根据文件名或表头自动区分家电和数码，
并按照 `config/payment_brands.yaml` 进行品类映射和品牌识别。

门店国补上传及回款情况表不读取原始数据，而是读取本程序自己已生成的
`output/审核明细.xlsx` 与 `output/回款明细.xlsx`，按品牌和品类聚合后填入空白模
板；因此需要先运行“审核明细”和“回款明细”两个模式。模板本身不会被修改。

## 字体

程序会依次查找 Maple Mono NF CN、微软雅黑、Noto Sans CJK SC 和苹方。找不到时，
可通过环境变量指定字体：

```powershell
$env:UPLOAD_DATA_FONT_PATH = "C:\Windows\Fonts\msyh.ttc"
$env:UPLOAD_DATA_FONT_NAME = "微软雅黑"
uv run python main.py
```

## 安全与失败处理

- 原始文件不会被覆盖。
- 单个输出先写入临时文件并验证，成功后才替换正式文件。
- 组合任务后续步骤失败时，已经更新的输出会回滚到运行前版本。
- 输出可能包含商户编号、SN/IMEI、发票号和金额，请妥善保管。

## 更新

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
