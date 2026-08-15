# Subsidy Data Processing

统一处理家电、数码国补 Excel 数据：已上传数据、收款单、回款明细、审核明细、
门店国补上传及回款情况表。

**仅支持 Linux**（含 WSL）。以源码方式运行，不打包、不安装为全局命令。

## 使用

安装 [Git](https://git-scm.com/) 和 [uv](https://docs.astral.sh/uv/) 后：

```bash
git clone https://github.com/pinume/subsidy-data-processing.git
cd subsidy-data-processing
uv sync --locked
uv run python main.py
```

首次运行会创建 `data` 目录。把原始 Excel **直接放进 `data/`**（不要子目录），
再重新运行并选择处理模式。结果写入 `output/`。

### 启动方式

入口始终是仓库里的 `main.py`，用 `uv run python` 执行：

| 命令 | 说明 |
|------|------|
| `uv run python main.py` | 交互菜单（默认） |
| `uv run python main.py --all` | 按顺序处理全部模式（等同菜单「全部处理」） |
| `uv run python main.py --mode N` | 只处理模式 `N`（编号与菜单一致） |
| `uv run python main.py -h` | 显示帮助 |

规则：

- `--all` 与 `--mode` **不能同时使用**。
- `--mode` 取值为 **1–5**；全部处理请用 `--all`，不要使用 `--mode 6`。
- 非法参数退出码为 `2`。

### 处理模式（菜单 / `--mode` 编号）

| 编号 | 模式 | 主要输出 |
|------|------|----------|
| 1 | 已上传数据（家电+数码） | `output/家电_已上传.xlsx`、`output/数码_已上传.xlsx` |
| 2 | 收款单统计 | `output/收款单统计.xlsx` |
| 3 | 回款明细（家电+数码） | `output/回款明细.xlsx` |
| 4 | 审核明细（销售用券情况统计） | `output/审核明细.xlsx` |
| 5 | 门店国补上传及回款情况表 | `output/2026年门店国补上传及回款情况表（益庄店）.xlsx` |
| （菜单 6 / `--all`） | 全部处理 | 上述全部 |

示例：

```bash
# 交互选择
uv run python main.py

# 只跑已上传数据
uv run python main.py --mode 1

# 全量处理（适合脚本或定时任务里 cd 到项目目录后调用）
uv run python main.py --all
```

### 单实例与中断

- 同一时刻只允许一个实例写入 `output/`。锁文件：
  `/tmp/subsidy-data-processing.lock`。若已有实例在跑，新进程退出码为 `3`。
- `Ctrl+C` 与 `SIGTERM` 均视为取消：未提交的输出会回滚，原有 `output/` 文件保持不变。
  例外：进程刚启动、依赖尚未 import 完的极短窗口内（通常约 1 秒），`SIGTERM`
  仍走系统默认终止；此阶段尚未加锁、未写 `output/`，不会留下半成品。

### 环境变量（可选）

| 变量 | 作用 |
|------|------|
| `UPLOAD_DATA_VERBOSE=1` | 输出更详细的明细与核对信息 |
| `UPLOAD_DATA_DEBUG=1` | 失败时打印 traceback |
| `UPLOAD_DATA_FONT_PATH` | 指定测宽用字体文件路径 |
| `UPLOAD_DATA_FONT_NAME` | 写入工作簿的字体名（配合上一变量） |

## 配置

- `config/merchants.yaml`：家电、数码商户编号。回款明细按编号筛选行，已上传数据按
  编号定位 `MER_<商户编号>_*.xlsx`。**首次使用必须填写。**
- `config/brand_mapping.yaml`：品牌归一化（统帅→海尔、COLMO→美的等）。
- `config/payment_brands.yaml`：回款品类、品牌关键词、美的系归并等。

## 门店报表模板

模式 5（门店国补上传及回款情况表）需要 `data/` 目录里有一份**正式新版空白模板**。
`data/` 不进 Git，模板必须由操作员在每个运行环境手工放入：

- 文件名：`2026年门店国补上传及回款情况表（益庄店）.xlsx`
  （与文件名关键字匹配即可）
- **删除目录里的旧模板再放入新模板**；同目录存在多份含
  「门店国补上传及回款情况表」的 .xlsx 会因无法确定唯一模板而报错

### 版本校验

正式模板在 **A55** 单元格写有版本标记 `模板版本：2026-V3`（表 3 合计行之下，
该行已隐藏，不会出现在报表打印结果中）。程序启动模式 5 时按内容校验：

- 模板没有版本标记（旧模板或手工修改版）→ 明确报错
  「请更换为新版模板（含版本标记的正式模板）」；
- 模板结构被改动（行号、标签与代码契约不一致）→ 同样拒绝，防止
  在错误的模板上静默产出错位报表。

### 关键结构（代码按此校准，勿手工改动）

- 单工作表，名称为「益庄」，8 列（A..H）；
- 表 1 明细与总计（第 2–34 行）：列语义为 A 品类、B 序号、C 品牌、D 发生额、E 上传额、F 上传率（E/D）、G 回款额、H 回款率（G/D）。明细覆盖第 3–33 行（冰洗 3–14、电视 15–19、空调 20–26、厨卫 27–32、3C 数码 33），第 34 行「费用总计」（D/E/G 列金额总计，F/H 列总计比例）；
- 表 2 主要品牌汇总（第 39–48 行）：列语义为 C 品牌、D 发生额、E 回款额、F 回款率（E/D）。第 40 行表头，41–47 行各品牌组（海尔系、美的系、格力、博西、海信系、创维、TCL），第 48 行「合计」；
- 表 3 数量统计（第 50–54 行）：列语义为 C 项目、D 审核中数量、E 未上传数量。第 50–51 行表头，52 家电 / 53 数码 / 54 合计；
- 「方太冰箱」在第 14 行、「厨卫/方太」在第 28 行，是两个独立行。

## 年度迁移清单

补贴年度变更（例如进入 2027 年）时，先修改
`processors/coupons/report_contract.py` 的 `SUBSIDY_YEAR`，再逐项核对：

1. **外部用券导出表头与列号**：`processors/coupons/sources.py` 的
   `COUPON_FAMILY_SUBSIDY_HEADER` / `COUPON_DIGITAL_SUBSIDY_HEADER`（补贴列名）
   与 `COUPON_FAMILY_SUBSIDY_COLUMN` / `COUPON_DIGITAL_SUBSIDY_COLUMN`（列号）。
   这是源系统的导出契约，不随 `SUBSIDY_YEAR` 自动变化，必须人工同步。
2. **审核明细汇总表头**：由 `SUBSIDY_YEAR` 派生（`SUMMARY_SUBSIDY_HEADER`），
   确认生成结果无误即可。
3. **门店报表**：`data` 目录里的空白模板文件名与结构；`processors/store_report.py` 的
   结构校验常量（`TEMPLATE_STRUCTURE_CELLS` 中的年度表头标签）、输出文件名由
   `SUBSIDY_YEAR` 派生；报表标题由模板原样保留，需随模板人工更新（代码仅更新
   「更新时间：」之后的时间戳）。
4. **金额错误文案**：家电/数码明细的补贴金额校验文案引用补贴列名常量，随列名
   自动更新，无需单独修改。
5. **测试夹具与预期表头**：`tests/` 中硬编码的年度字段与表头。
6. **完整回归**：用真实的新年度文件完整运行一次全部模式（菜单 `6` 或
   `uv run python main.py --all`），核对各输出与门店报表。

## 更新

```bash
cd subsidy-data-processing
git pull --ff-only
uv sync --locked
uv run python main.py
```
