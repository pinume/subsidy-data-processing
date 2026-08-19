# Subsidy Data Processing

统一处理家电、数码国补 Excel 数据（仅支持 Linux / WSL）。以源码方式运行。

## 快速开始

```bash
git clone https://github.com/pinume/subsidy-data-processing.git
cd subsidy-data-processing
uv sync --locked
uv run python main.py
```

### 运行方式

| 命令 | 说明 |
|------|------|
| `uv run python main.py` | 交互菜单（默认） |
| `uv run python main.py --all` | 按顺序全量处理模式 1–5 |
| `uv run python main.py --mode N` | 运行指定模式（1–5） |

## 目录约定

- **输入目录**：将原始 Excel 文件**直接放入 `data/` 根目录**（不要创建子目录）。首次运行会自动创建 `data/` 目录。
- **输出目录**：所有处理结果生成在 `output/` 目录。
- **排他锁**：单实例运行锁位于 `/tmp/subsidy-data-processing.lock`，防止多实例并发冲突。

## 配置文件 (config/)

- `config/merchants.yaml`：**首次使用必填**。配置家电和数码的商户编号（用于筛选回款数据与定位上传文件）。
- `config/brand_mapping.yaml`：品牌归一化规则（如统帅→海尔、COLMO→美的等）。
- `config/payment_brands.yaml`：品类划分、品牌识别关键词优先级及型号归并规则。

## 门店报表模板（模式 5）

模式 5（门店国补上传及回款情况表）需要 `data/` 目录中放置一份空白正式模板：
- **文件名**：匹配 `2026年门店国补上传及回款情况表（益庄店）.xlsx`。
- **结构与口径**：表 1 明细为第 3–32 行，总计为第 33 行；方太冰箱与方太厨卫统一汇总至第 27 行（厨卫/方太）。
- **版本校验**：正式模板在 **A53** 单元格包含隐藏的版本标记 `模板版本：2026-V5`（第 53 行必须隐藏）。缺少该标记、改动关键表格结构或使用旧版（V3/V4）模板将拒绝生成报表。
