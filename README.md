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

## 更新

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
