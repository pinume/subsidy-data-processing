# Upload Data Processing

## Quick Start

Install Git and [uv](https://docs.astral.sh/uv/), then run:

```powershell
git clone https://github.com/pinume/Upload-data-processing.git
Set-Location Upload-data-processing
uv sync --locked
uv run python main.py
```

## Update

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
