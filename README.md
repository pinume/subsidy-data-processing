# Upload Data Processing

## Quick Start

Install Git and [uv](https://docs.astral.sh/uv/), then run:

```powershell
git clone https://github.com/pinume/Upload-data-processing.git
Set-Location Upload-data-processing
uv sync --locked
uv run python main.py
```

`uv` automatically manages the required Python environment. The minimum supported Python version is 3.12.


## Update

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
