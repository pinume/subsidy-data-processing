# Upload Data Processing

A Python tool for processing household appliance and digital business Excel data.

Supported workflows:

- submitted data processing
- receipt statistics
- subsidy coupon statistics

The program reads source files only and saves generated results atomically to:

```text
output/large_appliances/
output/digital/
```

## Quick Start

Install Git and [uv](https://docs.astral.sh/uv/), then run:

```powershell
git clone https://github.com/pinume/Upload-data-processing.git
Set-Location Upload-data-processing
uv sync --locked
uv run python main.py
```

`uv` automatically manages the required Python environment. The minimum supported Python version is 3.12.

## Input Directory

Create the following structure before the first run:

```text
data/
├── reference_number_supplement/
│   └── reference_number_supplement.xlsx
├── receipt_statistics/
│   └── receipt_statistics.XLS
├── subsidy_coupon_statistics/
│   └── subsidy_coupon_statistics.XLS
└── submitted/
```

Main input files:

| Data | Location |
| --- | --- |
| Submitted data | `data/submitted/*.xlsx` |
| Receipt statistics | `data/receipt_statistics/receipt_statistics.XLS` |
| Subsidy coupon statistics | `data/subsidy_coupon_statistics/subsidy_coupon_statistics.XLS` |
| Reference number supplement (optional, household appliances only) | `data/reference_number_supplement/reference_number_supplement.xlsx` |

Do not place temporary Excel files starting with `~$` in the input folders.

## Configuration

Brand normalization rules are stored in:

```text
config/brand_mapping.yaml
```

Update brand mappings in this file without modifying the source code.

## Update

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
