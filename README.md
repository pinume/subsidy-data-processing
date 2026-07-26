# Upload Data Processing

A Python tool for processing household appliance and digital business Excel data.

Supported workflows:

- submitted data processing
- receipt statistics
- subsidy coupon statistics

The program reads source files only and saves generated results atomically to
a flat `output/` directory (no subfolders):

```text
output/家电_已上传.xlsx
output/数码_已上传.xlsx
output/收款单统计.xlsx
output/审核明细.xlsx
```

`收款单统计.xlsx` and `审核明细.xlsx` are shared between both projects:
`收款单统计.xlsx` is produced with the household-appliance rules (including
same-model-replacement detection and the special remarks in
`config/receipt_special_remarks.yaml`) regardless of which project you run
it from, and `审核明细.xlsx` always contains both projects' processed data
(a `家电-明细总表` sheet, a `数码-明细总表` sheet, and a combined `数据汇总`
sheet), built by reading both projects' source files at once.

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

If the `data` directory does not exist, the program creates it (empty) and
ends the current run. All input files go directly into `data/` — there are
no subfolders, and large appliances and digital data live side by side so
both projects can be processed without swapping files in and out.

Each project finds its own files by filename keyword (and, where two
projects' files share a keyword, by header content):

| Data | How it's identified | Location |
| --- | --- | --- |
| Submitted data (large appliances) | Filename contains `MER_89813015722APT1` | `data/*.xlsx` |
| Submitted data (digital) | Filename contains `MER_89813014812B06R` | `data/*.xlsx` |
| Receipt statistics (shared by both projects) | Filename contains `收款单统计` | `data/*.XLS` |
| Subsidy coupon statistics | Filename contains `销售用券情况统计`; the project is determined by the last column header (`2026家电国补（计入收入）` for large appliances, `2026数码国补（计入收入）` for digital) | `data/*.XLS` |
| Reference number supplement (optional, household appliances only) | Filename contains `新建 Microsoft Excel 工作表`; columns are `参考号`/`单据号`/`单据日期` | `data/*.xlsx` |

Do not place temporary Excel files starting with `~$` in the input directory.
If more than one file matches a keyword where only one file is expected (or
more than one file matches the same coupon header), the program stops with
an error rather than silently picking one.

## Configuration

Brand normalization rules are stored in:

```text
config/brand_mapping.yaml
```

Fixed `退换货\倒票` special remarks (matched by date + document number) are
stored in:

```text
config/receipt_special_remarks.yaml
```

Update either file without modifying the source code.

## Subsidy Rules

The `补贴金额` column is computed as a percentage of `交易金额`, capped per
order. The rate is the same for both projects, but **the caps differ**:

| Project | Rate | Cap | Defined in |
| --- | --- | --- | --- |
| Household appliances | 15% | 1500 | `processors/large_appliances/submitted.py` |
| Digital | 15% | 500 | `processors/digital.py` |

Both are named `SUBSIDY_RATE` / `SUBSIDY_CAP` in those modules and are pinned
by tests in `tests/test_submitted_validation.py`. They were once both written
as 500, which silently understated 43% of the appliance rows — change them
only when the subsidy policy itself changes.

## Update

```powershell
git pull --ff-only
uv sync --locked
uv run python main.py
```
