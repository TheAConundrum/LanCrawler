# LanCrawler

LAN file-index crawler + Excel VBA search tool.

## What’s in this repo

- `src/crawler/` — parallel Python crawler (CSV / SQLite / Excel import)
- `src/vba/` — VBA modules for the Excel workbook
- `src/inject_vba.py` — inject VBA from disk into the `.xlsm`
- `Blank_LAN_Crawler Tool.xlsm` — blank workbook template (default target for VBA inject + crawl import; live workbooks stay gitignored)

## Not in git

Live `.xlsm` / `.xlsx` files (except the blank template), `crawl_output/`, `venv/`, and secrets are ignored.

## Quick start

See `src/crawler/README.md` and `src/vba/README.md`.
