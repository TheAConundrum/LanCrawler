# LanCrawler

LAN file-index crawler + Excel VBA search tool.

## What’s in this repo

- `src/crawler/` — parallel Python crawler (CSV / AccDB / Excel import + crawl GUI)
- `src/vba/` — VBA modules for the Excel workbook
- `src/inject_vba.py` — inject VBA from disk into the `.xlsm`
- `Blank_LAN_Crawler Tool.xlsm` — stable working template (leave untouched for AccDB work)
- `AccDB-Blank_LAN_Crawler Tool.xlsm` — Crawlerv2 AccDB workbook (default inject / crawl target)

## Index backends (exclusive, no hybrid)

On workbook open: if `{workbook}\DB\LAN_Search_Index.accdb` exists → AccDB (Ingestion restored from `tblIngested`); else if onboard `Database!tblFiles` is non-empty → sheet; else empty (Ingestion cleared). AccDB crawls clear the sheet so AccDB stays exclusive.

## Not in git

Live `.xlsm` / `.xlsx` files (except the blank templates), `*.accdb` / `DB/`, `crawl_output/`, `venv/`, and secrets are ignored.

## Quick start

See `src/crawler/README.md` and `src/vba/README.md`.
