# LanCrawler

LAN file-index crawler + Excel VBA search tool.

## What’s in this repo

- `src/crawler/` — parallel Python crawler (quick update + live progress GUI; AccDB / Excel import)
- `src/vba/` — VBA modules for the Excel workbook
- `src/inject_vba.py` — inject VBA from disk into the `.xlsm`
- `Lan_Search_Tool.xlsm` — AccDB search workbook (default inject / crawl target)
- `Blank_LAN_Crawler Tool.xlsm` — stable non-AccDB template (leave untouched for AccDB work)
- `pack_portable.cmd` — zip a colleague kit (`dist/Lan_Search_Tool.zip`) with offline wheels

## Index backends (exclusive, no hybrid)

On workbook open: newest `{workbook}\DB\SearchIndex-M-D-YYYY.accdb` (else `LAN_Search_Index.accdb`) → AccDB (Ingestion restored from `tblIngested`); else if onboard `Database!tblFiles` is non-empty → sheet; else empty (Ingestion cleared). AccDB crawls clear the sheet so AccDB stays exclusive. Additional folder crawls **ReplaceRoot-merge** into the same AccDB (no rename). Delete the AccDB (or choose Start fresh) to begin a new snapshot.

## Not in git

Live `.xlsm` / `.xlsx` files (except the templates), `*.accdb` / `DB/`, `crawl_output/`, `venv/`, `vendor/wheels/`, and secrets are ignored.

## Quick start

See `src/crawler/README.md` and `src/vba/README.md`.

**Portable zip (colleague PC):** run `pack_portable.cmd` on a machine that can `pip download` via JFrog, then unzip `dist/Lan_Search_Tool.zip` and double-click `run_crawl.cmd`.
