# LAN Search Tool — VBA import guide

## Preferred: VS Code task (silent inject)

1. Trust access to the VBA project object model enabled in Excel  
   (File → Options → Trust Center → Trust Center Settings → Macro Settings)
2. Run task **Import VBA code to Excel**  
   - Closes `AccDB-Blank_LAN_Crawler Tool.xlsm` if open  
   - Injects everything from `src/vba/`  
   - Removes old `Module1`–`Module4` if present  
   - Leaves Excel open with the updated workbook  

   Companion task: **Export from Excel to Disk** (VBA → `src/vba`)  

Or from a terminal:

```text
python src\inject_vba.py "AccDB-Blank_LAN_Crawler Tool.xlsm"
```

Keep `Blank_LAN_Crawler Tool.xlsm` as the stable non-AccDB template. AccDB work uses `AccDB-Blank_LAN_Crawler Tool.xlsm`.

Manual import below is only a fallback.

## Sheet rename map

| Old | New |
|-----|-----|
| HomePage | Dashboard |
| Feed The Beast | Ingestion |
| Data | Database |

Run `RunSetupWorkbook` once after import to rename sheets (if still old names), apply labels, and set dropdown validation.

## Modules to import

| File | Type |
|------|------|
| `modPathUtil.bas` | Standard |
| `modEntryPoints.bas` | Standard |
| `modIndexCache.bas` | Standard |
| `clsSheetMap.cls` | Class |
| `clsExcelAppState.cls` | Class |
| `clsFileIndexer.cls` | Class |
| `clsFileSearch.cls` | Class |

Remove old `Module1`–`Module4` (and `Macro9`) after import.

## Button rebinds (only these two are public)

| Old macro | New macro |
|-----------|-----------|
| `ListFiles` | `RunIndexFromIngestion` |
| `GetSearchResults` | `RunSearchFromDashboard` |

Class methods and `modPathUtil` helpers are hidden from the Assign Macro list (`Option Private Module`).  
If you still see `ListFiles` / `GetSearchResults` / many other names, **delete the old Module1–Module4** from the VBA project.

## Exclusive index backend (no hybrid)

On open / warm, `modIndexCache.ResolveIndexBackend` picks **one** backend for the session:

1. If onboard `Database!tblFiles` is non-empty → **Sheet**
2. Else if `{workbook}\DB\LAN_Search_Index.accdb` exists → **AccDB** (late-bound ADODB / ACE)
3. Else → empty (search prompts to crawl)

AccDB crawls clear onboard `tblFiles` so leftover sheet rows cannot shadow AccDB. Search does not merge backends.

## Cell layout

### Dashboard (user owns rows 1–6 chrome)

| Row | Purpose |
|-----|---------|
| 1 | Instructions (yours) |
| 2 | Criteria headers |
| 3–5 | Criteria inputs |
| 3–6 | Options / stats chrome (right side; yours) |
| 6 | Divider |
| 7 | Results headers |
| 8+ | Results data (macro-drawn; not an Excel ListObject) |

Criteria (`A2:F5`): Order | Key Word | Operator | Key Word | Size | Size (MB)  
Flexible match: label **`H3`**, value **`H4`**  
Search folders: label **`I3:J3`**, value **`I4`** = `Yes` (default) / `No` / `Only`  
- **Yes** — files + folders  
- **No** — files only  
- **Only** — folders only  

Drive filter: value **`H5`** — dropdown from the active backend’s distinct UNC shares. Default **`All`**.

Search stats: macro writes **`M3`** file count, **`M4`** folder count, **`M5`** total Size MB.

Results: **A:K** File (blue; **double-click** opens folder) | **L** File Type | **M** Size MB | **N** Date Created | **O** Drive Address | **P:T** Folder Path | hidden **U** UNC target.

**UI lock:** free-text editable — Key Word **B**/**D** and Size MB **F** (rows 3–5). Dropdowns — Operator **C**, Size op **E**, Flexible **H4**, Search folders **I4**, Drive **H5**. **Database** is sheet-protected and **`xlSheetVeryHidden`**.

**VBA project password:** set once as admin in VBE. Python does **not** set or bypass it.

### Ingestion

- **No path text box** — `RunIndexFromIngestion` opens a folder picker
- `B2` — UncRoot override
- `B4` — Mode: ReplaceRoot / Append / RebuildAll
- **Summary** **E2:E5** — Scans, Files, Folders, Total Size MB
- `tblIngested` (from **A6**): RootPath | DateIngested | FileCount | FolderCount | TotalSizeMB

### Database

- `tblFiles`: **FilePath** | **FileDate** | **SizeMB** | **EntryType** (`FILE` / `FOLDER`)
- Used when the session backend is Sheet. Empty when AccDB is the exclusive index.

### AccDB (beside workbook)

- Path: `{workbook folder}\DB\LAN_Search_Index.accdb`
- `tblFiles`: same schema as sheet; indexes on `FilePath` and `EntryType`
- `tblIngested`: one row per crawl root (RootPath | DateIngested | FileCount | FolderCount | TotalSizeMB); accumulates across AccDB ReplaceRoot crawls
- On AccDB resolve / cache warm, VBA replaces **Ingestion!tblIngested** + **E2:E5** with AccDB’s scan list only
- If AccDB is deleted, next warm/search clears Ingestion history
- Requires ACE on each PC that searches AccDB

## Index filter (docs + images + parent archives)

Crawl indexes only these extensions (see `modPathUtil.ALLOWED_INDEX_EXTS`):

- Documents: `pdf doc docx xls xlsx xlsm xlsb ppt pptx txt rtf csv odt ods odp msg eml`
- Images: `jpg jpeg png gif tif tiff bmp webp`
- Archives (parent only): `zip 7z rar tar gz tgz bz2 xz cab`
- Video: `mp4 mov avi mkv wmv webm m4v mpg mpeg flv 3gp ts mts m2ts`

Also skips junk folders/files and split archive volumes.

## Parallel Python crawler (recommended for large LAN roots)

See [`src/crawler/README.md`](../crawler/README.md). Example:

```bat
cd /d "D:\path\to\LAN Search Tool\src"
python -m crawler "O:\Some Folder" -w 16 --import-excel "..\AccDB-Blank_LAN_Crawler Tool.xlsm" --target Auto
```

## Search behavior

- Keywords match **FilePath** (`LIKE` / `InStr`), including folder segments — not name-only.
- Operators: AND / OR / NOT; optional size `>` / `<`; drive share filter.
- Flexible match normalizes terms (and sheet paths) before compare; AccDB applies a post-filter when Flexible is on.
- Folder hit descendant counts are not precomputed (display may show `0` files/folders).

## Search performance

`Workbook_Open` schedules deferred `ResolveIndexBackend` (~2s). Sheet backend loads a slim `tblFiles` cache; AccDB opens on demand via late-bound ADODB. Results write into the prebuilt 20000-row grid. Immediate Window shows `SEARCH` / `INDEX` diagnostics.

## Notes

- Re-ingest the same root to refresh counts/sizes (ReplaceRoot) for sheet mode.
- Size filter is optional; leave Size / Size (MB) blank to ignore size.
- Flexible match strips ` .,;-_` etc.; keeps `\` and `/`.
- On ingest failure: read the MsgBox **Stage** / **Error #**, and open VBA Immediate Window (**Ctrl+G**) for the full log.
