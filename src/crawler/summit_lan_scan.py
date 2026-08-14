"""
SummitLANScan: crawl every mapped network drive on this PC in one run.

No folder picker. Each drive is ReplaceRoot-merged into the workbook AccDB
(other drives already in the index are kept). Re-run anytime — no reset.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import (  # noqa: E402
    accdb_path_for_workbook,
    crawl_parallel,
    write_accdb,
    write_csv,
)
from crawler.paths import (  # noqa: E402
    build_drive_map,
    bytes_to_size_mb,
    list_mapped_network_drives,
    to_unc_path,
)

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_DEFAULT_WORKBOOK = _PROJECT_ROOT / "AccDB-Blank_LAN_Crawler Tool.xlsm"


def _default_out_dir() -> Path:
    return _PROJECT_ROOT / "crawl_output"


def run_summit_lan_scan(
    *,
    workers: int = 16,
    workbook: str | Path | None = None,
    progress_every: int = 50,
) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

    print("SummitLANScan — all mapped network drives", flush=True)

    wb_path = Path(workbook) if workbook else _DEFAULT_WORKBOOK
    if not wb_path.is_file():
        print(f"Workbook not found: {wb_path}", flush=True)
        return 1

    drives = list_mapped_network_drives()
    if not drives:
        print("No reachable mapped network drives found on this PC.", flush=True)
        return 1

    accdb_out = accdb_path_for_workbook(wb_path)
    out_dir = _default_out_dir()
    stamp = time.strftime("%Y%m%d_%H%M%S")

    print(f"Workbook: {wb_path}", flush=True)
    print(f"AccDB:    {accdb_out}", flush=True)
    print(f"Workers:  {workers} per drive", flush=True)
    print(f"Drives:   {len(drives)}", flush=True)
    for letter, root, unc in drives:
        print(f"  {letter}  {root}  {unc}", flush=True)
    print("---", flush=True)

    def on_progress(msg: str) -> None:
        print(msg, flush=True)

    t_all = time.time()
    failed: list[str] = []
    ok: list[str] = []

    for i, (letter, root, unc_hint) in enumerate(drives, start=1):
        print(f"[{i}/{len(drives)}] Crawling {letter} ...", flush=True)
        drive_map = build_drive_map([root])
        unc = to_unc_path(root, drive_map, "") or unc_hint
        print(f"  UNC: {unc}", flush=True)

        t0 = time.time()
        try:
            rows, stats = crawl_parallel(
                root,
                workers=workers,
                drive_map=drive_map,
                progress_every=max(1, progress_every),
                on_progress=on_progress,
            )
        except Exception as exc:  # noqa: BLE001 — keep remaining drives
            print(f"  FAILED {letter}: {exc}", flush=True)
            failed.append(letter)
            continue

        csv_path = out_dir / f"{stamp}_{letter[0]}_tblFiles.csv"
        write_csv(rows, csv_path)
        print(
            f"  CSV {csv_path.name}: {len(rows):,} rows  "
            f"files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
            f"errors={stats.errors:,} disk~{bytes_to_size_mb(stats.bytes_all_files):,.2f} MB "
            f"in {time.time() - t0:.1f}s",
            flush=True,
        )
        print(f"  Writing AccDB ReplaceRoot for {unc} ...", flush=True)
        try:
            write_accdb(rows, accdb_out, crawl_root=unc, replace_root=True)
        except Exception as exc:  # noqa: BLE001
            print(f"  FAILED AccDB {letter}: {exc}", flush=True)
            failed.append(letter)
            continue
        ok.append(letter)
        print(f"  Done {letter}.", flush=True)
        print("---", flush=True)

    if ok:
        try:
            from crawler.import_to_excel import clear_onboard_tblfiles

            print("Clearing onboard Database!tblFiles so AccDB is exclusive...", flush=True)
            clear_onboard_tblfiles(wb_path)
        except Exception as exc:  # noqa: BLE001
            print(f"  (sheet clear skipped: {exc})", flush=True)

    print("SummitLANScan complete.", flush=True)
    print(f"  OK:     {', '.join(ok) if ok else '(none)'}", flush=True)
    print(f"  Failed: {', '.join(failed) if failed else '(none)'}", flush=True)
    print(f"  AccDB:  {accdb_out}", flush=True)
    print(f"  Total:  {time.time() - t_all:.1f}s", flush=True)
    print("Re-open the workbook to search.", flush=True)
    return 1 if failed or not ok else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="SummitLANScan: crawl all mapped network drives into the LAN Search AccDB."
    )
    parser.add_argument(
        "--workbook",
        default=None,
        help="Target .xlsm (default: AccDB-Blank_LAN_Crawler Tool.xlsm next to this project)",
    )
    parser.add_argument(
        "-w",
        "--workers",
        type=int,
        default=16,
        help="Parallel folder workers per drive (default 16)",
    )
    args = parser.parse_args(argv)
    return run_summit_lan_scan(workers=args.workers, workbook=args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
