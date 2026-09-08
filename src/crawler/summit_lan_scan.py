"""
SummitLANScan: crawl every mapped network drive on this PC in one run.

Each drive is ReplaceRoot-merged into the workbook AccDB (other drives
already in the index are kept). Re-run anytime — no reset.
"""
from __future__ import annotations

import argparse
import sys
import threading
import time
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import (  # noqa: E402
    accdb_path_for_workbook,
    crawl_parallel,
    load_previous_index,
    purge_workbook_accdbs,
    write_accdb,
)
from crawler.paths import (  # noqa: E402
    build_drive_map,
    bytes_to_size_mb,
    list_mapped_network_drives,
    to_unc_path,
)

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_DEFAULT_WORKBOOK = _PROJECT_ROOT / "Lan_Search_Tool.xlsm"


def run_summit_lan_scan(
    *,
    workers: int = 16,
    workbook: str | Path | None = None,
    progress_every: int = 50,
    accdb_fresh: bool = False,
    full_crawl: bool = False,
    use_gui: bool = True,
) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

    def say(msg: str) -> None:
        if not use_gui:
            print(msg, flush=True)

    wb_path = Path(workbook) if workbook else _DEFAULT_WORKBOOK
    if not wb_path.is_file():
        msg = f"Workbook not found: {wb_path}"
        if use_gui:
            from crawler.crawl_gui import report_fatal

            report_fatal("SummitLANScan", msg)
        else:
            print(msg, flush=True)
        return 1

    drives = list_mapped_network_drives()
    if not drives:
        msg = "No reachable mapped network drives found on this PC."
        if use_gui:
            from crawler.crawl_gui import report_fatal

            report_fatal("SummitLANScan", msg)
        else:
            print(msg, flush=True)
        return 1

    want_fresh = bool(accdb_fresh)
    want_incremental = not full_crawl and not want_fresh
    workers = max(1, min(48, int(workers)))

    if use_gui:
        from crawler.crawl_gui import run_crawl_gui, run_progress_window

        gui = run_crawl_gui(initial_workbook=wb_path, summit_drives=drives)
        if gui.cancelled:
            return 0
        wb_path = Path(gui.workbook)
        workers = gui.workers
        want_fresh = gui.accdb_fresh
        want_incremental = gui.incremental and not want_fresh
        if not wb_path.is_file():
            from crawler.crawl_gui import report_fatal

            report_fatal("SummitLANScan", f"Workbook not found: {wb_path}")
            return 1

    if want_fresh:
        say("Starting fresh: deleting existing AccDB files in DB\\ ...")
        try:
            purge_workbook_accdbs(wb_path)
        except RuntimeError as exc:
            if use_gui:
                from crawler.crawl_gui import report_fatal

                report_fatal("SummitLANScan", str(exc))
            else:
                print(str(exc), flush=True)
            return 1

    accdb_out = accdb_path_for_workbook(wb_path)
    crawl_mode_label = (
        "Start fresh (wipe index)"
        if want_fresh
        else ("Quick update" if want_incremental else "Full recrawl")
    )
    drive_label = ", ".join(letter for letter, _root, _unc in drives)

    say("SummitLANScan — all mapped network drives")
    say(f"Workbook: {wb_path}")
    say(f"AccDB:    {accdb_out}")
    say(f"Workers:  {workers} per drive")
    say(f"Update:   {crawl_mode_label}")
    say(f"Drives:   {len(drives)}")
    for letter, root, unc in drives:
        say(f"  {letter}  {root}  {unc}")
    say("---")

    job: dict[str, object] = {"failed": [], "ok": [], "error": None}

    def run_job(
        cancel_event: threading.Event | None,
        on_stats: object,
        on_log: object,
    ) -> None:
        def emit(msg: str) -> None:
            say(msg)
            if callable(on_log):
                on_log(msg)

        t_all = time.time()
        failed: list[str] = []
        ok: list[str] = []

        for i, (letter, root, unc_hint) in enumerate(drives, start=1):
            if cancel_event is not None and cancel_event.is_set():
                emit("Cancelled. Remaining drives were not crawled.")
                break
            emit(f"[{i}/{len(drives)}] Crawling {letter} ...")
            drive_map = build_drive_map([root])
            unc = to_unc_path(root, drive_map, "") or unc_hint
            emit(f"  UNC: {unc}")

            t0 = time.time()
            previous = None
            use_inc = want_incremental
            if use_inc and accdb_out.is_file():
                try:
                    previous = load_previous_index(accdb_out, unc)
                except Exception as exc:  # noqa: BLE001
                    emit(f"  (previous index skipped: {exc})")
                    previous = None
                use_inc = bool(previous is not None and previous.has_meta)
                if use_inc:
                    emit(f"  Quick update: {len(previous.folders):,} known folders")
                else:
                    emit("  No folder timestamps yet — full listing this drive")
            try:
                rows, stats = crawl_parallel(
                    root,
                    workers=workers,
                    drive_map=drive_map,
                    progress_every=max(1, progress_every),
                    on_progress=emit,
                    on_stats=on_stats if callable(on_stats) else None,
                    incremental=use_inc,
                    previous=previous,
                    cancel_event=cancel_event,
                )
            except Exception as exc:  # noqa: BLE001 — keep remaining drives
                emit(f"  FAILED {letter}: {exc}")
                failed.append(letter)
                continue

            if stats.cancelled or (cancel_event is not None and cancel_event.is_set()):
                emit("Cancelled. Index was not updated for the remaining drives.")
                break

            emit(
                f"  Crawl {letter}: {len(rows):,} rows  "
                f"files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
                f"scan={stats.folders_scanned:,} quick={stats.folders_quick:,} "
                f"errors={stats.errors:,} disk~{bytes_to_size_mb(stats.bytes_all_files):,.2f} MB "
                f"in {time.time() - t0:.1f}s"
            )
            emit(f"  Writing AccDB ReplaceRoot for {unc} ...")
            try:
                write_accdb(rows, accdb_out, crawl_root=unc, replace_root=True)
            except Exception as exc:  # noqa: BLE001
                emit(f"  FAILED AccDB {letter}: {exc}")
                failed.append(letter)
                continue
            ok.append(letter)
            emit(f"  Done {letter}.")

        if ok:
            try:
                from crawler.import_to_excel import clear_onboard_tblfiles

                emit("Clearing onboard Database!tblFiles so AccDB is exclusive...")
                clear_onboard_tblfiles(wb_path)
            except Exception as exc:  # noqa: BLE001
                emit(f"  (sheet clear skipped: {exc})")

        emit("SummitLANScan complete.")
        emit(f"  OK:     {', '.join(ok) if ok else '(none)'}")
        emit(f"  Failed: {', '.join(failed) if failed else '(none)'}")
        emit(f"  AccDB:  {accdb_out}")
        emit(f"  Total:  {time.time() - t_all:.1f}s")
        emit("Re-open the workbook to search.")
        job["failed"] = failed
        job["ok"] = ok
        if failed or not ok:
            job["error"] = "One or more drives failed."

    if use_gui:
        from crawler.crawl_gui import run_progress_window

        ok = run_progress_window(
            root_label=drive_label,
            mode_label=f"{crawl_mode_label}  ·  {workers} workers  ·  {len(drives)} drives",
            work=run_job,
        )
        failed = job.get("failed") or []
        ok_drives = job.get("ok") or []
        if not ok:
            return 1
        return 1 if failed or not ok_drives else 0

    try:
        run_job(None, None, None)
    except Exception as exc:  # noqa: BLE001
        print(f"Index write failed: {exc}", flush=True)
        return 1
    failed = job.get("failed") or []
    ok_drives = job.get("ok") or []
    return 1 if failed or not ok_drives else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="SummitLANScan: crawl all mapped network drives into the LAN Search AccDB."
    )
    parser.add_argument(
        "--workbook",
        default=None,
        help="Target .xlsm (default: Lan_Search_Tool.xlsm next to this project)",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Delete existing AccDB files and write a new snapshot",
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help="Re-list every folder (skip quick update)",
    )
    parser.add_argument(
        "-w",
        "--workers",
        type=int,
        default=16,
        help="Parallel folder workers per drive (default 16)",
    )
    parser.add_argument(
        "--no-gui",
        action="store_true",
        help="Console-only (no setup or progress window)",
    )
    args = parser.parse_args(argv)
    return run_summit_lan_scan(
        workers=args.workers,
        workbook=args.workbook,
        accdb_fresh=args.fresh,
        full_crawl=args.full,
        use_gui=not args.no_gui,
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        try:
            from crawler.crawl_gui import report_fatal

            report_fatal("SummitLANScan", f"{exc}")
        except Exception:
            print(f"CRASH: {exc}", flush=True)
        raise SystemExit(1) from exc
