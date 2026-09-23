"""Spawn-safe AccDB shard worker (Windows multiprocessing re-imports this module)."""
from __future__ import annotations

from typing import Any


def run_shard(payload: dict[str, Any], progress_q: Any) -> None:
    """Load pickled rows into a throwaway AccDB. Must stay importable as crawler.accdb_shard."""
    com_inited = False
    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        com_inited = True
    except Exception:  # noqa: BLE001
        com_inited = False

    shard_index = int(payload.get("shard_index", 0))
    worker_total = int(payload.get("worker_total", 1))

    def note(msg: str) -> None:
        try:
            progress_q.put(("log", shard_index, msg))
        except Exception:  # noqa: BLE001
            pass

    try:
        from crawler.crawl import _run_accdb_shard_job

        n = _run_accdb_shard_job(payload, note)
        progress_q.put(("ok", shard_index, n))
    except Exception as exc:  # noqa: BLE001
        progress_q.put(
            (
                "err",
                shard_index,
                f"worker {shard_index + 1}/{worker_total}: {type(exc).__name__}: {exc}",
            )
        )
    finally:
        if com_inited and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
