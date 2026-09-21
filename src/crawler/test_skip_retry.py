"""Classify folder errors and retry network skips, not no-access."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import (  # noqa: E402
    SKIP_KIND_MISSING,
    SKIP_KIND_NETWORK,
    SKIP_KIND_NO_ACCESS,
    SKIP_KIND_OTHER,
    SkipRecord,
    classify_folder_error,
    crawl_parallel,
    visit_folder,
    write_skip_log,
)


def _winerr(code: int, msg: str) -> OSError:
    exc = OSError(code, msg)
    exc.winerror = code
    return exc


class ClassifyFolderErrorTests(unittest.TestCase):
    def test_access_denied(self) -> None:
        kind, _msg, code = classify_folder_error(_winerr(5, "Access is denied"))
        self.assertEqual(kind, SKIP_KIND_NO_ACCESS)
        self.assertEqual(code, 5)

    def test_network_access_denied(self) -> None:
        kind, _msg, code = classify_folder_error(
            _winerr(65, "Network access is denied")
        )
        self.assertEqual(kind, SKIP_KIND_NO_ACCESS)
        self.assertEqual(code, 65)

    def test_unexpected_network(self) -> None:
        kind, _msg, code = classify_folder_error(
            _winerr(59, "An unexpected network error occurred")
        )
        self.assertEqual(kind, SKIP_KIND_NETWORK)
        self.assertEqual(code, 59)

    def test_location_unreachable(self) -> None:
        kind, _msg, code = classify_folder_error(
            _winerr(1232, "The network location cannot be reached")
        )
        self.assertEqual(kind, SKIP_KIND_NETWORK)
        self.assertEqual(code, 1232)

    def test_missing_path(self) -> None:
        kind, _msg, code = classify_folder_error(_winerr(3, "The system cannot find the path"))
        self.assertEqual(kind, SKIP_KIND_MISSING)
        self.assertEqual(code, 3)

    def test_parses_winerror_from_message(self) -> None:
        kind, _msg, code = classify_folder_error(
            Exception("[WinError 1232] The network location cannot be reached")
        )
        self.assertEqual(kind, SKIP_KIND_NETWORK)
        self.assertEqual(code, 1232)

    def test_sharing_violation_is_retryable_other(self) -> None:
        kind, _msg, _code = classify_folder_error(_winerr(32, "The process cannot access"))
        self.assertEqual(kind, SKIP_KIND_OTHER)
        rec = SkipRecord(path=r"K:\x", kind=kind, error="x")
        self.assertTrue(rec.retryable)
        denied = SkipRecord(path=r"K:\y", kind=SKIP_KIND_NO_ACCESS, error="x")
        self.assertFalse(denied.retryable)


class SkipRetryCrawlTests(unittest.TestCase):
    def test_retries_network_and_leaves_no_access(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            ok = root / "ok"
            flaky = root / "flaky"
            denied = root / "denied"
            ok.mkdir()
            flaky.mkdir()
            denied.mkdir()
            (ok / "a.txt").write_text("a", encoding="utf-8")
            (flaky / "b.txt").write_text("b", encoding="utf-8")
            (denied / "c.txt").write_text("c", encoding="utf-8")

            flaky_tries = {"n": 0}
            real_visit = visit_folder

            def wrapped(path: str, *args, **kwargs):
                name = Path(path).name.lower()
                if name == "flaky":
                    flaky_tries["n"] += 1
                    if flaky_tries["n"] == 1:
                        res = real_visit(path, *args, **kwargs)
                        res.error = "[WinError 59] An unexpected network error occurred"
                        res.error_kind = SKIP_KIND_NETWORK
                        res.winerror = 59
                        res.file_rows = []
                        res.subfolders = []
                        return res
                if name == "denied":
                    res = real_visit(path, *args, **kwargs)
                    res.error = "[WinError 5] Access is denied"
                    res.error_kind = SKIP_KIND_NO_ACCESS
                    res.winerror = 5
                    res.file_rows = []
                    res.subfolders = []
                    return res
                return real_visit(path, *args, **kwargs)

            skip_log = root / "crawl_skips.txt"
            with patch("crawler.crawl.visit_folder", wrapped):
                rows, stats = crawl_parallel(
                    str(root),
                    workers=2,
                    progress_every=1,
                    skip_log_path=skip_log,
                    retry_rounds=2,
                    retry_delay_sec=0.0,
                )

            names = {Path(r.path).name.lower() for r in rows if r.entry_type == "FILE"}
            self.assertIn("a.txt", names)
            self.assertIn("b.txt", names)
            self.assertNotIn("c.txt", names)
            self.assertGreaterEqual(flaky_tries["n"], 2)
            self.assertEqual(stats.skip_no_access, 1)
            self.assertEqual(stats.skip_network, 0)
            self.assertGreaterEqual(stats.skip_recovered, 1)
            self.assertTrue(skip_log.is_file())
            log_text = skip_log.read_text(encoding="utf-8")
            self.assertIn("no_access", log_text)
            self.assertIn("recovered", log_text)
            self.assertIn("denied", log_text.lower())

    def test_write_skip_log_columns(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "skips.txt"
            write_skip_log(
                path,
                root=r"K:\ops",
                records=[
                    SkipRecord(
                        path=r"K:\ops\secret",
                        kind=SKIP_KIND_NO_ACCESS,
                        error="[WinError 5] Access is denied",
                    ),
                    SkipRecord(
                        path=r"K:\ops\camp",
                        kind=SKIP_KIND_NETWORK,
                        error="[WinError 59] net",
                        attempts=3,
                        recovered=True,
                    ),
                ],
            )
            text = path.read_text(encoding="utf-8")
            self.assertIn("status\tkind\tattempts\tpath\terror", text)
            self.assertIn("skipped\tno_access\t1\t", text)
            self.assertIn("recovered\tnetwork\t3\t", text)


if __name__ == "__main__":
    unittest.main()
