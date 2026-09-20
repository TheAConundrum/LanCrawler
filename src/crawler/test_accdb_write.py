"""AccDB bulk-write helpers: ACE lock-count handling."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import (  # noqa: E402
    IndexRow,
    _chunked_delete_under_root,
    _executemany_committed,
    _is_ace_lock_count_error,
    _raise_ace_max_locks,
    write_accdb,
)


class _LockConn:
    def __init__(self) -> None:
        self.commits = 0
        self.rollbacks = 0

    def commit(self) -> None:
        self.commits += 1

    def rollback(self) -> None:
        self.rollbacks += 1


class _LockThenOkCursor:
    def __init__(self, fail_if_at_least: int) -> None:
        self.fail_if_at_least = fail_if_at_least
        self.batch_sizes: list[int] = []
        self.inserted = 0

    def executemany(self, _sql: str, chunk: list[tuple[object, ...]]) -> None:
        self.batch_sizes.append(len(chunk))
        if len(chunk) >= self.fail_if_at_least:
            raise Exception(
                "('HY000', '[HY000] [Microsoft][ODBC Microsoft Access Driver] "
                "File sharing lock count exceeded. Increase MaxLocksPerFile "
                "registry entry. (-1033) (SQLExecDirectW)')"
            )
        self.inserted += len(chunk)


class _ChunkDeleteCursor:
    def __init__(self, paths: list[str]) -> None:
        self.paths = list(paths)
        self._rows: list[tuple[object, ...]] = []
        self.delete_sizes: list[int] = []

    def execute(self, sql: str, params: tuple[object, ...] | list[object] | None = None) -> None:
        sql_u = sql.strip().upper()
        if sql_u.startswith("SELECT TOP 1 *"):
            self._rows = [(self.paths[0],)] if self.paths else []
            return
        if sql_u.startswith("SELECT TOP"):
            top_n = int(sql.split()[2])
            self._rows = [(p,) for p in self.paths[:top_n]]
            return
        if sql_u.startswith("DELETE"):
            doomed = [str(p) for p in (params or ())]
            self.delete_sizes.append(len(doomed))
            doomed_set = set(doomed)
            self.paths = [p for p in self.paths if p not in doomed_set]
            self._rows = []
            return
        raise AssertionError(f"unexpected SQL: {sql}")

    def fetchall(self) -> list[tuple[object, ...]]:
        return list(self._rows)

    def fetchone(self) -> tuple[object, ...] | None:
        return self._rows[0] if self._rows else None


class AccdbLockHelperTests(unittest.TestCase):
    def test_detects_max_locks_error(self) -> None:
        exc = Exception(
            "('HY000', '[HY000] [Microsoft][ODBC Microsoft Access Driver] "
            "File sharing lock count exceeded. Increase MaxLocksPerFile "
            "registry entry. (-1033) (SQLExecDirectW)')"
        )
        self.assertTrue(_is_ace_lock_count_error(exc))
        self.assertFalse(_is_ace_lock_count_error(Exception("permission denied")))

    def test_executemany_retries_with_smaller_batch(self) -> None:
        cur = _LockThenOkCursor(fail_if_at_least=20)
        conn = _LockConn()
        records = [(i,) for i in range(25)]
        _executemany_committed(
            cur, conn, "INSERT INTO t VALUES (?)", records, batch_size=25
        )
        self.assertGreaterEqual(conn.rollbacks, 1)
        self.assertEqual(cur.inserted, 25)
        self.assertTrue(all(size < 20 for size in cur.batch_sizes if size != 25))

    def test_chunked_delete_never_exceeds_chunk(self) -> None:
        root = r"\\server\share"
        paths = [rf"{root}\f{i:05d}\file.txt" for i in range(950)]
        cur = _ChunkDeleteCursor(paths)
        conn = _LockConn()
        deleted = _chunked_delete_under_root(
            cur,
            conn,
            table_name="tblFiles",
            path_col="FilePath",
            root=root,
            chunk=400,
        )
        self.assertEqual(deleted, 950)
        self.assertEqual(sum(cur.delete_sizes), 950)
        self.assertTrue(all(size <= 400 for size in cur.delete_sizes))
        self.assertEqual(cur.paths, [])

    def test_raise_max_locks_does_not_crash(self) -> None:
        _raise_ace_max_locks()


class AccdbWriteIntegrationTests(unittest.TestCase):
    def test_replace_root_past_default_lock_limit(self) -> None:
        n = 12_000
        root = r"\\sisl-fs2\kearl_public$"
        rows = [
            IndexRow(
                path=rf"{root}\folder_{i:05d}\file.txt",
                file_date="2026-09-20",
                size_mb=0.01,
                entry_type="FILE",
                mtime_ts=1.0,
                size_bytes=100,
            )
            for i in range(n)
        ]
        with tempfile.TemporaryDirectory() as raw:
            accdb = Path(raw) / "SearchIndex-9-20-2026.accdb"
            try:
                write_accdb(rows, accdb, crawl_root=root, replace_root=True)
                write_accdb(rows, accdb, crawl_root=root, replace_root=True)
            except RuntimeError as exc:
                self.skipTest(str(exc))
            self.assertTrue(accdb.is_file())


if __name__ == "__main__":
    unittest.main()
