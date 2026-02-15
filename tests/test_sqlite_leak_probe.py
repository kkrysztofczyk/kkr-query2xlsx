import sqlite3
import unittest
from unittest import mock


@unittest.skip(
    "TODO: connect leak probe to a real sqlite-creating flow (demo export/self-test) and then unskip."
)
class SQLiteLeakProbeTests(unittest.TestCase):
    def test_no_leaked_sqlite_connections_in_selected_paths(self):
        opened = []
        closed = set()

        real_connect = sqlite3.connect

        def tracking_connect(*args, **kwargs):
            conn = real_connect(*args, **kwargs)
            opened.append(conn)
            real_close = conn.close

            def tracking_close():
                closed.add(id(conn))
                return real_close()

            conn.close = tracking_close  # type: ignore[assignment]
            return conn

        with mock.patch("sqlite3.connect", side_effect=tracking_connect):
            # TODO: wywołaj 1-2 najczęściej używane ścieżki “demo/sqlite”
            # np. uruchom eksport demo, self-test albo inny przepływ, który tworzy sqlite DB.
            # Jeśli nie ma prostego entrypointu do importu, ten test można na razie zostawić z `pass`
            # i aktywować dopiero przy realnym incydencie.
            pass

        leaked = [c for c in opened if id(c) not in closed]
        self.assertEqual(leaked, [], f"Leaked sqlite connections: {len(leaked)}")
