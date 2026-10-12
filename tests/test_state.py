from pathlib import Path
import tempfile
import unittest

from belsito_mail.state import LocalState


class LocalStateTests(unittest.TestCase):
    def test_completion_survives_restart_is_separate_and_can_be_reopened(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.sqlite3"
            state = LocalState(path)
            state.set_completed("sample", "opaque-123", True)
            state.set_completed("sample", "opaque-123", True)
            self.assertEqual(state.completed_ids("sample"), {"opaque-123"})
            self.assertEqual(state.completed_ids("outlook"), set())
            state.close()
            state = LocalState(path)
            self.assertEqual(state.completed_ids("sample"), {"opaque-123"})
            state.set_completed("sample", "opaque-123", False)
            self.assertEqual(state.completed_ids("sample"), set())
            state.close()

    def test_sample_gate_defaults_to_false_and_requires_acknowledgement(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.sqlite3"
            state = LocalState(path)
            self.assertFalse(state.sample_test_passed)
            state.acknowledge_sample_test()
            state.close()
            state = LocalState(path)
            self.assertTrue(state.sample_test_passed)
            self.assertEqual(state.completed_ids("outlook"), set())
            state.close()

    def test_schema_stores_no_mail_content_and_ids_are_parameterized(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.sqlite3"
            state = LocalState(path)
            malicious = "'; DROP TABLE completed; --"
            state.set_completed("sample", malicious, True)
            self.assertEqual(state.completed_ids("sample"), {malicious})
            columns = [row[1] for row in state.connection.execute("PRAGMA table_info(completed)")]
            self.assertEqual(columns, ["source", "item_id", "completed_at"])
            state.close()


if __name__ == "__main__":
    unittest.main()
