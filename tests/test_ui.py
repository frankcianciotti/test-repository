"""Sample GUI integration tests. Needs a desktop or a temporary test display."""

from datetime import datetime, timezone
from pathlib import Path
import tempfile
import time
import tkinter as tk
from unittest.mock import patch
import unittest
from zipfile import ZipFile

from belsito_mail.outlook_reader import OutlookReadResult
from belsito_mail.sample_data import sample_messages
from belsito_mail.state import LocalState
from belsito_mail.ui import EmailAssistant


class DesktopTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("No desktop display; run on Windows or a test display.")
        self.root.withdraw()
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name)
        self.state = LocalState(self.path / "state.sqlite3")
        self.app = EmailAssistant(self.root, self.state)

    def tearDown(self):
        if hasattr(self, "app"):
            for identifier in self.root.tk.call("after", "info"):
                self.root.after_cancel(identifier)
            self.state.close()
            self.root.destroy()
            self.directory.cleanup()

    def wait_review(self):
        deadline = time.monotonic() + 5
        while self.app.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.busy, "Sample review did not finish")

    def review_sample(self):
        self.app._review()
        self.wait_review()
        self.assertEqual(len(self.app.items), 11)

    def test_default_is_sample_and_sample_review_never_calls_outlook(self):
        self.assertEqual(self.app.mode.get(), "sample")
        self.assertEqual(self.app.review_button.cget("text"), "Review My Emails")
        with patch("belsito_mail.outlook_reader.read_outlook_messages") as reader:
            self.review_sample()
            reader.assert_not_called()
        self.assertEqual(len(self.app.table.get_children()), 11)
        self.assertIn("SAMPLE", self.app.status.get())
        self.assertIn("disabled", self.app.stage_button.state())

    def test_completion_hides_item_and_reopen_restores_it(self):
        self.review_sample()
        identifier = self.app.items[0].id
        self.app.table.selection_set(identifier)
        self.app._set_completed(True)
        self.assertIn(identifier, self.state.completed_ids("sample"))
        self.assertNotIn(identifier, self.app.table.get_children())
        self.assertEqual(self.state.completed_ids("outlook"), set())
        self.app.show_completed.set(True)
        self.app._refresh()
        self.app.table.selection_set(identifier)
        self.app._set_completed(False)
        self.assertNotIn(identifier, self.state.completed_ids("sample"))
        self.assertIn(identifier, self.app.table.get_children())

    def test_all_sample_functions_and_acknowledgement_required_to_enable_outlook(self):
        self.review_sample()
        self.app.table.selection_set(self.app.items[0].id)
        self.app._set_completed(True)
        self.assertIn("disabled", self.app.stage_button.state())
        report = self.path / "sample.xlsx"
        with patch("belsito_mail.ui.filedialog.asksaveasfilename", return_value=str(report)):
            self.app._export()
        self.assertTrue(report.is_file())
        self.assertNotIn("disabled", self.app.stage_button.state())
        self.assertFalse(self.state.sample_test_passed)
        with patch("belsito_mail.ui.messagebox.askyesno", return_value=False):
            self.app._acknowledge_sample()
        self.assertFalse(self.state.sample_test_passed)
        with patch("belsito_mail.ui.messagebox.askyesno", return_value=True), patch("belsito_mail.outlook_reader.read_outlook_messages") as reader:
            self.app._acknowledge_sample()
            reader.assert_not_called()
        self.assertTrue(self.state.sample_test_passed)
        self.assertEqual(self.app.mode.get(), "sample")

    def test_live_mode_blocked_before_sample_and_cancelled_read_never_calls_com(self):
        with patch("belsito_mail.ui.messagebox.showinfo"), patch("belsito_mail.outlook_reader.read_outlook_messages") as reader:
            self.app.mode.set("outlook")
            self.app._review()
            self.assertEqual(self.app.mode.get(), "sample")
            reader.assert_not_called()
        self.state.acknowledge_sample_test()
        self.app.mode.set("outlook")
        with patch("belsito_mail.ui.messagebox.askyesno", return_value=False), patch("belsito_mail.outlook_reader.read_outlook_messages") as reader:
            self.app._review()
            reader.assert_not_called()
        self.assertFalse(self.app.busy)

    def test_explicit_live_consent_passes_expected_scope_and_warns_in_export(self):
        # Fake only; this test never imports Windows COM or opens Outlook.
        now = datetime.now(timezone.utc)
        fake_result = OutlookReadResult(sample_messages(now), 16, ("Test incomplete folder coverage.",))
        self.state.acknowledge_sample_test()
        self.app.mode.set("outlook")
        with patch("belsito_mail.ui.messagebox.askyesno", return_value=True) as consent, patch("belsito_mail.ui.messagebox.showwarning") as warning, patch("belsito_mail.outlook_reader.read_outlook_messages", return_value=fake_result) as reader:
            self.app._review()
            self.wait_review()
            self.app._review()
            self.wait_review()
            self.assertEqual(consent.call_count, 2)
            self.assertEqual(reader.call_count, 2)
            self.assertEqual(warning.call_count, 2)
            kwargs = reader.call_args.kwargs
            self.assertTrue(kwargs["authorization"].user_approved)
            self.assertTrue(kwargs["authorization"].sample_test_passed)
            self.assertTrue(kwargs["include_subfolders"])
        report = self.path / "incomplete.xlsx"
        with patch("belsito_mail.ui.filedialog.asksaveasfilename", return_value=str(report)):
            self.app._export()
        with ZipFile(report) as workbook:
            xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn("INCOMPLETE COVERAGE", xml)
        self.assertIn("Test incomplete folder coverage", xml)


if __name__ == "__main__":
    unittest.main()
