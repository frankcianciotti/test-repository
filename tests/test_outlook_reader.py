"""Mailbox safety tests use fake COM objects and never access a real mailbox."""

from datetime import datetime, timedelta, timezone
import sys
import types
import unittest
from unittest import mock

from belsito_mail.outlook_reader import (
    OutlookReadError,
    ReadAuthorization,
    _read_from_namespace,
    _read_with_com,
    read_outlook_messages,
)


NOW = datetime(2026, 10, 12, 15, 30, tzinfo=timezone.utc)
FORBIDDEN = {
    "Delete", "Move", "Send", "Save", "Reply", "ReplyAll", "Forward",
    "Display", "Sort", "Logon", "Logoff", "Quit", "SendAndReceive",
    "Add", "Remove", "Copy", "Dispatch",
}


class ReadOnlyFake:
    """Reject every property assignment and every unapproved COM operation."""

    def __init__(self, name, log, **members):
        object.__setattr__(self, "_name", name)
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_members", members)

    def __setattr__(self, name, value):
        raise AssertionError(f"COM property assignment attempted: {name}")

    def __getattr__(self, name):
        if name in FORBIDDEN or name not in self._members:
            raise AssertionError(f"Unapproved COM operation attempted: {name}")
        value = self._members[name]
        if callable(value):
            def invoke(*args):
                self._log.append((self._name, name, args))
                return value(*args)
            return invoke
        self._log.append((self._name, name, None))
        if isinstance(value, Exception):
            raise value
        return value


def collection(name, log, elements, *, restrict=False):
    members = {"Count": len(elements), "Item": lambda index: elements[index - 1]}
    if restrict:
        selected = collection(name + "/selected", log, elements)
        members["Restrict"] = lambda query: selected
    return ReadOnlyFake(name, log, **members)


def mail(log, entry_id="message-1", **changes):
    members = {
        "Class": 43,
        "EntryID": entry_id,
        "Subject": "Please repair a leaking pipe",
        "SenderName": "Sample customer",
        "SenderEmailAddress": "customer@example.test",
        "ReceivedTime": NOW - timedelta(days=1),
        "Body": "Please send a plumber.",
    }
    members.update(changes)
    return ReadOnlyFake(entry_id, log, **members)


def folder(log, name="Inbox", *, messages=None, children=None, **changes):
    members = {
        "StoreID": "sample-store",
        "Items": collection(name + "/items", log, messages or [], restrict=True),
        "Folders": collection(name + "/folders", log, children or []),
    }
    members.update(changes)
    return ReadOnlyFake(name, log, **members)


def namespace(log, inbox):
    def get_default_folder(number):
        if number != 6:
            raise AssertionError("Only the default Inbox is allowed")
        return inbox
    return ReadOnlyFake("namespace", log, GetDefaultFolder=get_default_folder)


class OutlookReaderTests(unittest.TestCase):
    def read(self, inbox, log, **options):
        return _read_from_namespace(namespace(log, inbox), now=NOW, **options)

    def test_authorization_is_required_before_any_com_import(self):
        cases = [None, ReadAuthorization(False, True), ReadAuthorization(True, False)]
        for authorization in cases:
            with self.subTest(authorization=authorization):
                with mock.patch("belsito_mail.outlook_reader.sys.platform", "win32"):
                    with mock.patch("builtins.__import__") as import_spy:
                        with self.assertRaises(OutlookReadError):
                            read_outlook_messages(authorization=authorization, now=NOW)
                        self.assertEqual(import_spy.call_count, 0)

    def test_non_boolean_authorization_does_not_grant_access(self):
        with self.assertRaises(OutlookReadError):
            read_outlook_messages(authorization=ReadAuthorization(1, 1), now=NOW)

    def test_linux_has_clear_error_without_windows_imports(self):
        with mock.patch("belsito_mail.outlook_reader.sys.platform", "linux"):
            with mock.patch("builtins.__import__") as import_spy:
                with self.assertRaisesRegex(OutlookReadError, "requires Windows"):
                    read_outlook_messages(authorization=ReadAuthorization(True, True), now=NOW)
                self.assertEqual(import_spy.call_count, 0)

    def test_default_inbox_and_children_are_read_without_mutation(self):
        log = []
        child = folder(log, "child", messages=[mail(log, "child-message")])
        inbox = folder(log, messages=[mail(log)], children=[child])
        result = self.read(inbox, log)
        self.assertEqual(len(result.messages), 2)
        self.assertEqual(result.scanned_count, 2)
        self.assertFalse(result.warnings)
        operations = {entry[1] for entry in log}
        self.assertTrue(operations.isdisjoint(FORBIDDEN))
        self.assertEqual(
            operations,
            {"GetDefaultFolder", "StoreID", "Items", "Restrict", "Count", "Item",
             "Class", "ReceivedTime", "EntryID", "Subject", "SenderName",
             "SenderEmailAddress", "Body", "Folders"},
        )
        message = next(message for message in result.messages if message.subject)
        self.assertEqual(message.source, "outlook")
        self.assertTrue(message.id.startswith("outlook:"))
        self.assertNotIn("sample-store", message.id)
        self.assertNotIn("message-1", message.id)

    def test_subfolders_can_be_excluded_without_reading_their_collection(self):
        log = []
        child = folder(log, "child", messages=[mail(log, "child-message")])
        result = self.read(folder(log, messages=[mail(log)], children=[child]), log,
                           include_subfolders=False)
        self.assertEqual(len(result.messages), 1)
        self.assertNotIn("Folders", {entry[1] for entry in log})

    def test_dasl_query_uses_utc_and_independent_exact_date_bounds(self):
        log = []
        inbox = folder(log, messages=[
            mail(log, "old", ReceivedTime=NOW - timedelta(days=30, microseconds=1)),
            mail(log, "boundary", ReceivedTime=NOW - timedelta(days=30)),
            mail(log, "current", ReceivedTime=NOW),
            mail(log, "future", ReceivedTime=NOW + timedelta(microseconds=1)),
        ])
        result = self.read(inbox, log)
        self.assertEqual(len(result.messages), 2)
        self.assertEqual(result.scanned_count, 4)
        query = next(args[0] for _, name, args in log if name == "Restrict")
        self.assertEqual(query, (
            '@SQL="http://schemas.microsoft.com/mapi/proptag/0x0E060040" '
            ">= '2026-09-12 15:30:00' AND "
            '"http://schemas.microsoft.com/mapi/proptag/0x0E060040" '
            "<= '2026-10-12 15:30:00'"
        ))
        self.assertFalse(any(name == "Body" and tag in {"old", "future"}
                             for tag, name, _ in log))

    def test_query_converts_an_aware_non_utc_review_time(self):
        log = []
        review_time = NOW.astimezone(timezone(timedelta(hours=-4)))
        _read_from_namespace(namespace(log, folder(log)), now=review_time)
        query = next(args[0] for _, name, args in log if name == "Restrict")
        self.assertIn("2026-10-12 15:30:00", query)

    def test_partial_second_query_does_not_omit_eligible_messages(self):
        log = []
        now = NOW + timedelta(microseconds=900_000)
        inbox = folder(log, messages=[mail(log, ReceivedTime=NOW + timedelta(microseconds=800_000))])
        result = _read_from_namespace(namespace(log, inbox), now=now)
        self.assertEqual(len(result.messages), 1)
        query = next(args[0] for _, name, args in log if name == "Restrict")
        self.assertIn("2026-10-12 15:30:01", query)

    def test_non_mail_items_are_skipped_before_reading_their_content(self):
        log = []
        meeting = ReadOnlyFake("meeting", log, Class=26)
        result = self.read(folder(log, messages=[meeting]), log)
        self.assertEqual(result.messages, [])
        self.assertEqual(result.scanned_count, 1)
        self.assertEqual([name for tag, name, _ in log if tag == "meeting"], ["Class"])

    def test_timezone_aware_dates_are_retained_and_naive_dates_are_localized(self):
        log = []
        aware = (NOW - timedelta(days=1)).astimezone(timezone(timedelta(hours=-4)))
        naive = (NOW - timedelta(days=2)).astimezone().replace(tzinfo=None)
        result = self.read(folder(log, messages=[
            mail(log, "aware", ReceivedTime=aware), mail(log, "naive", ReceivedTime=naive),
        ]), log)
        self.assertEqual(len(result.messages), 2)
        self.assertEqual(result.messages[0].received_at, aware)
        self.assertEqual(result.messages[0].received_at.tzinfo, aware.tzinfo)
        self.assertIsNotNone(result.messages[1].received_at.tzinfo)
        self.assertEqual(result.messages[1].received_at, naive.astimezone())

    def test_ids_are_stable_and_store_specific_duplicates_are_removed(self):
        log = []
        message = mail(log)
        inbox = folder(log, messages=[message, message])
        first = self.read(inbox, log)
        second = self.read(inbox, log)
        third = self.read(folder(log, messages=[message], StoreID="different-store"), log)
        self.assertEqual(len(first.messages), 1)
        self.assertEqual(first.messages[0].id, second.messages[0].id)
        self.assertNotEqual(first.messages[0].id, third.messages[0].id)

    def test_long_bodies_have_explicit_incomplete_classification_warning(self):
        log = []
        result = self.read(folder(log, messages=[mail(log, Body="x" * 100_001)]), log)
        self.assertEqual(len(result.messages[0].body), 100_000)
        self.assertTrue(any("classification may be incomplete" in warning for warning in result.warnings))

    def test_item_limit_is_global_and_incomplete_is_explicit(self):
        log = []
        child = folder(log, "child", messages=[mail(log, "child-1"), mail(log, "child-2")])
        inbox = folder(log, messages=[mail(log)], children=[child])
        result = self.read(inbox, log, maximum_items=2)
        self.assertEqual(result.scanned_count, 2)
        self.assertEqual(len(result.messages), 2)
        self.assertTrue(any("report is incomplete" in warning for warning in result.warnings))
        self.assertFalse(any(tag == "child-2" for tag, _, _ in log))

    def test_exact_item_limit_without_more_messages_is_complete(self):
        log = []
        result = self.read(folder(log, messages=[mail(log)]), log, maximum_items=1)
        self.assertFalse(result.warnings)

    def test_folder_failure_does_not_prevent_reading_other_folders(self):
        log = []
        secret_error = RuntimeError("sensitive mailbox contents should never appear")
        blocked = folder(log, "blocked", Items=secret_error)
        good = folder(log, "good", messages=[mail(log)])
        result = self.read(folder(log, children=[blocked, good]), log)
        self.assertEqual(len(result.messages), 1)
        self.assertTrue(any("report is incomplete" in warning for warning in result.warnings))
        self.assertNotIn("sensitive", str(result.warnings))

    def test_restrict_failure_has_no_unfiltered_fallback(self):
        log = []
        def fail(query):
            raise RuntimeError("private COM diagnostic")
        items = ReadOnlyFake("broken-items", log, Restrict=fail)
        result = self.read(folder(log, Items=items), log)
        self.assertEqual(result.messages, [])
        self.assertTrue(result.warnings)
        self.assertNotIn("private", str(result.warnings))
        self.assertNotIn("Item", {name for _, name, _ in log})

    def test_one_unreadable_message_does_not_drop_other_messages(self):
        log = []
        result = self.read(folder(log, messages=[
            mail(log, "bad", Body=RuntimeError("private content")), mail(log, "good"),
        ]), log)
        self.assertEqual(len(result.messages), 1)
        self.assertEqual(result.scanned_count, 2)
        self.assertTrue(result.warnings)
        self.assertNotIn("private", str(result.warnings))

    def test_invalid_dates_and_missing_ids_are_not_silently_accepted(self):
        log = []
        result = self.read(folder(log, messages=[
            mail(log, "bad-date", ReceivedTime="not a date"), mail(log, ""),
            mail(log, "null-id", EntryID=None),
        ]), log)
        self.assertFalse(result.messages)
        self.assertTrue(result.warnings)

    def test_folder_safety_limit_bounds_discovered_folders(self):
        log = []
        children = [folder(log, f"child-{index}") for index in range(10)]
        with mock.patch("belsito_mail.outlook_reader._MAX_FOLDERS", 3):
            result = self.read(folder(log, children=children), log)
        self.assertTrue(any("folder safety limit" in warning for warning in result.warnings))
        visited = {tag for tag, name, _ in log if name == "StoreID"}
        self.assertEqual(visited, {"Inbox", "child-0", "child-1"})

    def test_worker_uses_only_active_instance_and_always_releases_com(self):
        log = []
        ns = namespace(log, folder(log, messages=[mail(log)]))
        application = ReadOnlyFake("application", log, GetNamespace=lambda name: ns)
        lifecycle = []
        def get_active_object(name):
            lifecycle.append(("active", name))
            return application
        result = _read_with_com(
            get_active_object=get_active_object,
            co_initialize=lambda: lifecycle.append("initialize"),
            co_uninitialize=lambda: lifecycle.append("uninitialize"),
            now=NOW, days=30, include_subfolders=True, maximum_items=5_000,
        )
        self.assertEqual(len(result.messages), 1)
        self.assertEqual(lifecycle, ["initialize", ("active", "Outlook.Application"), "uninitialize"])
        self.assertIn(("application", "GetNamespace", ("MAPI",)), log)

    def test_connection_failure_releases_com_and_has_safe_error(self):
        cleanup = mock.Mock()
        with self.assertRaisesRegex(OutlookReadError, "Open Outlook Classic") as caught:
            _read_with_com(
                get_active_object=mock.Mock(side_effect=RuntimeError("private account")),
                co_initialize=mock.Mock(), co_uninitialize=cleanup,
                now=NOW, days=30, include_subfolders=True, maximum_items=5_000,
            )
        cleanup.assert_called_once_with()
        self.assertNotIn("private", str(caught.exception))
        self.assertTrue(caught.exception.__suppress_context__)

    def test_inbox_failure_releases_com(self):
        cleanup = mock.Mock()
        log = []
        ns = ReadOnlyFake("namespace", log, GetDefaultFolder=mock.Mock(side_effect=RuntimeError("blocked")))
        app = ReadOnlyFake("application", log, GetNamespace=lambda name: ns)
        with self.assertRaisesRegex(OutlookReadError, "default Inbox"):
            _read_with_com(
                get_active_object=lambda name: app,
                co_initialize=mock.Mock(), co_uninitialize=cleanup,
                now=NOW, days=30, include_subfolders=True, maximum_items=5_000,
            )
        cleanup.assert_called_once_with()

    def test_failed_initialization_does_not_uninitialize_uninitialized_com(self):
        cleanup = mock.Mock()
        active = mock.Mock()
        with self.assertRaisesRegex(OutlookReadError, "could not be initialized"):
            _read_with_com(
                get_active_object=active,
                co_initialize=mock.Mock(side_effect=RuntimeError("COM unavailable")),
                co_uninitialize=cleanup,
                now=NOW, days=30, include_subfolders=True, maximum_items=5_000,
            )
        active.assert_not_called()
        cleanup.assert_not_called()

    def test_public_reader_lazy_imports_windows_components_only_after_consent(self):
        log = []
        ns = namespace(log, folder(log, messages=[mail(log)]))
        app = ReadOnlyFake("application", log, GetNamespace=lambda name: ns)
        pythoncom = types.ModuleType("pythoncom")
        pythoncom.CoInitialize = mock.Mock()
        pythoncom.CoUninitialize = mock.Mock()
        client = types.ModuleType("win32com.client")
        client.GetActiveObject = mock.Mock(return_value=app)
        win32com = types.ModuleType("win32com")
        win32com.client = client
        with mock.patch.dict(sys.modules, {"pythoncom": pythoncom, "win32com": win32com,
                                           "win32com.client": client}):
            with mock.patch("belsito_mail.outlook_reader.sys.platform", "win32"):
                result = read_outlook_messages(authorization=ReadAuthorization(True, True), now=NOW)
        self.assertEqual(len(result.messages), 1)
        client.GetActiveObject.assert_called_once_with("Outlook.Application")
        pythoncom.CoInitialize.assert_called_once_with()
        pythoncom.CoUninitialize.assert_called_once_with()

    def test_missing_windows_dependency_has_safe_installation_guidance(self):
        original_import = __import__
        def safe_import(name, *args, **kwargs):
            if name == "pythoncom":
                raise ImportError("private machine path")
            return original_import(name, *args, **kwargs)
        with mock.patch("belsito_mail.outlook_reader.sys.platform", "win32"):
            with mock.patch("builtins.__import__", side_effect=safe_import):
                with self.assertRaisesRegex(OutlookReadError, "installation steps for pywin32") as caught:
                    read_outlook_messages(authorization=ReadAuthorization(True, True), now=NOW)
        self.assertNotIn("private machine path", str(caught.exception))

    def test_options_are_validated_before_com_import(self):
        cases = [{"now": NOW.replace(tzinfo=None)}, {"days": 0}, {"days": True},
                 {"maximum_items": 0}, {"include_subfolders": "yes"}, {"days": 10**20}]
        for invalid in cases:
            options = {"now": NOW, **invalid}
            with self.subTest(options=options):
                with mock.patch("builtins.__import__") as import_spy:
                    with self.assertRaises(OutlookReadError):
                        read_outlook_messages(authorization=ReadAuthorization(True, True), **options)
                    self.assertEqual(import_spy.call_count, 0)


if __name__ == "__main__":
    unittest.main()
