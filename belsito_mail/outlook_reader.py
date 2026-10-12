"""Read messages from the default Outlook Classic Inbox without changing Outlook.

This adapter uses only an already-running Outlook Classic instance. It reads the
default profile's Inbox and, optionally, that Inbox's descendants. It does not
enumerate other accounts, other PST files, sent mail, or archived folders.
Email content stays in memory; this module has no network or persistence code.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import hashlib
import sys
from typing import Any, Callable

from .models import EmailMessage


_INBOX = 6
_MAIL_ITEM_CLASS = 43
_MAX_BODY_CHARACTERS = 100_000
_MAX_FOLDERS = 1_000
_RECEIVED_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"


class OutlookReadError(RuntimeError):
    """A safe, user-facing error that contains no mailbox or COM error details."""


@dataclass(frozen=True)
class ReadAuthorization:
    sample_test_passed: bool
    user_approved: bool


@dataclass(frozen=True)
class OutlookReadResult:
    messages: list[EmailMessage]
    scanned_count: int
    warnings: tuple[str, ...]


def _validate_options(
    *, now: datetime, days: int, include_subfolders: bool, maximum_items: int
) -> None:
    if not isinstance(now, datetime) or now.tzinfo is None or now.utcoffset() is None:
        raise OutlookReadError("The review time must include a time zone.")
    if type(days) is not int or days < 1:
        raise OutlookReadError("The review period must be a positive number of days.")
    if type(maximum_items) is not int or maximum_items < 1:
        raise OutlookReadError("The message limit must be a positive whole number.")
    if type(include_subfolders) is not bool:
        raise OutlookReadError("The subfolder setting must be true or false.")
    try:
        now - timedelta(days=days)
    except (OverflowError, ValueError):
        raise OutlookReadError("The review period is too large.") from None


def _restriction(now: datetime, days: int) -> str:
    """DASL date comparisons use UTC, independent of Outlook's display locale."""
    end = now.astimezone(timezone.utc)
    start = end - timedelta(days=days)
    # Restrict accepts whole seconds. Include the current partial second here;
    # the Python datetime check below enforces the exact bounds.
    upper = end.replace(microsecond=0)
    if end.microsecond:
        upper += timedelta(seconds=1)
    date_format = "%Y-%m-%d %H:%M:%S"
    return (
        f'@SQL="{_RECEIVED_PROPERTY}" >= \'{start.strftime(date_format)}\' '
        f'AND "{_RECEIVED_PROPERTY}" <= \'{upper.strftime(date_format)}\''
    )


def _received_datetime(value: Any) -> datetime:
    if not isinstance(value, datetime):
        raise ValueError("Missing received time")
    if value.tzinfo is None or value.utcoffset() is None:
        # astimezone() interprets a naive value in Windows' configured local
        # time zone, including the offset for that value's date.
        return value.astimezone()
    return value


def _read_from_namespace(
    namespace: Any,
    *,
    now: datetime,
    days: int = 30,
    include_subfolders: bool = True,
    maximum_items: int = 5_000,
) -> OutlookReadResult:
    """Internal reader; namespace injection permits safe fake-COM tests on Linux."""
    _validate_options(
        now=now, days=days, include_subfolders=include_subfolders,
        maximum_items=maximum_items,
    )
    try:
        inbox = namespace.GetDefaultFolder(_INBOX)
    except Exception:
        raise OutlookReadError(
            "Outlook's default Inbox could not be read. Check that Outlook Classic "
            "is open with the intended mailbox."
        ) from None

    end = now.astimezone(timezone.utc)
    start = end - timedelta(days=days)
    query = _restriction(now, days)
    messages: list[EmailMessage] = []
    seen_ids: set[str] = set()
    warnings: list[str] = []
    warning_set: set[str] = set()
    scanned_count = 0
    folder_count = 0
    pending = [(inbox, "Inbox")]

    def warn(message: str) -> None:
        if message not in warning_set:
            warning_set.add(message)
            warnings.append(message)

    def cap_warning() -> None:
        warn(
            f"The review reached its {maximum_items:,}-item safety limit. "
            "The report is incomplete; increase the limit or review fewer folders."
        )

    while pending:
        folder, label = pending.pop()
        folder_count += 1
        if folder_count > _MAX_FOLDERS:
            warn("The folder safety limit was reached; the report is incomplete.")
            break

        try:
            store_id = folder.StoreID
            if not isinstance(store_id, str) or not store_id:
                raise ValueError("Missing store identity")
            selected = folder.Items.Restrict(query)
            selected_count = int(selected.Count)
            if selected_count < 0:
                raise ValueError("Invalid item count")
        except Exception:
            warn(
                f"Messages in {label} could not be read or filtered by date. "
                "The report is incomplete."
            )
            selected_count = 0
            selected = None

        available = maximum_items - scanned_count
        if selected_count > available:
            cap_warning()
        for index in range(1, min(selected_count, available) + 1):
            scanned_count += 1
            try:
                item = selected.Item(index)
                if item.Class != _MAIL_ITEM_CLASS:
                    continue
                received_at = _received_datetime(item.ReceivedTime)
                if not start <= received_at.astimezone(timezone.utc) <= end:
                    continue
                entry_id = item.EntryID
                if not isinstance(entry_id, str) or not entry_id:
                    raise ValueError("Missing message identity")
                message_id = "outlook:" + hashlib.sha256(
                    (store_id + "\0" + entry_id).encode("utf-8")
                ).hexdigest()
                if message_id in seen_ids:
                    continue
                subject = str(item.Subject or "")
                sender_name = str(item.SenderName or "")
                sender_email = str(item.SenderEmailAddress or "")
                body = str(item.Body or "")
                if len(body) > _MAX_BODY_CHARACTERS:
                    body = body[:_MAX_BODY_CHARACTERS]
                    warn(
                        "One or more email bodies were shortened to 100,000 "
                        "characters; classification may be incomplete."
                    )
                messages.append(EmailMessage(
                    id=message_id,
                    subject=subject,
                    sender_name=sender_name,
                    sender_email=sender_email,
                    received_at=received_at,
                    body=body,
                    source="outlook",
                ))
                seen_ids.add(message_id)
            except Exception:
                warn(
                    f"One or more messages in {label} could not be read. "
                    "The report is incomplete. Outlook may require permission "
                    "to read protected message fields."
                )

        if selected_count > available:
            break
        if include_subfolders:
            try:
                children = folder.Folders
                child_count = int(children.Count)
                if child_count < 0:
                    raise ValueError("Invalid folder count")
                # Keep traversal itself bounded even if a collection is huge.
                folder_capacity = max(0, _MAX_FOLDERS - folder_count - len(pending))
                if child_count > folder_capacity:
                    warn("The folder safety limit was reached; the report is incomplete.")
                    child_count = folder_capacity
                child_folders = []
                for child_index in range(1, child_count + 1):
                    try:
                        child_folders.append((
                            children.Item(child_index), f"{label}/subfolder {child_index}"
                        ))
                    except Exception:
                        warn(
                            f"A subfolder of {label} could not be read. "
                            "The report is incomplete."
                        )
                pending.extend(reversed(child_folders))
            except Exception:
                warn(
                    f"Subfolders of {label} could not be read. "
                    "The report is incomplete."
                )

    messages.sort(key=lambda message: (message.received_at, message.id), reverse=True)
    return OutlookReadResult(messages, scanned_count, tuple(warnings))


def _read_with_com(
    *,
    get_active_object: Callable[[str], Any],
    co_initialize: Callable[[], None],
    co_uninitialize: Callable[[], None],
    now: datetime,
    days: int,
    include_subfolders: bool,
    maximum_items: int,
) -> OutlookReadResult:
    """Initialize and release COM on the same thread used for all mailbox reads."""
    try:
        co_initialize()
    except Exception:
        raise OutlookReadError("Windows COM could not be initialized for this review.") from None
    application = None
    namespace = None
    try:
        try:
            application = get_active_object("Outlook.Application")
        except Exception:
            raise OutlookReadError(
                "Open Outlook Classic with your existing mailbox, then try again. "
                "The assistant only attaches to an already-running Outlook instance."
            ) from None
        try:
            namespace = application.GetNamespace("MAPI")
        except Exception:
            raise OutlookReadError(
                "Outlook's local mailbox could not be accessed. Check that Outlook "
                "Classic has finished opening."
            ) from None
        return _read_from_namespace(
            namespace, now=now, days=days, include_subfolders=include_subfolders,
            maximum_items=maximum_items,
        )
    finally:
        # Release apartment-bound references before leaving this worker's COM
        # apartment. Mailbox objects are never passed back to the UI thread.
        namespace = None
        application = None
        co_uninitialize()


def read_outlook_messages(
    *,
    authorization: ReadAuthorization,
    now: datetime,
    days: int = 30,
    include_subfolders: bool = True,
    maximum_items: int = 5_000,
) -> OutlookReadResult:
    """Read authorized local messages; never launch Outlook or change its contents.

    The caller must run this on a worker thread and obtain fresh user permission
    after the sample test passes. No Outlook imports or COM calls occur unless
    both authorization requirements are satisfied.
    """
    if not isinstance(authorization, ReadAuthorization):
        raise OutlookReadError("Explicit authorization is required before reading Outlook.")
    if authorization.sample_test_passed is not True:
        raise OutlookReadError("Pass and review the sample-data test before reading Outlook.")
    if authorization.user_approved is not True:
        raise OutlookReadError("Your permission is required before reading Outlook.")
    _validate_options(
        now=now, days=days, include_subfolders=include_subfolders,
        maximum_items=maximum_items,
    )
    if sys.platform != "win32":
        raise OutlookReadError("Local Outlook review requires Windows and Outlook Classic.")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise OutlookReadError(
            "The Windows Outlook component is missing. Follow the approved "
            "installation steps for pywin32, then try again."
        ) from None
    try:
        return _read_with_com(
            get_active_object=win32com.client.GetActiveObject,
            co_initialize=pythoncom.CoInitialize,
            co_uninitialize=pythoncom.CoUninitialize,
            now=now, days=days, include_subfolders=include_subfolders,
            maximum_items=maximum_items,
        )
    except OutlookReadError:
        raise
    except Exception:
        raise OutlookReadError(
            "Outlook review could not finish. Close the assistant and try again "
            "with Outlook Classic already open."
        ) from None
