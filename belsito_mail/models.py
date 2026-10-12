"""Shared, immutable data passed between the local mailbox and review UI."""

from dataclasses import dataclass
from datetime import date, datetime


@dataclass(frozen=True)
class EmailMessage:
    id: str
    subject: str
    sender_name: str
    sender_email: str
    received_at: datetime
    body: str
    source: str = "sample"

    def __post_init__(self) -> None:
        if self.received_at.tzinfo is None or self.received_at.utcoffset() is None:
            raise ValueError("Email received_at must include a timezone.")


@dataclass(frozen=True)
class FollowUpItem:
    id: str
    message: EmailMessage
    categories: tuple[str, ...]
    priority: str
    action: str
    reasons: tuple[str, ...]
    due_date: date | None = None
