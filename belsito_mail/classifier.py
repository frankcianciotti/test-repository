"""Transparent, local rules for identifying messages that merit human review.

These rules are a triage aid, not a complete reading of a conversation. They do
not open attachments, infer payment status from QuickBooks, or contact a service.
"""

from datetime import date, datetime, timedelta, timezone
import re
from typing import Iterable

from .models import EmailMessage, FollowUpItem


SERVICE = "Customer service / new job"
ESTIMATE = "Estimate / proposal follow-up"
COLLECTIONS = "Outstanding invoice / collections"
REGULATORY = "Regulatory notice"
URGENT = "Urgent"
CATEGORIES = (SERVICE, ESTIMATE, COLLECTIONS, REGULATORY, URGENT)
PRIORITY_ORDER = {"Immediate": 0, "High": 1, "Normal": 2, "Low": 3}


def _pattern(expression: str) -> re.Pattern[str]:
    return re.compile(expression, re.IGNORECASE)


_SERVICE_EXPLICIT = _pattern(
    r"\b(?:service request|new job|new project|work order|need (?:a )?plumber|"
    r"request(?:ing)? (?:a )?(?:repair|service)|"
    r"schedule (?:a |an )?(?:repair|service|visit|installation)|"
    r"please schedule|plumbing (?:renovation|installation)|"
    r"request (?:a )?(?:quote|estimate|proposal))\b"
)
_PLUMBING_PROBLEM = _pattern(
    r"\b(?:leaking|active leak|burst pipe|broken pipe|sewage backup|"
    r"(?:blocked|clogged) (?:drain|toilet|sink)|no (?:hot )?water|"
    r"water heater (?:failure|failed|broken)|overflowing|flooding|"
    r"gas leak|plumbing (?:problem|issue|repair)|"
    r"(?:drain|toilet|sink) (?:is )?(?:blocked|clogged)|"
    r"toilet (?:is )?not flushing|water heater (?:is )?not working)\b"
)
_SERVICE_ASSET = _pattern(
    r"\b(?:plumbing|pipes?|drains?|sinks?|toilets?|water heater|boiler|"
    r"backflow|sewer|faucet|shower|sprinkler)\b"
)
_SERVICE_INTENT = _pattern(
    r"\b(?:need|request(?:ing)?|can you|could you|please|schedule|repair request)\b"
)
_SERVICE_RESOLVED = _pattern(
    r"\b(?:job (?:is )?(?:complete|completed|finished)|"
    r"(?:leak|repair|issue|problem|(?:burst |broken )?pipe|flooding|gas leak|"
    r"sewage backup|drain|toilet|sink|water heater) "
    r"(?:is |was |has been |has )?(?:fixed|repaired|resolved|stopped)|"
    r"service (?:is )?completed|no (?:further )?(?:service|repair) (?:is )?needed)\b"
)
_ESTIMATE = _pattern(r"\b(?:estimate|proposal|quote|bid)\b")
_ESTIMATE_PENDING = _pattern(
    r"\b(?:follow[ -]?up|awaiting|waiting for|pending|no response|"
    r"have(?:n['’]t| not) heard|checking in|for (?:your )?review|"
    r"please (?:review|approve|confirm|provide|send)|"
    r"request(?:ed|ing)? (?:a |an )?(?:estimate|proposal|quote|bid)|"
    r"need (?:a |an )?(?:estimate|proposal|quote|bid)|"
    r"status (?:of|update)|let (?:us|me) know|decision)\b"
)
_ESTIMATE_RESOLVED = _pattern(
    r"\b(?:(?:estimate|proposal|quote|bid) (?:is |was |has been )?"
    r"(?:approved|accepted|signed|declined|rejected|cancelled|canceled)|"
    r"(?:approved|accepted|signed|declined|rejected) (?:the |your |our )?"
    r"(?:estimate|proposal|quote|bid))\b"
)
_INVOICE = _pattern(r"\b(?:invoice|invoices|payment|balance|account|collections?)\b")
_UNSETTLED = _pattern(
    r"\b(?:outstanding|overdue|past[ -]due|unpaid|not (?:yet )?paid|"
    r"has(?:n['’]t| not) been paid|never paid|balance due|payment (?:is )?(?:due|reminder)|"
    r"invoice(?: #?[\w-]+)? (?:is )?(?:due|reminder)|collections?|demand for payment|"
    r"unsettled|remaining balance)\b"
)
_PAID = _pattern(
    r"(?<!not )(?<!not yet )\b(?:paid in full|fully paid|payment received|"
    r"balance (?:is )?(?:zero|cleared)|account settled|"
    r"invoice(?: #?[\w-]+)? (?:is |was |has been )?paid)\b"
)
_AGENCY = _pattern(
    r"\b(?:DOB|DEP|FDNY|OSHA|EPA|department of buildings|"
    r"department of environmental protection|fire department|"
    r"department of health|health department|building department|"
    r"environmental protection agency)\b"
)
_REGULATORY_ACTION = _pattern(
    r"\b(?:notice|violation|inspection|permit|summons|hearing|citation|"
    r"compliance|consent order|stop[ -]work order|regulatory)\b"
)
_REGULATORY_EXPLICIT = _pattern(
    r"\b(?:regulatory notice|code violation|compliance notice|"
    r"stop[ -]work order|notice of violation)\b"
)
_URGENCY = _pattern(
    r"\b(?:urgent|urgently|emergency|immediate(?:ly)?|asap|"
    r"burst pipe|active flooding|flooding|gas leak|sewage backup|"
    r"(?:respond|reply|action required|deadline|due|by) today|"
    r"no (?:hot )?water)\b"
)
_NEGATED_URGENCY = _pattern(
    r"\b(?:not (?:an? )?urgent|not (?:an? )?emergency|non[ -]urgent|"
    r"no (?:urgent|immediate) (?:action|response) (?:is )?(?:required|needed)|"
    r"no emergency|no urgency|(?:urgent|immediate) (?:action|response) "
    r"(?:is )?not (?:required|needed)|not needed immediately)\b"
)
_DEADLINE = _pattern(
    r"\b(?:due(?: date)?|deadline|respond by|reply by|no later than|"
    r"inspection on|hearing on)\s*(?:is\s+|:\s*)?(\d{4}-\d{2}-\d{2})\b"
)


def _current_body(body: str) -> str:
    """Avoid treating a quoted previous request as a fresh request."""
    lines: list[str] = []
    for line in body.splitlines():
        stripped = line.strip()
        if re.match(r"^-+\s*Original Message\s*-+$", stripped, re.I):
            break
        if re.match(r"^On .+ wrote:$", stripped, re.I):
            break
        if not stripped.startswith(">"):
            lines.append(line)
    return "\n".join(lines)


def _clauses(text: str) -> list[str]:
    return [
        part.strip()
        for part in re.split(r"[.!?;\n]+|\b(?:but|however)\b|\band (?:now|currently)\b", text, flags=re.I)
        if part.strip()
    ]


def _evidence(pattern: re.Pattern[str], text: str) -> str | None:
    match = pattern.search(text)
    return match.group(0) if match else None


def _active_invoice_clause(clause: str) -> bool:
    if not (_INVOICE.search(clause) and _UNSETTLED.search(clause)):
        return False
    # A confirmation of payment can repeat the earlier "overdue" wording.
    return _PAID.search(clause) is None or bool(
        re.search(r"\b(?:unpaid|not (?:yet )?paid|remaining balance)\b", clause, re.I)
    )


def _service_evidence(clause: str) -> str | None:
    phrase = _evidence(_SERVICE_EXPLICIT, clause) or _evidence(_PLUMBING_PROBLEM, clause)
    if phrase:
        return phrase
    if _SERVICE_ASSET.search(clause) and _SERVICE_INTENT.search(clause):
        return f"{_evidence(_SERVICE_INTENT, clause)} / {_evidence(_SERVICE_ASSET, clause)}"
    return None


def _urgency_evidence(clauses: Iterable[str]) -> str | None:
    for clause in clauses:
        # Mentioning a past emergency in a resolution does not make it active.
        if _SERVICE_RESOLVED.search(clause) or _ESTIMATE_RESOLVED.search(clause):
            continue
        if _PAID.search(clause) and not _active_invoice_clause(clause):
            continue
        phrase = _evidence(_URGENCY, _NEGATED_URGENCY.sub("", clause))
        if phrase:
            return phrase
    return None


def _due_date(text: str) -> date | None:
    dates: list[date] = []
    for match in _DEADLINE.finditer(text):
        try:
            dates.append(date.fromisoformat(match.group(1)))
        except ValueError:
            continue
    return min(dates) if dates else None


def _review_one(message: EmailMessage, now: datetime) -> FollowUpItem | None:
    body = _current_body(message.body)
    text = f"{message.subject}\n{body}"
    body_clauses = _clauses(body)
    clauses = _clauses(text)
    categories: list[str] = []
    reasons: list[str] = []

    service_resolved = bool(_SERVICE_RESOLVED.search(body))
    active_body_service = any(
        _service_evidence(clause)
        and not _SERVICE_RESOLVED.search(clause)
        for clause in body_clauses
    )
    service_phrase = None
    if not service_resolved or active_body_service:
        for clause in body_clauses if service_resolved else clauses:
            if _SERVICE_RESOLVED.search(clause):
                continue
            service_phrase = _service_evidence(clause)
            if service_phrase:
                break
    if service_phrase:
        categories.append(SERVICE)
        reasons.append(f"Service/new-job wording: '{service_phrase}'.")

    estimate_resolved = bool(_ESTIMATE_RESOLVED.search(body))
    estimate_phrase = _evidence(_ESTIMATE, text)
    pending_phrase = _evidence(_ESTIMATE_PENDING, text)
    if estimate_phrase and pending_phrase and not estimate_resolved:
        categories.append(ESTIMATE)
        reasons.append(
            f"Estimate/proposal wording: '{estimate_phrase}' with '{pending_phrase}'."
        )

    body_unsettled = any(_active_invoice_clause(clause) for clause in body_clauses)
    body_paid = bool(_PAID.search(body))
    invoice_phrase = None
    if not body_paid or body_unsettled:
        for clause in clauses:
            if _active_invoice_clause(clause):
                invoice_phrase = _evidence(_UNSETTLED, clause)
                break
    if invoice_phrase:
        categories.append(COLLECTIONS)
        reasons.append(f"Possible unpaid-account wording: '{invoice_phrase}'. Verify payment status.")

    agency_phrase = _evidence(_AGENCY, text)
    regulatory_phrase = _evidence(_REGULATORY_EXPLICIT, text)
    if regulatory_phrase or (agency_phrase and _REGULATORY_ACTION.search(text)):
        categories.append(REGULATORY)
        if agency_phrase:
            reasons.append(f"Agency '{agency_phrase}' appears with notice/compliance wording.")
        else:
            reasons.append(f"Regulatory wording: '{regulatory_phrase}'.")

    body_urgency = _urgency_evidence(body_clauses)
    urgency_phrase = body_urgency or _urgency_evidence(_clauses(message.subject))
    # A subject can repeat an old emergency even when the latest reply resolves
    # it, records payment, or explicitly says that urgent action is unnecessary.
    resolved_body = (
        service_resolved
        or (body_paid and not body_unsettled)
        or estimate_resolved
    )
    if resolved_body or (_NEGATED_URGENCY.search(body) and not body_urgency):
        urgency_phrase = body_urgency
    if urgency_phrase:
        categories.append(URGENT)
        reasons.append(f"Urgency wording: '{urgency_phrase}'.")

    if not categories:
        return None

    deadline = _due_date(text)
    priority = "Normal"
    if URGENT in categories:
        priority = "Immediate"
    elif REGULATORY in categories or COLLECTIONS in categories:
        priority = "High"
    elif now.astimezone(timezone.utc) - message.received_at.astimezone(timezone.utc) >= timedelta(days=7):
        priority = "High"
        reasons.append("Message is at least 7 days old; check whether a response is still needed.")
    if deadline is not None:
        reasons.append(f"Explicit date in message: {deadline.isoformat()}; verify the context.")
        local_today = now.date()
        if deadline <= local_today and REGULATORY in categories:
            priority = "Immediate"
        elif deadline <= local_today + timedelta(days=2) and priority != "Immediate":
            priority = "High"

    actions = {
        URGENT: "Review promptly and decide the next action",
        REGULATORY: "Verify the notice, stated deadline, and responsible person",
        COLLECTIONS: "Verify the balance and decide whether payment follow-up is needed",
        SERVICE: "Confirm the customer's request and next scheduling step",
        ESTIMATE: "Check the estimate/proposal status and decide the follow-up",
    }
    action = "; ".join(
        actions[category]
        for category in (URGENT, REGULATORY, COLLECTIONS, SERVICE, ESTIMATE)
        if category in categories
    ) + "."
    return FollowUpItem(
        id=message.id,
        message=message,
        categories=tuple(categories),
        priority=priority,
        action=action,
        reasons=tuple(reasons),
        due_date=deadline,
    )


def review_messages(
    messages: Iterable[EmailMessage], now: datetime, days: int = 30
) -> list[FollowUpItem]:
    """Review an inclusive rolling window; sort by urgency, then oldest first.

    Only message text is used. Results are candidates for human follow-up, not
    confirmation of a job, a debt, an official notice, or a real-world deadline.
    The timezone of ``now`` determines today's calendar date for deadline review.
    """
    if now.tzinfo is None or now.utcoffset() is None:
        raise ValueError("Review time must include a timezone.")
    if isinstance(days, bool) or not isinstance(days, int) or days <= 0:
        raise ValueError("Review days must be a positive integer.")
    utc_now = now.astimezone(timezone.utc)
    cutoff = utc_now - timedelta(days=days)
    items: list[FollowUpItem] = []
    for message in messages:
        if cutoff <= message.received_at.astimezone(timezone.utc) <= utc_now:
            item = _review_one(message, now)
            if item is not None:
                items.append(item)
    return sorted(items, key=lambda item: (
        PRIORITY_ORDER[item.priority], item.message.received_at.astimezone(timezone.utc), item.id
    ))
