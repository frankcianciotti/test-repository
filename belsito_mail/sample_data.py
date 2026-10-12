"""Entirely synthetic mailbox data for testing before Outlook access."""

from datetime import datetime, timedelta

from .models import EmailMessage


def sample_messages(now: datetime) -> list[EmailMessage]:
    """Use reserved .invalid addresses; no customer or mailbox data is included."""
    tomorrow = (now + timedelta(days=1)).date().isoformat()
    next_week = (now + timedelta(days=7)).date().isoformat()
    examples = [
        ("burst-pipe", 0.04, "Emergency: burst pipe at Example Building", "Sample Property Manager",
         "An active burst pipe is flooding the basement. Please send help immediately."),
        ("new-job", 2, "New job: plumbing renovation", "Sample General Contractor",
         "We are starting a plumbing renovation at 100 Example Street. Can you schedule a visit?"),
        ("proposal", 9, "Follow-up on proposal BP-SAMPLE-14", "Sample Building Owner",
         "Our proposal for the boiler work is awaiting your review. Please confirm the next step."),
        ("estimate", 3, "Estimate for water heater replacement", "Sample Homeowner",
         "Checking in on the estimate. Please let me know when you can provide it."),
        ("invoice", 12, "Outstanding invoice SAMPLE-1042", "Sample Accounts Contact",
         "The invoice remains unpaid and is now overdue. Please review the remaining balance."),
        ("dob", 1, "DOB notice: plumbing inspection", "Sample DOB Notice",
         f"This synthetic Department of Buildings notice requests inspection documentation. Respond by {tomorrow}."),
        ("dep", 4, "DEP compliance notice", "Sample DEP Notice",
         f"This synthetic Department of Environmental Protection notice concerns discharge permit records. Deadline: {next_week}."),
        ("fdny", 6, "FDNY inspection notice", "Sample FDNY Notice",
         "This synthetic Fire Department notice requests inspection records for review."),
        ("other-regulatory", 5, "OSHA compliance notice", "Sample Safety Notice",
         "This synthetic OSHA compliance notice requests a response from the business owner."),
        ("multiple", 0.2, "Urgent new job and quote request", "Sample Commercial Customer",
         "We need a plumber for a new job. Please send a proposal for review ASAP."),
        ("not-urgent", 1.5, "Service request: routine appointment, not urgent", "Sample Customer",
         "This is not an emergency. Please schedule a service visit for next week."),
        ("paid", 2.5, "Re: overdue invoice SAMPLE-900", "Sample Payer",
         "The invoice is paid in full. Thank you for confirming payment received."),
        ("resolved", 2.8, "Re: Emergency leaking pipe", "Sample Customer",
         "The leak is repaired and the job is complete. Thank you."),
        ("routine", 1, "Weekly plumbing industry newsletter", "Sample Newsletter",
         "Welcome to this week's industry news and product updates. No reply is necessary."),
        ("old", 36, "Urgent service request from last month", "Sample Old Request",
         "We need a plumber to repair a burst pipe."),
        ("future", -1, "Service request with a future timestamp", "Sample Future Message",
         "Please schedule a repair visit."),
    ]
    return [
        EmailMessage(
            id=f"sample-{identifier}", subject=subject,
            sender_name=sender, sender_email=f"{identifier}@example.invalid",
            received_at=now - timedelta(days=age), body=body,
        )
        for identifier, age, subject, sender, body in examples
    ]
