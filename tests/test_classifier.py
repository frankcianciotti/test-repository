import unittest
from datetime import datetime, timedelta, timezone

from belsito_mail.classifier import (
    COLLECTIONS, ESTIMATE, REGULATORY, SERVICE, URGENT, review_messages,
)
from belsito_mail.models import EmailMessage
from belsito_mail.sample_data import sample_messages


class ClassifierTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime(2026, 10, 12, 14, 0, tzinfo=timezone.utc)

    def message(self, subject, body="", age=1, identifier="test"):
        return EmailMessage(
            identifier, subject, "Sample Sender", "sender@example.invalid",
            self.now - timedelta(days=age), body,
        )

    def review(self, subject, body="", age=1):
        return review_messages([self.message(subject, body, age)], self.now)

    def test_inclusive_window_and_future_exclusion(self):
        messages = [
            self.message("Service request", age=30, identifier="boundary"),
            self.message("Service request", age=30.0001, identifier="too-old"),
            self.message("Service request", age=0, identifier="current"),
            self.message("Service request", age=-0.0001, identifier="future"),
        ]
        self.assertEqual([item.id for item in review_messages(messages, self.now)], ["boundary", "current"])

    def test_service_request_and_multi_label(self):
        items = self.review("Urgent new job", "Need a plumber. Please send a proposal for review ASAP.")
        self.assertEqual(set(items[0].categories), {SERVICE, ESTIMATE, URGENT})
        self.assertEqual(items[0].priority, "Immediate")
        self.assertTrue(items[0].reasons)

    def test_common_plumbing_requests_are_identified(self):
        for subject in ("Need water heater replacement", "Could you repair our boiler?", "Toilet is not flushing", "Sink clogged"):
            with self.subTest(subject=subject):
                self.assertIn(SERVICE, self.review(subject)[0].categories)
        self.assertEqual(self.review("Water heater industry newsletter", "Replacement technology update."), [])

    def test_estimate_requires_pending_signal(self):
        self.assertEqual(self.review("Industry estimates for 2026", "Average prices increased."), [])
        item = self.review("Proposal", "Checking in on the proposal for your review.")[0]
        self.assertEqual(item.categories, (ESTIMATE,))
        self.assertEqual(item.priority, "Normal")
        self.assertEqual(self.review("Follow-up on proposal", "The proposal is approved."), [])

    def test_unpaid_invoice_and_payment_confirmation(self):
        item = self.review("Invoice 123", "This invoice is not yet paid and remains outstanding.")[0]
        self.assertEqual(item.categories, (COLLECTIONS,))
        self.assertEqual(item.priority, "High")
        self.assertEqual(self.review("Re: overdue invoice", "The invoice is paid in full."), [])
        self.assertEqual(self.review("Urgent overdue invoice", "The invoice is paid in full."), [])
        self.assertEqual(self.review("Past-due invoice paid", "Invoice 123 was paid in full."), [])

    def test_paid_one_invoice_does_not_hide_another_unpaid_invoice(self):
        item = self.review("Account status", "Invoice 12 is paid in full; invoice 13 remains overdue.")[0]
        self.assertIn(COLLECTIONS, item.categories)

    def test_regulatory_notices_and_acronym_context(self):
        for agency in ("DOB", "DEP", "FDNY", "OSHA", "EPA"):
            with self.subTest(agency=agency):
                item = self.review(f"{agency} inspection notice")[0]
                self.assertIn(REGULATORY, item.categories)
                self.assertEqual(item.priority, "High")
        self.assertEqual(self.review("Please record DOB", "The patient's date of birth is provided."), [])
        self.assertIn(REGULATORY, self.review("Notice of violation", "Please review.")[0].categories)

    def test_negated_urgent_words_do_not_escalate_service(self):
        for wording in ("not urgent", "not an emergency", "no immediate action required", "non-urgent"):
            with self.subTest(wording=wording):
                item = self.review(f"Service request: {wording}", "Please schedule a service visit.")[0]
                self.assertNotIn(URGENT, item.categories)
                self.assertEqual(item.priority, "Normal")
        item = self.review("Urgent service request", "This is not urgent. Please schedule a service visit.")[0]
        self.assertNotIn(URGENT, item.categories)

    def test_resolved_request_and_quoted_thread_are_excluded(self):
        self.assertEqual(self.review("Re: Emergency leaking pipe", "The leak is repaired and the job is complete."), [])
        self.assertEqual(self.review("Re: Burst pipe", "The burst pipe was repaired and the job is complete."), [])
        self.assertEqual(self.review("Re: Burst pipe", "The burst pipe was repaired."), [])
        self.assertEqual(self.review("Thank you", "Thanks for completing the work.\nOn Monday, Customer wrote:\nUrgent burst pipe!"), [])
        self.assertEqual(self.review("Thanks", "All good.\n> Urgent service request"), [])

    def test_resolved_clause_does_not_hide_separate_active_emergency(self):
        for body in (
            "The burst pipe was repaired. There is a new gas leak. Please help.",
            "The burst pipe was repaired, but there is a new gas leak. Please help.",
        ):
            with self.subTest(body=body):
                item = self.review("Re: Burst pipe", body)[0]
                self.assertIn(URGENT, item.categories)
                self.assertIn(SERVICE, item.categories)
                self.assertEqual(item.priority, "Immediate")
                self.assertIn("Urgency wording: 'gas leak'.", item.reasons)

    def test_old_follow_up_escalates_and_stable_sort_uses_priority_age_id(self):
        messages = [
            self.message("Service request", age=1, identifier="new"),
            self.message("Follow-up on proposal", age=9, identifier="old"),
            self.message("Emergency burst pipe", age=0, identifier="urgent"),
            self.message("Invoice overdue", age=5, identifier="b"),
            self.message("Invoice overdue", age=5, identifier="a"),
        ]
        items = review_messages(messages, self.now)
        self.assertEqual([item.id for item in items], ["urgent", "old", "a", "b", "new"])
        self.assertEqual(items[1].priority, "High")

    def test_only_explicit_valid_iso_deadlines_are_used(self):
        item = self.review("DOB compliance notice", "Respond by 2026-10-12.")[0]
        self.assertEqual(item.due_date.isoformat(), "2026-10-12")
        self.assertEqual(item.priority, "Immediate")
        item = self.review("DEP permit notice", "Received 2026-10-10; deadline next Friday.")[0]
        self.assertIsNone(item.due_date)
        item = self.review("DEP permit notice", "Deadline: 2026-02-30.")[0]
        self.assertIsNone(item.due_date)

    def test_deadline_uses_local_calendar_date_near_midnight(self):
        local_now = datetime(2026, 10, 12, 22, 30, tzinfo=timezone(timedelta(hours=-4)))
        message = EmailMessage(
            "local-deadline", "DOB inspection notice", "Sample Agency",
            "agency@example.invalid", local_now - timedelta(hours=1),
            "Respond by 2026-10-13.",
        )
        item = review_messages([message], local_now)[0]
        self.assertEqual(item.priority, "High")
        self.assertEqual(item.due_date.isoformat(), "2026-10-13")

    def test_sample_contains_expected_inclusions_and_exclusions(self):
        originals = sample_messages(self.now)
        items = review_messages(originals, self.now)
        ids = {item.id for item in items}
        self.assertEqual(len(originals), 16)
        self.assertEqual(len(items), 11)
        self.assertTrue({"sample-old", "sample-future", "sample-paid", "sample-resolved", "sample-routine"}.isdisjoint(ids))
        self.assertEqual(items[0].priority, "Immediate")
        self.assertTrue(all(message.sender_email.endswith(".invalid") for message in originals))

    def test_invalid_time_and_days_are_rejected(self):
        with self.assertRaises(ValueError):
            self.message("Service request").__class__("id", "subject", "sender", "sender@example.invalid", datetime(2026, 10, 12), "body")
        with self.assertRaises(ValueError):
            review_messages([], datetime(2026, 10, 12))
        for days in (0, -1, True, 1.5):
            with self.subTest(days=days), self.assertRaises(ValueError):
                review_messages([], self.now, days=days)


if __name__ == "__main__":
    unittest.main()
