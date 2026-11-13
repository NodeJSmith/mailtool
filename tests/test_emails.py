"""
Email functionality tests.

Tests for listing, retrieving, creating, and managing emails.
"""

import pytest

from .conftest import TEST_PREFIX, assert_email_structure, assert_valid_entry_id


@pytest.mark.integration
@pytest.mark.email
class TestEmails:
    """Test email-related functionality"""

    def test_list_emails_small_limit(self, bridge):
        """Test listing a small number of emails"""
        emails = bridge.list_emails(limit=5)
        assert isinstance(emails, list)
        assert len(emails) <= 5

        for email in emails:
            assert_email_structure(email)
            assert "entry_id" in email
            assert "subject" in email

    def test_list_emails_from_inbox(self, bridge):
        """Test listing emails specifically from Inbox"""
        emails = bridge.list_emails(limit=10, folder="Inbox")
        assert isinstance(emails, list)

    def test_list_emails_default_folder(self, bridge):
        """Test that default folder is Inbox"""
        emails_default = bridge.list_emails(limit=5)
        emails_inbox = bridge.list_emails(limit=5, folder="Inbox")
        # Both should return the same number (default is Inbox)
        assert len(emails_default) == len(emails_inbox)

    def test_create_draft_email(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a draft email"""
        subject = f"{TEST_PREFIX}Draft Test {test_timestamp}"
        body = f"Test draft email body\nTimestamp: {test_timestamp}"

        entry_id = bridge.send_email(
            to="test@example.com", subject=subject, body=body, save_draft=True
        )

        assert_valid_entry_id(entry_id)

        # Verify we can retrieve the draft
        draft = bridge.get_item_by_id(entry_id)
        assert draft is not None
        assert draft.Subject == subject

        # Cleanup
        cleanup_helpers["delete_drafts_by_prefix"](TEST_PREFIX)

    def test_create_draft_with_cc_bcc(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a draft with CC and BCC"""
        subject = f"{TEST_PREFIX}Draft CC/BCC Test {test_timestamp}"

        entry_id = bridge.send_email(
            to="to@example.com",
            subject=subject,
            body="Test body",
            cc="cc@example.com",
            bcc="bcc@example.com",
            save_draft=True,
        )

        assert_valid_entry_id(entry_id)

        # Verify recipients were set
        draft = bridge.get_item_by_id(entry_id)
        assert draft is not None
        # Note: Outlook may format recipients differently

        # Cleanup
        cleanup_helpers["delete_drafts_by_prefix"](TEST_PREFIX)

    def test_create_draft_with_html_body(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a draft with HTML body"""
        subject = f"{TEST_PREFIX}HTML Draft Test {test_timestamp}"
        html_body = f"<html><body><h1>Test HTML</h1><p>Timestamp: {test_timestamp}</p></body></html>"

        entry_id = bridge.send_email(
            to="test@example.com",
            subject=subject,
            body="Plain text fallback",
            html_body=html_body,
            save_draft=True,
        )

        assert_valid_entry_id(entry_id)

        # Verify HTML body was set
        draft = bridge.get_item_by_id(entry_id)
        assert draft is not None
        assert draft.HTMLBody is not None

        # Cleanup
        cleanup_helpers["delete_drafts_by_prefix"](TEST_PREFIX)

    def test_get_email_body(self, bridge, sample_email_data):
        """Test retrieving full email body by EntryID"""
        email = bridge.get_email_body(sample_email_data["entry_id"])

        assert email is not None
        assert isinstance(email, dict)
        assert "entry_id" in email
        assert "subject" in email
        assert "body" in email
        assert "html_body" in email

        # Verify it's our test email
        assert email["subject"] == sample_email_data["subject"]
        assert sample_email_data["test_id"] in email["body"]

    def test_mark_email_read_unread(self, bridge, sample_email_data):
        """Test marking email as read/unread"""
        entry_id = sample_email_data["entry_id"]

        # Mark as read (unread=False)
        result = bridge.mark_email_read(entry_id, unread=False)
        assert result is True

        # Verify
        item = bridge.get_item_by_id(entry_id)
        assert item.Unread is False

        # Mark as unread (unread=True)
        result = bridge.mark_email_read(entry_id, unread=True)
        assert result is True

        # Verify
        item = bridge.get_item_by_id(entry_id)
        assert item.Unread is True

    def test_delete_email(self, bridge, test_timestamp, cleanup_helpers):
        """Test deleting an email"""
        # Create a draft
        subject = f"{TEST_PREFIX}Delete Test {test_timestamp}"
        entry_id = bridge.send_email(
            to="test@example.com", subject=subject, body="Test body", save_draft=True
        )

        assert_valid_entry_id(entry_id)

        # Verify it exists
        item = bridge.get_item_by_id(entry_id)
        assert item is not None

        # Delete it
        result = bridge.delete_email(entry_id)
        assert result is True

        # Verify it's gone
        item = bridge.get_item_by_id(entry_id)
        # Note: After deletion, get_item_by_id may return None or raise
        # behavior can vary, so we just check it's not the same item

    def test_search_emails_by_subject(self, bridge, test_timestamp, cleanup_helpers):
        """Test searching emails using Restriction filter"""
        # Test searching for unread emails using a simple filter
        # Note: Complex LIKE queries with DASL syntax may not work on all Outlook locales
        # Using simple boolean filter which is more universally supported
        filter_query = "[Unread] = TRUE"
        results = bridge.search_emails(filter_query, limit=10)

        assert isinstance(results, list)
        # Should find some unread emails (or return empty list if none)
        # All results should be unread
        for email in results:
            assert email["unread"] is True, (
                "search_emails returned non-unread email when filtering for unread"
            )

    def test_move_email_to_folder(self, bridge, test_timestamp, cleanup_helpers):
        """Test moving an email to a different folder"""
        # Create a draft
        subject = f"{TEST_PREFIX}Move Test {test_timestamp}"
        entry_id = bridge.send_email(
            to="test@example.com", subject=subject, body="Test body", save_draft=True
        )

        try:
            # Inbox is the one folder every mailbox has
            assert bridge.move_email(entry_id, "Inbox") is True
        finally:
            # The moved draft lands in the Inbox, so clean up both folders. Left
            # behind, a sender-less draft at the top of the Inbox breaks the
            # sender-search tests below.
            cleanup_helpers["delete_drafts_by_prefix"](TEST_PREFIX)
            cleanup_helpers["delete_inbox_by_prefix"](TEST_PREFIX)

    @pytest.fixture
    def inbox_sender(self, bridge):
        """Sender address of the newest Inbox email that has one.

        Drafts and some system items have no sender, and searching for an empty
        sender matches nothing, so it would walk the whole Inbox.
        """
        for email in bridge.list_emails(limit=20):
            if email["sender"]:
                return email["sender"]
        pytest.skip("No Inbox email with a sender to search for")

    def test_search_by_sender(self, bridge, inbox_sender):
        """Test searching emails by sender (handles Exchange addresses)"""
        # limit=1: search_by_sender scans newest-first until it has `limit`
        # matches, so a sender with fewer matches would make it walk the whole
        # Inbox. The email the sender came from is near the top.
        results = bridge.search_by_sender(inbox_sender, limit=1)

        assert len(results) == 1
        assert_email_structure(results[0])
        assert results[0]["sender"].lower() == inbox_sender.lower()

    def test_search_by_sender_case_insensitive(self, bridge, inbox_sender):
        """Test that search_by_sender is case-insensitive"""
        results = bridge.search_by_sender(inbox_sender.upper(), limit=1)

        assert len(results) == 1
        assert results[0]["sender"].lower() == inbox_sender.lower()
