"""
Core bridge connectivity and basic functionality tests.

These tests verify that the Outlook COM bridge can connect and perform basic operations.
"""

import pytest


@pytest.mark.integration
class TestOutlookBridge:
    """Test basic Outlook bridge connectivity and operations"""

    def test_bridge_initialization(self, bridge):
        """Test that bridge can be initialized"""
        assert bridge is not None
        assert bridge.outlook is not None
        assert bridge.namespace is not None

    def test_get_inbox(self, bridge):
        """Test that we can access the Inbox folder"""
        inbox = bridge.get_inbox()
        assert inbox is not None
        # Force a COM call to verify it works
        item_count = inbox.Items.Count
        assert item_count >= 0

    def test_get_calendar(self, bridge):
        """Test that we can access the Calendar folder"""
        calendar = bridge.get_calendar()
        assert calendar is not None
        item_count = calendar.Items.Count
        assert item_count >= 0

    def test_get_tasks(self, bridge):
        """Test that we can access the Tasks folder"""
        tasks = bridge.get_tasks()
        assert tasks is not None
        item_count = tasks.Items.Count
        assert item_count >= 0

    def test_get_current_user(self, bridge):
        """Test that we can get the current user information"""
        user = bridge.namespace.CurrentUser
        assert user is not None
        # Should have at least a name or address
        assert hasattr(user, "Name") or hasattr(user, "Address")

    def test_get_item_by_id_with_invalid_id(self, bridge):
        """Test that get_item_by_id handles invalid IDs gracefully"""
        result = bridge.get_item_by_id("invalid_id_12345")
        assert result is None

    @pytest.mark.slow
    def test_list_emails_returns_list(self, bridge):
        """Test that list_emails returns a list (even if empty)"""
        emails = bridge.list_emails(limit=5)
        assert isinstance(emails, list)

    @pytest.mark.slow
    def test_list_calendar_events_returns_list(self, bridge):
        """Test that list_calendar_events returns a list (even if empty)"""
        events = bridge.list_calendar_events(days=7)
        assert isinstance(events, list)

    @pytest.mark.slow
    def test_list_tasks_returns_list(self, bridge):
        """Test that list_tasks returns a list (even if empty)"""
        tasks = bridge.list_tasks()
        assert isinstance(tasks, list)
