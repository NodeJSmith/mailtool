"""
Calendar functionality tests.

Tests for listing, creating, and managing calendar appointments.
"""

from datetime import datetime, timedelta

import pytest

from mailtool.bridge import MailtoolError

from .conftest import TEST_PREFIX, assert_calendar_structure, assert_valid_entry_id


@pytest.mark.integration
@pytest.mark.calendar
class TestCalendar:
    """Test calendar-related functionality"""

    def test_list_calendar_events_default(self, bridge):
        """Test listing calendar events with default parameters"""
        events = bridge.list_calendar_events()
        assert isinstance(events, list)

    def test_list_calendar_events_custom_days(self, bridge):
        """Test listing calendar events for custom day range"""
        events = bridge.list_calendar_events(days=30)
        assert isinstance(events, list)

        for event in events:
            assert_calendar_structure(event)
            # Verify event is within date range
            # (allowing some margin for recurrence)

    @pytest.mark.slow
    def test_list_calendar_events_all(self, bridge):
        """Test listing all calendar events without date filtering

        Note: This test is marked as slow because it iterates over ALL calendar items,
        which can include corrupted items that cause COM errors. In production,
        use date-filtered queries (days parameter) instead of all_events=True.
        """
        events = bridge.list_calendar_events(all_events=True)
        assert isinstance(events, list)

        # Should return more events than the filtered version
        events_filtered = bridge.list_calendar_events(days=7)
        assert len(events) >= len(events_filtered)

    def test_create_appointment_basic(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a basic appointment"""
        subject = f"{TEST_PREFIX}Appt Test {test_timestamp}"

        start = datetime.now() + timedelta(hours=1)
        end = start + timedelta(minutes=30)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
        )

        assert_valid_entry_id(entry_id)

        # Verify we can retrieve it
        appt = bridge.get_appointment(entry_id)
        assert appt is not None
        assert appt["subject"] == subject

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_create_appointment_with_location(
        self, bridge, test_timestamp, cleanup_helpers
    ):
        """Test creating an appointment with location"""
        subject = f"{TEST_PREFIX}Location Test {test_timestamp}"
        location = "Conference Room B"

        start = datetime.now() + timedelta(hours=2)
        end = start + timedelta(hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
            location=location,
        )

        assert_valid_entry_id(entry_id)

        # Verify location was set
        appt = bridge.get_appointment(entry_id)
        assert appt is not None
        assert appt["location"] == location

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_create_appointment_with_body(
        self, bridge, test_timestamp, cleanup_helpers
    ):
        """Test creating an appointment with body/description"""
        subject = f"{TEST_PREFIX}Body Test {test_timestamp}"
        body = f"Agenda:\n1. Review project status\n2. Discuss timeline\n3. Action items\n\nTest ID: {test_timestamp}"

        start = datetime.now() + timedelta(hours=3)
        end = start + timedelta(minutes=45)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
            body=body,
        )

        assert_valid_entry_id(entry_id)

        # Verify body was set
        appt = bridge.get_appointment(entry_id)
        assert appt is not None
        assert test_timestamp in appt["body"]

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_create_all_day_event(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating an all-day event"""
        subject = f"{TEST_PREFIX}All Day Test {test_timestamp}"

        # For all-day events, start and end should be dates (no time)
        start = datetime.now() + timedelta(days=1)
        start = start.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
            all_day=True,
        )

        assert_valid_entry_id(entry_id)

        # Verify all_day flag was set
        appt = bridge.get_appointment(entry_id)
        assert appt is not None
        assert appt["all_day"] is True

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_create_appointment_with_attendees(
        self, bridge, test_timestamp, cleanup_helpers
    ):
        """Test creating an appointment with attendees"""
        subject = f"{TEST_PREFIX}Attendees Test {test_timestamp}"

        start = datetime.now() + timedelta(days=2)
        end = start + timedelta(hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
            required_attendees="required1@example.com; required2@example.com",
            optional_attendees="optional1@example.com",
        )

        assert_valid_entry_id(entry_id)

        # Verify attendees were set
        appt = bridge.get_appointment(entry_id)
        assert appt is not None
        assert "required1@example.com" in appt["required_attendees"]
        assert "optional1@example.com" in appt["optional_attendees"]

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_get_appointment(self, bridge, sample_calendar_data):
        """Test retrieving appointment by EntryID"""
        appt = bridge.get_appointment(sample_calendar_data["entry_id"])

        assert appt is not None
        assert isinstance(appt, dict)
        assert appt["entry_id"] == sample_calendar_data["entry_id"]
        assert appt["subject"] == sample_calendar_data["subject"]

        # Verify all expected fields are present
        expected_fields = [
            "entry_id",
            "subject",
            "start",
            "end",
            "location",
            "body",
            "all_day",
            "required_attendees",
            "optional_attendees",
            "response_status",
            "meeting_status",
            "response_requested",
        ]
        for field in expected_fields:
            assert field in appt

    def test_edit_appointment(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing an existing appointment"""
        # Create appointment
        subject = f"{TEST_PREFIX}Edit Test {test_timestamp}"
        start = datetime.now() + timedelta(days=3)
        end = start + timedelta(hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
            location="Original Location",
        )

        # Edit it
        new_location = "Updated Location"
        result = bridge.edit_appointment(entry_id, location=new_location)

        assert result is True

        # Verify changes
        appt = bridge.get_appointment(entry_id)
        assert appt["location"] == new_location

        # Cleanup
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)

    def test_delete_appointment(self, bridge, test_timestamp):
        """Test deleting an appointment"""
        # Create appointment
        subject = f"{TEST_PREFIX}Delete Test {test_timestamp}"
        start = datetime.now() + timedelta(days=4)
        end = start + timedelta(hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.strftime("%Y-%m-%d %H:%M:%S"),
            end=end.strftime("%Y-%m-%d %H:%M:%S"),
        )

        # Verify it exists
        appt = bridge.get_appointment(entry_id)
        assert appt is not None

        # Delete it
        result = bridge.delete_appointment(entry_id)
        assert result is True

        # Verify it's gone
        appt = bridge.get_appointment(entry_id)
        assert appt is None

    def test_get_free_busy_current_user(self, bridge):
        """Test getting free/busy information for current user"""
        # Get free/busy for today (defaults to current user if no email provided)
        result = bridge.get_free_busy()

        assert result is not None
        assert isinstance(result, dict)
        assert "email" in result
        assert "start_date" in result
        assert "end_date" in result

    def test_get_free_busy_with_dates(self, bridge):
        """Test getting free/busy with specific date range"""
        start_date = "2025-01-01"
        end_date = "2025-01-02"

        result = bridge.get_free_busy(start_date=start_date, end_date=end_date)

        assert result is not None
        assert isinstance(result, dict)
        assert result["start_date"] == start_date
        assert result["end_date"] == end_date

    def test_free_busy_unavailable_raises(self, bridge):
        """A resolvable address with no free/busy data raises MailtoolError.

        Any SMTP address resolves, so an external one gets as far as FreeBusy,
        which fails with "Cannot load free/busy information".
        """
        with pytest.raises(MailtoolError, match="free/busy"):
            bridge.get_free_busy(email_address="nonexistent@example.com")

    def test_free_busy_unresolvable_name(self, bridge):
        """A name Outlook can't resolve is a normal answer, not a failure"""
        result = bridge.get_free_busy(email_address="zz no such person qq")

        assert result["resolved"] is False
        assert result["error"] == "Could not resolve email address"
