"""
Calendar functionality tests.

Tests for listing, creating, and managing calendar appointments.
"""

import pytest
from whenever import Time

from mailtool.bridge import (
    DATE_PATTERN,
    WALL_CLOCK_PATTERN,
    MailtoolError,
    now_wall_clock,
    parse_wall_clock,
    to_outlook,
)

from .conftest import (
    TEST_PREFIX,
    assert_calendar_structure,
    assert_valid_entry_id,
    from_now,
    shifted,
)

# Meetings in the daily_series fixture.
OCCURRENCES = 4


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

    def test_month_window_includes_the_week(self, bridge):
        """A 30-day listing contains every event the 7-day listing does.

        30 days rather than all_events: the 365-day window takes ~40s on a busy
        calendar, and its bound is covered by the unit tests in test_filters.
        """
        month = {e["entry_id"] for e in bridge.list_calendar_events(days=30)}
        week = {e["entry_id"] for e in bridge.list_calendar_events(days=7)}

        assert week <= month

    def test_create_appointment_basic(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a basic appointment"""
        subject = f"{TEST_PREFIX}Appt Test {test_timestamp}"

        start = from_now(hours=1)
        end = shifted(start, minutes=30)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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

        start = from_now(hours=2)
        end = shifted(start, hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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

        start = from_now(hours=3)
        end = shifted(start, minutes=45)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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
        start = now_wall_clock().date().add(days=1).at(Time())
        end = shifted(start, days=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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

        start = from_now(days=2)
        end = shifted(start, hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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
        start = from_now(days=3)
        end = shifted(start, hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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
        start = from_now(days=4)
        end = shifted(start, hours=1)

        entry_id = bridge.create_appointment(
            subject=subject,
            start=start.format(WALL_CLOCK_PATTERN),
            end=end.format(WALL_CLOCK_PATTERN),
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


@pytest.fixture
def daily_series(bridge, test_timestamp):
    """A [TEST] daily series of four 15-minute meetings starting tomorrow 06:00.

    Built over COM because create_appointment has no recurrence option.
    Yields (series_id, subject); deletes the whole series afterwards.
    """
    subject = f"{TEST_PREFIX}Series {test_timestamp}"
    start = now_wall_clock().date().add(days=1).at(Time(6))
    appt = bridge.get_calendar().Items.Add()
    appt.Subject = subject
    appt.Start = to_outlook(start)
    appt.End = to_outlook(shifted(start, minutes=15))
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = 0  # olRecurDaily
    pattern.Occurrences = OCCURRENCES
    appt.Save()
    series_id = appt.EntryID
    yield series_id, subject
    # A test may already have deleted the series, which moves it to Deleted
    # Items where a second Delete fails, so only delete it from the calendar.
    item = bridge.get_item_by_id(series_id)
    if item is not None and item.Parent.EntryID == bridge.get_calendar().EntryID:
        item.Delete()


def listed(bridge, subject_prefix):
    """This week's listed events whose subject starts with `subject_prefix`."""
    return [
        e
        for e in bridge.list_calendar_events(days=7)
        if e["subject"].startswith(subject_prefix)
    ]


@pytest.mark.integration
@pytest.mark.calendar
class TestRecurringOccurrences:
    """One occurrence of a recurring meeting is addressable on its own."""

    def test_each_occurrence_has_its_own_id(self, bridge, daily_series):
        series_id, subject = daily_series

        events = listed(bridge, subject)

        assert len(events) == OCCURRENCES
        assert len({e["entry_id"] for e in events}) == OCCURRENCES
        assert {e["series_id"] for e in events} == {series_id}

    def test_get_occurrence(self, bridge, daily_series):
        _, subject = daily_series
        second = listed(bridge, subject)[1]

        appt = bridge.get_appointment(second["entry_id"])

        assert appt["entry_id"] == second["entry_id"]
        assert appt["start"] == second["start"]

    def test_edit_changes_only_that_occurrence(self, bridge, daily_series):
        _, subject = daily_series
        second = listed(bridge, subject)[1]

        bridge.edit_appointment(second["entry_id"], subject=f"{subject} edited")

        subjects = [e["subject"] for e in listed(bridge, subject)]
        assert subjects.count(f"{subject} edited") == 1
        assert subjects.count(subject) == OCCURRENCES - 1

    def test_moved_occurrence_is_found_by_its_new_id(self, bridge, daily_series):
        _, subject = daily_series
        second = listed(bridge, subject)[1]
        moved_start = shifted(parse_wall_clock(second["start"]), hours=2)

        bridge.edit_appointment(
            second["entry_id"],
            start=moved_start.format(WALL_CLOCK_PATTERN),
            end=shifted(moved_start, minutes=15).format(WALL_CLOCK_PATTERN),
        )

        new_ids = {e["entry_id"] for e in listed(bridge, subject)}
        assert second["entry_id"] not in new_ids
        for entry_id in new_ids:
            assert bridge.get_appointment(entry_id) is not None

    def test_delete_removes_only_that_occurrence(self, bridge, daily_series):
        _, subject = daily_series
        second = listed(bridge, subject)[1]

        bridge.delete_appointment(second["entry_id"])

        remaining = listed(bridge, subject)
        assert len(remaining) == OCCURRENCES - 1
        assert second["entry_id"] not in {e["entry_id"] for e in remaining}
        assert bridge.get_appointment(second["entry_id"]) is None

    def test_series_id_needs_the_series_flag(self, bridge, daily_series):
        series_id, subject = daily_series

        with pytest.raises(MailtoolError, match="recurring series"):
            bridge.delete_appointment(series_id)
        assert len(listed(bridge, subject)) == OCCURRENCES

        bridge.delete_appointment(series_id, series=True)
        assert listed(bridge, subject) == []

    def test_series_flag_on_an_occurrence_id_changes_the_series(
        self, bridge, daily_series
    ):
        _, subject = daily_series
        second = listed(bridge, subject)[1]

        bridge.edit_appointment(
            second["entry_id"], subject=f"{subject} renamed", series=True
        )

        subjects = [e["subject"] for e in listed(bridge, f"{subject} renamed")]
        assert len(subjects) == OCCURRENCES


@pytest.mark.integration
@pytest.mark.calendar
class TestWallClockTimes:
    """A time given to mailtool is the time Outlook shows: local wall clock."""

    START = "07:00:00"
    END = "07:30:00"

    def tomorrow_at(self, wall_time):
        tomorrow = from_now(days=1).date().format(DATE_PATTERN)
        return f"{tomorrow} {wall_time}"

    def test_create_keeps_the_requested_time(
        self, bridge, test_timestamp, cleanup_helpers
    ):
        start, end = self.tomorrow_at(self.START), self.tomorrow_at(self.END)

        entry_id = bridge.create_appointment(
            subject=f"{TEST_PREFIX}Wall clock {test_timestamp}", start=start, end=end
        )

        appt = bridge.get_appointment(entry_id)
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)
        assert appt is not None
        assert (appt["start"], appt["end"]) == (start, end)

    def test_edit_keeps_the_requested_time(
        self, bridge, test_timestamp, cleanup_helpers
    ):
        entry_id = bridge.create_appointment(
            subject=f"{TEST_PREFIX}Wall clock edit {test_timestamp}",
            start=self.tomorrow_at("10:00:00"),
            end=self.tomorrow_at("10:30:00"),
        )
        start, end = self.tomorrow_at(self.START), self.tomorrow_at(self.END)

        bridge.edit_appointment(entry_id, start=start, end=end)

        appt = bridge.get_appointment(entry_id)
        cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)
        assert appt is not None
        assert (appt["start"], appt["end"]) == (start, end)
