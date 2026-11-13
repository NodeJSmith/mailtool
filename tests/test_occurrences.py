"""
Unit tests for occurrence ids: one occurrence of a recurring meeting is
addressable on its own, and changing a whole series takes an explicit
series=True.

Outlook gives every occurrence the series' EntryID, so mailtool identifies an
occurrence as "<series EntryID>@<start>". Pure-Python (fake COM objects).
"""

from datetime import UTC, datetime

import pytest

from mailtool.bridge import (
    OL_APPOINTMENT,
    MailtoolError,
    occurrence_id,
    split_appointment_id,
)

from .test_fail_loudly import FakeComError, FakeFolder, ListItems, make_bridge

pytestmark = pytest.mark.unit

SERIES_ID = "00000000ABCDEF"
MON = datetime(2026, 11, 16, 9, 0, tzinfo=UTC)
TUE = datetime(2026, 11, 17, 9, 0, tzinfo=UTC)


class FakeMissingOccurrenceError(FakeComError):
    """What GetOccurrence raises when no occurrence starts at that time:
    DISP_E_EXCEPTION wrapping E_FAIL (verified against live Outlook)."""

    hresult = -2147352567
    excepinfo = (0, "Microsoft Outlook", "no occurrence", None, 0, -2147467259)


class FakeAppointment:
    Class = OL_APPOINTMENT
    Location = ""
    Organizer = "me"
    Body = ""
    AllDayEvent = False
    RequiredAttendees = ""
    OptionalAttendees = ""
    ResponseStatus = 0
    MeetingStatus = 0
    ResponseRequested = False

    def __init__(self, start, state, parent=None, subject="standup"):
        self.EntryID = SERIES_ID
        self.Start = start
        self.End = start
        self.Subject = subject
        self.RecurrenceState = state
        self.IsRecurring = state != 0
        self.Parent = parent
        self.calls = []

    def Save(self):  # noqa: N802 - mirrors Outlook COM API
        self.calls.append("Save")

    def Delete(self):  # noqa: N802 - mirrors Outlook COM API
        self.calls.append("Delete")

    def Response(self, code):  # noqa: N802 - mirrors Outlook COM API
        self.calls.append(("Response", code))

    def Send(self):  # noqa: N802 - mirrors Outlook COM API
        self.calls.append("Send")


class FakePattern:
    def __init__(self, occurrences):
        self.occurrences = occurrences
        self.asked = []

    def GetOccurrence(self, start):  # noqa: N802 - mirrors Outlook COM API
        self.asked.append(start)
        for occ in self.occurrences:
            if occ.Start == start:
                return occ
        raise FakeMissingOccurrenceError("no occurrence")


class FakeSeries(FakeAppointment):
    def __init__(self):
        super().__init__(MON, state=1)
        self.mon = FakeAppointment(MON, state=2, parent=self)
        self.tue = FakeAppointment(TUE, state=3, parent=self, subject="moved")
        self.pattern = FakePattern([self.mon, self.tue])

    def GetRecurrencePattern(self):  # noqa: N802 - mirrors Outlook COM API
        return self.pattern


def series_bridge():
    series = FakeSeries()
    return make_bridge(items={SERIES_ID: series}), series


class TestOccurrenceIdFormat:
    def test_round_trip(self):
        oid = occurrence_id(SERIES_ID, MON)

        assert oid == f"{SERIES_ID}@20261116T090000"
        assert split_appointment_id(oid) == (SERIES_ID, datetime(2026, 11, 16, 9, 0))

    def test_plain_entry_id_has_no_start(self):
        assert split_appointment_id(SERIES_ID) == (SERIES_ID, None)

    def test_unparsable_suffix_is_treated_as_a_plain_id(self):
        """It then fails the EntryID lookup like any other bad id."""
        assert split_appointment_id("abc@nope") == ("abc@nope", None)


class TestListing:
    def test_occurrences_get_their_own_ids_and_share_a_series_id(self):
        series = FakeSeries()
        single = FakeAppointment(MON, state=0, subject="one-off")
        single.EntryID = "single-id"
        bridge = make_bridge()

        events = bridge.collect_calendar_events(
            FakeFolder(ListItems([series.mon, series.tue, single])),
            days=7,
            all_events=False,
        )

        by_subject = {e["subject"]: e for e in events}
        assert by_subject["standup"]["entry_id"] == occurrence_id(SERIES_ID, MON)
        assert by_subject["moved"]["entry_id"] == occurrence_id(SERIES_ID, TUE)
        assert by_subject["standup"]["series_id"] == SERIES_ID
        assert by_subject["moved"]["series_id"] == SERIES_ID
        assert by_subject["one-off"]["entry_id"] == "single-id"
        assert by_subject["one-off"]["series_id"] is None


@pytest.fixture(autouse=True)
def window_covers_fakes(monkeypatch):
    """Pin "now" so the listing's overlap re-check keeps the fake dates."""

    class FixedDatetime(datetime):
        @classmethod
        def now(cls, tz=None):
            return datetime(2026, 11, 15, 0, 0)

    monkeypatch.setattr("mailtool.bridge.datetime", FixedDatetime)


class TestLookup:
    def test_get_appointment_by_occurrence_id_reads_that_occurrence(self):
        bridge, _ = series_bridge()

        appt = bridge.get_appointment(occurrence_id(SERIES_ID, TUE))

        assert appt["subject"] == "moved"
        assert appt["entry_id"] == occurrence_id(SERIES_ID, TUE)
        assert appt["series_id"] == SERIES_ID

    def test_occurrence_is_looked_up_by_wall_clock_start_labelled_utc(self):
        """pywin32 converts a naive datetime from local time; labelling the
        wall-clock start as UTC passes it to Outlook unchanged."""
        bridge, series = series_bridge()

        bridge.get_appointment(occurrence_id(SERIES_ID, MON))

        assert series.pattern.asked == [MON]

    def test_get_appointment_by_series_id_reads_the_series(self):
        bridge, _ = series_bridge()

        appt = bridge.get_appointment(SERIES_ID)

        assert appt["entry_id"] == SERIES_ID
        assert appt["series_id"] == SERIES_ID

    def test_missing_occurrence_is_none(self):
        bridge, _ = series_bridge()

        assert bridge.get_appointment(f"{SERIES_ID}@20261118T090000") is None

    def test_occurrence_id_on_a_non_recurring_item_is_none(self):
        single = FakeAppointment(MON, state=0)
        bridge = make_bridge(items={SERIES_ID: single})

        assert bridge.get_appointment(occurrence_id(SERIES_ID, MON)) is None


class TestMutatorsTargetOneOccurrence:
    @pytest.mark.parametrize(
        ("method", "extra", "expected"),
        [
            ("delete_appointment", (), "Delete"),
            ("edit_appointment", (), "Save"),
            ("respond_to_meeting", ("accept",), "Send"),
        ],
    )
    def test_occurrence_id_changes_only_that_occurrence(self, method, extra, expected):
        bridge, series = series_bridge()

        getattr(bridge, method)(occurrence_id(SERIES_ID, TUE), *extra)

        assert expected in series.tue.calls
        assert series.calls == []
        assert series.mon.calls == []

    @pytest.mark.parametrize(
        ("method", "extra"),
        [
            ("delete_appointment", ()),
            ("edit_appointment", ()),
            ("respond_to_meeting", ("accept",)),
        ],
    )
    def test_series_id_without_series_flag_raises(self, method, extra):
        bridge, series = series_bridge()

        with pytest.raises(MailtoolError, match="recurring series"):
            getattr(bridge, method)(SERIES_ID, *extra)

        assert series.calls == []

    @pytest.mark.parametrize("target", [SERIES_ID, occurrence_id(SERIES_ID, TUE)])
    def test_series_flag_changes_the_whole_series(self, target):
        """series=True acts on the series from either kind of id."""
        bridge, series = series_bridge()

        bridge.delete_appointment(target, series=True)

        assert series.calls == ["Delete"]
        assert series.tue.calls == []

    def test_non_recurring_appointment_needs_no_flag(self):
        single = FakeAppointment(MON, state=0)
        bridge = make_bridge(items={SERIES_ID: single})

        assert bridge.delete_appointment(SERIES_ID) is True
        assert single.calls == ["Delete"]
