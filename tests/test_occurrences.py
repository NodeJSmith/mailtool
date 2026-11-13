"""
Unit tests for occurrence ids: one occurrence of a recurring meeting is
addressable on its own, and changing a whole series takes an explicit
series=True.

Outlook gives every occurrence the series' EntryID, so mailtool identifies an
occurrence as "<series EntryID>@<start>". Pure-Python (fake COM objects).
"""

import pytest
from whenever import PlainDateTime, Time

from mailtool.bridge import (
    OL_APPOINTMENT,
    OL_APPT_EXCEPTION,
    OL_APPT_MASTER,
    OL_APPT_NOT_RECURRING,
    OL_APPT_OCCURRENCE,
    MailtoolError,
    now_wall_clock,
    occurrence_id,
    split_appointment_id,
    to_outlook,
)

from .test_fail_loudly import FakeComError, FakeFolder, ListItems, make_bridge

pytestmark = pytest.mark.unit

SERIES_ID = "00000000ABCDEF"
# Tomorrow and the day after at 09:00 local wall clock, inside the listing's
# 7-day window. The *_COM values are what pywin32 hands back for those times.
MON = now_wall_clock().date().add(days=1).at(Time(9))
TUE = MON.add(days=1)
MON_COM = to_outlook(MON)
TUE_COM = to_outlook(TUE)


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

    def __init__(self, start, state, subject="standup"):
        self.EntryID = SERIES_ID
        self.Start = start
        self.End = start
        self.Subject = subject
        self.RecurrenceState = state
        self.IsRecurring = state != OL_APPT_NOT_RECURRING
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
        super().__init__(MON_COM, state=OL_APPT_MASTER)
        self.mon = FakeAppointment(MON_COM, state=OL_APPT_OCCURRENCE)
        self.tue = FakeAppointment(TUE_COM, state=OL_APPT_EXCEPTION, subject="moved")
        self.pattern = FakePattern([self.mon, self.tue])

    def GetRecurrencePattern(self):  # noqa: N802 - mirrors Outlook COM API
        return self.pattern


def series_bridge():
    series = FakeSeries()
    return make_bridge(items={SERIES_ID: series}), series


class TestOccurrenceIdFormat:
    def test_round_trip(self):
        start = PlainDateTime(2026, 11, 16, 9, 0)
        oid = occurrence_id(SERIES_ID, start)

        assert oid == f"{SERIES_ID}@20261116T090000"
        assert split_appointment_id(oid) == (SERIES_ID, start)

    def test_plain_entry_id_has_no_start(self):
        assert split_appointment_id(SERIES_ID) == (SERIES_ID, None)

    def test_unparsable_suffix_is_treated_as_a_plain_id(self):
        """It then fails the EntryID lookup like any other bad id."""
        assert split_appointment_id("abc@nope") == ("abc@nope", None)


class TestListing:
    def test_occurrences_get_their_own_ids_and_share_a_series_id(self):
        series = FakeSeries()
        single = FakeAppointment(
            MON_COM, state=OL_APPT_NOT_RECURRING, subject="one-off"
        )
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


class TestLookup:
    def test_get_appointment_by_occurrence_id_reads_that_occurrence(self):
        bridge, _ = series_bridge()

        appt = bridge.get_appointment(occurrence_id(SERIES_ID, TUE))

        assert appt is not None

        assert appt["subject"] == "moved"
        assert appt["entry_id"] == occurrence_id(SERIES_ID, TUE)
        assert appt["series_id"] == SERIES_ID

    def test_occurrence_is_looked_up_by_wall_clock_start_labelled_utc(self):
        """Pins only the argument shape. That pywin32 then hands Outlook the
        right wall-clock time is proven by the live tests in test_calendar."""
        bridge, series = series_bridge()

        bridge.get_appointment(occurrence_id(SERIES_ID, MON))

        assert series.pattern.asked == [MON_COM]

    def test_get_appointment_by_series_id_reads_the_series(self):
        bridge, _ = series_bridge()

        appt = bridge.get_appointment(SERIES_ID)

        assert appt is not None

        assert appt["entry_id"] == SERIES_ID
        assert appt["series_id"] == SERIES_ID

    def test_missing_occurrence_is_none(self):
        bridge, _ = series_bridge()

        later = MON.add(days=5)

        assert bridge.get_appointment(occurrence_id(SERIES_ID, later)) is None

    def test_occurrence_id_on_a_non_recurring_item_is_none(self):
        single = FakeAppointment(MON_COM, state=OL_APPT_NOT_RECURRING)
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
        single = FakeAppointment(MON_COM, state=OL_APPT_NOT_RECURRING)
        bridge = make_bridge(items={SERIES_ID: single})

        assert bridge.delete_appointment(SERIES_ID) is True
        assert single.calls == ["Delete"]
