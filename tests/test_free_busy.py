"""
Unit tests for free/busy: Outlook's slot string becomes busy blocks for the
requested days, start_date through end_date inclusive.

Outlook's Recipient.FreeBusy(start, minutes_per_char, complete_format) returns
one digit per slot from midnight of the start date, for a fixed span (29 days
on a live Exchange account), whatever end is wanted. With complete_format the
digits are 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
Pure-Python (fake COM objects).
"""

import pytest
from whenever import PlainDateTime

from mailtool.bridge import MailtoolError, free_busy_blocks, to_outlook

from .test_fail_loudly import FakeItem, make_bridge

pytestmark = pytest.mark.unit

SLOT_MINUTES = 30
SLOTS_PER_DAY = 24 * 60 // SLOT_MINUTES
EXCHANGE_DN = "/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=me"
MY_SMTP = "me@example.com"


def day_of(*busy):
    """One day of 30-minute slots, free except `busy` = ("HH:MM", digit) pairs."""
    slots = ["0"] * SLOTS_PER_DAY
    for wall_time, digit in busy:
        hour, minute = map(int, wall_time.split(":"))
        slots[(hour * 60 + minute) // SLOT_MINUTES] = digit
    return "".join(slots)


class RecordingRecipient:
    """A resolvable recipient that returns `slots` and records each FreeBusy call."""

    def __init__(self, slots):
        self.slots = slots
        self.asked = []
        self.address = None

    def created_for(self, address):
        self.address = address
        return self

    def Resolve(self):  # noqa: N802 - mirrors Outlook COM API
        return True

    def FreeBusy(self, start, minutes_per_char, complete_format=False):  # noqa: N802 - mirrors Outlook COM API
        self.asked.append((start, minutes_per_char, complete_format))
        return self.slots


class AttendeelessMeeting(FakeItem):
    RequiredAttendees = ""


class FakeExchangeUser:
    PrimarySmtpAddress = MY_SMTP


class FakePropertyAccessor:
    def GetProperty(self, _tag):  # noqa: N802 - mirrors Outlook COM API
        return MY_SMTP


class FakeAddressEntry:
    """An Exchange address entry; cached=True mimics cached Exchange mode,
    where GetExchangeUser() returns None."""

    def __init__(self, cached=False):
        self.cached = cached
        self.PropertyAccessor = FakePropertyAccessor()

    def GetExchangeUser(self):  # noqa: N802 - mirrors Outlook COM API
        return None if self.cached else FakeExchangeUser()


class FakeCurrentUser:
    Address = EXCHANGE_DN

    def __init__(self, cached=False):
        self.AddressEntry = FakeAddressEntry(cached)


def free_busy_bridge(slots, items=None, cached=False):
    bridge = make_bridge(items=items)
    recipient = RecordingRecipient(slots)
    bridge.namespace.CreateRecipient = lambda addr: recipient.created_for(addr)
    bridge.namespace.CurrentUser = FakeCurrentUser(cached)
    return bridge, recipient


class TestBlocks:
    def test_merges_runs_of_the_same_status(self):
        start = PlainDateTime(2026, 1, 5, 0, 0)

        blocks = free_busy_blocks("0022300", start, 30)

        assert blocks == [
            {
                "start": "2026-01-05 01:00:00",
                "end": "2026-01-05 02:00:00",
                "status": "Busy",
            },
            {
                "start": "2026-01-05 02:00:00",
                "end": "2026-01-05 02:30:00",
                "status": "Out of Office",
            },
        ]

    def test_all_statuses_are_labelled(self):
        blocks = free_busy_blocks("012345", PlainDateTime(2026, 1, 5, 0, 0), 60)

        assert [b["status"] for b in blocks] == [
            "Tentative",
            "Busy",
            "Out of Office",
            "Working Elsewhere",
            "Unknown (5)",
        ]

    def test_block_running_to_the_end_closes_at_the_window_end(self):
        blocks = free_busy_blocks("02", PlainDateTime(2026, 1, 5, 23, 0), 30)

        assert blocks == [
            {
                "start": "2026-01-05 23:30:00",
                "end": "2026-01-06 00:00:00",
                "status": "Busy",
            }
        ]

    def test_all_free_is_empty(self):
        assert free_busy_blocks("0000", PlainDateTime(2026, 1, 5, 0, 0), 30) == []


class TestGetFreeBusy:
    def test_asks_outlook_for_full_detail_from_local_midnight(self):
        bridge, recipient = free_busy_bridge(day_of() * 3)

        bridge.get_free_busy(email_address="a@example.com", start_date="2026-01-05")

        expected_start = to_outlook(PlainDateTime(2026, 1, 5, 0, 0))
        assert recipient.asked == [(expected_start, 30, True)]

    def test_end_date_is_inclusive(self):
        monday = day_of(("09:30", "2"))
        tuesday = day_of(("10:00", "3"))
        wednesday_outside_window = day_of(("00:30", "2"))
        bridge, _ = free_busy_bridge(monday + tuesday + wednesday_outside_window)

        result = bridge.get_free_busy(
            email_address="a@example.com",
            start_date="2026-01-05",
            end_date="2026-01-06",
        )

        assert result["busy"] == [
            {
                "start": "2026-01-05 09:30:00",
                "end": "2026-01-05 10:00:00",
                "status": "Busy",
            },
            {
                "start": "2026-01-06 10:00:00",
                "end": "2026-01-06 10:30:00",
                "status": "Out of Office",
            },
        ]
        assert result["end_date"] == "2026-01-06"
        assert result["interval_minutes"] == 30

    def test_default_window_is_the_start_day(self):
        next_day_outside_window = day_of(("00:00", "2"))
        bridge, _ = free_busy_bridge(day_of() + next_day_outside_window)

        result = bridge.get_free_busy(
            email_address="a@example.com", start_date="2026-01-05"
        )

        assert result["end_date"] == "2026-01-05"
        assert result["busy"] == []

    def test_interval_sets_the_slot_size(self):
        bridge, recipient = free_busy_bridge("0" * (24 * 60 // 15))

        result = bridge.get_free_busy(
            email_address="a@example.com", start_date="2026-01-05", interval_minutes=15
        )

        assert recipient.asked[0][1] == 15
        assert result["interval_minutes"] == 15

    def test_range_past_what_outlook_returns_raises(self):
        bridge, _ = free_busy_bridge(day_of() * 29)

        with pytest.raises(MailtoolError, match="29 days"):
            bridge.get_free_busy(
                email_address="a@example.com",
                start_date="2026-01-01",
                end_date="2026-01-30",
            )

    def test_last_day_outlook_covers_is_allowed(self):
        bridge, _ = free_busy_bridge(day_of() * 29)

        result = bridge.get_free_busy(
            email_address="a@example.com",
            start_date="2026-01-01",
            end_date="2026-01-29",
        )

        assert result["busy"] == []

    def test_end_before_start_raises(self):
        bridge, _ = free_busy_bridge(day_of())

        with pytest.raises(MailtoolError, match="before start_date"):
            bridge.get_free_busy(
                email_address="a@example.com",
                start_date="2026-01-05",
                end_date="2026-01-04",
            )

    def test_empty_answer_from_outlook_raises(self):
        bridge, _ = free_busy_bridge("")

        with pytest.raises(MailtoolError, match="no free/busy data"):
            bridge.get_free_busy(email_address="a@example.com")

    @pytest.mark.parametrize("interval", [0, -30, 7])
    def test_interval_must_divide_a_day(self, interval):
        bridge, _ = free_busy_bridge(day_of())

        with pytest.raises(MailtoolError, match="interval"):
            bridge.get_free_busy(
                email_address="a@example.com", interval_minutes=interval
            )

    def test_legacy_id_without_attendees_raises(self):
        """The deprecated entry_id path must not fall back to the current user."""
        item = AttendeelessMeeting()
        bridge, recipient = free_busy_bridge(day_of(), items={"id-1": item})

        with pytest.raises(MailtoolError, match="no required attendees"):
            bridge.get_free_busy(entry_id="id-1")

        assert recipient.asked == []


class TestDefaultAddress:
    @pytest.mark.parametrize("cached", [False, True], ids=["online", "cached"])
    def test_current_user_is_their_smtp_address(self, cached):
        """Not the Exchange DN that CurrentUser.Address holds."""
        bridge, recipient = free_busy_bridge(day_of(), cached=cached)

        result = bridge.get_free_busy()

        assert result["email"] == MY_SMTP
        assert recipient.address == MY_SMTP
