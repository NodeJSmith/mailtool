"""
Unit tests for free/busy: Outlook's slot string becomes busy blocks for the
requested [start_date, end_date) window.

Outlook's Recipient.FreeBusy(start, minutes_per_char, complete_format) returns
one digit per slot from midnight of the start date, for a fixed span (29 days
on a live Exchange account), whatever end is wanted. With complete_format the
digits are 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
Pure-Python (fake COM objects).
"""

import pytest
from whenever import PlainDateTime

from mailtool.bridge import MailtoolError, free_busy_blocks, to_outlook

from .test_fail_loudly import make_bridge

pytestmark = pytest.mark.unit

SLOTS_PER_DAY = 48  # at 30 minutes


def day_of(*busy):
    """One day of 30-minute slots, free except `busy` = (slot index, digit)."""
    slots = ["0"] * SLOTS_PER_DAY
    for index, digit in busy:
        slots[index] = digit
    return "".join(slots)


class FakeRecipient:
    def __init__(self, slots):
        self.slots = slots
        self.asked = []

    def Resolve(self):  # noqa: N802 - mirrors Outlook COM API
        return True

    def FreeBusy(self, start, minutes_per_char, complete_format=False):  # noqa: N802 - mirrors Outlook COM API
        self.asked.append((start, minutes_per_char, complete_format))
        return self.slots


def free_busy_bridge(slots):
    bridge = make_bridge()
    recipient = FakeRecipient(slots)
    bridge.namespace.CreateRecipient = lambda _addr: recipient
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
        blocks = free_busy_blocks("01234", PlainDateTime(2026, 1, 5, 0, 0), 60)

        assert [b["status"] for b in blocks] == [
            "Tentative",
            "Busy",
            "Out of Office",
            "Working Elsewhere",
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

        assert recipient.asked == [
            (to_outlook(PlainDateTime(2026, 1, 5, 0, 0)), 30, True)
        ]

    def test_only_the_requested_days_are_returned(self):
        day1 = day_of((19, "2"))
        day2 = day_of((20, "3"))
        bridge, _ = free_busy_bridge(day1 + day2 + day_of((1, "2")))

        result = bridge.get_free_busy(
            email_address="a@example.com",
            start_date="2026-01-05",
            end_date="2026-01-07",
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
        assert result["interval_minutes"] == 30
        assert "free_busy" not in result

    def test_default_window_is_one_day(self):
        bridge, _ = free_busy_bridge(day_of() + day_of((0, "2")))

        result = bridge.get_free_busy(
            email_address="a@example.com", start_date="2026-01-05"
        )

        assert result["end_date"] == "2026-01-06"
        assert result["busy"] == []

    def test_interval_sets_the_slot_size(self):
        bridge, recipient = free_busy_bridge("0" * 96)

        result = bridge.get_free_busy(
            email_address="a@example.com", start_date="2026-01-05", interval_minutes=15
        )

        assert recipient.asked[0][1] == 15
        assert result["interval_minutes"] == 15

    def test_end_past_what_outlook_returns_raises(self):
        bridge, _ = free_busy_bridge(day_of() * 29)

        with pytest.raises(MailtoolError, match="29 days"):
            bridge.get_free_busy(
                email_address="a@example.com",
                start_date="2026-01-01",
                end_date="2026-02-01",
            )

    def test_end_not_after_start_raises(self):
        bridge, _ = free_busy_bridge(day_of())

        with pytest.raises(MailtoolError, match="after"):
            bridge.get_free_busy(
                email_address="a@example.com",
                start_date="2026-01-05",
                end_date="2026-01-05",
            )

    @pytest.mark.parametrize("interval", [0, -30, 7])
    def test_interval_must_divide_a_day(self, interval):
        bridge, _ = free_busy_bridge(day_of())

        with pytest.raises(MailtoolError, match="interval"):
            bridge.get_free_busy(
                email_address="a@example.com", interval_minutes=interval
            )
