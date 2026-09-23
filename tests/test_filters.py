"""
Unit tests for the filter fixes in bridge.py:

- translate_filter: SQL-style LIKE filters are translated to DASL (@SQL=)
  with field mapping, [Unread] literal inversion, boolean normalization
- search_emails: DASL translation applied before Restrict + Restrict failures
  raise instead of silently returning []
- list_calendar_events: single combined Restrict, recurrence expansion kept,
  bounded horizon for all_events, hard iteration cap, overlap re-check

These tests are pure-Python (no Outlook / pywin32 required).
"""

from datetime import datetime, timedelta
from types import SimpleNamespace

import pytest

import mailtool.bridge as bridge_module
from mailtool.bridge import (
    CALENDAR_ALL_EVENTS_MAX_DAYS,
    CALENDAR_MAX_ITEMS,
    OutlookBridge,
    translate_filter,
)

pytestmark = pytest.mark.unit


# =============================================================================
# translate_filter
# =============================================================================


class TestTranslateFilter:
    def test_like_becomes_dasl(self):
        assert (
            translate_filter("[Subject] LIKE '%BMT%'")
            == "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%BMT%'"
        )

    def test_wildcard_patterns_kept_verbatim(self):
        # SQL LIKE semantics: contains / starts-with / ends-with all pass through
        assert (
            translate_filter("[Subject] LIKE 'BMT%'")
            == "@SQL=\"urn:schemas:httpmail:subject\" LIKE 'BMT%'"
        )
        assert (
            translate_filter("[Subject] LIKE '%verslag'")
            == "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%verslag'"
        )

    def test_no_like_left_unchanged(self):
        query = "[Unread] = TRUE AND [ReceivedTime] >= '05/01/2026 00:00'"
        assert translate_filter(query) == query

    def test_like_keyword_case_insensitive(self):
        # keyword case is preserved in the output; DASL keywords are
        # case-insensitive
        assert (
            translate_filter("[SenderName] like '%john%'")
            == "@SQL=\"urn:schemas:httpmail:sendername\" like '%john%'"
        )

    def test_unread_literal_inverted_for_read_property(self):
        result = translate_filter("[SenderName] LIKE '%John%' AND [Unread] = TRUE")
        assert result == (
            "@SQL=\"urn:schemas:httpmail:sendername\" LIKE '%John%' "
            'AND "urn:schemas:httpmail:read" = 0'
        )

    def test_unread_false_becomes_read_one(self):
        result = translate_filter("[Subject] LIKE '%x%' AND [Unread] = FALSE")
        assert '"urn:schemas:httpmail:read" = 1' in result

    def test_other_boolean_literals_normalized(self):
        result = translate_filter("[Subject] LIKE '%x%' AND [HasAttachments] = TRUE")
        assert '"urn:schemas:httpmail:hasattachment" = 1' in result

    def test_date_and_sender_fields_mapped(self):
        result = translate_filter(
            "[SenderEmailAddress] LIKE '%utwente%' AND "
            "[ReceivedTime] >= '08/01/2026 00:00'"
        )
        assert result.startswith("@SQL=")
        # Proptag, not urn:schemas:httpmail:senderemail — see
        # test_senderemailaddress_mapped_via_proptag for why.
        assert (
            '"http://schemas.microsoft.com/mapi/proptag/0x0C1F001F" LIKE'
            in result
        )
        assert '"urn:schemas:httpmail:date" >= ' in result

    def test_senderemailaddress_mapped_via_proptag(self):
        # Mirrors test_messageclass_mapped_via_proptag: the httpmail schema
        # property for sender email is unreliably populated, so this field
        # is mapped via the MAPI proptag for PR_SENDER_EMAIL_ADDRESS instead.
        result = translate_filter("[SenderEmailAddress] LIKE '%databricks.com%'")
        assert result == (
            '@SQL="http://schemas.microsoft.com/mapi/proptag/0x0C1F001F" '
            "LIKE '%databricks.com%'"
        )

    def test_unknown_field_raises_value_error(self):
        with pytest.raises(ValueError, match="Cannot translate"):
            translate_filter("[NoSuchField] LIKE '%x%'")

    def test_messageclass_mapped_via_proptag(self):
        result = translate_filter("[MessageClass] LIKE '%IPM.Note%'")
        assert result == (
            '@SQL="http://schemas.microsoft.com/mapi/proptag/0x001A001F" '
            "LIKE '%IPM.Note%'"
        )

    def test_bracketed_text_inside_like_literal_not_treated_as_field(self):
        # '%[External]%' is a literal LIKE search pattern, not a [Field]
        # reference — must not raise and must survive untouched.
        result = translate_filter("[Subject] LIKE '%[External]%'")
        assert result == (
            '@SQL="urn:schemas:httpmail:subject" LIKE \'%[External]%\''
        )

    def test_true_false_text_inside_like_literal_not_mangled(self):
        # '%TRUE%' is literal text to search for, not the TRUE keyword —
        # must not be rewritten to '%1%'.
        result = translate_filter("[Subject] LIKE '%TRUE%'")
        assert result == '@SQL="urn:schemas:httpmail:subject" LIKE \'%TRUE%\''

        result = translate_filter("[Subject] LIKE '%FALSE%'")
        assert result == '@SQL="urn:schemas:httpmail:subject" LIKE \'%FALSE%\''

    def test_literal_masking_does_not_block_real_unread_substitution(self):
        # A literal containing bracketed text must be preserved verbatim
        # while a real [Unread] = TRUE clause elsewhere in the same filter
        # still gets translated and inverted normally.
        result = translate_filter(
            "[Subject] LIKE '%[External]%' AND [Unread] = TRUE"
        )
        assert result == (
            '@SQL="urn:schemas:httpmail:subject" LIKE \'%[External]%\' '
            'AND "urn:schemas:httpmail:read" = 0'
        )


# =============================================================================
# Fakes
# =============================================================================


class FakeItems:
    """Minimal stand-in for an Outlook Items collection."""

    def __init__(self, items=None, restrict_error=None):
        self._items = list(items or [])
        self.restrict_error = restrict_error
        self.restrict_calls = []
        self.sort_calls = []
        self.include_recurrences = None

    @property
    def IncludeRecurrences(self):  # noqa: N802 - mirrors Outlook COM API
        return self.include_recurrences

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):  # noqa: N802 - mirrors Outlook COM API
        self.include_recurrences = value

    def Restrict(self, filter_str):  # noqa: N802 - mirrors Outlook COM API
        self.restrict_calls.append(filter_str)
        if self.restrict_error is not None:
            raise self.restrict_error
        return self

    def Sort(self, prop, descending=False):  # noqa: N802 - mirrors Outlook COM API
        self.sort_calls.append((prop, descending))

    def __iter__(self):
        return iter(self._items)


class FakeFolder:
    def __init__(self, items):
        self.Items = items


class FakeCalendar:
    def __init__(self, items):
        self.Items = items


def make_appointment(start, end, subject="Test Appointment"):
    """Fake appointment item exposing the attributes _safe_get_attr reads."""
    return SimpleNamespace(
        EntryID=f"apt-{subject}-{start:%Y%m%d%H%M}",
        Subject=subject,
        Start=start,
        End=end,
        Location="",
        Organizer="Organizer",
        AllDayEvent=False,
        RequiredAttendees="",
        OptionalAttendees="",
        ResponseStatus=3,
        MeetingStatus=1,
        ResponseRequested=False,
    )


def make_bridge_with_inbox(items):
    bridge = OutlookBridge.__new__(OutlookBridge)  # skip COM __init__
    bridge.get_inbox = lambda: FakeFolder(items)
    return bridge, items


def make_bridge_with_calendar(items):
    bridge = OutlookBridge.__new__(OutlookBridge)  # skip COM __init__
    bridge.get_calendar = lambda: FakeCalendar(items)
    return bridge, items


# =============================================================================
# search_emails
# =============================================================================


class TestSearchEmailsFilterHandling:
    def test_like_translated_to_dasl_before_restrict(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        bridge.search_emails("[Subject] LIKE '%test%'", limit=5)

        assert len(items.restrict_calls) == 1
        applied = items.restrict_calls[0]
        assert applied.startswith("@SQL=")
        assert "\"urn:schemas:httpmail:subject\" LIKE '%test%'" in applied
        # DASL mail-only scoping (MessageClass proptag) ANDed in
        assert "0x001A001F" in applied
        assert bridge_module.MAIL_ONLY_FILTER not in applied

    def test_like_with_unread_inverts_read_literal(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        bridge.search_emails("[Unread] = TRUE AND [Subject] LIKE '%x%'", limit=5)

        applied = items.restrict_calls[0]
        assert '"urn:schemas:httpmail:read" = 0' in applied

    def test_like_unknown_field_raises_value_error(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        with pytest.raises(ValueError, match="Cannot translate"):
            bridge.search_emails("[NoSuchField] LIKE '%x%'", limit=5)
        assert items.restrict_calls == []

    def test_invalid_filter_raises_runtime_error(self):
        items = FakeItems(restrict_error=ValueError("Condition is not valid"))
        bridge, _ = make_bridge_with_inbox(items)

        with pytest.raises(RuntimeError, match="Restrict failed"):
            bridge.search_emails("[Subject] = 'x'")

    def test_empty_query_uses_mail_only_filter(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        bridge.search_emails("", limit=5)

        assert items.restrict_calls == [bridge_module.MAIL_ONLY_FILTER]

    def test_messageclass_in_query_not_double_scoped(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        bridge.search_emails("[MessageClass] = 'IPM.Note'", limit=5)

        assert items.restrict_calls == ["[MessageClass] = 'IPM.Note'"]

    def test_messageclass_like_query_not_double_scoped(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_inbox(items)

        bridge.search_emails("[MessageClass] LIKE '%IPM.Note%'", limit=5)

        applied = items.restrict_calls[0]
        assert applied.startswith("@SQL=")
        assert "0x001A001F" in applied
        # mail-only scope must NOT be ANDed a second time
        assert "IPM.Note{" not in applied

    def test_single_combined_restrict_and_recurrence_kept(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_calendar(items)

        bridge.list_calendar_events(days=7)

        assert len(items.restrict_calls) == 1
        assert items.include_recurrences is True
        # Sort must happen on the original collection (ascending) before Restrict
        assert items.sort_calls == [("[Start]", False)]
        applied = items.restrict_calls[0]
        assert applied.startswith("@SQL=")
        assert "0x001A001F" in applied  # appointments-only MessageClass scope
        assert '"urn:schemas:calendar:dtstart" <=' in applied
        assert '"urn:schemas:calendar:dtend" >=' in applied

    def test_days_window_in_filter(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_calendar(items)
        before = datetime.now()

        bridge.list_calendar_events(days=14)

        applied = items.restrict_calls[0]
        expected_end = (before + timedelta(days=14)).strftime("%Y-%m-%d")
        assert f"dtstart\" <= '{expected_end}" in applied

    def test_days_clamped_to_horizon(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_calendar(items)
        before = datetime.now()

        bridge.list_calendar_events(days=99999)

        applied = items.restrict_calls[0]
        expected_end = before + timedelta(days=CALENDAR_ALL_EVENTS_MAX_DAYS)
        assert f"dtstart\" <= '{expected_end.strftime('%Y-%m-%d')}" in applied

    def test_all_events_uses_bounded_horizon(self):
        items = FakeItems()
        bridge, _ = make_bridge_with_calendar(items)
        before = datetime.now()

        bridge.list_calendar_events(all_events=True)

        applied = items.restrict_calls[0]
        expected_end = before + timedelta(days=CALENDAR_ALL_EVENTS_MAX_DAYS)
        assert f"dtstart\" <= '{expected_end.strftime('%Y-%m-%d')}" in applied

    def test_overlapping_events_returned_outside_filtered(self):
        now = datetime.now()
        in_window = make_appointment(
            now + timedelta(days=1), now + timedelta(days=1, hours=1), "In Window"
        )
        past = make_appointment(
            now - timedelta(days=10), now - timedelta(days=10, hours=-1), "Past"
        )
        future = make_appointment(
            now + timedelta(days=100), now + timedelta(days=100, hours=1), "Far Future"
        )
        # Overlapping event that STARTS before the window but ends inside it
        overlapping = make_appointment(
            now - timedelta(hours=2), now + timedelta(hours=2), "Overlap"
        )
        items = FakeItems([in_window, past, future, overlapping])
        bridge, _ = make_bridge_with_calendar(items)

        events = bridge.list_calendar_events(days=7)

        subjects = [e["subject"] for e in events]
        assert "In Window" in subjects
        assert "Overlap" in subjects
        assert "Past" not in subjects
        assert "Far Future" not in subjects

    def test_iteration_cap_respected(self, monkeypatch):
        now = datetime.now()
        many = [
            make_appointment(
                now + timedelta(days=1, minutes=i * 30),
                now + timedelta(days=1, minutes=i * 30 + 15),
                f"Capped {i}",
            )
            for i in range(10)
        ]
        items = FakeItems(many)
        bridge, _ = make_bridge_with_calendar(items)
        monkeypatch.setattr(bridge_module, "CALENDAR_MAX_ITEMS", 3)

        events = bridge.list_calendar_events(days=7)

        assert len(events) <= 3


# =============================================================================
# _to_naive_datetime / _calendar_range_filter
# =============================================================================


class TestHelpers:
    def test_to_naive_datetime_with_python_datetime(self):
        dt = datetime(2026, 8, 18, 12, 30, 45)
        assert OutlookBridge._to_naive_datetime(dt) == dt

    def test_to_naive_datetime_with_com_style_object(self):
        com_dt = SimpleNamespace(
            Year=2026, Month=8, Day=18, Hour=12, Minute=30, Second=45
        )
        assert OutlookBridge._to_naive_datetime(com_dt) == datetime(
            2026, 8, 18, 12, 30, 45
        )

    def test_to_naive_datetime_garbage_returns_none(self):
        assert OutlookBridge._to_naive_datetime("not a date") is None
        assert OutlookBridge._to_naive_datetime(None) is None

    def test_calendar_range_filter_format(self):
        start = datetime(2026, 8, 18, 0, 0)
        end = datetime(2026, 9, 1, 0, 0)
        flt = OutlookBridge._calendar_range_filter(start, end)
        assert flt == (
            '@SQL=("http://schemas.microsoft.com/mapi/proptag/0x001A001F" '
            ">= 'IPM.Appointment' "
            'AND "http://schemas.microsoft.com/mapi/proptag/0x001A001F" '
            "< 'IPM.Appointment{') "
            "AND (\"urn:schemas:calendar:dtstart\" <= '2026-09-01 00:00') "
            "AND (\"urn:schemas:calendar:dtend\" >= '2026-08-18 00:00')"
        )

    def test_calendar_constants_sane(self):
        assert CALENDAR_ALL_EVENTS_MAX_DAYS == 365
        assert CALENDAR_MAX_ITEMS >= 100
