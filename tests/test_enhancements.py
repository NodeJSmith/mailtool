"""Unit tests for the inbox-triage enhancements.

These tests exercise the pure-Python logic added in the v0.11.0 triage work
(MessageClass filtering, sender SMTP salvage, body cleaning, and richer
metadata). They do NOT require Outlook or pywin32 to be running; the bridge module imports defensively so its static helpers and module
constants can be tested on any platform.
"""

from datetime import datetime

import pytest

from mailtool.bridge import _SMTP_REGEX, MAIL_ONLY_FILTER, OutlookBridge

# =============================================================================
# Body cleaning (Phase 5)
# =============================================================================


@pytest.mark.unit
class TestCleanBodyTop:
    def test_plain_body_returned_unchanged(self):
        body = "Hello,\n\nThis is a new message.\n\nCheers."
        assert OutlookBridge._clean_body_top(body) == body

    def test_strips_outlook_original_message_block(self):
        body = "My reply here.\n\n-----Original Message-----\nFrom: x@y.com\nOld"
        assert OutlookBridge._clean_body_top(body) == "My reply here."

    def test_strips_dutch_doorgestuurd_block(self):
        body = "Mijn antwoord.\n\n----- doorgestuurd bericht -----\nVan: x@y.com"
        assert OutlookBridge._clean_body_top(body) == "Mijn antwoord."

    def test_strips_origineel_bericht_block(self):
        body = "Antwoord.\n\n-----origineel bericht-----\nVan: x@y.com"
        assert OutlookBridge._clean_body_top(body) == "Antwoord."

    def test_strips_quoted_lines(self):
        body = "Top message\n> quoted line\n> more quoted"
        assert OutlookBridge._clean_body_top(body) == "Top message"

    def test_strips_signature_separator(self):
        body = "Hello\n\n_____\n\nsignature block"
        assert OutlookBridge._clean_body_top(body) == "Hello"

    def test_strips_from_header_with_email(self):
        body = "Reply text\n\nFrom: Stella <s.a.kruit@utwente.nl>\nSent: ..."
        assert OutlookBridge._clean_body_top(body) == "Reply text"

    def test_strips_wrote_footer(self):
        body = "Thanks!\n\nOn 1 July 2026 John wrote:\n> old"
        assert OutlookBridge._clean_body_top(body) == "Thanks!"

    def test_empty_body(self):
        assert OutlookBridge._clean_body_top("") == ""
        assert OutlookBridge._clean_body_top(None) == ""

    def test_collapses_blank_runs(self):
        assert OutlookBridge._clean_body_top("a\n\n\n\n\nb") == "a\n\nb"

    def test_caps_length(self):
        long_body = "x" * 5000
        assert len(OutlookBridge._clean_body_top(long_body, max_chars=100)) == 100

    def test_does_not_truncate_normal_to_line(self):
        # "To whom it may concern" must not be mistaken for a reply header.
        body = "To whom it may concern,\n\nHello."
        assert OutlookBridge._clean_body_top(body) == body


# =============================================================================
# Sender SMTP salvage regex (Phase 2)
# =============================================================================


@pytest.mark.unit
class TestSmtpRegex:
    def test_extracts_from_angle_brackets(self):
        m = _SMTP_REGEX.search("Kruit <stella.kruit@utwente.nl>")
        assert m is not None
        assert m.group(0) == "stella.kruit@utwente.nl"

    def test_extracts_embedded_smtp(self):
        assert _SMTP_REGEX.search("SMTP:s.a.kruit@utwente.nl").group(0) == (
            "s.a.kruit@utwente.nl"
        )

    def test_no_false_positive_on_exchange_dn(self):
        # A bare EX DN has no SMTP-shaped token; salvage must not invent one.
        dn = "/o=ExchangeLabs/ou=Group/cn=Recipients/cn=stella-kruit"
        assert _SMTP_REGEX.search(dn) is None

    def test_no_match_plain_text(self):
        assert _SMTP_REGEX.search("just some words") is None


# =============================================================================
# MessageClass filter constant (Phase 1)
# =============================================================================


@pytest.mark.unit
class TestMailOnlyFilter:
    def test_targets_ipm_note_range(self):
        assert "[MessageClass]" in MAIL_ONLY_FILTER
        assert "IPM.Note" in MAIL_ONLY_FILTER
        assert ">=" in MAIL_ONLY_FILTER
        assert "<" in MAIL_ONLY_FILTER

    def test_filter_well_formed(self):
        # Balanced single quotes and bracketed field references.
        assert MAIL_ONLY_FILTER.count("'") % 2 == 0
        assert MAIL_ONLY_FILTER.count("[") == MAIL_ONLY_FILTER.count("]")


class _FakeAttachment:
    def __init__(self, filename, size):
        self.FileName = filename
        self.Size = size
        self.DisplayName = filename
        self.ContentType = "application/pdf"
        self.IsInline = False


class _FakeAttachments:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):  # noqa: N802 - mirrors COM Attachments.Count (PascalCase)
        return len(self._items)

    def Item(self, i):  # noqa: N802 - mirrors COM Attachments.Item(i)
        return self._items[i - 1]


class _FakeMailItem:
    MessageClass = "IPM.Note"
    EntryID = "eid-1"
    Subject = "Hello"
    SenderName = "Alice"
    SenderEmailType = "SMTP"
    SenderEmailAddress = "alice@example.com"
    To = "bob@example.com"
    CC = "carol@example.com"
    BCC = ""
    Unread = True
    ReceivedTime = datetime(2026, 7, 7, 10, 0, 0)
    SentOn = datetime(2026, 7, 7, 9, 55, 0)
    ConversationID = "conv-1"
    ConversationTopic = "Hello"
    Body = "Body text\n\n-----Original Message-----\nFrom: old@x.com"
    HTMLBody = "<p>Body text</p>"

    def __init__(self):
        self.Attachments = _FakeAttachments([_FakeAttachment("a.pdf", 1234)])


@pytest.mark.unit
class TestMailItemToDict:
    def test_builds_full_dict_for_fake_mail_item(self):
        # __new__ skips __init__ so we don't touch COM / pywin32.
        bridge = OutlookBridge.__new__(OutlookBridge)
        d = bridge._mail_item_to_dict(_FakeMailItem(), include_body=True)

        assert d["entry_id"] == "eid-1"
        assert d["message_class"] == "IPM.Note"
        assert d["sender"] == "alice@example.com"
        assert d["to"] == "bob@example.com"
        assert d["cc"] == "carol@example.com"
        assert d["received_time"] == "2026-07-07 10:00:00"
        assert d["sent_time"] == "2026-07-07 09:55:00"
        assert d["unread"] is True
        assert d["has_attachments"] is True
        assert d["conversation_id"] == "conv-1"
        assert d["conversation_topic"] == "Hello"
        # body_top should be the cleaned new message only.
        assert d["body_top"] == "Body text"
        assert d["body"] == _FakeMailItem.Body
        assert len(d["attachments"]) == 1
        assert d["attachments"][0]["filename"] == "a.pdf"
        assert d["attachments"][0]["size"] == 1234

    def test_summary_mode_omits_body_fields(self):
        bridge = OutlookBridge.__new__(OutlookBridge)
        d = bridge._mail_item_to_dict(_FakeMailItem(), include_body=False)
        assert "body" not in d
        assert "attachments" not in d
        assert d["message_class"] == "IPM.Note"
