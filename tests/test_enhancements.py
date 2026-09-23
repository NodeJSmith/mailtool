"""Unit tests for the inbox-triage enhancements.

These tests exercise the pure-Python logic added in the v0.11.0 triage work
(MessageClass filtering, sender SMTP salvage, body cleaning, richer metadata,
and the new bulk/stats tools). They do NOT require Outlook or pywin32 to be
running; the bridge module imports defensively so its static helpers and module
constants can be tested on any platform.
"""

from datetime import datetime
from unittest.mock import MagicMock

import pytest

from mailtool.bridge import _SMTP_REGEX, MAIL_ONLY_FILTER, OutlookBridge
from mailtool.mcp.models import (
    AttachmentInfo,
    EmailDetails,
    EmailSummary,
    InboxStats,
)
from mailtool.mcp.server import (
    _email_details_from_dict,
    _email_summary_from_dict,
    get_email,
    get_emails,
    get_inbox_stats,
)

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


# =============================================================================
# Pydantic models (Phases 1, 3, 6)
# =============================================================================


@pytest.mark.unit
class TestModels:
    def test_email_summary_new_fields_default(self):
        e = EmailSummary(
            entry_id="x",
            subject="s",
            sender="a@b.com",
            sender_name="A",
            unread=True,
            has_attachments=False,
        )
        assert e.message_class == "IPM.Note"
        assert e.to == ""
        assert e.cc == ""
        assert e.sent_time is None
        d = EmailDetails(
            entry_id="x",
            subject="s",
            sender="a@b.com",
            sender_name="A",
            body="b",
            html_body="<b/>",
            has_attachments=False,
        )
        assert d.attachments == []
        assert d.body_top == ""
        assert d.bcc == ""
        assert d.message_class == "IPM.Note"

    def test_attachment_info_defaults(self):
        a = AttachmentInfo()
        assert a.filename == ""
        assert a.size == 0
        assert a.is_inline is False
        assert a.content_type is None

    def test_inbox_stats_defaults(self):
        s = InboxStats(folder="Inbox")
        assert s.total == 0
        assert s.unread == 0

    def test_email_summary_from_minimal_legacy_dict(self):
        # Old bridge dicts (only the original 7 keys) must still construct.
        e = _email_summary_from_dict(
            {
                "entry_id": "1",
                "subject": "s",
                "sender": "a@b.com",
                "sender_name": "A",
                "received_time": None,
                "unread": True,
                "has_attachments": False,
            }
        )
        assert e.message_class == "IPM.Note"
        assert e.entry_id == "1"

    def test_email_details_from_dict_with_attachments(self):
        d = _email_details_from_dict(
            {
                "entry_id": "1",
                "subject": "s",
                "sender": "a@b.com",
                "sender_name": "A",
                "body": "b",
                "html_body": "<b/>",
                "received_time": "2026-07-07 10:00:00",
                "has_attachments": True,
                "body_top": "top message",
                "message_class": "IPM.Note.SMIME",
                "attachments": [
                    {
                        "filename": "report.pdf",
                        "size": 1234,
                        "display_name": "report.pdf",
                        "content_type": "application/pdf",
                        "is_inline": False,
                    }
                ],
            }
        )
        assert len(d.attachments) == 1
        assert d.attachments[0].filename == "report.pdf"
        assert d.attachments[0].size == 1234
        assert d.body_top == "top message"
        assert d.message_class == "IPM.Note.SMIME"


# =============================================================================
# _mail_item_to_dict refactor (Phase 3) — uses a fake COM item, no Outlook
# =============================================================================


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


# =============================================================================
# New tools: get_emails, get_inbox_stats, get_email on non-mail (Phases 1, 4, 6)
# =============================================================================


@pytest.fixture
def mock_bridge():
    return MagicMock()


@pytest.fixture
def server_with_mock(mock_bridge):
    from mailtool.mcp import server

    server._bridge = mock_bridge
    yield server
    server._bridge = None


@pytest.mark.unit
class TestGetEmailsTool:
    def test_bulk_fetch_returns_details(self, server_with_mock, mock_bridge):
        mock_bridge.get_email_bodies.return_value = [
            {
                "entry_id": "a",
                "subject": "A",
                "sender": "a@x.com",
                "sender_name": "A",
                "body": "ba",
                "html_body": "<p>ba</p>",
                "received_time": "2026-07-07 10:00:00",
                "has_attachments": False,
            },
            {
                "entry_id": "b",
                "subject": "B",
                "sender": "b@x.com",
                "sender_name": "B",
                "body": "bb",
                "html_body": "<p>bb</p>",
                "received_time": None,
                "has_attachments": False,
            },
        ]
        result = get_emails(["a", "b", "missing"])
        assert len(result) == 2
        assert all(isinstance(r, EmailDetails) for r in result)
        assert [r.entry_id for r in result] == ["a", "b"]
        mock_bridge.get_email_bodies.assert_called_once_with(
            ["a", "b", "missing"], include_body=True
        )

    def test_bulk_fetch_no_body_flag_forwarded(self, server_with_mock, mock_bridge):
        mock_bridge.get_email_bodies.return_value = []
        get_emails(["a"], include_body=False)
        mock_bridge.get_email_bodies.assert_called_once_with(["a"], include_body=False)


@pytest.mark.unit
class TestGetInboxStatsTool:
    def test_returns_inbox_stats(self, server_with_mock, mock_bridge):
        mock_bridge.get_inbox_stats.return_value = {
            "folder": "Inbox",
            "total": 421,
            "unread": 166,
        }
        result = get_inbox_stats()
        assert isinstance(result, InboxStats)
        assert result.folder == "Inbox"
        assert result.total == 421
        assert result.unread == 166
        mock_bridge.get_inbox_stats.assert_called_once_with(folder="Inbox")


@pytest.mark.unit
class TestGetEmailOnNonMailItem:
    def test_non_mail_item_returns_details_with_message_class(
        self, server_with_mock, mock_bridge
    ):
        # A meeting notification that previously caused "Email not found" now
        # comes back as a populated dict with its MessageClass set.
        mock_bridge.get_email_body.return_value = {
            "entry_id": "mid-1",
            "subject": "Canceled: Standup",
            "sender": "organizer@x.com",
            "sender_name": "Org",
            "body": "",
            "html_body": "",
            "received_time": "2026-07-07 08:00:00",
            "has_attachments": False,
            "message_class": "IPM.Schedule.Meeting.Canceled",
            "to": "me@x.com",
            "cc": "",
            "bcc": "",
            "sent_time": None,
            "conversation_id": None,
            "conversation_topic": "Standup",
            "body_top": "",
            "attachments": [],
        }
        result = get_email("mid-1")
        assert isinstance(result, EmailDetails)
        assert result.message_class == "IPM.Schedule.Meeting.Canceled"
        assert result.subject == "Canceled: Standup"

    def test_truly_missing_raises(self, server_with_mock, mock_bridge):
        from mcp import McpError

        mock_bridge.get_email_body.return_value = None
        with pytest.raises(McpError):
            get_email("does-not-exist")
