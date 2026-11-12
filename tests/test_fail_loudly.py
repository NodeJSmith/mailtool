"""
Unit tests for the fail-loudly contract: a failed Outlook/COM operation raises
MailtoolError instead of returning [] / False / None, while a fetch of an
entry ID that doesn't exist still returns None.

Pure-Python (fake COM objects, no Outlook / pywin32 required).
"""

import json
import sys
from typing import ClassVar

import pytest

from mailtool import cli
from mailtool.bridge import MailtoolError, OutlookBridge

pytestmark = pytest.mark.unit


class FakeComError(Exception):
    """Stands in for pywintypes.com_error."""


class FakeItem:
    """A COM item whose listed operations fail."""

    EntryID = "id-1"
    Subject = "subject"
    Complete = False
    Status = 0
    PercentComplete = 0
    Importance = 1
    Body = ""
    DueDate = None

    def __init__(self, fail_on=()):
        self.fail_on = set(fail_on)
        self.calls = []

    def fail_if_requested(self, name):
        self.calls.append(name)
        if name in self.fail_on:
            raise FakeComError(f"{name} failed")

    def Save(self):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Save")

    def Delete(self):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Delete")

    def Move(self, _folder):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Move")

    def Send(self):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Send")

    def Response(self, _code):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Response")


class FakeNamespace:
    def __init__(self, items):
        self.items = items

    def GetItemFromID(self, entry_id):  # noqa: N802 - mirrors Outlook COM API
        if entry_id not in self.items:
            raise FakeComError("not found")
        return self.items[entry_id]


class ExplodingItems:
    """An Items collection whose COM calls fail."""

    def Restrict(self, _query):  # noqa: N802 - mirrors Outlook COM API
        raise FakeComError("restrict failed")

    def Sort(self, *_args):  # noqa: N802 - mirrors Outlook COM API
        raise FakeComError("sort failed")

    def __iter__(self):
        raise FakeComError("iteration failed")

    @property
    def Count(self):  # noqa: N802 - mirrors Outlook COM API
        raise FakeComError("count failed")


class ListItems:
    def __init__(self, items):
        self.items = list(items)

    def Restrict(self, _query):  # noqa: N802 - mirrors Outlook COM API
        return self

    def Sort(self, *_args):  # noqa: N802 - mirrors Outlook COM API
        pass

    def __iter__(self):
        return iter(self.items)


class FakeFolder:
    def __init__(self, items):
        self.Items = items


class GoodMail:
    EntryID = "mail-good"
    Subject = "hello"
    MessageClass = "IPM.Note"
    SenderEmailType = "SMTP"
    SenderEmailAddress = "a@example.com"
    SenderName = "A"
    Unread = True
    To = ""
    CC = ""


class BadMail:
    """Reading any attribute fails like a broken COM item."""

    def __getattr__(self, name):
        raise ValueError(f"cannot read {name}")


def make_bridge(items=None, inbox_items=None, tasks_items=None):
    bridge = OutlookBridge.__new__(OutlookBridge)
    bridge.namespace = FakeNamespace(items or {})
    bridge.default_account_name = None
    if inbox_items is not None:
        inbox = FakeFolder(inbox_items)
        bridge.get_inbox = lambda: inbox
    if tasks_items is not None:
        tasks = FakeFolder(tasks_items)
        bridge.get_tasks = lambda: tasks
    return bridge


class TestListMethodsRaise:
    def test_list_emails_com_failure_raises_instead_of_empty_list(self):
        bridge = make_bridge(inbox_items=ExplodingItems())

        with pytest.raises(MailtoolError, match="list emails") as exc_info:
            bridge.list_emails()

        assert isinstance(exc_info.value.__cause__, FakeComError)

    def test_search_emails_com_failure_raises(self):
        bridge = make_bridge(inbox_items=ExplodingItems())

        with pytest.raises(MailtoolError):
            bridge.search_emails("[Unread] = TRUE")

    def test_search_by_sender_com_failure_raises(self):
        bridge = make_bridge(inbox_items=ExplodingItems())

        with pytest.raises(MailtoolError, match="search emails by sender"):
            bridge.search_by_sender("a@example.com")

    def test_get_inbox_stats_com_failure_raises(self):
        bridge = make_bridge(inbox_items=ExplodingItems())

        with pytest.raises(MailtoolError, match="inbox stats"):
            bridge.get_inbox_stats()

    def test_list_calendar_events_com_failure_raises(self):
        bridge = OutlookBridge.__new__(OutlookBridge)
        calendar = FakeFolder(ExplodingItems())
        bridge.get_calendar = lambda: calendar

        with pytest.raises(MailtoolError, match="calendar events"):
            bridge.list_calendar_events()

    def test_list_calendar_events_missing_folder_raises(self):
        bridge = OutlookBridge.__new__(OutlookBridge)
        bridge.get_calendar = lambda: None

        with pytest.raises(MailtoolError, match="Calendar folder not found"):
            bridge.list_calendar_events()

    def test_list_tasks_com_failure_raises(self):
        bridge = make_bridge(tasks_items=ExplodingItems())

        with pytest.raises(MailtoolError, match="list tasks"):
            bridge.list_tasks()

    def test_list_tasks_missing_folder_raises(self):
        bridge = OutlookBridge.__new__(OutlookBridge)
        bridge.get_tasks = lambda: None

        with pytest.raises(MailtoolError, match="Tasks folder not found"):
            bridge.list_tasks()

    def test_mailtool_error_is_a_runtime_error(self):
        assert issubclass(MailtoolError, RuntimeError)


class TestPerItemSkips:
    def test_unreadable_item_is_skipped_with_a_warning(self, capsys):
        bridge = make_bridge(inbox_items=ListItems([BadMail(), GoodMail()]))

        emails = bridge.list_emails()

        assert [e["entry_id"] for e in emails] == ["mail-good"]
        assert "skipped an unreadable email" in capsys.readouterr().err

    def test_empty_folder_is_an_empty_list(self):
        bridge = make_bridge(inbox_items=ListItems([]))

        assert bridge.list_emails() == []


class TestMutatorsRaise:
    @pytest.mark.parametrize(
        ("method", "args", "fail_on"),
        [
            ("delete_email", ("id-1",), "Delete"),
            ("mark_email_read", ("id-1",), "Save"),
            ("delete_appointment", ("id-1",), "Delete"),
            ("respond_to_meeting", ("id-1", "accept"), "Response"),
            ("delete_task", ("id-1",), "Delete"),
            ("complete_task", ("id-1",), "Save"),
            ("edit_task", ("id-1", "new subject"), "Save"),
            ("edit_appointment", ("id-1",), "Save"),
        ],
    )
    def test_com_failure_raises(self, method, args, fail_on):
        bridge = make_bridge(items={"id-1": FakeItem(fail_on=[fail_on])})

        with pytest.raises(MailtoolError):
            getattr(bridge, method)(*args)

    @pytest.mark.parametrize(
        ("method", "args"),
        [
            ("delete_email", ("nope",)),
            ("mark_email_read", ("nope",)),
            ("move_email", ("nope", "Archive")),
            ("reply_email", ("nope", "body")),
            ("forward_email", ("nope", "a@example.com")),
            ("delete_appointment", ("nope",)),
            ("respond_to_meeting", ("nope", "accept")),
            ("edit_appointment", ("nope",)),
            ("delete_task", ("nope",)),
            ("complete_task", ("nope",)),
            ("edit_task", ("nope",)),
            ("download_attachments", ("nope", "/tmp/x")),
        ],
    )
    def test_missing_item_raises(self, method, args):
        bridge = make_bridge()

        with pytest.raises(MailtoolError, match="not found: nope"):
            getattr(bridge, method)(*args)

    def test_success_returns_true(self):
        item = FakeItem()
        bridge = make_bridge(items={"id-1": item})

        assert bridge.delete_email("id-1") is True
        assert bridge.complete_task("id-1") is True
        assert bridge.mark_email_read("id-1") is True
        assert item.calls == ["Delete", "Save", "Save"]

    def test_move_to_missing_folder_raises(self, monkeypatch):
        bridge = make_bridge(items={"id-1": FakeItem()})
        monkeypatch.setattr(bridge, "get_folder_by_name", lambda _name: None)

        with pytest.raises(MailtoolError, match="Folder not found: Archive"):
            bridge.move_email("id-1", "Archive")

    def test_move_failure_raises(self, monkeypatch):
        bridge = make_bridge(items={"id-1": FakeItem(fail_on=["Move"])})
        monkeypatch.setattr(bridge, "get_folder_by_name", lambda _name: object())

        with pytest.raises(MailtoolError, match="move email"):
            bridge.move_email("id-1", "Archive")

    def test_unknown_meeting_response_raises(self):
        bridge = make_bridge(items={"id-1": FakeItem()})

        with pytest.raises(MailtoolError, match="Unknown meeting response"):
            bridge.respond_to_meeting("id-1", "maybe")


class FakeMail(FakeItem):
    def __init__(self, fail_on=()):
        super().__init__(fail_on)
        self.attachments = []

    @property
    def Attachments(self):  # noqa: N802 - mirrors Outlook COM API
        return self

    def Add(self, path):  # noqa: N802 - mirrors Outlook COM API
        self.fail_if_requested("Add")
        self.attachments.append(path)


class FakeOutlook:
    def __init__(self, mail):
        self.mail = mail

    def CreateItem(self, _kind):  # noqa: N802 - mirrors Outlook COM API
        return self.mail


class TestCreateAndSend:
    def test_send_failure_raises(self):
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(FakeMail(fail_on=["Send"]))
        bridge._get_root = lambda: None

        with pytest.raises(MailtoolError, match="send email"):
            bridge.send_email("a@example.com", "s", "b")

    def test_attachment_failure_raises_before_sending(self):
        mail = FakeMail(fail_on=["Add"])
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(mail)
        bridge._get_root = lambda: None

        with pytest.raises(MailtoolError, match="send email"):
            bridge.send_email("a@example.com", "s", "b", file_paths=["/nope"])

        assert "Send" not in mail.calls

    def test_send_success_returns_true_and_draft_returns_entry_id(self):
        mail = FakeMail()
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(mail)
        bridge._get_root = lambda: None

        assert bridge.send_email("a@example.com", "s", "b") is True
        assert bridge.send_email("a@example.com", "s", "b", save_draft=True) == "id-1"

    def test_create_task_failure_raises(self):
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(FakeItem(fail_on=["Save"]))
        bridge.get_tasks = lambda: None

        with pytest.raises(MailtoolError, match="create task"):
            bridge.create_task("subject")

    def test_create_appointment_failure_raises(self):
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(FakeItem(fail_on=["Save"]))
        bridge.get_calendar = lambda: None

        with pytest.raises(MailtoolError, match="create appointment"):
            bridge.create_appointment("s", "2026-01-01 10:00:00", "2026-01-01 11:00:00")

    def test_create_appointment_bad_date_raises(self):
        bridge = make_bridge()
        bridge.outlook = FakeOutlook(FakeItem())
        bridge.get_calendar = lambda: None

        with pytest.raises(MailtoolError, match="create appointment"):
            bridge.create_appointment("s", "not a date", "also not")


class TestNotFoundReturnsNone:
    def test_get_email_body_missing_is_none(self):
        assert make_bridge().get_email_body("nope") is None

    def test_get_task_missing_is_none(self):
        assert make_bridge().get_task("nope") is None

    def test_get_appointment_missing_is_none(self):
        assert make_bridge().get_appointment("nope") is None

    def test_get_email_body_read_failure_raises(self):
        bridge = make_bridge(items={"id-1": BadMail()})

        with pytest.raises(MailtoolError, match="read email id-1"):
            bridge.get_email_body("id-1")

    def test_get_email_bodies_omits_missing_ids(self):
        bridge = make_bridge(items={"mail-good": GoodMail()})

        results = bridge.get_email_bodies(["nope", "mail-good"], include_body=False)

        assert [r["entry_id"] for r in results] == ["mail-good"]


class TestFreeBusy:
    def test_com_failure_raises(self):
        bridge = make_bridge()
        bridge.namespace.CreateRecipient = lambda _addr: (_ for _ in ()).throw(
            FakeComError("boom")
        )

        with pytest.raises(MailtoolError, match="free/busy"):
            bridge.get_free_busy(email_address="a@example.com")

    def test_unresolvable_address_is_a_normal_answer(self):
        bridge = make_bridge()

        class Recipient:
            def Resolve(self):  # noqa: N802 - mirrors Outlook COM API
                return False

        bridge.namespace.CreateRecipient = lambda _addr: Recipient()

        result = bridge.get_free_busy(email_address="a@example.com")

        assert result["resolved"] is False


class FakeCliBridge:
    """Bridge double for cli.main: every method named in `behaviors` runs it."""

    behaviors: ClassVar[dict] = {}

    def __getattr__(self, name):
        return self.behaviors[name]


@pytest.fixture
def run_cli(monkeypatch, capsys):
    monkeypatch.setattr(cli, "_check_platform", lambda: None)
    monkeypatch.setattr(cli, "_check_pywin32", lambda: None)

    def run(argv, **behaviors):
        FakeCliBridge.behaviors = behaviors
        monkeypatch.setattr(cli, "OutlookBridge", FakeCliBridge)
        monkeypatch.setattr(sys, "argv", ["mailtool", *argv])
        code = 0
        try:
            cli.main()
        except SystemExit as e:
            code = e.code
        captured = capsys.readouterr()
        return code, captured.out, captured.err

    return run


def raises(exc):
    def raiser(*_args, **_kwargs):
        raise exc

    return raiser


class TestCli:
    def test_mailtool_error_prints_error_json_and_exits_1(self, run_cli):
        code, out, _ = run_cli(
            ["emails"], list_emails=raises(MailtoolError("Inbox folder not found"))
        )

        assert code == 1
        assert json.loads(out) == {
            "status": "error",
            "message": "Inbox folder not found",
        }

    def test_unexpected_exception_prints_error_json_and_exits_1(self, run_cli):
        code, out, err = run_cli(["tasks"], list_tasks=raises(ValueError("surprise")))

        assert code == 1
        payload = json.loads(out)
        assert payload["status"] == "error"
        assert "surprise" in payload["message"]
        assert "Traceback" in err

    def test_mutator_error_exits_1(self, run_cli):
        code, out, _ = run_cli(
            ["delete-email", "--id", "x"],
            delete_email=raises(MailtoolError("Email not found: x")),
        )

        assert code == 1
        assert json.loads(out)["message"] == "Email not found: x"

    def test_empty_list_is_still_success(self, run_cli):
        code, out, _ = run_cli(["emails"], list_emails=lambda **_kw: [])

        assert code == 0
        assert json.loads(out) == []

    @pytest.mark.parametrize(
        ("argv", "behavior"),
        [
            (["email", "--id", "x"], "get_email_body"),
            (["appointment", "--id", "x"], "get_appointment"),
            (["task", "--id", "x"], "get_task"),
        ],
    )
    def test_not_found_exits_1_with_not_found_message(self, run_cli, argv, behavior):
        code, out, _ = run_cli(argv, **{behavior: lambda *_a, **_kw: None})

        assert code == 1
        assert "not found" in json.loads(out)["message"]

    def test_success_prints_success_json(self, run_cli):
        code, out, _ = run_cli(
            ["delete-email", "--id", "x"], delete_email=lambda _id: True
        )

        assert code == 0
        assert json.loads(out) == {"status": "success", "message": "Email deleted"}
