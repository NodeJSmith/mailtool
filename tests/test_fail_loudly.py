"""
Unit tests for the fail-loudly contract: a failed Outlook/COM operation raises
MailtoolError instead of returning [] / False / None, while a fetch of an
entry ID that doesn't exist still returns None.

Pure-Python (fake COM objects, no Outlook / pywin32 required).
"""

import json
import sys

import pytest

from mailtool import cli
from mailtool.bridge import MailtoolError, OutlookBridge, _com_operation

pytestmark = pytest.mark.unit


class FakeComError(Exception):
    """Stands in for pywintypes.com_error."""


class FakeNotFoundError(FakeComError):
    """Outlook's "no item has this EntryID" error (MAPI_E_NOT_FOUND)."""

    hresult = -2147221233


class FakeWrappedNotFoundError(FakeComError):
    """The same error as pywin32 usually delivers it: DISP_E_EXCEPTION with the
    real code in excepinfo[5]."""

    hresult = -2147352567
    excepinfo = (0, "Microsoft Outlook", "not found", None, 0, -2147221241)


class AnyClass:
    """An Outlook Class value that matches whatever class a method asks for."""

    def __eq__(self, other):
        return True

    def __ne__(self, other):
        return False

    def __hash__(self):
        return 0


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

    def __init__(self, fail_on=(), olclass=None):
        self.fail_on = set(fail_on)
        self.calls = []
        self.Class = AnyClass() if olclass is None else olclass

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
    def __init__(self, items, error=None):
        self.items = items
        self.error = error

    def GetItemFromID(self, entry_id):  # noqa: N802 - mirrors Outlook COM API
        if self.error is not None:
            raise self.error
        if entry_id not in self.items:
            raise FakeNotFoundError("not found")
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


class NoAttachments:
    Count = 0


class GoodMail:
    EntryID = "mail-good"
    Attachments = NoAttachments()
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


def make_bridge(items=None, inbox_items=None, tasks_items=None, lookup_error=None):
    bridge = OutlookBridge.__new__(OutlookBridge)
    bridge.namespace = FakeNamespace(items or {}, error=lookup_error)
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
    """Bridge double for cli.main: only the methods named in `behaviors` exist."""

    def __init__(self, behaviors):
        self.behaviors = behaviors

    def __getattr__(self, name):
        behaviors = self.__dict__.get("behaviors", {})
        if name not in behaviors:
            raise AttributeError(name)
        return behaviors[name]


@pytest.fixture
def run_cli(monkeypatch, capsys):
    monkeypatch.setattr(cli, "_check_platform", lambda: None)
    monkeypatch.setattr(cli, "_check_pywin32", lambda: None)

    def run(argv, real_bridge=None, **behaviors):
        bridge = real_bridge or FakeCliBridge(behaviors)
        monkeypatch.setattr(cli, "OutlookBridge", lambda: bridge)
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


class TestEditTaskCompleteFlag:
    @pytest.mark.parametrize(("raw", "expected"), [("false", False), ("true", True)])
    def test_complete_flag_is_parsed_as_a_boolean(self, run_cli, raw, expected):
        seen = {}

        def edit_task(*args):
            seen["complete"] = args[-1]
            return True

        code, _, _ = run_cli(
            ["edit-task", "--id", "x", "--complete", raw], edit_task=edit_task
        )

        assert code == 0
        assert seen["complete"] is expected

    def test_invalid_value_is_a_usage_error(self, run_cli):
        code, _, _ = run_cli(
            ["edit-task", "--id", "x", "--complete", "maybe"],
            edit_task=lambda *_a: True,
        )

        assert code == 2


class TestGetItemById:
    def test_not_found_hresult_returns_none(self):
        assert make_bridge().get_item_by_id("nope") is None

    def test_not_found_code_in_excepinfo_returns_none(self):
        bridge = make_bridge(lookup_error=FakeWrappedNotFoundError("wrapped"))

        assert bridge.get_item_by_id("nope") is None

    def test_other_com_error_raises_instead_of_reading_as_not_found(self):
        error = FakeComError("RPC server unavailable")
        bridge = make_bridge(lookup_error=error)

        with pytest.raises(MailtoolError, match="Could not look up item x") as exc_info:
            bridge.get_item_by_id("x")

        assert exc_info.value.__cause__ is error

    def test_outage_propagates_through_getters(self):
        bridge = make_bridge(lookup_error=FakeComError("outage"))

        with pytest.raises(MailtoolError):
            bridge.get_email_body("x")
        with pytest.raises(MailtoolError):
            bridge.delete_email("x")

    def test_legacy_free_busy_entry_id_not_found_raises(self):
        bridge = make_bridge()

        with pytest.raises(MailtoolError, match="Item not found: nope"):
            bridge.get_free_busy(entry_id="nope")

    def test_get_email_bodies_warns_for_each_missing_id(self, capsys):
        bridge = make_bridge(items={"mail-good": GoodMail()})

        results = bridge.get_email_bodies(["a", "mail-good", "b"], include_body=False)

        assert [r["entry_id"] for r in results] == ["mail-good"]
        err = capsys.readouterr().err
        assert "email a: not found" in err
        assert "email b: not found" in err

    def test_get_email_bodies_raises_on_outage(self):
        bridge = make_bridge(lookup_error=FakeComError("outage"))

        with pytest.raises(MailtoolError):
            bridge.get_email_bodies(["a"])


class TestItemClassChecks:
    """An EntryID for the wrong kind of item is treated as not found."""

    OL_MAIL = 43

    @pytest.mark.parametrize("method", ["get_appointment", "get_task"])
    def test_getter_returns_none_for_wrong_class(self, method):
        bridge = make_bridge(items={"id-1": FakeItem(olclass=self.OL_MAIL)})

        assert getattr(bridge, method)("id-1") is None

    @pytest.mark.parametrize(
        ("method", "args"),
        [
            ("edit_appointment", ("id-1", None, None, "new")),
            ("delete_appointment", ("id-1",)),
            ("respond_to_meeting", ("id-1", "accept")),
            ("edit_task", ("id-1", "new")),
            ("complete_task", ("id-1",)),
            ("delete_task", ("id-1",)),
        ],
    )
    def test_mutator_raises_for_wrong_class_without_touching_the_item(
        self, method, args
    ):
        item = FakeItem(olclass=self.OL_MAIL)
        bridge = make_bridge(items={"id-1": item})

        with pytest.raises(MailtoolError, match="not found: id-1"):
            getattr(bridge, method)(*args)

        assert item.calls == []

    @pytest.mark.parametrize("method", ["get_appointment", "get_task"])
    def test_unreadable_class_raises_instead_of_not_found(self, method):
        class UnreadableClassItem(FakeItem):
            @property
            def Class(self):  # noqa: N802 - mirrors Outlook COM API
                raise FakeComError("store offline")

            @Class.setter
            def Class(self, _value):  # noqa: N802 - mirrors Outlook COM API
                pass

        bridge = make_bridge(items={"id-1": UnreadableClassItem()})

        with pytest.raises(MailtoolError, match="read the class"):
            getattr(bridge, method)("id-1")

    def test_task_class_is_accepted_for_task_methods(self):
        bridge = make_bridge(items={"id-1": FakeItem(olclass=48)})

        assert bridge.complete_task("id-1") is True

    def test_appointment_class_is_accepted_for_appointment_methods(self):
        bridge = make_bridge(items={"id-1": FakeItem(olclass=26)})

        assert bridge.delete_appointment("id-1") is True


class BrokenFolder:
    Name = "Broken"

    @property
    def Folders(self):  # noqa: N802 - mirrors Outlook COM API
        raise FakeComError("store unavailable")


class FakeFolderNode:
    def __init__(self, name, children=()):
        self.Name = name
        self.children = list(children)

    @property
    def Folders(self):  # noqa: N802 - mirrors Outlook COM API
        return FakeChildren(self.children)


class FakeChildren:
    def __init__(self, children):
        self.children = children

    def __getitem__(self, name):
        for child in self.children:
            if child.Name == name:
                return child
        raise FakeComError("not found")

    def __iter__(self):
        return iter(self.children)


class FakeRoots:
    def __init__(self, roots):
        self.roots = roots
        self.Folders = self

    @property
    def Count(self):  # noqa: N802 - mirrors Outlook COM API
        return len(self.roots)

    def Item(self, i):  # noqa: N802 - mirrors Outlook COM API
        return self.roots[i - 1]


class TestFolderLookups:
    def test_calendar_absent_is_none(self):
        bridge = make_bridge()
        bridge._get_root = lambda: FakeFolderNode("root", [FakeFolderNode("Inbox")])
        bridge.namespace = FakeRoots([])

        assert bridge.get_calendar() is None

    def test_calendar_found_case_insensitively(self):
        calendar = FakeFolderNode("calendar")
        bridge = make_bridge()
        bridge._get_root = lambda: FakeFolderNode("root", [calendar])

        assert bridge.get_calendar() is calendar

    def test_calendar_busy_outlook_raises_instead_of_not_found(self):
        bridge = make_bridge()
        bridge._get_root = lambda: BrokenFolder()

        with pytest.raises(MailtoolError, match="Could not list subfolders") as exc:
            bridge.get_calendar()

        assert isinstance(exc.value.__cause__, FakeComError)

    def test_calendar_root_listing_failure_raises(self):
        class DeadNamespace:
            @property
            def Folders(self):  # noqa: N802 - mirrors Outlook COM API
                raise FakeComError("outage")

        bridge = make_bridge()
        bridge._get_root = lambda: None
        bridge.namespace = DeadNamespace()

        with pytest.raises(MailtoolError, match="list account roots"):
            bridge.get_calendar()

    def test_tasks_and_inbox_default_folder_failure_raises(self):
        class DeadNamespace:
            def GetDefaultFolder(self, _kind):  # noqa: N802 - mirrors Outlook COM API
                raise FakeComError("outage")

        bridge = make_bridge()
        bridge._get_root = lambda: None
        bridge.namespace = DeadNamespace()

        with pytest.raises(MailtoolError, match="Tasks folder"):
            bridge.get_tasks()
        with pytest.raises(MailtoolError, match="Inbox folder"):
            bridge.get_inbox()

    def test_failing_account_root_is_warned_and_raises_when_nothing_found(self, capsys):
        bridge = make_bridge()
        bridge._get_root = lambda: None
        bridge.namespace = FakeRoots([BrokenFolder()])

        with pytest.raises(MailtoolError, match="account root failed"):
            bridge.get_folder_by_name("Archive")

        assert "skipped account root 1" in capsys.readouterr().err

    def test_failing_account_root_ignored_when_folder_found_elsewhere(self):
        archive = FakeFolderNode("Archive")
        bridge = make_bridge()
        bridge._get_root = lambda: None
        bridge.namespace = FakeRoots(
            [BrokenFolder(), FakeFolderNode("acct", [archive])]
        )

        assert bridge.get_folder_by_name("Archive") is archive

    def test_failing_subtree_warns_and_raises_when_folder_not_found(self, capsys):
        root = FakeFolderNode("root", [BrokenFolder(), FakeFolderNode("Other")])
        bridge = make_bridge()
        bridge._get_root = lambda: root

        with pytest.raises(MailtoolError, match="could not be searched"):
            bridge.get_folder_by_name("Archive")

        assert "can't be listed" in capsys.readouterr().err

    def test_failing_subtree_ignored_when_folder_found(self):
        archive = FakeFolderNode("Archive")
        root = FakeFolderNode("root", [archive, BrokenFolder()])
        bridge = make_bridge()
        bridge._get_root = lambda: root

        assert bridge.get_folder_by_name("Archive") is archive

    def test_path_lookup_failure_raises(self):
        root = FakeFolderNode("root", [BrokenFolder()])
        bridge = make_bridge()
        bridge._get_root = lambda: root

        with pytest.raises(MailtoolError):
            bridge.get_folder_by_name("Broken\\Child")


class FakeAccountFolder:
    """An account root with one child folder, shaped for list_folders."""

    def __init__(self, name):
        self.Name = name
        self.EntryID = f"id-{name}"
        self.FolderPath = f"\\\\{name}"
        self.Items = type("I", (), {"Count": 3})()
        self.Folders = []


class TestListFolders:
    def test_lists_every_account_not_just_the_first_six(self):
        bridge = make_bridge()
        bridge.namespace = FakeRoots([FakeAccountFolder(f"acct{i}") for i in range(8)])

        result = bridge.list_folders()

        assert len(result) == 8

    def test_unknown_account_raises(self):
        bridge = make_bridge()
        bridge.namespace = FakeRoots([FakeAccountFolder("work")])

        with pytest.raises(MailtoolError, match="Account not found: nope"):
            bridge.list_folders("nope")

    def test_filters_to_the_named_account(self):
        bridge = make_bridge()
        bridge.namespace = FakeRoots(
            [FakeAccountFolder("work"), FakeAccountFolder("home")]
        )

        assert list(bridge.list_folders("home")) == ["home"]

    def test_failure_raises(self):
        bridge = make_bridge()
        bridge.namespace = FakeRoots([BrokenFolder()])

        with pytest.raises(MailtoolError, match="list folders"):
            bridge.list_folders()

    def test_cli_folders_stdout_is_pure_json(self, run_cli):
        bridge = make_bridge()
        bridge.namespace = FakeRoots([FakeAccountFolder("work")])

        code, out, _ = run_cli(["folders"], real_bridge=bridge)

        assert code == 0
        assert list(json.loads(out)) == ["work"]


class TestSearchEmailsErrors:
    def test_items_access_failure_raises_mailtool_error(self):
        class BrokenFolderItems:
            @property
            def Items(self):  # noqa: N802 - mirrors Outlook COM API
                raise FakeComError("items unavailable")

        bridge = make_bridge()
        bridge.get_inbox = lambda: BrokenFolderItems()

        with pytest.raises(MailtoolError, match="search emails"):
            bridge.search_emails("[Unread] = TRUE")

    def test_untranslatable_like_field_raises_mailtool_error(self):
        bridge = make_bridge(inbox_items=ListItems([]))

        with pytest.raises(MailtoolError, match="Cannot translate"):
            bridge.search_emails("[NoSuchField] LIKE '%x%'")

    def test_missing_inbox_raises(self):
        bridge = make_bridge()
        bridge.get_inbox = lambda: None

        with pytest.raises(MailtoolError, match="Inbox folder not found"):
            bridge.search_emails("")


class TestMailOnlyFallback:
    def test_failed_restrict_warns_and_returns_unfiltered_items(self, capsys):
        items = ExplodingItems()

        result = OutlookBridge.restrict_to_mail(items, include_non_mail=False)

        assert result is items
        assert "mail-only filter failed" in capsys.readouterr().err

    def test_include_non_mail_skips_the_filter(self):
        items = ExplodingItems()

        assert OutlookBridge.restrict_to_mail(items, include_non_mail=True) is items


class TestComOperation:
    def test_mailtool_error_passes_through_unwrapped(self):
        original = MailtoolError("specific message")

        with (
            pytest.raises(MailtoolError) as exc_info,
            _com_operation("outer"),
            _com_operation("inner"),
        ):
            raise original

        assert exc_info.value is original

    def test_other_exceptions_are_wrapped_and_chained(self):
        cause = FakeComError("boom")

        with (
            pytest.raises(MailtoolError, match="doing x failed: boom") as exc_info,
            _com_operation("doing x"),
        ):
            raise cause

        assert exc_info.value.__cause__ is cause


class TestCliExtras:
    def test_email_with_no_attachments_is_success(self, run_cli):
        code, out, _ = run_cli(
            ["attachments", "--id", "x", "--dir", "/tmp/out"],
            download_attachments=lambda *_a: [],
        )

        assert code == 0
        assert json.loads(out) == {"status": "success", "attachments": []}

    @pytest.mark.parametrize(
        ("response", "message"),
        [
            ("accept", "Meeting accepted"),
            ("decline", "Meeting declined"),
            ("tentative", "Meeting tentatively accepted"),
        ],
    )
    def test_respond_message(self, run_cli, response, message):
        code, out, _ = run_cli(
            ["respond", "--id", "x", "--response", response],
            respond_to_meeting=lambda *_a: True,
        )

        assert code == 0
        assert json.loads(out)["message"] == message

    def test_unconfigured_bridge_method_raises_attribute_error(self):
        with pytest.raises(AttributeError):
            FakeCliBridge({}).list_emails()


class BrokenAttachmentsItem:
    @property
    def Attachments(self):  # noqa: N802 - mirrors Outlook COM API
        raise FakeComError("store offline")


class TestAttachmentReadFailures:
    """A failed attachment read must not pass for "no attachments"."""

    def test_count_failure_raises(self):
        with pytest.raises(MailtoolError, match="count attachments"):
            OutlookBridge._attachment_count(BrokenAttachmentsItem())

    def test_extract_failure_raises(self):
        bridge = make_bridge()

        with pytest.raises(MailtoolError, match="read attachments"):
            bridge._extract_attachments(BrokenAttachmentsItem())


class TestFreeBusyCli:
    def test_unresolvable_address_exits_1(self, run_cli):
        code, out, _ = run_cli(
            ["freebusy", "--email", "typo@example.com"],
            get_free_busy=lambda **_kw: {
                "email": "typo@example.com",
                "resolved": False,
                "error": "Could not resolve email address",
            },
        )

        assert code == 1
        assert json.loads(out) == {
            "status": "error",
            "message": "Could not resolve email address: typo@example.com",
        }
