"""
Unit tests for get_folder_by_name's subfolder recursion (_folder_by_path,
_find_folder_by_name) and the not-found warning added to every caller that
resolves a folder by name (list_emails, search_by_sender, get_inbox_stats),
via the shared _resolve_folder helper.

get_folder_by_name previously only checked an account root's *direct*
children, so a real subfolder (e.g. "Databricks Alerts" nested under
"Inbox") could never be resolved by name — and its callers silently fell
back to Inbox instead of surfacing that the folder wasn't found.

These tests are pure-Python (no Outlook / pywin32 required).
"""

import pytest

from mailtool.bridge import (
    FOLDER_SEARCH_MAX_DEPTH,
    MailtoolError,
    OutlookBridge,
    _find_folder_by_name,
    _folder_by_path,
)

pytestmark = pytest.mark.unit


class FakeFoldersCollection:
    """Minimal stand-in for Outlook's Folders collection: supports
    subscripting by name (case-sensitive, mirroring the real COM object)
    and iteration."""

    def __init__(self, children):
        self._children = list(children)

    def __getitem__(self, name):
        for c in self._children:
            if c.Name == name:
                return c
        raise KeyError(name)

    def __iter__(self):
        return iter(self._children)


class FakeMailFolder:
    def __init__(self, name, children=None):
        self.Name = name
        self._children = list(children or [])

    @property
    def Folders(self):  # noqa: N802 - mirrors Outlook COM API
        return FakeFoldersCollection(self._children)


def make_tree():
    """Inbox
    ├── Databricks Alerts
    │   └── Quality Alerts
    └── Unimportant
    """
    quality_alerts = FakeMailFolder("Quality Alerts")
    databricks_alerts = FakeMailFolder("Databricks Alerts", [quality_alerts])
    unimportant = FakeMailFolder("Unimportant")
    inbox = FakeMailFolder("Inbox", [databricks_alerts, unimportant])
    return inbox, databricks_alerts, quality_alerts, unimportant


class TestFindFolderByName:
    def test_finds_direct_child(self):
        inbox, _, _, unimportant = make_tree()
        assert _find_folder_by_name(inbox, "Unimportant", 8) is unimportant

    def test_finds_nested_grandchild(self):
        inbox, _, quality_alerts, _ = make_tree()
        assert _find_folder_by_name(inbox, "Quality Alerts", 8) is quality_alerts

    def test_case_insensitive(self):
        inbox, databricks_alerts, _, _ = make_tree()
        assert _find_folder_by_name(inbox, "databricks alerts", 8) is databricks_alerts

    def test_not_found_returns_none(self):
        inbox, _, _, _ = make_tree()
        assert _find_folder_by_name(inbox, "Nope", 8) is None

    def test_depth_bound_stops_search(self):
        # Quality Alerts is 2 levels below inbox; max_depth=1 only reaches
        # inbox's direct children (Databricks Alerts, Unimportant).
        inbox, _, _, _ = make_tree()
        assert _find_folder_by_name(inbox, "Quality Alerts", 1) is None

    def test_default_depth_constant_is_generous_enough(self):
        inbox, _, quality_alerts, _ = make_tree()
        assert (
            _find_folder_by_name(inbox, "Quality Alerts", FOLDER_SEARCH_MAX_DEPTH)
            is quality_alerts
        )

    def test_max_visited_bounds_a_wide_search(self):
        # 10 direct children; target is the last one visited. A max_visited
        # of 3 must give up before reaching it, even though max_depth (8)
        # would otherwise allow the search to continue.
        children = [FakeMailFolder(f"Child{i}") for i in range(10)]
        root = FakeMailFolder("Root", children)
        assert _find_folder_by_name(root, "Child9", 8, max_visited=3) is None
        assert _find_folder_by_name(root, "Child9", 8, max_visited=20) is children[9]


class TestFolderByPath:
    def test_walks_multi_segment_path(self):
        inbox, _, quality_alerts, _ = make_tree()
        result = _folder_by_path(inbox, ["Databricks Alerts", "Quality Alerts"])
        assert result is quality_alerts

    def test_single_segment_path(self):
        inbox, databricks_alerts, _, _ = make_tree()
        assert _folder_by_path(inbox, ["Databricks Alerts"]) is databricks_alerts

    def test_case_insensitive_segments(self):
        inbox, _, quality_alerts, _ = make_tree()
        result = _folder_by_path(inbox, ["databricks alerts", "quality alerts"])
        assert result is quality_alerts

    def test_missing_intermediate_segment_returns_none(self):
        inbox, _, _, _ = make_tree()
        assert _folder_by_path(inbox, ["Nope", "Quality Alerts"]) is None

    def test_missing_final_segment_returns_none(self):
        inbox, _, _, _ = make_tree()
        assert _folder_by_path(inbox, ["Databricks Alerts", "Nope"]) is None


def make_bridge_with_root(root_folder):
    bridge = OutlookBridge.__new__(OutlookBridge)  # skip COM __init__
    bridge._get_root = lambda: root_folder
    bridge.namespace = FakeFoldersCollection([])  # .Folders.Count -> error path
    return bridge


class TestGetFolderByName:
    def test_resolves_nested_subfolder_by_bare_name(self):
        inbox, _, quality_alerts, _ = make_tree()
        bridge = make_bridge_with_root(inbox)
        assert bridge.get_folder_by_name("Quality Alerts") is quality_alerts

    def test_resolves_backslash_path(self):
        inbox, _, quality_alerts, _ = make_tree()
        bridge = make_bridge_with_root(inbox)
        result = bridge.get_folder_by_name("Databricks Alerts\\Quality Alerts")
        assert result is quality_alerts

    def test_returns_none_for_unknown_name(self):
        inbox, _, _, _ = make_tree()
        bridge = make_bridge_with_root(inbox)
        assert bridge.get_folder_by_name("Nope") is None

    def test_empty_name_returns_none(self):
        inbox, _, _, _ = make_tree()
        bridge = make_bridge_with_root(inbox)
        assert bridge.get_folder_by_name("") is None
        assert bridge.get_folder_by_name(None) is None

    def test_separators_only_path_returns_none(self):
        # "\\" or "\\\\" pass the truthiness check but every segment is
        # empty after filtering — must not raise IndexError on path_parts[0].
        inbox, _, _, _ = make_tree()
        bridge = make_bridge_with_root(inbox)
        assert bridge.get_folder_by_name("\\") is None
        assert bridge.get_folder_by_name("\\\\") is None


class FakeItems:
    def __init__(self, items=None):
        self._items = list(items or [])

    def Restrict(self, _filter_str):  # noqa: N802 - mirrors Outlook COM API
        return self

    def Sort(self, _prop, _descending=False):  # noqa: N802 - mirrors Outlook COM API
        pass

    def __iter__(self):
        return iter(self._items)


class TestFolderNotFoundRaises:
    """A folder that can't be found is an error, never a silent fallback to
    Inbox (which would report another folder's mail as the requested one)."""

    @staticmethod
    def make_bridge(monkeypatch):
        bridge = OutlookBridge.__new__(OutlookBridge)
        monkeypatch.setattr(bridge, "get_folder_by_name", lambda _name: None)
        inbox_folder = type("F", (), {"Items": FakeItems()})()
        bridge.get_inbox = lambda: inbox_folder
        return bridge

    def test_list_emails_raises(self, monkeypatch):
        bridge = self.make_bridge(monkeypatch)

        with pytest.raises(MailtoolError, match="Folder not found: Nope"):
            bridge.list_emails(limit=5, folder="Nope")

    def test_search_by_sender_raises(self, monkeypatch):
        bridge = self.make_bridge(monkeypatch)

        with pytest.raises(MailtoolError, match="Folder not found: Nope"):
            bridge.search_by_sender("someone@example.com", folder="Nope")

    def test_missing_inbox_raises(self):
        bridge = OutlookBridge.__new__(OutlookBridge)
        bridge.get_inbox = lambda: None

        with pytest.raises(MailtoolError, match="Inbox folder not found"):
            bridge.list_emails()


class FakeCountableItems(FakeItems):
    """FakeItems with a .Count, for get_inbox_stats (which never iterates)."""

    def __init__(self, count):
        super().__init__()
        self.Count = count

    def Restrict(self, _filter_str):  # noqa: N802 - mirrors Outlook COM API
        return self


class TestGetInboxStatsNotFoundReporting:
    def test_unknown_folder_raises_instead_of_reporting_inbox_counts(self, monkeypatch):
        bridge = OutlookBridge.__new__(OutlookBridge)
        monkeypatch.setattr(bridge, "get_folder_by_name", lambda _name: None)
        inbox_folder = type("F", (), {"Items": FakeCountableItems(7)})()
        bridge.get_inbox = lambda: inbox_folder

        with pytest.raises(MailtoolError, match="Folder not found: Nope"):
            bridge.get_inbox_stats(folder="Nope")

    def test_resolved_folder_reports_its_own_name(self, monkeypatch):
        bridge = OutlookBridge.__new__(OutlookBridge)
        target_folder = type("F", (), {"Items": FakeCountableItems(3)})()
        monkeypatch.setattr(bridge, "get_folder_by_name", lambda _name: target_folder)

        result = bridge.get_inbox_stats(folder="Databricks Alerts")

        assert result["folder"] == "Databricks Alerts"
        assert result["total"] == 3
