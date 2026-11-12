"""
Unit tests for O(1) item lookup in edit_appointment and for the narrowed
exception handling in _safe_get_attr.

Pure-Python: fake COM objects stand in for Outlook, so no pywin32 is needed.
"""

import pytest

from mailtool import bridge as bridge_module
from mailtool.bridge import COM_ATTR_ERRORS, MailtoolError, OutlookBridge

pytestmark = pytest.mark.unit


class FakeComError(Exception):
    """Stands in for pywintypes.com_error."""


class FakeAppointment:
    Class = 26

    def __init__(self, entry_id):
        self.EntryID = entry_id
        self.Subject = "old"
        self.Location = "old"
        self.saved = False

    def Save(self):  # noqa: N802 - mirrors Outlook COM API
        self.saved = True


class ScanForbiddenItems:
    """Items collection that fails the test if anything iterates it."""

    def __iter__(self):
        raise AssertionError("calendar.Items must not be scanned for an EntryID")


class FakeCalendar:
    Items = ScanForbiddenItems()


class FakeNamespace:
    def __init__(self, items):
        self.items = items
        self.lookups = []

    def GetItemFromID(self, entry_id):  # noqa: N802 - mirrors Outlook COM API
        self.lookups.append(entry_id)
        if entry_id not in self.items:
            raise FakeComError("not found")
        return self.items[entry_id]


def make_bridge(items):
    bridge = OutlookBridge.__new__(OutlookBridge)
    bridge.namespace = FakeNamespace(items)
    bridge.get_calendar = lambda: FakeCalendar()
    return bridge


def test_edit_appointment_uses_get_item_from_id():
    appt = FakeAppointment("abc")
    bridge = make_bridge({"abc": appt})

    assert bridge.edit_appointment("abc", subject="new", location="Room 1") is True

    assert bridge.namespace.lookups == ["abc"]
    assert appt.Subject == "new"
    assert appt.Location == "Room 1"
    assert appt.saved is True


def test_edit_appointment_unknown_id_raises():
    bridge = make_bridge({})

    with pytest.raises(MailtoolError, match="Appointment not found: missing"):
        bridge.edit_appointment("missing", subject="new")


def test_edit_appointment_rejects_other_item_types():
    email = FakeAppointment("mail-1")
    email.Class = 43  # olMail
    bridge = make_bridge({"mail-1": email})

    assert bridge.edit_appointment("mail-1", subject="new") is False

    assert email.Subject == "old"
    assert email.saved is False


class RaisingObj:
    def __init__(self, exc):
        self.exc = exc

    @property
    def Field(self):  # noqa: N802 - mirrors Outlook COM API
        raise self.exc


def test_com_attr_errors_includes_attribute_error():
    assert AttributeError in COM_ATTR_ERRORS


def test_safe_get_attr_returns_default_for_attribute_error():
    obj = RaisingObj(AttributeError("nope"))

    assert OutlookBridge._safe_get_attr(obj, "Field", "dflt") == "dflt"
    assert OutlookBridge._safe_get_attr(object(), "Missing", "dflt") == "dflt"


def test_safe_get_attr_returns_default_for_com_error(monkeypatch):
    monkeypatch.setattr(
        bridge_module, "COM_ATTR_ERRORS", (FakeComError, AttributeError)
    )
    obj = RaisingObj(FakeComError("com failure"))

    assert OutlookBridge._safe_get_attr(obj, "Field", "dflt") == "dflt"


@pytest.mark.parametrize(
    "exc", [TypeError("bug"), ValueError("bug"), RuntimeError("bug")]
)
def test_safe_get_attr_propagates_unexpected_errors(exc):
    obj = RaisingObj(exc)

    with pytest.raises(type(exc)):
        OutlookBridge._safe_get_attr(obj, "Field", "dflt")


def test_safe_get_attr_returns_real_value():
    class HasValue:
        Present = 5

    assert OutlookBridge._safe_get_attr(HasValue(), "Present", 0) == 5
