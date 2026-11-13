"""
Outlook COM Automation Bridge
Runs on Windows, callable from WSL2
Provides CLI interface to Outlook email and calendar via COM

Requirements (Windows):
    pip install pywin32

Usage:
    # List recent emails
    python outlook_com_bridge.py emails --limit 10

    # List calendar events
    python outlook_com_bridge.py calendar --days 7

    # Get email body
    python outlook_com_bridge.py email --id <entry_id>
"""

import contextlib
import os
import re
import sys
from collections import deque
from datetime import datetime
from typing import Any

from whenever import SYSTEM_TZ, Date, PlainDateTime, Time, ZonedDateTime

try:
    import win32com.client
except ImportError:
    # pywin32 is only required for live COM access. Allowing the module to import
    # without it means the pure-Python helpers (e.g. _clean_body_top, _SMTP_REGEX,
    # MAIL_ONLY_FILTER) can be unit-tested on any platform. Instantiating
    # OutlookBridge still requires pywin32 + a running Outlook on Windows.
    win32com = None

try:
    import pywintypes
except ImportError:
    pywintypes = None

# Exception types a COM attribute read can legitimately raise: a COM failure
# (pywintypes.com_error, when pywin32 is present) or a missing attribute.
# Anything else (e.g. a TypeError from a coding mistake) is a real defect and
# should propagate rather than look like an empty field.
COM_ATTR_ERRORS: tuple[type[BaseException], ...] = (
    (pywintypes.com_error, AttributeError) if pywintypes else (AttributeError,)
)

# Regex used to salvage an SMTP address out of a raw Exchange DN or display string
# when both GetExchangeUser() and the PropertyAccessor lookup fail.
_SMTP_REGEX = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

# Outlook OlMeetingStatus enum. There is no 2, and an invite received from
# someone else is 3, not "cancelled": cancellations are 5 and 7.
MEETING_STATUS_LABELS = {
    0: "NonMeeting",
    1: "Meeting",
    3: "Received",
    5: "Canceled",
    7: "ReceivedAndCanceled",
}

# Maximum depth (levels below an account root) get_folder_by_name will
# recurse into when searching for a bare folder name. Real mailbox trees
# rarely nest more than a handful of levels deep; this just bounds
# pathological cases.
FOLDER_SEARCH_MAX_DEPTH = 8

# Maximum total folders _find_folder_by_name will visit across the whole
# search, independent of depth. A depth bound alone doesn't limit the work
# for a wide tree — a mailbox with thousands of folders spread across a few
# levels can still visit all of them, and each visit is a synchronous COM
# call on the main thread. This caps the worst case regardless of shape.
FOLDER_SEARCH_MAX_VISITED = 500

# Outlook OlObjectClass values used to check that an EntryID points at the
# kind of item the caller expects.
OL_APPOINTMENT = 26
OL_TASK = 48

# Error codes that mean "no item has this EntryID". get_item_by_id returns None
# for these and raises for every other failure.
# - MAPI_E_NOT_FOUND (0x8004010F) and MAPI_E_INVALID_ENTRYID (0x80040107).
# - E_INVALIDARG (0x80070057): what a live Outlook raises for a string that isn't
#   a well-formed EntryID at all, such as "invalid_id_12345".
NOT_FOUND_HRESULTS = frozenset({-2147221233, -2147221241, -2147024809})

# What RecurrencePattern.GetOccurrence raises when no occurrence starts at the
# requested time (a deleted occurrence, or a time the series never had): E_FAIL
# inside DISP_E_EXCEPTION. Only checked around GetOccurrence, after the series
# itself was found, because E_FAIL elsewhere is a real failure.
GET_OCCURRENCE_MISSING_HRESULT = -2147467259

# Outlook OlRecurrenceState values.
OL_APPT_NOT_RECURRING = 0
OL_APPT_MASTER = 1
OL_APPT_OCCURRENCE = 2
OL_APPT_EXCEPTION = 3

# Every occurrence of a recurring meeting carries the series' EntryID, so an
# EntryID alone can only address the whole series. mailtool addresses one
# occurrence as "<series EntryID>@<start>", with the start as Outlook reports
# it. EntryIDs are hex, so "@" can't occur in one.
OCCURRENCE_ID_SEPARATOR = "@"

# whenever format patterns. Times in and out of mailtool are local wall clock,
# the same time Outlook shows.
WALL_CLOCK_PATTERN = "YYYY-MM-DD HH:mm:ss"
DATE_PATTERN = "YYYY-MM-DD"
# The start inside an occurrence id.
OCCURRENCE_START_PATTERN = "YYYYMMDD'T'HHmmss"
# Calendar Restrict filters compare to the minute.
DASL_DATETIME_PATTERN = "YYYY-MM-DD HH:mm"

# The time fields of a COM date value: lowercase on a stdlib datetime (what
# pywin32 returns), capitalised on older pywintypes Time objects.
COM_TIME_FIELDS = ("year", "month", "day", "hour", "minute", "second")

# Outlook OlResponseStatus enum.
RESPONSE_STATUS_LABELS = {
    0: "None",
    1: "Organizer",
    2: "Tentative",
    3: "Accepted",
    4: "Declined",
    5: "NotResponded",
}

# Outlook Restrict filter that captures IPM.Note and its subtypes (e.g. IPM.Note.SMIME)
# while excluding meeting notifications (IPM.Schedule.Meeting.*), post items, reports,
# etc. Uses the same half-open MessageClass range trick as list_calendar_events().
MAIL_ONLY_FILTER = "[MessageClass] >= 'IPM.Note' AND [MessageClass] < 'IPM.Note{'"

# Mail-only scoping in DASL syntax: MessageClass is MAPI property 0x001A (string
# flavour 0x001A001F). Verified against a live Exchange mailbox.
DASL_MAIL_ONLY_FILTER = (
    "(\"http://schemas.microsoft.com/mapi/proptag/0x001A001F\" >= 'IPM.Note' "
    "AND \"http://schemas.microsoft.com/mapi/proptag/0x001A001F\" < 'IPM.Note{')"
)

# Jet [Field] name -> DASL schema reference. Only fields with a verified DASL
# mapping are translated; anything else raises instead of silently matching
# nothing (bare names like "Subject" are NOT resolvable in DASL).
DASL_FIELD_MAP = {
    "subject": '"urn:schemas:httpmail:subject"',
    "sendername": '"urn:schemas:httpmail:sendername"',
    # NOTE: mapped via MAPI proptag (PR_SENDER_EMAIL_ADDRESS, 0x0C1F), not
    # urn:schemas:httpmail:senderemail. The httpmail property is unreliably
    # populated on Exchange Online mailboxes — LIKE queries against it can
    # silently match zero items even when [SenderEmailAddress] = '...'
    # (plain Jet, same underlying MAPI property) finds the message. Verified
    # against a live mailbox: same fix already applied to messageclass below
    # for the same reason.
    "senderemailaddress": '"http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"',
    "receivedtime": '"urn:schemas:httpmail:date"',
    # NOTE: inverted semantics — httpmail:read = 0 means UNREAD.
    "unread": '"urn:schemas:httpmail:read"',
    "hasattachments": '"urn:schemas:httpmail:hasattachment"',
    "messageclass": '"http://schemas.microsoft.com/mapi/proptag/0x001A001F"',
    "body": '"urn:schemas:httpmail:textdescription"',
    "to": '"urn:schemas:httpmail:to"',
    "cc": '"urn:schemas:httpmail:cc"',
    "importance": '"urn:schemas:httpmail:importance"',
}

# Maximum horizon (days) used when list_calendar_events(all_events=True) is
# requested. A truly unbounded scan expands every recurring series forever
# ("calendar bomb"), blocking Outlook's single-threaded COM apartment for
# minutes, and every other COM call queues behind it. all_events is therefore capped to a forward 1-year window.
CALENDAR_ALL_EVENTS_MAX_DAYS = 365

# Hard cap on the number of (expanded) occurrences list_calendar_events will
# iterate. Belt-and-braces protection against pathological calendars.
CALENDAR_MAX_ITEMS = 2000

# Items.Restrict uses Jet syntax by default, which has no LIKE operator (and
# this Outlook build also rejects the CI_* operators outright). LIKE is native
# to DASL, so any filter containing a LIKE clause is translated wholesale to
# DASL ("@SQL=" prefix): field references mapped via DASL_FIELD_MAP, boolean
# literals normalized, and [Unread] comparison values inverted to match
# httpmail:read semantics. Filters without LIKE pass through unchanged (pure
# Jet), so existing behaviour is untouched.
_JET_FIELD_RE = re.compile(r"\[\s*([A-Za-z][A-Za-z0-9_ ]*?)\s*\]")
_JET_LIKE_RE = re.compile(
    r"\[\s*[A-Za-z][A-Za-z0-9_ ]*?\s*\]\s+LIKE\s+'", re.IGNORECASE
)
# Jet fields this Outlook build rejects outright even without a LIKE clause
# — e.g. HasAttachments ("The property \"HasAttachments\" is unknown"),
# verified against a live mailbox, unlike most Jet fields, which just work.
# A filter referencing one of these must be routed through DASL translation
# regardless of whether it also contains a LIKE clause. Add a field name
# here (lowercase, matching DASL_FIELD_MAP's keys) if another one is found
# to have the same quirk — no new regex needed.
_JET_FORCE_DASL_FIELDS = frozenset({"hasattachments"})
_BOOL_CMP_RE = re.compile(r"(TRUE|FALSE)\b", re.IGNORECASE)
_UNREAD_CMP_RE = re.compile(r'("urn:schemas:httpmail:read")\s*(=|<>)\s*(TRUE|FALSE)')

# Single-quoted string literals (LIKE search patterns, date/string comparison
# values) must never be touched by the field/boolean substitutions above —
# e.g. '%[External]%' or '%TRUE%' is literal text to search for, not a field
# reference or a boolean keyword. Jet/DASL filters in this codebase don't use
# escaped quotes inside literals, so a non-greedy match to the next quote is
# sufficient.
_STRING_LITERAL_RE = re.compile(r"'[^']*'")
_LITERAL_PLACEHOLDER = "\x00LITERAL{}\x00"


def _jet_to_dasl(filter_query):
    """Translate a Jet-style filter string to DASL ("@SQL=" prefixed)."""

    def _field(match):
        name = re.sub(r"\s+", "", match.group(1)).lower()
        try:
            return DASL_FIELD_MAP[name]
        except KeyError:
            supported = ", ".join(f"[{k.title()}]" for k in DASL_FIELD_MAP)
            raise ValueError(
                f"Cannot translate [Field] {match.group(1)!r} to DASL for LIKE "
                f"filters. Supported fields: {supported}"
            ) from None

    # Mask out quoted literals before running the field/boolean substitutions
    # so bracketed text or TRUE/FALSE keywords inside a LIKE pattern (e.g.
    # '%[External]%', '%TRUE%') are never mistaken for filter syntax.
    literals = []

    def _mask(match):
        literals.append(match.group(0))
        return _LITERAL_PLACEHOLDER.format(len(literals) - 1)

    dasl = _STRING_LITERAL_RE.sub(_mask, filter_query)

    dasl = _JET_FIELD_RE.sub(_field, dasl)
    # [Unread] mapped to httpmail:read has INVERTED semantics: flip the literal.
    dasl = _UNREAD_CMP_RE.sub(
        lambda m: (
            f"{m.group(1)} {m.group(2)} {'0' if m.group(3).upper() == 'TRUE' else '1'}"
        ),
        dasl,
    )
    # Remaining TRUE/FALSE literals (HasAttachments etc.) become 1/0, which
    # DASL accepts for boolean properties.
    dasl = _BOOL_CMP_RE.sub(
        lambda m: "1" if m.group(1).upper() == "TRUE" else "0", dasl
    )

    # Restore the original literal text, untouched by the substitutions above.
    for index, literal in enumerate(literals):
        dasl = dasl.replace(_LITERAL_PLACEHOLDER.format(index), literal)

    return f"@SQL={dasl}"


def translate_filter(filter_query):
    """
    Prepare a user filter for Items.Restrict.

    Filters without a LIKE clause, and without a reference to a field in
    _JET_FORCE_DASL_FIELDS, are returned unchanged (Jet syntax, the
    historical path). Filters containing "LIKE 'pattern'" are translated to
    DASL so the wildcard pattern actually works: '%x%' contains, 'x%' starts
    with, '%x' ends with — SQL LIKE semantics, case-insensitive. A filter
    referencing a _JET_FORCE_DASL_FIELDS field (e.g. [HasAttachments]) is
    also routed through DASL translation even without a LIKE clause, since
    Jet rejects that property outright on this Outlook build.

    Raises ValueError when a LIKE filter references a field without a known
    DASL mapping.
    """
    if _JET_LIKE_RE.search(filter_query):
        return _jet_to_dasl(filter_query)
    # Mask quoted literals before scanning for forced fields, mirroring
    # _jet_to_dasl's own masking — otherwise bracketed text inside a literal
    # (e.g. [Body] = 'note: [HasAttachments] mentioned') would be mistaken
    # for a real field reference and force an unnecessary DASL translation.
    masked_query = _STRING_LITERAL_RE.sub("", filter_query)
    for match in _JET_FIELD_RE.finditer(masked_query):
        name = re.sub(r"\s+", "", match.group(1)).lower()
        if name in _JET_FORCE_DASL_FIELDS:
            return _jet_to_dasl(filter_query)
    return filter_query


def _child_folder_by_name(parent, name):
    """Return the direct child of `parent` named `name` (case-insensitive), or
    None if there isn't one.

    Raises MailtoolError if the children can't be listed: a failed direct
    lookup alone can't tell "absent" from "Outlook is busy", but a successful
    enumeration can.
    """
    try:
        return parent.Folders[name]
    except Exception:
        pass
    target = str(name).strip().lower()
    try:
        for child in parent.Folders:
            if str(child.Name).strip().lower() == target:
                return child
    except Exception as e:
        raise MailtoolError(
            f"Could not list subfolders while looking for folder '{name}': {e}"
        ) from e
    return None


def _folder_by_path(root, path_parts):
    """Walk a list of folder-name segments from `root` (e.g. ["Inbox",
    "Databricks Alerts"]). Each segment is matched case-insensitively.
    Returns the resolved folder, or None if any segment doesn't exist.
    Raises MailtoolError if a level can't be listed."""
    current = root
    for part in path_parts:
        current = _child_folder_by_name(current, part)
        if current is None:
            return None
    return current


def _find_folder_by_name(
    folder, target_name, max_depth, max_visited=FOLDER_SEARCH_MAX_VISITED
):
    """Breadth-first, case-insensitive search for `target_name` among the
    descendants of `folder`, bounded to `max_depth` levels below it and
    `max_visited` folders visited in total.

    The depth bound alone doesn't cap the amount of work for a wide tree: a
    mailbox with many folders at one level would still visit all of them.
    `max_visited` bounds the total regardless of shape, since each visit is a
    synchronous COM call on the main thread.

    Returns None only when the searched tree has no such folder. A subfolder
    that can't be listed or read is warned about on stderr, and if the search
    finds nothing, MailtoolError is raised (chained from the first such
    failure), because the folder may be in the part that couldn't be searched.
    Reaching `max_visited` without a match raises for the same reason.
    """
    target_lower = str(target_name).strip().lower()
    frontier = deque([(folder, 0)])
    visited = 0
    first_failure = None
    failures = 0
    while frontier and visited < max_visited:
        current, depth = frontier.popleft()
        try:
            children = list(current.Folders)
        except Exception as e:
            _warn_skipped("a folder whose subfolders can't be listed", e)
            first_failure = first_failure or e
            failures += 1
            continue
        for child in children:
            visited += 1
            try:
                if str(child.Name).strip().lower() == target_lower:
                    return child
            except Exception as e:
                _warn_skipped("a subfolder whose name can't be read", e)
                first_failure = first_failure or e
                failures += 1
                continue
            if depth + 1 < max_depth:
                frontier.append((child, depth + 1))
            if visited >= max_visited:
                break
    if visited >= max_visited:
        # Stopped at the cap: the folder may exist in the part not searched.
        raise MailtoolError(
            f"Folder '{target_name}' not found in the first {max_visited} "
            "folders searched; the search stopped there"
        )
    if first_failure is not None:
        raise MailtoolError(
            f"Folder '{target_name}' not found, but {failures} folder(s) could "
            f"not be searched (first error: {first_failure})"
        ) from first_failure
    return None


def _com_error_codes(error):
    """The error's own hresult plus the real Outlook code, when there is one.

    pywin32 reports Outlook's errors as DISP_E_EXCEPTION (-2147352567) with the
    real code in excepinfo[5].
    """
    codes = {getattr(error, "hresult", None)}
    excepinfo = getattr(error, "excepinfo", None)
    if excepinfo is not None and len(excepinfo) > 5:
        codes.add(excepinfo[5])
    return codes


def _is_not_found_error(error):
    """True if a COM error says no item has the requested EntryID."""
    return bool(_com_error_codes(error) & NOT_FOUND_HRESULTS)


def occurrence_id(series_id: str, start: PlainDateTime) -> str:
    """The id of the occurrence of series `series_id` that starts at `start`."""
    return (
        f"{series_id}{OCCURRENCE_ID_SEPARATOR}{start.format(OCCURRENCE_START_PATTERN)}"
    )


def to_outlook(wall: PlainDateTime) -> datetime:
    """Convert a local wall-clock time into the value an Outlook COM call takes.

    Outlook reads a time as local wall clock, but pywin32 shifts a naive
    datetime from local time to UTC on the way in, so the value comes out
    off by the UTC offset. Handing pywin32 the wall-clock time labelled UTC
    makes it pass through unchanged (verified against live Outlook).
    """
    return wall.assume_utc().to_stdlib()


def date_to_outlook(day: Date) -> datetime:
    """A date for an Outlook COM call, as midnight local wall clock.

    Midnight is safe: to_outlook passes the wall-clock time through unchanged,
    so the date can't slip across a day boundary.
    """
    return to_outlook(day.at(Time(0, 0)))


def from_outlook(value) -> PlainDateTime | None:
    """Read an Outlook COM time as local wall clock, or None if unreadable.

    Outlook returns local wall-clock time, which pywin32 labels with a UTC
    tzinfo; reading only the fields drops the label.
    """
    if not value:
        return None
    if isinstance(value, datetime):
        names = COM_TIME_FIELDS
    else:
        names = tuple(name.capitalize() for name in COM_TIME_FIELDS)
    try:
        return PlainDateTime(*(getattr(value, name) for name in names))
    except (AttributeError, TypeError, ValueError):
        return None


def format_outlook_time(value):
    """Output string for an Outlook COM time: "YYYY-MM-DD HH:MM:SS", or None."""
    wall = from_outlook(value)
    return wall.format(WALL_CLOCK_PATTERN) if wall else None


def format_outlook_date(value):
    """Output string for the date of an Outlook COM time: "YYYY-MM-DD", or None."""
    wall = from_outlook(value)
    return wall.date().format(DATE_PATTERN) if wall else None


def parse_wall_clock(text) -> PlainDateTime:
    """Parse a "YYYY-MM-DD HH:MM:SS" local wall-clock time (raises ValueError)."""
    return PlainDateTime.parse(text, pattern=WALL_CLOCK_PATTERN)


def parse_date(value) -> Date:
    """A "YYYY-MM-DD" string parsed into a Date (raises ValueError).

    A Date passes through, for library callers of get_free_busy.
    """
    if isinstance(value, Date):
        return value
    return Date.parse(value, pattern=DATE_PATTERN)


def now_wall_clock() -> PlainDateTime:
    """The current local wall-clock time, to compare with Outlook's times."""
    return ZonedDateTime.now(SYSTEM_TZ).to_plain()


def split_appointment_id(entry_id):
    """Split an appointment id into (EntryID, occurrence start or None).

    A plain EntryID has no start. A suffix that doesn't parse as a start leaves
    the whole string as the EntryID, so it fails the lookup like any bad id.
    """
    series_id, sep, start = entry_id.partition(OCCURRENCE_ID_SEPARATOR)
    if not sep:
        return entry_id, None
    try:
        return series_id, PlainDateTime.parse(start, pattern=OCCURRENCE_START_PATTERN)
    except ValueError:
        return entry_id, None


class MailtoolError(RuntimeError):
    """An Outlook/COM operation failed.

    Raised instead of returning an empty result or a False sentinel, so a
    failure can never be mistaken for "nothing found". A fetch by an entry ID
    that doesn't exist is not a failure: those methods return None.
    """


@contextlib.contextmanager
def _com_operation(action):
    """Re-raise any failure inside the block as MailtoolError("<action> failed: ...").

    MailtoolError raised inside the block passes through unchanged so its
    more specific message isn't wrapped twice.
    """
    try:
        yield
    except MailtoolError:
        raise
    except Exception as e:
        raise MailtoolError(f"{action} failed: {e}") from e


def _warn_skipped(what, error):
    """Report one item skipped inside a list loop; the call as a whole goes on."""
    print(f"Warning: skipped {what}: {error}", file=sys.stderr)


class OutlookBridge:
    """Bridge to Outlook application via COM"""

    @staticmethod
    def _safe_get_attr(obj, attr, default=None) -> Any:
        """
        Safely get an attribute from a COM object, handling COM errors gracefully

        Args:
            obj: COM object
            attr: Attribute name to get
            default: Default value if attribute access fails

        Returns:
            Attribute value or default
        """
        try:
            return getattr(obj, attr, default)
        except COM_ATTR_ERRORS:
            return default

    def __init__(self, default_account: str | None = None):
        """
        Connect to running Outlook instance or start it

        Launch Logic:
        1. Try GetActiveObject first (if Outlook is running)
        2. Fall back to Dispatch if Outlook is closed
        """
        if win32com is None:
            raise RuntimeError("pywin32 is not importable; OutlookBridge needs it")
        try:
            self.outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as e:
            # Outlook might not be running - try to launch it
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as launch_error:
                raise MailtoolError(
                    "Could not connect to or launch Outlook "
                    f"(attach failed: {e}; launch failed: {launch_error}). "
                    "Make sure Outlook is installed and you can launch it manually."
                ) from launch_error

        self.namespace = self.outlook.GetNamespace("MAPI")
        # Default account name and root folder (set by set_default_account or via init param)
        self.default_account_name = None
        self.default_root_folder = None

        # If provided, attempt to set the default account/store on init
        if default_account:
            self.set_default_account(default_account)

    # -- Helper methods for account/folder resolution -----------------
    def _find_root_by_name(self, acc_name: str):
        """Find and return the root folder object for an account by name (case-insensitive).

        Returns None if not found.
        """
        if not acc_name:
            return None
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    root = self.namespace.Folders.Item(i)
                    if str(root.Name).strip().lower() == str(acc_name).strip().lower():
                        return root
                except Exception:
                    continue

        # Fallback: try a reasonable range if Count isn't available
        for i in range(1, 10):
            try:
                root = self.namespace.Folders.Item(i)
                if str(root.Name).strip().lower() == str(acc_name).strip().lower():
                    return root
            except Exception:
                continue

        return None

    def _get_root(self):
        """Return the active root folder to use (default account root if set, else the first mailbox/root)."""
        if self.default_root_folder:
            return self.default_root_folder
        # Try DefaultStore if set
        try:
            default_store = getattr(self.namespace, "DefaultStore", None)
            if default_store:
                try:
                    root = default_store.GetRootFolder()
                    return root
                except Exception:
                    pass
        except Exception:
            pass

        # Fallback: first available root
        try:
            return self.namespace.Folders.Item(1)
        except Exception:
            # try a small range as a last resort
            for i in range(1, 10):
                try:
                    return self.namespace.Folders.Item(i)
                except Exception:
                    continue
        return None

    def _find_account_by_name(self, name: str):
        """Find an Outlook Account object by SMTP address or display name (case-insensitive)."""
        if not name:
            return None
        try:
            accounts = self.namespace.Accounts
            for acc in accounts:
                try:
                    if (
                        hasattr(acc, "SmtpAddress")
                        and str(acc.SmtpAddress).strip().lower()
                        == str(name).strip().lower()
                    ):
                        return acc
                    if (
                        hasattr(acc, "DisplayName")
                        and str(acc.DisplayName).strip().lower()
                        == str(name).strip().lower()
                    ):
                        return acc
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def set_default_account(self, acc_name: str):
        """
        Set the default account by name

        Args:
            acc_name: Account name to set as default

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If no account with that name exists.
        """
        root = self._find_root_by_name(acc_name)
        if not root:
            raise MailtoolError(f"Account not found: {acc_name}")
        self.default_account_name = acc_name
        self.default_root_folder = root
        # Best effort by design: DefaultStore only helps other COM calls that
        # rely on it; the bridge itself uses the root folder set above, so a
        # failure here must not undo a successful account selection.
        with contextlib.suppress(Exception):
            self.namespace.DefaultStore = root.Store
        return True

    def get_item_of_class(self, entry_id, olclass):
        """Like get_item_by_id, but None also when the item is another type.

        An EntryID for an email must not be read or edited as an appointment.
        """
        item = self.get_item_by_id(entry_id)
        if item is None:
            return None
        # Read Class directly: a failed read must raise, not pass for "wrong type".
        with _com_operation(f"read the class of item {entry_id}"):
            item_class = item.Class
        return item if item_class == olclass else None

    def _require_item(self, entry_id, kind, olclass=None):
        """Return the item for entry_id, raising MailtoolError if it doesn't exist
        (or isn't of Outlook class `olclass`, when given).

        For mutators: acting on a missing item is a failure, unlike a fetch,
        which returns None.
        """
        if olclass is None:
            item = self.get_item_by_id(entry_id)
        else:
            item = self.get_item_of_class(entry_id, olclass)
        if item is None:
            raise MailtoolError(f"{kind} not found: {entry_id}")
        return item

    def get_appointment_item(self, entry_id):
        """Resolve an appointment id to its AppointmentItem, or None if missing.

        `entry_id` is a plain EntryID (a single appointment, or a whole
        recurring series) or an occurrence id from list_calendar_events, which
        resolves to just that occurrence.
        """
        item_id, start = split_appointment_id(entry_id)
        item = self.get_item_of_class(item_id, OL_APPOINTMENT)
        if item is None or start is None:
            return item
        with _com_operation(f"look up occurrence {entry_id}"):
            # GetRecurrencePattern on a single appointment would start turning
            # it into a series, so check first.
            if not item.IsRecurring:
                return None
            pattern = item.GetRecurrencePattern()
            try:
                return pattern.GetOccurrence(to_outlook(start))
            except Exception as e:
                if GET_OCCURRENCE_MISSING_HRESULT in _com_error_codes(e):
                    return None
                raise

    def require_appointment(self, entry_id, kind, series):
        """Resolve `entry_id` for a mutator: the item to change, or raise.

        A recurring series is changed only when `series` is True, so a series
        id can't change every occurrence by accident. With `series`, an
        occurrence id also resolves to its whole series.
        """
        item = self.get_appointment_item(entry_id)
        if item is None:
            raise MailtoolError(f"{kind} not found: {entry_id}")
        with _com_operation(f"read the recurrence of {entry_id}"):
            state = item.RecurrenceState
        if state == OL_APPT_MASTER and not series:
            raise MailtoolError(
                f"{entry_id} is a recurring series, so this would change every "
                "occurrence. Use one occurrence's entry_id from `mailtool calendar`, "
                "or pass --series to change the whole series."
            )
        if series and state in (OL_APPT_OCCURRENCE, OL_APPT_EXCEPTION):
            # An occurrence id's EntryID part is the series' own EntryID.
            series_id, _ = split_appointment_id(entry_id)
            return self._require_item(series_id, kind, OL_APPOINTMENT)
        return item

    def appointment_ids(self, item):
        """(entry_id, series_id) for an appointment.

        entry_id addresses exactly this item: an occurrence id for one
        occurrence of a series. series_id is the series' EntryID for anything
        recurring, None otherwise.
        """
        entry_id = item.EntryID
        state = item.RecurrenceState
        if state in (OL_APPT_OCCURRENCE, OL_APPT_EXCEPTION):
            start = from_outlook(item.Start)
            if start is None:
                raise MailtoolError(
                    f"Can't read the start of an occurrence of {entry_id}"
                )
            return occurrence_id(entry_id, start), entry_id
        if state == OL_APPT_MASTER:
            return entry_id, entry_id
        return entry_id, None

    def get_inbox(self):
        """Get the inbox folder.

        Raises:
            MailtoolError: If Outlook can't produce it.
        """
        # Prefer default account root when set
        root = self._get_root()
        if root:
            found = _child_folder_by_name(root, "Inbox")
            if found is not None:
                return found

        # Fallback to namespace default
        with _com_operation("get the Inbox folder"):
            return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    def list_folders(self, acc_name: str | None = None) -> dict[str, list[dict]]:
        """
        Recursively list all folders for all accounts.
        Args:
            acc_name: Specific account name to list folders for (optional)
        Returns:
            Dict with account names as keys and list of folder info dicts as values

        Raises:
            MailtoolError: If the account isn't found or reading a folder's
                details fails.
        """

        def retrieve_folder_details(folder, parent_folder, depth):
            all_items = [
                {
                    "name": folder.Name,
                    "id": folder.EntryID,
                    "parent_name": parent_folder.Name if parent_folder else None,
                    "parent_id": parent_folder.EntryID if parent_folder else None,
                    "number_of_items": folder.Items.Count,
                    "path": folder.FolderPath,
                    "depth": depth,
                    "account": parent_folder.Name if parent_folder else None,
                }
            ]
            for subfolder in folder.Folders:
                all_items.extend(retrieve_folder_details(subfolder, folder, depth + 1))
            return all_items

        final = {}
        with _com_operation("list folders"):
            accounts = self.namespace.Folders
            for i in range(1, accounts.Count + 1):
                parent_folder = accounts.Item(i)
                if acc_name and parent_folder.Name != acc_name:
                    continue
                final[parent_folder.Name] = retrieve_folder_details(
                    parent_folder, None, 0
                )

        if acc_name and not final:
            raise MailtoolError(f"Account not found: {acc_name}")
        return final

    def search_account_roots(self, search):
        """Call `search(root)` on each account root until it returns a folder.

        Returns None if no root has it. A root that fails is warned about; if
        nothing is found and any root failed, raises MailtoolError chained from
        the first failure, since the folder may live in the root that failed.
        """
        with _com_operation("list account roots"):
            count = self.namespace.Folders.Count
        first_failure = None
        for i in range(1, count + 1):
            try:
                found = search(self.namespace.Folders.Item(i))
            except Exception as e:
                _warn_skipped(f"account root {i}", e)
                first_failure = first_failure or e
                continue
            if found is not None:
                return found
        if first_failure is not None:
            raise MailtoolError(
                f"Folder not found, and an account root failed: {first_failure}"
            ) from first_failure
        return None

    def get_calendar(self):
        """Get the calendar folder.

        Returns:
            The folder, or None if no account has a Calendar folder

        Raises:
            MailtoolError: If the lookup itself fails (Outlook busy, store
                unavailable), which is not the same as the folder being absent.
        """
        # Prefer default account root when set
        root = self._get_root()
        if root:
            found = _child_folder_by_name(root, "Calendar")
            if found is not None:
                return found

        # Fallback: search all accounts
        return self.search_account_roots(
            lambda account_root: _child_folder_by_name(account_root, "Calendar")
        )

    def get_tasks(self):
        """Get the tasks folder.

        Raises:
            MailtoolError: If Outlook can't produce it.
        """
        root = self._get_root()
        if root:
            found = _child_folder_by_name(root, "Tasks")
            if found is not None:
                return found

        with _com_operation("get the Tasks folder"):
            return self.namespace.GetDefaultFolder(13)  # 13 = olFolderTasks

    def get_folder_by_name(self, folder_name):
        """
        Get a folder by name or by a "\\"-separated path from an account
        root (e.g. "Sent Items", "Archive", or "Inbox\\Databricks Alerts"
        for a subfolder).

        A bare name is resolved with a bounded recursive search across the
        whole folder tree (see FOLDER_SEARCH_MAX_DEPTH) — not just an
        account root's direct children. Use a "\\"-separated path instead
        when two folders share a name under different parents.

        Args:
            folder_name: Folder name, or a "\\"-separated path

        Returns:
            Folder object, or None if there is no such folder

        Raises:
            MailtoolError: If part of the folder tree can't be searched and the
                folder wasn't found elsewhere.
        """
        if not folder_name:
            return None

        path_parts = [p for p in str(folder_name).split("\\") if p]
        if not path_parts:
            # e.g. folder_name == "\\" or "\\\\" — passes the truthiness
            # check above but every segment is empty after filtering.
            return None
        if len(path_parts) > 1:
            root = self._get_root()
            if root:
                found = _folder_by_path(root, path_parts)
                if found:
                    return found

            return self.search_account_roots(
                lambda parent: _folder_by_path(parent, path_parts)
            )

        target_name = path_parts[0]

        root = self._get_root()
        root_error = None
        if root:
            # A failed search of the preferred root mustn't stop the other
            # accounts from being searched; it's raised only if none has it.
            try:
                found = _find_folder_by_name(
                    root,
                    target_name,
                    FOLDER_SEARCH_MAX_DEPTH,
                    max_visited=FOLDER_SEARCH_MAX_VISITED,
                )
            except MailtoolError as e:
                root_error = e
            else:
                if found:
                    return found

        found = self.search_account_roots(
            lambda parent: _find_folder_by_name(
                parent,
                target_name,
                FOLDER_SEARCH_MAX_DEPTH,
                max_visited=FOLDER_SEARCH_MAX_VISITED,
            )
        )
        if found is None and root_error is not None:
            raise root_error
        return found

    def _resolve_folder(self, folder):
        """
        Resolve `folder` to a mail folder.

        Needed by every caller that accepts a folder name (list_emails,
        search_by_sender, get_inbox_stats) — centralized here so the lookup
        and its error can't drift between them. A missing folder is an error
        rather than a silent fallback to Inbox, which would report another
        folder's mail as if it were the one asked for.

        Args:
            folder: Folder name (or "Inbox")

        Returns:
            (mail_folder, folder_name)

        Raises:
            MailtoolError: If the folder (or the Inbox) can't be found.
        """
        if folder == "Inbox":
            inbox = self.get_inbox()
            if inbox is None:
                raise MailtoolError("Inbox folder not found")
            return inbox, "Inbox"

        mail_folder = self.get_folder_by_name(folder)
        if mail_folder is None:
            raise MailtoolError(f"Folder not found: {folder}")
        return mail_folder, folder

    def get_item_by_id(self, entry_id):
        """
        Get any Outlook item by EntryID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID

        Returns:
            Outlook item (MailItem, AppointmentItem, TaskItem, etc.), or None
            if Outlook says no item has this ID

        Raises:
            MailtoolError: If the lookup fails for any other reason (Outlook
                or RPC unavailable, ...), which is not the same as "not found".
        """
        try:
            return self.namespace.GetItemFromID(entry_id)
        except Exception as e:
            if _is_not_found_error(e):
                return None
            raise MailtoolError(f"Could not look up item {entry_id}: {e}") from e

    def resolve_smtp_address(self, mail_item):
        """
        Get the SMTP address of a mail item's sender.

        Robust against cached Exchange mode where Sender.GetExchangeUser() returns
        None. Resolution order for EX-type senders:
          1. Sender.GetExchangeUser().PrimarySmtpAddress   (original path)
          2. PropertyAccessor -> PidTagSenderSmtpAddress   (0x5D01001F)
          3. regex salvage of an SMTP token from SenderEmailAddress
          4. the raw SenderEmailAddress (may be an Exchange DN)

        Args:
            mail_item: Outlook MailItem (or compatible item with a Sender)

        Returns:
            SMTP email address string ("" if nothing could be resolved)
        """
        try:
            # Non-EX senders already carry an SMTP address; return as-is.
            sender_email_type = self._safe_get_attr(mail_item, "SenderEmailType", "")
            raw_address = self._safe_get_attr(mail_item, "SenderEmailAddress", "") or ""
            if sender_email_type and sender_email_type != "EX":
                return raw_address

            # EX path.
            sender = self._safe_get_attr(mail_item, "Sender")
            if sender is not None:
                try:
                    exchange_user = sender.GetExchangeUser()
                except Exception:
                    exchange_user = None
                if exchange_user is not None:
                    primary = self._safe_get_attr(exchange_user, "PrimarySmtpAddress")
                    if primary:
                        return primary

            # PropertyAccessor: PidTagSenderSmtpAddress (reliable on Outlook 2007+).
            try:
                prop_accessor = mail_item.PropertyAccessor
                smtp = prop_accessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
                )
                if smtp:
                    return smtp
            except Exception:
                pass

            # Last-ditch salvage: pull an SMTP-shaped token out of whatever we have.
            match = _SMTP_REGEX.search(raw_address)
            if match:
                return match.group(0)

            return raw_address
        except Exception:
            return self._safe_get_attr(mail_item, "SenderEmailAddress", "") or ""

    @staticmethod
    def _attachment_count(item):
        """Return the number of attachments on a COM item.

        A failed read raises rather than reporting 0, so it can't pass for "no
        attachments"; list loops skip that item with a warning.
        """
        with _com_operation("count attachments"):
            return item.Attachments.Count

    def _extract_attachments(self, item):
        """Build a list of attachment-metadata dicts from a COM item."""
        out = []
        with _com_operation("read attachments"):
            attachments = item.Attachments
            count = attachments.Count
        for i in range(1, count + 1):
            with _com_operation(f"read attachment {i}"):
                a = attachments.Item(i)
            out.append(
                {
                    "filename": self._safe_get_attr(a, "FileName", "") or "",
                    "size": self._safe_get_attr(a, "Size", 0) or 0,
                    "display_name": self._safe_get_attr(a, "DisplayName", "") or "",
                    "content_type": self._safe_get_attr(a, "ContentType", None),
                    "is_inline": bool(self._safe_get_attr(a, "IsInline", False)),
                }
            )
        return out

    @staticmethod
    def _clean_body_top(body, max_chars=1000):
        """Return the 'new' portion of an email body: text before quoted reply
        chains, forwarded headers, and signature blocks.

        Pure-Python heuristic so it is unit-testable without Outlook. Trims and
        collapses runs of blank lines, then caps to max_chars.
        """
        if not body:
            return ""
        text = body.replace("\r\n", "\n").replace("\r", "\n")
        kept = []
        for line in text.split("\n"):
            stripped = line.strip()
            low = stripped.lower()
            # Outlook/Outlook-style forwarded or original-message header blocks.
            if low.startswith(
                (
                    "-----original message",
                    "-----origineel bericht",
                    "-----message réenvoyé",
                    "----- doorgestuurd bericht",
                    "-----transcribed message",
                )
            ):
                break
            # Quoted line.
            if line.lstrip().startswith(">"):
                break
            # Signature separator (a run of underscores).
            if len(stripped) >= 5 and set(stripped) <= {"_"}:
                break
            # Reply header lines carrying an address.
            if low.startswith(("from:", "van:")) and ("@" in line or "<" in line):
                break
            if low.startswith(("to:", "cc:", "bcc:")) and "@" in line and kept:
                break
            if low.startswith(("sent:", "verzonden:")) and kept:
                break
            # "On <date> X wrote:" / "Op <date> schreef X:" footer lines.
            if (low.endswith("wrote:") or low.endswith("schreef:")) and len(low) < 120:
                break
            kept.append(line)
        cleaned = "\n".join(kept).strip()
        cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
        return cleaned[:max_chars]

    def _mail_item_to_dict(self, item, *, include_body=False):
        """Build an email dict from a COM item using safe accessors throughout.

        Works for MailItem as well as non-mail items (meeting notifications, etc.):
        fields that don't exist on the item type come back as defaults instead of
        raising, so callers can branch on 'message_class' rather than catch errors.
        """
        message_class = (
            self._safe_get_attr(item, "MessageClass", "IPM.Note") or "IPM.Note"
        )
        d = {
            "entry_id": self._safe_get_attr(item, "EntryID", "") or "",
            "subject": self._safe_get_attr(item, "Subject", "") or "",
            "sender": self.resolve_smtp_address(item),
            "sender_name": self._safe_get_attr(item, "SenderName", "") or "",
            "received_time": format_outlook_time(
                self._safe_get_attr(item, "ReceivedTime")
            ),
            "sent_time": format_outlook_time(self._safe_get_attr(item, "SentOn")),
            "unread": bool(self._safe_get_attr(item, "Unread", False)),
            "has_attachments": self._attachment_count(item) > 0,
            "message_class": message_class,
            "to": self._safe_get_attr(item, "To", "") or "",
            "cc": self._safe_get_attr(item, "CC", "") or "",
            "conversation_id": self._safe_get_attr(item, "ConversationID", None),
            "conversation_topic": self._safe_get_attr(item, "ConversationTopic", None),
        }
        if include_body:
            body = self._safe_get_attr(item, "Body", "") or ""
            html_body = self._safe_get_attr(item, "HTMLBody", "") or ""
            d["body"] = body
            d["html_body"] = html_body
            d["body_top"] = self._clean_body_top(body)
            d["bcc"] = self._safe_get_attr(item, "BCC", "") or ""
            d["attachments"] = self._extract_attachments(item)
        return d

    @staticmethod
    def restrict_to_mail(items, include_non_mail):
        """Filter `items` to real emails (IPM.Note*) unless the caller opts out.

        Restrict can fail on unusual folders; the unfiltered collection is
        returned then (a superset, never an empty result), with a warning.
        """
        if include_non_mail:
            return items
        try:
            return items.Restrict(MAIL_ONLY_FILTER)
        except Exception as e:
            print(
                f"Warning: mail-only filter failed ({e}); results may include "
                "non-mail items",
                file=sys.stderr,
            )
            return items

    def list_emails(self, limit=10, folder="Inbox", include_non_mail=False):
        """
        List emails from the specified folder.

        By default only real email items (MessageClass IPM.Note and subtypes) are
        returned; meeting notifications and other inbox item types are filtered out
        at the COM level for efficiency. Pass include_non_mail=True to include them.

        Args:
            limit: Maximum number of emails to return
            folder: Folder name (default: Inbox)
            include_non_mail: If True, also return non-mail items (meeting
                notifications, post items, etc.)

        Returns:
            List of email dictionaries (empty only when the folder has no
            matching mail)

        Raises:
            MailtoolError: If the folder can't be found or the COM calls fail.
        """
        inbox, _ = self._resolve_folder(folder)

        with _com_operation(f"list emails in {folder}"):
            items = inbox.Items

            items = self.restrict_to_mail(items, include_non_mail)

            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break
                try:
                    emails.append(self._mail_item_to_dict(item, include_body=False))
                    count += 1
                except Exception as e:
                    # One unreadable item must not sink the whole listing.
                    _warn_skipped("an unreadable email", e)

            return emails

    def get_email_body(self, entry_id):
        """
        Get the full body and metadata of an item by entry ID (O(1) direct access).

        Works for any item type that lives in the mailbox: a real email returns
        body/html_body and full headers; a non-mail item (e.g. a meeting
        notification) returns everything that is accessible plus its message_class
        so the caller can branch. Returns None only when no item matches the ID.

        Args:
            entry_id: Outlook EntryID of the email

        Returns:
            Email dictionary with body, or None if the item is not found

        Raises:
            MailtoolError: If the item exists but can't be read.
        """
        item = self.get_item_by_id(entry_id)
        if item is None:
            return None
        with _com_operation(f"read email {entry_id}"):
            return self._mail_item_to_dict(item, include_body=True)

    def get_email_bodies(self, entry_ids, include_body=True):
        """
        Bulk-fetch full details for many EntryIDs in a single call (O(1) each).

        Avoids the N+1 round-trip pattern of calling get_email_body once per item.

        Args:
            entry_ids: Iterable of Outlook EntryIDs
            include_body: If False, fetch only summary fields (faster)

        Returns:
            List of email dictionaries for items that were found (missing IDs are
            omitted, each with a warning on stderr; an item that exists but
            can't be read is skipped with a warning too)

        Raises:
            MailtoolError: If a lookup fails for a reason other than the item
                not existing.
        """
        results = []
        for entry_id in entry_ids or []:
            item = self.get_item_by_id(entry_id)
            if item is None:
                _warn_skipped(f"email {entry_id}", "not found")
                continue
            try:
                results.append(self._mail_item_to_dict(item, include_body=include_body))
            except Exception as e:
                _warn_skipped(f"unreadable email {entry_id}", e)
        return results

    @staticmethod
    def _calendar_range_filter(start_dt, end_dt):
        """
        Build the combined DASL filter for appointments overlapping
        [start_dt, end_dt]: the MessageClass range (appointments only, so
        meeting requests/responses cannot poison iteration) AND the date
        window, applied in ONE Restrict call.

        Dates use DASL with ISO literals ("YYYY-MM-DD HH:MM") because Jet
        date literals are parsed according to the user's Windows locale: on a
        Dutch-locale system "09/01/2026 12:00" silently parses as 9 January,
        yielding mysteriously empty results for windows whose day component
        is <= 12. DASL + ISO is locale-independent (verified live).
        """
        return (
            '@SQL=("http://schemas.microsoft.com/mapi/proptag/0x001A001F" '
            ">= 'IPM.Appointment' "
            'AND "http://schemas.microsoft.com/mapi/proptag/0x001A001F" '
            "< 'IPM.Appointment{') "
            'AND ("urn:schemas:calendar:dtstart" <= '
            f"'{end_dt.format(DASL_DATETIME_PATTERN)}') "
            'AND ("urn:schemas:calendar:dtend" >= '
            f"'{start_dt.format(DASL_DATETIME_PATTERN)}')"
        )

    def list_calendar_events(self, days=7, all_events=False):
        """
        List calendar events for the next N days

        Args:
            days: Number of days ahead to look
            all_events: If True, look ahead further than `days`: from now
                through +CALENDAR_ALL_EVENTS_MAX_DAYS (365). A truly unbounded
                scan would expand every recurring series forever and block
                Outlook's COM apartment, wedging the server.

        Returns:
            List of event dictionaries

        Raises:
            MailtoolError: If the calendar folder can't be found or the COM
                calls fail.
        """
        calendar = self.get_calendar()
        if calendar is None:
            raise MailtoolError("Calendar folder not found")
        with _com_operation("list calendar events"):
            return self.collect_calendar_events(calendar, days, all_events)

    def collect_calendar_events(self, calendar, days, all_events):
        """Run the restricted calendar query and build the event dicts."""
        items = calendar.Items

        # Canonical order per Microsoft docs: Sort -> IncludeRecurrences ->
        # a single Restrict. The old code applied a second Restrict AFTER
        # setting IncludeRecurrences, which drops recurrence expansion on the
        # re-derived collection and yields flaky (often empty) results for
        # some date windows.
        items.Sort("[Start]")  # ascending, required for recurrence expansion
        items.IncludeRecurrences = True

        start_dt = now_wall_clock()
        if all_events:
            end_dt = start_dt.add(days=CALENDAR_ALL_EVENTS_MAX_DAYS)
        else:
            days = max(1, min(days, CALENDAR_ALL_EVENTS_MAX_DAYS))
            end_dt = start_dt.add(days=days)

        # Single COM-level Restrict: appointments only + overlap window.
        # This is also what prevents the "Calendar Bomb" (infinite recurring
        # items) from ever reaching Python iteration.
        items = items.Restrict(self._calendar_range_filter(start_dt, end_dt))

        events = []
        scanned = 0
        for item in items:
            scanned += 1
            if scanned > CALENDAR_MAX_ITEMS:
                print(
                    f"Warning: calendar listing truncated at {CALENDAR_MAX_ITEMS} items",
                    file=sys.stderr,
                )
                break
            try:
                # Use safe attribute access to handle COM errors
                start = self._safe_get_attr(item, "Start")
                end = self._safe_get_attr(item, "End")

                # Python-level overlap re-check (defence in depth; the COM
                # filter above should already have scoped correctly).
                start_wall = from_outlook(start)
                end_wall = from_outlook(end)
                if start_wall is None:
                    _warn_skipped("a calendar item with no readable start", start)
                    continue
                if start_wall > end_dt:
                    continue
                if end_wall is not None and end_wall < start_dt:
                    continue

                # Get attendees (safe access)
                required_attendees = self._safe_get_attr(item, "RequiredAttendees", "")
                optional_attendees = self._safe_get_attr(item, "OptionalAttendees", "")

                # Get meeting status
                response_status = self._safe_get_attr(item, "ResponseStatus")

                meeting_status = self._safe_get_attr(item, "MeetingStatus")
                entry_id, series_id = self.appointment_ids(item)

                event = {
                    "entry_id": entry_id,
                    "series_id": series_id,
                    "subject": self._safe_get_attr(item, "Subject", "(No Subject)"),
                    "start": start_wall.format(WALL_CLOCK_PATTERN),
                    "end": end_wall.format(WALL_CLOCK_PATTERN) if end_wall else None,
                    "location": self._safe_get_attr(item, "Location", ""),
                    "organizer": self._safe_get_attr(item, "Organizer"),
                    "all_day": self._safe_get_attr(item, "AllDayEvent", False),
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": RESPONSE_STATUS_LABELS.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": MEETING_STATUS_LABELS.get(
                        meeting_status, "Unknown"
                    ),
                    "response_requested": self._safe_get_attr(
                        item, "ResponseRequested", False
                    ),
                }
                events.append(event)
            except Exception as e:
                # One bad occurrence must not sink the whole listing.
                _warn_skipped("an unreadable calendar item", e)

        return events

    def use_default_send_account(self, mail):
        """Send `mail` from the default account when one is set.

        Best effort by design: if the account can't be found or applied, Outlook
        sends from its own default account rather than the send failing.
        """
        if not self.default_account_name:
            return
        try:
            acc = self._find_account_by_name(self.default_account_name)
            if acc:
                mail.SendUsingAccount = acc
        except Exception:
            pass

    def send_email(
        self,
        to,
        subject,
        body,
        cc=None,
        bcc=None,
        html_body=None,
        file_paths=None,
        save_draft=False,
    ):
        """
        Send an email (or save as draft)

        Args:
            to: Recipient email address
            subject: Email subject
            body: Email body (plain text)
            cc: CC recipients (optional)
            bcc: BCC recipients (optional)
            html_body: HTML body (optional)
            file_paths: List of file paths to attach (optional)
            save_draft: If True, save to Drafts instead of sending

        Returns:
            Draft entry ID if saved, True if sent

        Raises:
            MailtoolError: If the email can't be built, attached, saved or sent.
        """
        with _com_operation("send email"):
            # If saving to drafts, create the mail in the default account's Drafts folder when available
            root = self._get_root()
            mail = None
            if save_draft and root:
                try:
                    drafts = None
                    try:
                        drafts = root.Folders["Drafts"]
                    except Exception:
                        for f in root.Folders:
                            if str(f.Name).strip().lower() == "drafts":
                                drafts = f
                                break
                    if drafts:
                        mail = drafts.Items.Add()
                except Exception:
                    mail = None

            if mail is None:
                mail = self.outlook.CreateItem(0)  # 0 = olMailItem

            mail.To = to
            mail.Subject = subject
            if html_body:
                mail.HTMLBody = html_body
            else:
                mail.Body = body
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            # Add attachments (a missing attachment must not be sent without it)
            if file_paths:
                for file_path in file_paths:
                    mail.Attachments.Add(file_path)

            self.use_default_send_account(mail)

            if save_draft:
                mail.Save()
                return mail.EntryID
            else:
                mail.Send()
                return True

    def reply_email(self, entry_id, body, reply_all=False):
        """
        Reply to an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            body: Reply body
            reply_all: True to reply all, False to reply sender only

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the email doesn't exist or the reply can't be sent.
        """
        item = self._require_item(entry_id, "Email")
        with _com_operation("reply to email"):
            if reply_all:
                reply = item.ReplyAll()
            else:
                reply = item.Reply()
            reply.Body = body
            self.use_default_send_account(reply)
            reply.Send()
            return True

    def forward_email(self, entry_id, to, body=""):
        """
        Forward an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            to: Recipient to forward to
            body: Optional additional body text

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the email doesn't exist or can't be forwarded.
        """
        item = self._require_item(entry_id, "Email")
        with _com_operation("forward email"):
            forward = item.Forward()
            forward.To = to
            if body:
                forward.Body = body + "\n\n" + forward.Body
            self.use_default_send_account(forward)
            forward.Send()
            return True

    def mark_email_read(self, entry_id, unread=False):
        """
        Mark an email as read or unread (O(1) direct access)

        Args:
            entry_id: Email entry ID
            unread: True to mark as unread, False to mark as read

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the email doesn't exist or can't be updated.
        """
        item = self._require_item(entry_id, "Email")
        with _com_operation("mark email"):
            item.Unread = unread
            item.Save()
            return True

    def move_email(self, entry_id, folder_name):
        """
        Move an email to a different folder (O(1) direct access)

        Args:
            entry_id: Email entry ID
            folder_name: Target folder name

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the email or the target folder doesn't exist, or
                the move fails.
        """
        item = self._require_item(entry_id, "Email")
        target_folder = self.get_folder_by_name(folder_name)
        if target_folder is None:
            raise MailtoolError(f"Folder not found: {folder_name}")
        with _com_operation("move email"):
            item.Move(target_folder)
            return True

    def delete_email(self, entry_id):
        """
        Delete an email (O(1) direct access)

        Args:
            entry_id: Email entry ID

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the email doesn't exist or can't be deleted.
        """
        item = self._require_item(entry_id, "Email")
        with _com_operation("delete email"):
            item.Delete()
            return True

    def download_attachments(self, entry_id, download_dir):
        """
        Download all attachments from an email

        Args:
            entry_id: Email entry ID
            download_dir: Directory to save attachments

        Returns:
            List of downloaded file paths (empty when the email has no
            attachments)

        Raises:
            MailtoolError: If the email doesn't exist or a download fails.
        """
        item = self._require_item(entry_id, "Email")
        with _com_operation("download attachments"):
            if item.Attachments.Count == 0:
                return []

            downloaded = []
            os.makedirs(download_dir, exist_ok=True)

            for i in range(item.Attachments.Count):
                attachment = item.Attachments.Item(i + 1)  # COM is 1-indexed
                filename = attachment.FileName
                filepath = os.path.join(download_dir, filename)
                attachment.SaveAsFile(filepath)
                downloaded.append(filepath)
            return downloaded

    def create_appointment(
        self,
        subject,
        start,
        end,
        location="",
        body="",
        all_day=False,
        required_attendees=None,
        optional_attendees=None,
    ):
        """
        Create a calendar appointment

        Args:
            subject: Appointment subject
            start: Start time (YYYY-MM-DD HH:MM:SS)
            end: End time (YYYY-MM-DD HH:MM:SS)
            location: Location
            body: Appointment body/description
            all_day: True for all-day event
            required_attendees: Semicolon-separated list of required attendees
            optional_attendees: Semicolon-separated list of optional attendees

        Returns:
            Appointment entry ID

        Raises:
            MailtoolError: If the appointment can't be created or saved.
        """
        with _com_operation("create appointment"):
            # Prefer creating the appointment in the default account's Calendar folder
            calendar = self.get_calendar()
            if calendar:
                try:
                    appointment = calendar.Items.Add()
                except Exception:
                    appointment = self.outlook.CreateItem(1)  # fallback
            else:
                appointment = self.outlook.CreateItem(1)  # 1 = olAppointmentItem
            appointment.Subject = subject
            appointment.Start = to_outlook(parse_wall_clock(start))
            appointment.End = to_outlook(parse_wall_clock(end))
            appointment.Location = location
            appointment.Body = body
            appointment.AllDayEvent = all_day
            if required_attendees:
                appointment.RequiredAttendees = required_attendees
            if optional_attendees:
                appointment.OptionalAttendees = optional_attendees
            appointment.Save()
            return appointment.EntryID

    def edit_appointment(
        self,
        entry_id,
        required_attendees=None,
        optional_attendees=None,
        subject=None,
        start=None,
        end=None,
        location=None,
        body=None,
        series=False,
    ):
        """
        Edit an existing appointment

        Moving an occurrence changes its entry_id, since the id holds its start.

        Args:
            entry_id: Appointment entry ID, or an occurrence id from
                list_calendar_events to edit just that occurrence
            required_attendees: Comma-separated list of required attendees
            optional_attendees: Comma-separated list of optional attendees
            subject: New subject (optional)
            start: New start time (optional)
            end: New end time (optional)
            location: New location (optional)
            body: New body (optional)
            series: Required to edit a whole recurring series

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the appointment doesn't exist, is a series and
                `series` is False, or can't be saved.
        """
        item = self.require_appointment(entry_id, "Appointment", series)
        with _com_operation("edit appointment"):
            if required_attendees:
                item.RequiredAttendees = required_attendees
            if optional_attendees:
                item.OptionalAttendees = optional_attendees
            if subject:
                item.Subject = subject
            if start:
                item.Start = to_outlook(parse_wall_clock(start))
            if end:
                item.End = to_outlook(parse_wall_clock(end))
            if location is not None:
                item.Location = location
            if body is not None:
                item.Body = body
            item.Save()
            return True

    def get_appointment(self, entry_id):
        """
        Get full appointment details by entry ID (O(1) direct access)

        Args:
            entry_id: Appointment entry ID, or an occurrence id from
                list_calendar_events

        Returns:
            Appointment dictionary with full details, or None if not found

        Raises:
            MailtoolError: If the appointment exists but can't be read.
        """
        item = self.get_appointment_item(entry_id)
        if item:
            with _com_operation("read appointment"):
                item_entry_id, series_id = self.appointment_ids(item)
                required_attendees = (
                    item.RequiredAttendees if hasattr(item, "RequiredAttendees") else ""
                )
                optional_attendees = (
                    item.OptionalAttendees if hasattr(item, "OptionalAttendees") else ""
                )
                response_status = self._safe_get_attr(item, "ResponseStatus")
                meeting_status = self._safe_get_attr(item, "MeetingStatus")

                return {
                    "entry_id": item_entry_id,
                    "series_id": series_id,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "start": format_outlook_time(getattr(item, "Start", None)),
                    "end": format_outlook_time(getattr(item, "End", None)),
                    "location": item.Location if hasattr(item, "Location") else "",
                    "organizer": item.Organizer if hasattr(item, "Organizer") else None,
                    "body": item.Body if hasattr(item, "Body") else "",
                    "all_day": item.AllDayEvent
                    if hasattr(item, "AllDayEvent")
                    else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": RESPONSE_STATUS_LABELS.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": MEETING_STATUS_LABELS.get(
                        meeting_status, "Unknown"
                    ),
                    "response_requested": item.ResponseRequested
                    if hasattr(item, "ResponseRequested")
                    else False,
                }
        return None

    def respond_to_meeting(self, entry_id, response, series=False):
        """
        Respond to a meeting invitation (O(1) direct access)

        Args:
            entry_id: Appointment entry ID, or an occurrence id from
                list_calendar_events to respond to just that occurrence
            response: Response - "accept", "decline", "tentative"
            series: Required to respond to a whole recurring series

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the meeting doesn't exist, the response isn't one
                of accept/decline/tentative, or sending it fails.
        """
        response_map = {
            "accept": 3,  # olResponseAccepted
            "decline": 4,  # olResponseDeclined
            "tentative": 2,  # olResponseTentative
        }
        if response.lower() not in response_map:
            raise MailtoolError(
                f"Unknown meeting response '{response}' "
                f"(expected one of: {', '.join(response_map)})"
            )
        item = self.require_appointment(entry_id, "Meeting", series)
        with _com_operation("respond to meeting"):
            item.Response(response_map[response.lower()])
            item.Send()
            return True

    def delete_appointment(self, entry_id, series=False):
        """
        Delete an appointment (O(1) direct access)

        Args:
            entry_id: Appointment entry ID, or an occurrence id from
                list_calendar_events to delete just that occurrence
            series: Required to delete a whole recurring series

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the appointment doesn't exist, is a series and
                `series` is False, or can't be deleted.
        """
        item = self.require_appointment(entry_id, "Appointment", series)
        with _com_operation("delete appointment"):
            item.Delete()
            return True

    def list_tasks(self, include_completed=False):
        """
        List tasks (only incomplete by default)

        Args:
            include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

        Returns:
            List of task dictionaries

        Raises:
            MailtoolError: If the tasks folder can't be found or the COM calls
                fail.
        """
        tasks_folder = self.get_tasks()
        if tasks_folder is None:
            raise MailtoolError("Tasks folder not found")
        with _com_operation("list tasks"):
            return self.collect_tasks(tasks_folder, include_completed)

    def collect_tasks(self, tasks_folder, include_completed):
        """Iterate the Tasks folder and build the task dicts."""
        items = tasks_folder.Items

        tasks = []
        for item in items:
            try:
                # Skip completed tasks unless include_completed is True
                if not include_completed and item.Complete:
                    continue

                task = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": format_outlook_date(getattr(item, "DueDate", None)),
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
                tasks.append(task)
            except Exception as e:
                # One unreadable task must not sink the whole listing.
                _warn_skipped("an unreadable task", e)

        return tasks

    def list_all_tasks(self):
        """
        List all tasks including completed ones

        Returns:
            List of task dictionaries

        Raises:
            MailtoolError: As for list_tasks.
        """
        return self.list_tasks(include_completed=True)

    def create_task(self, subject, body="", due_date=None, importance=1):
        """
        Create a new task

        Args:
            subject: Task subject
            body: Task description
            due_date: Due date (YYYY-MM-DD) or None
            importance: 0=Low, 1=Normal, 2=High

        Returns:
            Task entry ID

        Raises:
            MailtoolError: If the task can't be created or saved.
        """
        with _com_operation("create task"):
            # Prefer creating the task in the default account's Tasks folder
            tasks_folder = self.get_tasks()
            if tasks_folder:
                try:
                    task = tasks_folder.Items.Add()
                except Exception:
                    task = self.outlook.CreateItem(3)
            else:
                task = self.outlook.CreateItem(3)  # 3 = olTaskItem
            task.Subject = subject
            task.Body = body
            if due_date:
                task.DueDate = date_to_outlook(parse_date(due_date))
            task.Importance = importance
            task.Save()
            return task.EntryID

    def get_task(self, entry_id):
        """
        Get full task details by entry ID (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            Task dictionary with full details, or None if not found

        Raises:
            MailtoolError: If the task exists but can't be read.
        """
        item = self.get_item_of_class(entry_id, OL_TASK)
        if item:
            with _com_operation("read task"):
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": format_outlook_date(getattr(item, "DueDate", None)),
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
        return None

    def edit_task(
        self,
        entry_id,
        subject=None,
        body=None,
        due_date=None,
        importance=None,
        percent_complete=None,
        complete=None,
    ):
        """
        Edit an existing task (O(1) direct access)

        Args:
            entry_id: Task entry ID
            subject: New subject (optional)
            body: New body (optional)
            due_date: New due date YYYY-MM-DD (optional)
            importance: New importance 0=Low, 1=Normal, 2=High (optional)
            percent_complete: New percent complete 0-100 (optional)
            complete: Mark complete/incomplete True/False (optional)

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the task doesn't exist or can't be saved.
        """
        item = self._require_item(entry_id, "Task", OL_TASK)
        with _com_operation("edit task"):
            if subject:
                item.Subject = subject
            if body is not None:
                item.Body = body
            if due_date:
                item.DueDate = date_to_outlook(parse_date(due_date))
            if importance is not None:
                item.Importance = importance
            if percent_complete is not None:
                item.PercentComplete = percent_complete
                # Update status based on percent_complete
                if percent_complete == 100:
                    item.Status = 2  # Complete
                    item.Complete = True
                elif percent_complete == 0:
                    item.Status = 0  # Not started
                else:
                    item.Status = 1  # In progress
            if complete is not None:
                item.Complete = complete
                if complete:
                    item.PercentComplete = 100
                    item.Status = 2
                else:
                    item.PercentComplete = 0
                    item.Status = 0
            item.Save()
            return True

    def complete_task(self, entry_id):
        """
        Mark a task as complete (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the task doesn't exist or can't be saved.
        """
        item = self._require_item(entry_id, "Task", OL_TASK)
        with _com_operation("complete task"):
            item.Complete = True
            item.PercentComplete = 100
            item.Status = 2  # olTaskComplete
            item.Save()
            return True

    def delete_task(self, entry_id):
        """
        Delete a task (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True. Failure raises instead of returning False.

        Raises:
            MailtoolError: If the task doesn't exist or can't be deleted.
        """
        item = self._require_item(entry_id, "Task", OL_TASK)
        with _com_operation("delete task"):
            item.Delete()
            return True

    def search_emails(self, filter_query, limit=100, include_non_mail=False):
        """
        Search emails using Outlook Restrict filter (O(1) search, no iteration).

        By default only real emails (MessageClass IPM.Note and subtypes) are
        returned; pass include_non_mail=True to also match meeting items etc.

        SQL-style LIKE patterns are accepted and translated automatically:
        Items.Restrict's default Jet syntax has no LIKE, so any filter with a
        LIKE clause is rewritten to DASL ("[Subject] LIKE '%meeting%'" runs as
        an @SQL= query with SQL LIKE wildcard semantics, case-insensitive).

        NOTE on dates in plain (non-LIKE) filters: Jet date literals are
        parsed according to the user's Windows locale — on e.g. a Dutch
        locale, '07/01/2026' means 7 January. Prefer unambiguous literals for
        your locale (e.g. '2026-07-01 00:00' or '01-07-2026' on NL locales)
        when filtering by ReceivedTime.

        Args:
            filter_query: Filter query string for filtering (e.g. "[Unread] = TRUE",
                "[Subject] LIKE '%meeting%'", or a date range — see the locale
                note above)
            limit: Max results to return
            include_non_mail: If True, do not scope the filter to IPM.Note items

        Returns:
            List of email dictionaries (empty only when nothing matches)

        Raises:
            MailtoolError: If the Inbox can't be found, a LIKE filter references
                a field without a known DASL mapping (supported: the fields in
                DASL_FIELD_MAP), Outlook rejects the filter (invalid syntax),
                or a COM call fails. A bad filter must not be indistinguishable
                from "no matches".
        """
        # Translate SQL-style LIKE filters to DASL so wildcards actually work.
        raw_query = filter_query or ""
        mentions_message_class = "messageclass" in raw_query.lower()
        try:
            prepared = translate_filter(raw_query)
        except ValueError as e:
            raise MailtoolError(str(e)) from e
        dasl_mode = prepared.startswith("@SQL=")

        folder, _ = self._resolve_folder("Inbox")
        with _com_operation("search emails"):
            items = folder.Items
        # Compose the effective filter. Unless the caller opts out of the
        # mail-only scope (or already mentioned MessageClass themselves),
        # AND in the IPM.Note range so meeting/post items are excluded.
        # DASL needs the "@SQL=" prefix at the very start of the whole string.
        scope = DASL_MAIL_ONLY_FILTER if dasl_mode else MAIL_ONLY_FILTER
        scope_needed = (
            bool(prepared) and not include_non_mail and (not mentions_message_class)
        )
        if prepared and scope_needed:
            if dasl_mode:
                effective_filter = f"@SQL=({prepared[len('@SQL=') :]}) AND {scope}"
            else:
                effective_filter = f"({prepared}) AND {scope}"
        elif prepared:
            effective_filter = prepared
        elif not include_non_mail:
            effective_filter = scope
        else:
            effective_filter = None

        if effective_filter:
            try:
                # Apply restriction filter
                items = items.Restrict(effective_filter)
            except Exception as e:
                msg = (
                    f"Outlook Restrict failed (filter={effective_filter!r}): {e}. "
                    "Note: Jet syntax is the default; LIKE filters are "
                    "translated to DASL automatically, other SQL-isms "
                    "(IN/GLOB/underscore wildcards) are not."
                )
                raise MailtoolError(msg) from e

        with _com_operation("search emails"):
            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break

                try:
                    emails.append(self._mail_item_to_dict(item, include_body=False))
                    count += 1
                except Exception as e:
                    # One unreadable item must not sink the whole search.
                    _warn_skipped("an unreadable email", e)

            return emails

    def search_by_sender(
        self, sender_email, limit=100, folder="Inbox", include_non_mail=False
    ):
        """
        Search emails by sender email address (handles Exchange addresses).

        This method properly handles both SMTP and Exchange email addresses.
        For Exchange users (internal emails), it resolves the Exchange address
        to SMTP address before matching.

        Args:
            sender_email: Email address to search for
            limit: Max results to return (default: 100)
            folder: Folder name to search in (default: "Inbox")
            include_non_mail: If True, also consider non-mail items

        Returns:
            List of email dictionaries matching the sender (empty only when
            nothing matches)

        Raises:
            MailtoolError: If the folder can't be found or the COM calls fail.
        """
        mail_folder, _ = self._resolve_folder(folder)
        with _com_operation("search emails by sender"):
            items = mail_folder.Items
            items = self.restrict_to_mail(items, include_non_mail)
            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            target = (sender_email or "").lower()
            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break

                try:
                    d = self._mail_item_to_dict(item, include_body=False)
                    # Case-insensitive email match on the resolved SMTP address
                    if d["sender"] and d["sender"].lower() == target:
                        emails.append(d)
                        count += 1
                except Exception as e:
                    # One unreadable item must not sink the whole search.
                    _warn_skipped("an unreadable email", e)

            return emails

    def get_inbox_stats(self, folder="Inbox"):
        """
        Return cheap total/unread counts for a folder without fetching items.

        Uses Restrict+Count at the COM level so it is O(1)-ish regardless of
        folder size. Useful for pagination decisions and inbox monitoring.

        Args:
            folder: Folder name (default: "Inbox")

        Returns:
            Dict with 'folder', 'total', and 'unread' integer counts

        Raises:
            MailtoolError: If the folder can't be found or a count can't be read.
        """
        mail_folder, queried_folder = self._resolve_folder(folder)
        with _com_operation("get inbox stats"):
            items = mail_folder.Items
            total = items.Count
            unread = items.Restrict("[Unread] = TRUE").Count
            return {
                "folder": queried_folder,
                "total": int(total),
                "unread": int(unread),
            }

    def get_free_busy(
        self, email_address=None, start_date=None, end_date=None, entry_id=None
    ):
        """
        Get free/busy status for an email address

        Args:
            email_address: Email address to check (optional, defaults to current user)
            start_date: Start date ("YYYY-MM-DD" or whenever Date; defaults to today)
            end_date: End date ("YYYY-MM-DD" or whenever Date; defaults to start + 1 day)
            entry_id: DEPRECATED - Appointment entry ID (legacy, use email_address instead)

        Returns:
            Dictionary with free/busy information. An address Outlook can't
            resolve is a normal answer, not a failure: the dict has
            resolved=False and an "error" description.

        Raises:
            MailtoolError: If the COM calls fail or a date can't be parsed.
        """
        with _com_operation("get free/busy"):
            # Handle legacy entry_id parameter (extract first required attendee)
            if entry_id and not email_address:
                item = self.get_item_by_id(entry_id)
                if item is None:
                    raise MailtoolError(f"Item not found: {entry_id}")
                if hasattr(item, "RequiredAttendees") and item.RequiredAttendees:
                    attendees = item.RequiredAttendees.split(";")
                    if attendees:
                        email_address = attendees[0].strip()

            # Default to current user if no email provided
            if not email_address:
                email_address = self.namespace.CurrentUser.Address

            # Default to today if no dates provided
            start_date = (
                parse_date(start_date) if start_date else now_wall_clock().date()
            )
            end_date = parse_date(end_date) if end_date else start_date.add(days=1)

            # Create recipient and get free/busy
            recipient = self.namespace.CreateRecipient(email_address)
            if recipient.Resolve():
                # FreeBusy returns a string with time slots and status
                # 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
                freebusy = recipient.FreeBusy(
                    date_to_outlook(start_date), 60 * 24
                )  # 1440 minutes = 1 day
                return {
                    "email": email_address,
                    "start_date": start_date.format(DATE_PATTERN),
                    "end_date": end_date.format(DATE_PATTERN),
                    "free_busy": freebusy,
                    "resolved": True,
                }
            else:
                return {
                    "email": email_address,
                    "start_date": start_date.format(DATE_PATTERN),
                    "end_date": end_date.format(DATE_PATTERN),
                    "error": "Could not resolve email address",
                    "resolved": False,
                }
