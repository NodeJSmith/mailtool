#!/usr/bin/env python3
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

# Modified to test pre-commit hook

import contextlib
import re
import sys
import traceback
from datetime import datetime, timedelta

try:
    import win32com.client
except ImportError:
    # pywin32 is only required for live COM access. Allowing the module to import
    # without it means the pure-Python helpers (e.g. _clean_body_top, _SMTP_REGEX,
    # MAIL_ONLY_FILTER) can be unit-tested on any platform. Instantiating
    # OutlookBridge still requires pywin32 + a running Outlook on Windows.
    win32com = None

# Regex used to salvage an SMTP address out of a raw Exchange DN or display string
# when both GetExchangeUser() and the PropertyAccessor lookup fail.
_SMTP_REGEX = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

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
# minutes — which wedges the entire MCP server, since every other COM call
# queues behind it. all_events is therefore capped to a forward 1-year window.
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
        lambda m: f"{m.group(1)} {m.group(2)} {'0' if m.group(3).upper() == 'TRUE' else '1'}",
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


class OutlookBridge:
    """Bridge to Outlook application via COM"""

    @staticmethod
    def _safe_get_attr(obj, attr, default=None):
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
        # Catch pywintypes.com_error if available, otherwise fall back to Exception
        except Exception:
            return default

    def __init__(self, default_account: str | None = None):
        """
        Connect to running Outlook instance or start it

        Launch Logic:
        1. Try GetActiveObject first (if Outlook is running)
        2. Fall back to Dispatch if Outlook is closed
        """
        try:
            self.outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as e:
            # Outlook might not be running - try to launch it
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                print("Error: Could not connect to or launch Outlook.", file=sys.stderr)
                print(f"Details: {e}", file=sys.stderr)
                # output full traceback
                print(
                    "Hint: Make sure Outlook is installed and you can launch it manually.",
                    file=sys.stderr,
                )
                traceback.print_exc()
                sys.exit(1)

        self.namespace = self.outlook.GetNamespace("MAPI")
        # Default account name and root folder (set by set_default_account or via init param)
        self.default_account_name = None
        self.default_root_folder = None

        # If provided, attempt to set the default account/store on init
        if default_account:
            with contextlib.suppress(Exception):
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
            True if successful, False otherwise
        """
        root = self._find_root_by_name(acc_name)
        if not root:
            return False
        try:
            # Set attributes for bridge usage
            self.default_account_name = acc_name
            self.default_root_folder = root
            # Also set DefaultStore to help other COM calls that rely on it
            with contextlib.suppress(Exception):
                self.namespace.DefaultStore = root.Store
            return True
        except Exception:
            return False

    def get_inbox(self):
        """Get the inbox folder"""
        # Prefer default account root when set
        root = self._get_root()
        if root:
            try:
                # Try case-sensitive first
                return root.Folders["Inbox"]
            except Exception:
                # Try case-insensitive search across root subfolders
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "inbox":
                            return f
                except Exception:
                    pass

        # Fallback to namespace default
        try:
            return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        except Exception:
            return None

    def list_folders(self, acc_name: str | None = None) -> dict[str, list[dict]]:
        """
        Recursively list all folders for all accounts.
        Args:
            acc_name: Specific account name to list folders for (optional)
        Returns:
            Dict with account names as keys and list of folder info dicts as values

        """

        def retrieve_folder_details(folder, parent_folder, depth):
            print(f"{'  ' * depth}- {folder.Name} (Items: {folder.Items.Count})")
            all_items = []
            cur_folder_data = {
                "name": folder.Name,
                "id": folder.EntryID,
                "parent_name": parent_folder.Name if parent_folder else None,
                "parent_id": parent_folder.EntryID if parent_folder else None,
                "number_of_items": folder.Items.Count,
                "path": folder.FolderPath,
                "depth": depth,
                "account": parent_folder.Name if parent_folder else None,
            }
            all_items.append(cur_folder_data)
            for subfolder in folder.Folders:
                all_items.extend(retrieve_folder_details(subfolder, folder, depth + 1))
            return all_items

        final = {}
        for i in range(1, 7):
            try:
                parent_folder = self.namespace.Folders.Item(i)
                if acc_name and parent_folder.Name != acc_name:
                    print(
                        f"acc_name arg does not match found account: {parent_folder.Name}\n  skipping..."
                    )
                    continue
                print(f"Account: {parent_folder.Name}")
            except Exception:
                print(f"Finished listing accounts. Total accounts: {i - 1}")
                break  # No more accounts

            final[parent_folder.Name] = retrieve_folder_details(parent_folder, None, 0)

        return final

    def get_calendar(self):
        """Get the calendar folder"""
        # Prefer default account root when set
        root = self._get_root()
        if root:
            try:
                # direct access
                return root.Folders["Calendar"]
            except Exception:
                # case-insensitive search
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "calendar":
                            return f
                except Exception:
                    pass

        # Fallback: search all accounts by name
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    parent_folder = self.namespace.Folders.Item(i)
                    try:
                        cal = parent_folder.Folders["Calendar"]
                        return cal
                    except Exception:
                        # try case-insensitive
                        for f in parent_folder.Folders:
                            if str(f.Name).strip().lower() == "calendar":
                                return f
                except Exception:
                    continue

        return None

    def get_tasks(self):
        """Get the tasks folder"""
        root = self._get_root()
        if root:
            try:
                return root.Folders["Tasks"]
            except Exception:
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "tasks":
                            return f
                except Exception:
                    pass

        try:
            return self.namespace.GetDefaultFolder(13)  # 13 = olFolderTasks
        except Exception:
            return None

    def get_folder_by_name(self, folder_name):
        """
        Get a folder by name (e.g., "Sent Items", "Archive", etc.)

        Args:
            folder_name: Name of the folder

        Returns:
            Folder object or None
        """
        # Try default account root first
        if not folder_name:
            return None

        root = self._get_root()
        if root:
            try:
                return root.Folders[folder_name]
            except Exception:
                try:
                    for f in root.Folders:
                        if (
                            str(f.Name).strip().lower()
                            == str(folder_name).strip().lower()
                        ):
                            return f
                except Exception:
                    pass

        # Search across all account roots
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    parent = self.namespace.Folders.Item(i)
                    try:
                        return parent.Folders[folder_name]
                    except Exception:
                        # case-insensitive search in this parent
                        try:
                            for f in parent.Folders:
                                if (
                                    str(f.Name).strip().lower()
                                    == str(folder_name).strip().lower()
                                ):
                                    return f
                        except Exception:
                            pass
                except Exception:
                    continue

        # Last resort: try the first root's children
        try:
            root = self.namespace.Folders.Item(1)
            try:
                return root.Folders[folder_name]
            except Exception:
                for f in root.Folders:
                    if str(f.Name).strip().lower() == str(folder_name).strip().lower():
                        return f
        except Exception:
            pass

        return None

    def get_item_by_id(self, entry_id):
        """
        Get any Outlook item by EntryID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID

        Returns:
            Outlook item (MailItem, AppointmentItem, TaskItem, etc.) or None
        """
        try:
            return self.namespace.GetItemFromID(entry_id)
        except Exception:
            return None

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
    def _format_com_datetime(value):
        """Format a COM/pywintypes datetime to 'YYYY-MM-DD HH:MM:SS' or None.

        Handles both real datetime objects (strftime) and pywintypes Time objects
        (which expose Year/Month/.../Second properties).
        """
        if not value:
            return None
        try:
            if isinstance(value, datetime):
                return value.strftime("%Y-%m-%d %H:%M:%S")
            return datetime(
                value.Year,
                value.Month,
                value.Day,
                value.Hour,
                value.Minute,
                value.Second,
            ).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return None

    @staticmethod
    def _attachment_count(item):
        """Return the number of attachments on a COM item (0 on any failure)."""
        try:
            return item.Attachments.Count
        except Exception:
            return 0

    def _extract_attachments(self, item):
        """Build a list of attachment-metadata dicts from a COM item."""
        out = []
        try:
            attachments = item.Attachments
            count = attachments.Count
        except Exception:
            return out
        for i in range(1, count + 1):
            try:
                a = attachments.Item(i)
            except Exception:
                continue
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
            "received_time": self._format_com_datetime(
                self._safe_get_attr(item, "ReceivedTime")
            ),
            "sent_time": self._format_com_datetime(self._safe_get_attr(item, "SentOn")),
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
            List of email dictionaries
        """
        # Use get_inbox() for the default Inbox to ensure correct account
        if folder == "Inbox":
            inbox = self.get_inbox()
        else:
            inbox = self.get_folder_by_name(folder)
            if not inbox:
                inbox = self.get_inbox()

        if inbox is None:
            return []

        items = inbox.Items

        # Filter to real emails (IPM.Note*) unless the caller opts out.
        if not include_non_mail:
            # Restrict can fail on unusual folders; fall back to unfiltered.
            with contextlib.suppress(Exception):
                items = items.Restrict(MAIL_ONLY_FILTER)

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
            except Exception:
                # Skip items that can't be accessed
                continue

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
        """
        item = self.get_item_by_id(entry_id)
        if item is None:
            return None
        try:
            return self._mail_item_to_dict(item, include_body=True)
        except Exception:
            return None

    def get_email_bodies(self, entry_ids, include_body=True):
        """
        Bulk-fetch full details for many EntryIDs in a single call (O(1) each).

        Avoids the N+1 round-trip pattern of calling get_email_body once per item.

        Args:
            entry_ids: Iterable of Outlook EntryIDs
            include_body: If False, fetch only summary fields (faster)

        Returns:
            List of email dictionaries for items that were found (missing IDs are
            silently omitted)
        """
        results = []
        for entry_id in entry_ids or []:
            try:
                item = self.get_item_by_id(entry_id)
            except Exception:
                item = None
            if item is None:
                continue
            try:
                results.append(self._mail_item_to_dict(item, include_body=include_body))
            except Exception:
                continue
        return results

    @staticmethod
    def _to_naive_datetime(value):
        """
        Normalize a COM/pywintypes datetime (or a datetime-like with
        capitalized components) to a naive Python datetime.

        Returns None when the value cannot be normalized.
        """
        try:
            if isinstance(value, datetime):
                # drop tzinfo if present so we can compare with datetime.now()
                return datetime(
                    value.year,
                    value.month,
                    value.day,
                    value.hour,
                    value.minute,
                    value.second,
                )
            return datetime(
                value.Year,
                value.Month,
                value.Day,
                value.Hour,
                value.Minute,
                value.Second,
            )
        except Exception:
            return None

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
            f"'{end_dt.strftime('%Y-%m-%d %H:%M')}') "
            'AND ("urn:schemas:calendar:dtend" >= '
            f"'{start_dt.strftime('%Y-%m-%d %H:%M')}')"
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
        """
        calendar = self.get_calendar()
        items = calendar.Items

        # Canonical order per Microsoft docs: Sort -> IncludeRecurrences ->
        # a single Restrict. The old code applied a second Restrict AFTER
        # setting IncludeRecurrences, which drops recurrence expansion on the
        # re-derived collection and yields flaky (often empty) results for
        # some date windows.
        items.Sort("[Start]")  # ascending, required for recurrence expansion
        items.IncludeRecurrences = True

        start_dt = datetime.now()
        if all_events:
            end_dt = start_dt + timedelta(days=CALENDAR_ALL_EVENTS_MAX_DAYS)
        else:
            days = max(1, min(days, CALENDAR_ALL_EVENTS_MAX_DAYS))
            end_dt = start_dt + timedelta(days=days)

        # Single COM-level Restrict: appointments only + overlap window.
        # This is also what prevents the "Calendar Bomb" (infinite recurring
        # items) from ever reaching Python iteration.
        items = items.Restrict(self._calendar_range_filter(start_dt, end_dt))

        events = []
        scanned = 0
        for item in items:
            scanned += 1
            if scanned > CALENDAR_MAX_ITEMS:
                break
            try:
                # Use safe attribute access to handle COM errors
                start = self._safe_get_attr(item, "Start")
                end = self._safe_get_attr(item, "End")

                # Skip if no start time
                if not start:
                    continue

                # Python-level overlap re-check (defence in depth; the COM
                # filter above should already have scoped correctly).
                start_n = self._to_naive_datetime(start)
                end_n = self._to_naive_datetime(end)
                if start_n is None:
                    continue
                if start_n > end_dt:
                    continue
                if end_n is not None and end_n < start_dt:
                    continue

                # Get attendees (safe access)
                required_attendees = self._safe_get_attr(item, "RequiredAttendees", "")
                optional_attendees = self._safe_get_attr(item, "OptionalAttendees", "")

                # Get meeting status
                # ResponseStatus: 0=None, 1=Organizer, 2=Tentative, 3=Accepted, 4=Declined, 5=NotResponded
                response_status = self._safe_get_attr(item, "ResponseStatus")
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }

                # MeetingStatus: 0=Non-meeting, 1=Meeting, 2=Received, 3=Canceled
                meeting_status = self._safe_get_attr(item, "MeetingStatus")
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                event = {
                    "entry_id": self._safe_get_attr(item, "EntryID", ""),
                    "subject": self._safe_get_attr(item, "Subject", "(No Subject)"),
                    "start": start.strftime("%Y-%m-%d %H:%M:%S") if start else None,
                    "end": end.strftime("%Y-%m-%d %H:%M:%S") if end else None,
                    "location": self._safe_get_attr(item, "Location", ""),
                    "organizer": self._safe_get_attr(item, "Organizer"),
                    "all_day": self._safe_get_attr(item, "AllDayEvent", False),
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": self._safe_get_attr(
                        item, "ResponseRequested", False
                    ),
                }
                events.append(event)
            except (Exception, BaseException):
                # Skip items that cause errors (including COM fatal errors)
                continue

        return events

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
            Draft entry ID if saved, True if sent, False if failed
        """
        try:
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

            # Add attachments
            if file_paths:
                for file_path in file_paths:
                    with contextlib.suppress(Exception):
                        mail.Attachments.Add(file_path)

            # Ensure sending uses the default account when set
            try:
                acc = None
                if self.default_account_name:
                    acc = self._find_account_by_name(self.default_account_name)
                # If DefaultStore was set, try to find account by matching store owner
                if not acc:
                    try:
                        accounts = self.namespace.Accounts
                        for a in accounts:
                            try:
                                if hasattr(a, "SmtpAddress") and a.SmtpAddress in (
                                    self.default_account_name or ""
                                ):
                                    acc = a
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass

                if acc:
                    with contextlib.suppress(Exception):
                        mail.SendUsingAccount = acc
            except Exception:
                pass

            if save_draft:
                mail.Save()
                return mail.EntryID
            else:
                mail.Send()
                return True
        except Exception as e:
            print(f"Error sending email: {e}", file=sys.stderr)
            return False

    def reply_email(self, entry_id, body, reply_all=False):
        """
        Reply to an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            body: Reply body
            reply_all: True to reply all, False to reply sender only

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if reply_all:
                    reply = item.ReplyAll()
                else:
                    reply = item.Reply()
                reply.Body = body
                # try to enforce default send account
                try:
                    if self.default_account_name:
                        acc = self._find_account_by_name(self.default_account_name)
                        if acc:
                            with contextlib.suppress(Exception):
                                reply.SendUsingAccount = acc
                except Exception:
                    pass
                reply.Send()
                return True
            except Exception as e:
                print(f"Error replying to email: {e}", file=sys.stderr)
                return False
        return False

    def forward_email(self, entry_id, to, body=""):
        """
        Forward an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            to: Recipient to forward to
            body: Optional additional body text

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                forward = item.Forward()
                forward.To = to
                if body:
                    forward.Body = body + "\n\n" + forward.Body
                try:
                    if self.default_account_name:
                        acc = self._find_account_by_name(self.default_account_name)
                        if acc:
                            with contextlib.suppress(Exception):
                                forward.SendUsingAccount = acc
                except Exception:
                    pass
                forward.Send()
                return True
            except Exception as e:
                print(f"Error forwarding email: {e}", file=sys.stderr)
                return False
        return False

    def mark_email_read(self, entry_id, unread=False):
        """
        Mark an email as read or unread (O(1) direct access)

        Args:
            entry_id: Email entry ID
            unread: True to mark as unread, False to mark as read

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Unread = unread
                item.Save()
                return True
            except Exception:
                return False
        return False

    def move_email(self, entry_id, folder_name):
        """
        Move an email to a different folder (O(1) direct access)

        Args:
            entry_id: Email entry ID
            folder_name: Target folder name

        Returns:
            True if successful
        """
        try:
            item = self.get_item_by_id(entry_id)
            if not item:
                return False

            target_folder = self.get_folder_by_name(folder_name)
            if not target_folder:
                print(f"Error: Folder '{folder_name}' not found", file=sys.stderr)
                return False

            item.Move(target_folder)
            return True
        except Exception as e:
            print(f"Error moving email: {e}", file=sys.stderr)
            return False

    def delete_email(self, entry_id):
        """
        Delete an email (O(1) direct access)

        Args:
            entry_id: Email entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                return False
        return False

    def download_attachments(self, entry_id, download_dir):
        """
        Download all attachments from an email

        Args:
            entry_id: Email entry ID
            download_dir: Directory to save attachments

        Returns:
            List of downloaded file paths
        """
        item = self.get_item_by_id(entry_id)
        if not item or item.Attachments.Count == 0:
            return []

        downloaded = []
        try:
            import os

            os.makedirs(download_dir, exist_ok=True)

            for i in range(item.Attachments.Count):
                attachment = item.Attachments.Item(i + 1)  # COM is 1-indexed
                filename = attachment.FileName
                filepath = os.path.join(download_dir, filename)
                attachment.SaveAsFile(filepath)
                downloaded.append(filepath)
            return downloaded
        except Exception as e:
            print(f"Error downloading attachments: {e}", file=sys.stderr)
            return []

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
            Appointment entry ID if successful
        """
        try:
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
            appointment.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
            appointment.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
            appointment.Location = location
            appointment.Body = body
            appointment.AllDayEvent = all_day
            if required_attendees:
                appointment.RequiredAttendees = required_attendees
            if optional_attendees:
                appointment.OptionalAttendees = optional_attendees
            appointment.Save()
            return appointment.EntryID
        except Exception as e:
            print(f"Error creating appointment: {e}", file=sys.stderr)
            return None

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
    ):
        """
        Edit an existing appointment

        Args:
            entry_id: Appointment entry ID
            required_attendees: Comma-separated list of required attendees
            optional_attendees: Comma-separated list of optional attendees
            subject: New subject (optional)
            start: New start time (optional)
            end: New end time (optional)
            location: New location (optional)
            body: New body (optional)

        Returns:
            True if successful
        """
        try:
            calendar = self.get_calendar()
            for item in calendar.Items:
                if item.EntryID == entry_id:
                    if required_attendees:
                        item.RequiredAttendees = required_attendees
                    if optional_attendees:
                        item.OptionalAttendees = optional_attendees
                    if subject:
                        item.Subject = subject
                    if start:
                        item.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
                    if end:
                        item.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
                    if location is not None:
                        item.Location = location
                    if body is not None:
                        item.Body = body
                    item.Save()
                    return True
            return False
        except Exception as e:
            print(f"Error editing appointment: {e}", file=sys.stderr)
            return False

    def get_appointment(self, entry_id):
        """
        Get full appointment details by entry ID (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            Appointment dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                required_attendees = (
                    item.RequiredAttendees if hasattr(item, "RequiredAttendees") else ""
                )
                optional_attendees = (
                    item.OptionalAttendees if hasattr(item, "OptionalAttendees") else ""
                )
                response_status = (
                    item.ResponseStatus if hasattr(item, "ResponseStatus") else None
                )
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }
                meeting_status = (
                    item.MeetingStatus if hasattr(item, "MeetingStatus") else None
                )
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "start": item.Start.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "Start") and item.Start
                    else None,
                    "end": item.End.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "End") and item.End
                    else None,
                    "location": item.Location if hasattr(item, "Location") else "",
                    "organizer": item.Organizer if hasattr(item, "Organizer") else None,
                    "body": item.Body if hasattr(item, "Body") else "",
                    "all_day": item.AllDayEvent
                    if hasattr(item, "AllDayEvent")
                    else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": item.ResponseRequested
                    if hasattr(item, "ResponseRequested")
                    else False,
                }
            except Exception:
                pass
        return None

    def respond_to_meeting(self, entry_id, response):
        """
        Respond to a meeting invitation (O(1) direct access)

        Args:
            entry_id: Appointment entry ID
            response: Response - "accept", "decline", "tentative"

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                response_map = {
                    "accept": 3,  # olResponseAccepted
                    "decline": 4,  # olResponseDeclined
                    "tentative": 2,  # olResponseTentative
                }
                if response.lower() in response_map:
                    item.Response(response_map[response.lower()])
                    item.Send()
                    return True
            except Exception:
                pass
        return False

    def delete_appointment(self, entry_id):
        """
        Delete an appointment (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

    def list_tasks(self, include_completed=False):
        """
        List tasks (only incomplete by default)

        Args:
            include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

        Returns:
            List of task dictionaries
        """
        tasks_folder = self.get_tasks()
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
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
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
            except Exception:
                continue

        return tasks

    def list_all_tasks(self):
        """
        List all tasks including completed ones

        Returns:
            List of task dictionaries
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
            Task entry ID if successful
        """
        try:
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
                # Use noon to avoid timezone boundary issues
                task.DueDate = datetime.strptime(
                    f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                )
            task.Importance = importance
            task.Save()
            return task.EntryID
        except Exception as e:
            print(f"Error creating task: {e}", file=sys.stderr)
            return None

    def get_task(self, entry_id):
        """
        Get full task details by entry ID (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            Task dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
            except Exception:
                pass
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
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if subject:
                    item.Subject = subject
                if body is not None:
                    item.Body = body
                if due_date:
                    # Use noon to avoid timezone boundary issues
                    item.DueDate = datetime.strptime(
                        f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                    )
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
            except Exception as e:
                print(f"Error editing task: {e}", file=sys.stderr)
            pass
        return False

    def complete_task(self, entry_id):
        """
        Mark a task as complete (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Complete = True
                item.PercentComplete = 100
                item.Status = 2  # olTaskComplete
                item.Save()
                return True
            except Exception:
                pass
        return False

    def delete_task(self, entry_id):
        """
        Delete a task (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

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
            List of email dictionaries

        Raises:
            ValueError: If a LIKE filter references a field without a known
                DASL mapping (supported: the fields in DASL_FIELD_MAP).
            RuntimeError: If Outlook rejects the filter (invalid syntax) or the
                COM call fails. Callers previously saw a silent empty list here,
                which made bad filters indistinguishable from "no matches".
        """
        # Use get_inbox() directly to ensure correct account
        folder = self.get_inbox()

        # Translate SQL-style LIKE filters to DASL so wildcards actually work.
        raw_query = filter_query or ""
        mentions_message_class = "messageclass" in raw_query.lower()
        prepared = translate_filter(raw_query)
        dasl_mode = prepared.startswith("@SQL=")

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
                print(f"Error searching emails: {msg}", file=sys.stderr)
                raise RuntimeError(msg) from e

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
            except Exception:
                # Skip items that can't be accessed
                continue

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
            List of email dictionaries matching the sender
        """
        try:
            # Get the folder
            if folder == "Inbox":
                mail_folder = self.get_inbox()
            else:
                mail_folder = self.get_folder_by_name(folder)
                if not mail_folder:
                    mail_folder = self.get_inbox()

            items = mail_folder.Items
            # Filter to real emails unless the caller opts out.
            if not include_non_mail:
                with contextlib.suppress(Exception):
                    items = items.Restrict(MAIL_ONLY_FILTER)
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
                except Exception:
                    # Skip items that can't be accessed
                    continue

            return emails
        except Exception as e:
            print(f"Error searching emails by sender: {e}", file=sys.stderr)
            return []

    def get_inbox_stats(self, folder="Inbox"):
        """
        Return cheap total/unread counts for a folder without fetching items.

        Uses Restrict+Count at the COM level so it is O(1)-ish regardless of
        folder size. Useful for pagination decisions and inbox monitoring.

        Args:
            folder: Folder name (default: "Inbox")

        Returns:
            Dict with 'folder', 'total', and 'unread' integer counts
        """
        try:
            if folder == "Inbox":
                mail_folder = self.get_inbox()
            else:
                mail_folder = self.get_folder_by_name(folder)
                if not mail_folder:
                    mail_folder = self.get_inbox()

            if mail_folder is None:
                return {"folder": folder, "total": 0, "unread": 0}

            items = mail_folder.Items
            total = self._safe_get_attr(items, "Count", 0) or 0
            try:
                unread = items.Restrict("[Unread] = TRUE").Count
            except Exception:
                unread = 0
            return {"folder": folder, "total": int(total), "unread": int(unread)}
        except Exception:
            return {"folder": folder, "total": 0, "unread": 0}

    def get_free_busy(
        self, email_address=None, start_date=None, end_date=None, entry_id=None
    ):
        """
        Get free/busy status for an email address

        Args:
            email_address: Email address to check (optional, defaults to current user)
            start_date: Start date (YYYY-MM-DD) or datetime object (optional, defaults to today)
            end_date: End date (YYYY-MM-DD) or datetime object (optional, defaults to start + 1 day)
            entry_id: DEPRECATED - Appointment entry ID (legacy, use email_address instead)

        Returns:
            Dictionary with free/busy information
        """
        try:
            # Handle legacy entry_id parameter (extract first required attendee)
            if entry_id and not email_address:
                item = self.get_item_by_id(entry_id)
                if (
                    item
                    and hasattr(item, "RequiredAttendees")
                    and item.RequiredAttendees
                ):
                    attendees = item.RequiredAttendees.split(";")
                    if attendees:
                        email_address = attendees[0].strip()

            # Default to current user if no email provided
            if not email_address:
                email_address = self.namespace.CurrentUser.Address

            # Default to today if no dates provided
            if not start_date:
                start_date = datetime.now()
            elif isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%Y-%m-%d")

            if not end_date:
                end_date = start_date + timedelta(days=1)
            elif isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%Y-%m-%d")

            # Create recipient and get free/busy
            recipient = self.namespace.CreateRecipient(email_address)
            if recipient.Resolve():
                # FreeBusy returns a string with time slots and status
                # 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
                freebusy = recipient.FreeBusy(
                    start_date, 60 * 24
                )  # 1440 minutes = 1 day
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "free_busy": freebusy,
                    "resolved": True,
                }
            else:
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "error": "Could not resolve email address",
                    "resolved": False,
                }
        except Exception as e:
            return {
                "email": email_address if email_address else "unknown",
                "error": str(e),
                "resolved": False,
            }
