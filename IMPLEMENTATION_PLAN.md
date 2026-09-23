# Implementation Plan — Inbox-Triage Improvements

Stress-tested the mailtool MCP by pulling and classifying 166 unread inbox emails
(2026-07-07). This plan turns the findings in `C:/dev/inbox-cleanup-2026-07-07/MCP_IMPROVEMENTS.md`
into concrete, code-grounded work. Every item below cites the exact file/method/line
that has to change.

## Correction to the original review

- **Drop finding #6 ("concatenated JSON, not an array").** The tools already return
  proper `list[EmailSummary]` / `EmailDetails` Pydantic models (`src/mailtool/mcp/server.py`).
  The concatenation I parsed was a rendering artifact of the harness wrapping the MCP
  result, **not** a server bug. Removed from scope.

## How findings map to the code

| # | Finding | Root cause in code | Location |
|---|---------|--------------------|----------|
| 1 | `get_email` fails on calendar/meeting items that `list_unread_emails` returns | `list_emails`/`search_emails` iterate **all** inbox items with no `MessageClass` filter; `get_email_body` returns `None` on non-`MailItem` → `OutlookNotFoundError` | `bridge.py:445` (`list_emails`), `bridge.py:495` (`get_email_body`), `bridge.py:1380` (`search_emails`) |
| 2 | No To/CC, no thread ID, no sent time, no attachment details | Dict builders only read 7 fields | Same three builders |
| 3 | `sender` is a raw Exchange DN for internal mail | `resolve_smtp_address` falls back to `SenderEmailAddress` (which is the EX DN) when `GetExchangeUser()` returns `None` | `bridge.py:411` (`resolve_smtp_address`) |
| 4 | No bulk fetch | Only single-id `get_email` | `server.py:175` |
| 5 | No cleaned body; full quote chain returned | `body`/`html_body` returned verbatim | `bridge.py:513` |

The calendar code already does the right pattern and is the template to copy:
`bridge.py:540-542` filters with `[MessageClass] >= 'IPM.Appointment' AND [MessageClass] < 'IPM.Appointment{'`
before iterating.

---

## Phase 1 — Correctness: non-mail items in email listings (P0, bug)

**Problem.** `list_unread_emails` → `search_emails("[Unread] = TRUE")` returns every
unread inbox item regardless of type: `IPM.Note` (real mail), `IPM.Schedule.Meeting.Request`,
`.Canceled`, `.Response`, `IPM.Post`, etc. (8% of my inbox — 13/166). They look like
emails in the summary, then `get_email` blows up with "Email not found" because
`get_email_body` returns `None`.

### Changes

**`bridge.py`**
- `list_emails(self, limit=10, folder="Inbox", include_non_mail=False)` — when
  `include_non_mail` is False (default), add `.Restrict("[MessageClass] = 'IPM.Note'")`
  after `folder.Items` and before `.Sort(...)`. Use a range filter
  `"[MessageClass] >= 'IPM.Note' AND [MessageClass] < 'IPM.Note{'"` to also catch
  `IPM.Note.SMIME` etc.
- `search_emails(self, filter_query, limit=100, include_non_mail=False)` — same.
  When the caller passes an explicit MessageClass in `filter_query`, don't double-filter.
- `get_email_body(self, entry_id)` — on `MailItem` keep current behaviour. On a
  non-`MailItem`, instead of returning `None`, return a dict with `item_type` (from
  `MessageClass`), `entry_id`, `subject`, `sender`, `sender_name`, `received_time`,
  `message_class`, and `body`/`html_body` only if accessible. This lets `get_email`
  distinguish "truly missing" from "this is a meeting item" — see server change below.

**`server.py`**
- `get_email` — distinguish `result is None` (raise `OutlookNotFoundError` as today)
  from `result.get("item_type") not in (None, "IPM.Note")` (return the partial dict,
  or raise a new `OutlookItemTypeError`). Recommended: return the partial `EmailDetails`
  with `message_class` set so the caller can branch without a second round-trip.

**`models.py`**
- Add `message_class: str = Field(default="IPM.Note", description="Outlook MessageClass, e.g. IPM.Note / IPM.Schedule.Meeting.Request")` to both `EmailSummary` and `EmailDetails`.

### Tests
- **Unit (no Outlook):** the filter-string builder is pure — extract it and assert it
  produces the `IPM.Note` range and that `include_non_mail=True` skips it.
- **Integration:** on a live mailbox, `list_unread_emails(limit=1000)` returns zero
  items with `message_class != "IPM.Note"` (default), and the same call with
  `include_non_mail=True` returns the meeting items. `get_email(<meeting item id>)`
  returns a populated `EmailDetails` with `message_class` set instead of raising.
- Update `assert_email_structure` in `tests/conftest.py:275` to include `message_class`.

### Behaviour change to flag in release notes
Callers who today rely on `list_unread_emails` surfacing meeting notifications will
stop seeing them by default. This is desirable (it's an *email* tool) but technically
breaking — call it out and provide the `include_non_mail=True` escape hatch.

---

## Phase 2 — Correctness: robust sender SMTP resolution (P0, bug)

**Problem.** `resolve_smtp_address` (`bridge.py:411`) tries
`Sender.GetExchangeUser().PrimarySmtpAddress`, and on any failure returns
`SenderEmailAddress` — which for `SenderEmailType == "EX"` is the raw Exchange DN
(`/O=EXCHANGELABS/.../CN=...-MARSMANEM`). In cached Exchange mode, `GetExchangeUser()`
returns `None` for a meaningful fraction of internal senders (this is exactly what I
observed: ~7% of senders came back as DNs).

### Changes

**`bridge.py: resolve_smtp_address`** — add a `PropertyAccessor` path before the DN
fallback, and a regex sweep as the last resort:

```
1. If SenderEmailType == "EX":
   a. Try Sender.GetExchangeUser().PrimarySmtpAddress        (current)
   b. NEW: Try mail_item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F")   # PidTagSenderSmtpAddress
   c. NEW: Regex-extract first SMTP-shaped token from SenderEmailAddress
            (e.g. r"<([^<>@\s]+@[^<>@\s]+)>") as a last-ditch salvage
2. Else: SenderEmailAddress  (current, correct for SMTP)
3. On any exception: ""       (current)
```

The MAPI property `0x5D01001F` (`PidTagSenderSmtpAddress`) is the reliable source on
Outlook 2007+ and is what `search_by_sender` effectively needs to agree with. Verify
the tag resolves on your Outlook build during implementation.

### Tests
- **Unit:** the regex salvage against synthetic DN-with-embedded-SMTP strings.
- **Integration:** pick a known internal sender who currently returns a DN; assert the
  resolved `sender` matches `sender_name`'s parenthetical SMTP, and that
  `search_by_sender(<that smtp>)` now finds their mail (closing the loop that
  motivated the original `search_emails_by_sender` workaround).

---

## Phase 3 — Richer email metadata (P1)

**Problem.** For triage, To/CC, sent time, thread ID, and attachment list are all
load-bearing; today none are returned.

### Changes

**`models.py`**
- `EmailSummary` += `to: str`, `cc: str`, `sent_time: str | None`, `conversation_id: str | None`, `conversation_topic: str | None`.
- `EmailDetails` += `bcc: str` (only populated on sent items) and `attachments: list[AttachmentInfo]`.
- New model `AttachmentInfo(BaseModel)`: `filename: str`, `size: int`, `display_name: str`, `content_type: str | None`, `is_inline: bool`.

**`bridge.py`** — extend the three dict builders (`list_emails`, `get_email_body`,
`search_emails`, `search_by_sender`) to read the new properties **through
`_safe_get_attr`** (`bridge.py:35`, already exists) so a COM hiccup on one item can't
regress the whole listing:
- `item.To`, `item.CC`, `item.BCC` (semicolon-separated strings)
- `item.SentOn` → format like `received_time`
- `item.ConversationID`, `item.ConversationTopic`
- `item.MessageClass` (ties into Phase 1)
- Attachments: `[{"filename": a.FileName, "size": a.Size, "display_name": a.DisplayName,
  "content_type": _safe_get_attr(a, "ContentType"), "is_inline": bool(_safe_get_attr(a, "IsInline", False))} for a in item.Attachments]`

Factor the field extraction into one private helper, e.g. `_mail_item_to_dict(item, *, include_body)`, and have all four callers use it — kills the current 4× copy-paste and keeps Phase 1/3 changes in one place.

**`server.py`** — pass the new dict keys through to the models in `list_emails`,
`list_unread_emails`, `get_email`, `search_emails`, `search_emails_by_sender`.

### Tests
- **Unit:** `AttachmentInfo` and extended-model validation; `_mail_item_to_dict` shape
  with a fake item object (no Outlook needed).
- **Integration:** `get_email` on a known message with an attachment returns a non-empty
  `attachments` list with sane `filename`/`size`; `conversation_id` is identical across
  two messages in the same thread.

---

## Phase 4 — Bulk fetch (P1, performance)

**Problem.** Triage requires N bodies → N `get_email` round-trips (166 here). One batch
call cuts latency and tool-call overhead dramatically.

### Changes
**`bridge.py`** — `get_email_bodies(self, entry_ids: list[str], include_body: bool = True) -> list[dict]`.
Loop `get_item_by_id`, reuse `_mail_item_to_dict`, skip `None`s, cap input length
(e.g. `if len > 200: raise OutlookValidationError`). Return only items actually found.
**`server.py`** — new tool `get_emails(entry_ids: list[str]) -> list[EmailDetails]`.
**`models.py`** — no change (reuse `EmailDetails`).

### Tests
- **Integration:** fetch 3 known entry IDs in one call; assert 3 `EmailDetails`, order
  matches input, missing IDs are omitted (not errors).

---

## Phase 5 — Cleaned body convenience (P2)

**Problem.** Consumers all re-implement quote-chain stripping. Provide it once.

### Changes
- `bridge.py` — pure-Python `_clean_body_top(body: str) -> str`: cut at the first of
  `-----Original Message-----`, `-----Message réenvoyé-----`, `Van:`, `From:`,
  `Op … schreef`, `On … wrote:`, `> ` quote runs, `_____` sig separators; collapse
  blank runs; cap length (e.g. 1000 chars).
- `models.py` — `EmailDetails` += `body_top: str = Field(default="")`.
- `get_email_body`/server populate it.

### Tests
- **Unit (no Outlook):** feed synthetic bodies (plain, Dutch reply, quoted, no quote,
  signature-only) and assert the trimmed output. Deterministic and fast.

---

## Phase 6 — Listing/search ergonomics (P2)

- New tool `get_inbox_stats(folder: str = "Inbox") -> InboxStats(total: int, unread: int)`
  using a `Restrict`-then-`Count` (cheap, no Python iteration). Resolves the
  "am I paginated?" problem without changing the `list[EmailSummary]` return type.
- `list_unread_emails`: raise default `limit` from 10 to 50 in the docstring-recommended
  usage (keep signature default 10 to avoid surprising existing callers) — or just
  document that callers should pass a larger `limit` for cleanup workflows.
- `search_emails`: document date-range (`[ReceivedTime] >= '...' AND [ReceivedTime] <= '...'`)
  and attachment (`[HasAttachments] = TRUE`) filter examples in the docstring; no code change needed.

---

## Cross-cutting: testing strategy

The current suite is **integration-only** against a live mailbox
(`tests/conftest.py:40` session `bridge` fixture, `[TEST]` prefix cleanup, no mocking).
Several of the changes above are pure logic and should not require Outlook:

| Unit-testable (add `@pytest.mark.unit`, run in CI without Outlook) | Still integration |
|---|---|
| MessageClass filter-string builder (Phase 1) | `list_*` actually filters meeting items |
| sender SMTP regex salvage (Phase 2) | `resolve_smtp_address` against a real EX sender |
| `_mail_item_to_dict` with a fake item (Phase 3) | attachment/conversation fields on real mail |
| `_clean_body_top` heuristic (Phase 5) | — |
| new/extended Pydantic models (Phases 1, 3) | — |

Recommend a `@pytest.mark.unit` marker + a `pytest.ini` rule so unit tests run on every
platform and integration tests stay gated to Windows-with-Outlook (matches the existing
`@pytest.mark.integration` convention already used in `tests/test_emails.py:14`).

## Sequencing & versioning

1. **Phase 1 + 2** first — small, isolated, fix real data-loss bugs. Ship together.
2. **Phase 3** — the metadata win that most improves downstream triage; refactor the
   four dict builders into `_mail_item_to_dict` as the precondition.
3. **Phase 4 + 5** — DX/perf; independently shippable.
4. **Phase 6** + unit-test reorganisation — last.

Bump **2.3.0 → 2.4.0**. Phases 1 and 3 are additive on models (new fields) but Phase 1
changes `list_*`/`search_*` default output (meeting items disappear unless
`include_non_mail=True`) — that's the one breaking-ish note for the changelog. There is
no CHANGELOG.md today; consider adding one with this release.

## Out of scope (noted, not recommended now)
- Returning attachment bytes (security/storage cost; Outlook `SaveAsFile` already exists
  for the rare case it's needed).
- Per-folder enumeration in `list_*` results (low value; folders are an input param already).
