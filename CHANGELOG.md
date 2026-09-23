# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to
[Semantic Versioning](https://semver.org/).


## [0.12.0] - 2026-08-18

 Triage follow-up release: every fix below was found by running an agent-driven
 150-email inbox triage against the live MCP server, then verified with unit
 tests plus live Outlook smoke tests.

### Fixed

- **`search_emails` LIKE filters actually work.** `Items.Restrict` uses Jet
  syntax, which has no LIKE operator — every documented example such as
  `"[Subject] LIKE '%project%'"` raised a COM error that was swallowed and
  returned an empty list (indistinguishable from "no matches"). Filters
  containing LIKE are now translated wholesale to DASL (`@SQL=`) with a
  verified field map (Subject, SenderName, SenderEmailAddress, ReceivedTime,
  Unread, HasAttachments, MessageClass, Body, To, CC, Importance), SQL LIKE
  wildcard semantics, `[Unread]` literals inverted to `httpmail:read`
  semantics, and remaining TRUE/FALSE normalized to 1/0. Bonus: the documented
  `[HasAttachments] = TRUE` example also works now (it was rejected in Jet on
  the tested Outlook build). Unknown fields in LIKE filters raise a clear
  `ValueError` instead of matching nothing.
- **Restrict failures are no longer silent.** `search_emails` raises
  `RuntimeError` (surfaced by the MCP tool / CLI) when Outlook rejects a
  filter. The CLI `search` and `calendar` commands print a JSON error and exit 1.
- **`list_calendar_events` date windows were locale-broken.** Jet date
  literals are parsed per the user's Windows locale: on a Dutch-locale system
  `"09/01/2026 12:00"` silently parsed as 9 January, so windows whose end date
  had a day component <= 12 returned empty (e.g. `days=14` and `days=45`
  returned 0 events while `days=7`/`days=30` worked). The filter now uses
  DASL `urn:schemas:calendar:dtstart/dtend` with ISO `YYYY-MM-DD HH:MM`
  literals, which are locale-independent (verified: monotonic event counts
  across 7/14/30/45/60-day windows).
- **Calendar iteration order corrected.** The old code applied a second
  `Restrict` AFTER setting `IncludeRecurrences`/`Sort`, which drops recurrence
  expansion on the re-derived collection. Now: Sort → IncludeRecurrences → a
  single combined Restrict (Microsoft-documented order).

### Changed

- **`list_calendar_events(all_events=True)` is now bounded** (from now through
  +365 days, at most `CALENDAR_MAX_ITEMS = 2000` iterations). The previous
  unbounded expansion of every recurring series blocked Outlook's
  single-threaded COM apartment for minutes, which wedged the entire MCP
  server — every other tool call (even `get_inbox_stats`) timed out until the
  process was killed. `days` is clamped to [1, 365].
- Python-level re-check in `list_calendar_events` uses overlap semantics
  (matching the COM filter) instead of requiring the start to fall inside the
  window, so multi-day events straddling the window edge are kept.

### Added

- `translate_filter` / `DASL_FIELD_MAP` / `DASL_MAIL_ONLY_FILTER` /
  `CALENDAR_ALL_EVENTS_MAX_DAYS` / `CALENDAR_MAX_ITEMS` module constants and
  helpers (unit-testable without Outlook).

### Tests

- New `tests/test_filters.py` (28 unit tests, no Outlook): DASL translation
  (wildcards, field map, Unread inversion, boolean normalization, unknown
  fields), search_emails filter composition + error surfacing, and calendar
  single-Restrict / horizon / cap / overlap behaviour via fake COM items.

 triage-driven release. Every change below was motivated by stress-testing the MCP
server against a 166-email inbox cleanup and is verified by unit tests plus a live
read-only smoke test against Outlook.

### Fixed

- **`get_email` no longer fails on non-mail inbox items.** `list_emails` /
  `list_unread_emails` / `search_emails` previously returned meeting notifications
  (`IPM.Schedule.Meeting.*`), shares (`IPM.Sharing`), and post items alongside real
  mail, and `get_email` then raised `OutlookNotFoundError` on them ("Email not found").
  Listings now scope to `IPM.Note` (and subtypes) at the COM level by default; `get_email`
  on a non-mail item returns a populated `EmailDetails` with its `message_class` set so
  callers can branch instead of catching errors. Pass `include_non_mail=True` to opt back
  in to the old behaviour on `list_emails` / `search_emails` / `search_emails_by_sender`.
- **Sender SMTP resolution hardened for cached Exchange mode.** `resolve_smtp_address`
  now falls back to `PropertyAccessor` → `PidTagSenderSmtpAddress` (0x5D01001F) and a
  regex salvage before returning the raw Exchange DN, so internal senders no longer come
  back as `/O=EXCHANGELABS/...` strings when `GetExchangeUser()` returns `None`.

### Added

- **Richer email metadata.** `EmailSummary` and `EmailDetails` now carry `to`, `cc`,
  `sent_time`, `conversation_id`, `conversation_topic`, and `message_class`. `EmailDetails`
  additionally carries `bcc`, `attachments` (list of new `AttachmentInfo`:
  filename/size/display_name/content_type/is_inline), and `body_top` (the new-message
  portion of the body with quoted reply chains and signatures stripped).
- **`get_emails(entry_ids, include_body=True)`** tool + `bridge.get_email_bodies(...)`
  for bulk fetching (avoids the N+1 round-trip of calling `get_email` per item).
- **`get_inbox_stats(folder)`** tool + `bridge.get_inbox_stats(...)` returning cheap
  `{folder, total, unread}` counts via `Restrict` + `Count` (no item iteration).
- `body_top` cleaning helper, SMTP salvage regex, and the `IPM.Note` MessageClass filter
  are exposed as unit-testable module constants / static methods, and `import mailtool.bridge`
  no longer hard-requires `pywin32` (defensive import) so pure helpers test on any platform.
- `@pytest.mark.unit` marker registered in `pytest.ini`.

### Changed

- Refactored the four duplicated email dict builders in `bridge.py` into a single
  `_mail_item_to_dict(item, *, include_body=False)` helper that reads every field via
  `_safe_get_attr`, so one bad field on one item can no longer drop the whole listing.
- `search_emails` docstring now documents date-range and `HasAttachments` filter examples.

### Tests

- Added `tests/test_enhancements.py` (unit, no Outlook): body cleaning, SMTP regex,
  MessageClass filter constant, model defaults, `_mail_item_to_dict` with a fake COM item,
  and the new `get_emails` / `get_inbox_stats` / non-mail `get_email` tools via mocks.
- Updated `tests/mcp/test_tools.py` and `tests/mcp/test_models.py` for the new
  `include_non_mail` kwarg and the expanded model serializations.

## [0.10.0] - prior

- MCP SDK v2 / FastMCP migration baseline (see `CLAUDE.md` for the historical notes).
