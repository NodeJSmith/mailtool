---
audience: personal tool
developers: solo
data-sensitivity: internal
peers: exchangelib, python-o365 (O365), pywin32 Outlook examples
---

# mailtool: Outlook COM Automation Bridge

A Python library and CLI for Outlook (classic) email, calendar, and tasks via Windows COM.
It runs on Windows next to a running Outlook; WSL2 sessions call it as `mailtool.exe`
through Windows interop. Both work machines install it with
`uv tool install git+https://github.com/NodeJSmith/mailtool`, and `dfl check` flags when the
installed commit falls behind `main`.

Forked from utsmok/mailtool and since detached. The upstream MCP server was removed (it was
unused); this repo is the CLI and library only. The `mailtool` name on PyPI belongs to
upstream, so nothing here publishes.

## Layout

- `src/mailtool/bridge.py`: `OutlookBridge`, all COM access. Imports pywin32 defensively so
  its static helpers and constants can be tested on any platform.
- `src/mailtool/cli.py`: argparse CLI. Checks for Windows and pywin32 first so a bad
  environment gets a readable error (help text on stderr, error JSON on stdout, exit 1).
  stdout carries only JSON: every command prints its result or `{"status": "error", ...}`
  there and exits 1 on failure, and warnings go to stderr (`test_stdout_contract.py` fails on
  a bare `print` that isn't JSON). argparse usage errors are the exception: exit 2, text on
  stderr, nothing on stdout.
- `tests/`: unit tests (pure Python, fake COM objects) plus integration tests that drive a
  live Outlook through the session-scoped `bridge` fixture in `conftest.py`.

## Design decisions

- **O(1) lookups.** Items are fetched with `GetItemFromID(entry_id)`, never by iterating a
  folder. Large mailboxes depend on it.
- **Calendar bomb.** Calendar listing sets `IncludeRecurrences = True` and `Sort("[Start]")`,
  then applies a COM-level `Restrict` before Python iterates. `all_events` is capped to a
  one-year horizon (`CALENDAR_ALL_EVENTS_MAX_DAYS`), because an unbounded scan expands
  recurring series forever and blocks Outlook's single-threaded COM apartment.
- **Occurrence ids.** Outlook gives every occurrence of a recurring meeting the series'
  EntryID, so an EntryID alone can only address the series. Listings return
  `<EntryID>@<start>` for an occurrence (`occurrence_id`), resolved with
  `GetRecurrencePattern().GetOccurrence(start)`. Mutators refuse a series id unless
  `series=True`, so one occurrence is the easy target and the whole series is the deliberate
  one. Moving an occurrence changes its id.
- **Times are local wall clock, in whenever.** Outlook reads and returns appointment times
  and due dates as local wall-clock time, so mailtool holds them as whenever `PlainDateTime`
  and `Date` (stdlib `datetime` appears only at the COM boundary). Every value crosses that
  boundary through `to_outlook` and `from_outlook`. pywin32 shifts a naive datetime from
  local time to UTC on the way in, so `to_outlook` hands it the wall-clock time labelled UTC,
  which passes through unchanged; a naive datetime written straight to COM lands off by the
  UTC offset.
- **Filters.** `Items.Restrict` speaks Jet, which has no LIKE. Filters containing LIKE are
  translated to DASL (`@SQL=`) through a verified field map in `bridge.py`.
- **Free/busy** takes an email address directly and defaults to the current user.
- **Return values.** Failures raise `MailtoolError` (a `RuntimeError`), never an empty
  result or a `False` sentinel, because the consumers are unattended digests that read an
  empty list as "nothing new". Lists return `[]` only when nothing matches; a fetch by an
  EntryID that doesn't exist returns `None` (not a failure), and so does the wrong item type,
  but only appointment and task methods check Outlook's `Class`; email methods accept any
  item. Only Outlook's not-found errors
  (`NOT_FOUND_HRESULTS`) count as not found; any other lookup error raises, so an Outlook
  outage can't read as "not found". Mutators return `True` and creates return the new
  EntryID. A list loop may skip one unreadable item (with a stderr
  warning), but the whole call failing raises. The CLI turns `MailtoolError` (and any other
  exception) into `{"status": "error", "message": ...}` on stdout with exit 1. A sent email
  moves to Sent Items with a new EntryID, so the original can't be returned.

## Testing

```bash
uv run pytest                 # anywhere; live-Outlook tests skip off Windows
uv run pytest -m windows      # only the live-Outlook tests (Windows + Outlook)
```

- Any test using the `bridge` fixture (directly or through another fixture) is marked
  `windows` automatically by `conftest.py` and skipped where pywin32 isn't available. New
  live-Outlook tests need no manual marker; new pure-Python tests must not request `bridge`.
- CI (`tests.yml`, ubuntu) runs the pure-Python tests; the live-Outlook ones run only on a
  Windows machine with Outlook open.
- Integration tests use real Outlook data, no mocking. Items they create carry a `[TEST] `
  prefix and are cleaned up by prefix, in a `finally`, from whichever folder they end up
  in (a draft moved to the Inbox is cleaned from the Inbox).
- Each test has a 60-second timeout (`pytest.ini`). Searches that iterate items over COM
  (`search_by_sender`) need a small `limit` and a query that matches near the top, or they
  walk the whole folder.
- COM is apartment-threaded: all calls happen on one thread (the session-scoped fixture), so
  don't run integration tests in parallel.
- COM date filters are locale-specific (currently `MM/DD/YYYY HH:MM`).

## Development

```bash
mise install                              # prek, python, uv
uv sync
prek install                              # pre-commit + pre-push hooks
prek run --all-files --stage pre-push     # everything CI's lint job runs
```

Hooks live in `prek.toml`; Renovate (self-hosted, `renovate.json`) keeps their revs and the
locked dependencies current.
