---
audience: personal tool
developers: solo
data-sensitivity: internal
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
  environment gets a readable error (plain text on stderr, exit 1). Past those checks every
  command prints JSON and exits 1 on failure; argparse usage errors exit 2 with argparse's
  own text.
- `tests/`: unit tests (pure Python, fake COM objects) plus integration tests that drive a
  live Outlook through the session-scoped `bridge` fixture in `conftest.py`.

## Design decisions

- **O(1) lookups.** Items are fetched with `GetItemFromID(entry_id)`, never by iterating a
  folder. Large mailboxes depend on it.
- **Calendar bomb.** Calendar listing sets `IncludeRecurrences = True` and `Sort("[Start]")`,
  then applies a COM-level `Restrict` before Python iterates. `all_events` is capped to a
  one-year horizon (`CALENDAR_ALL_EVENTS_MAX_DAYS`), because an unbounded scan expands
  recurring series forever and blocks Outlook's single-threaded COM apartment.
- **Filters.** `Items.Restrict` speaks Jet, which has no LIKE. Filters containing LIKE are
  translated to DASL (`@SQL=`) through a verified field map in `bridge.py`.
- **Free/busy** takes an email address directly and defaults to the current user.
- **Return values.** Failures raise `MailtoolError` (a `RuntimeError`), never an empty
  result or a `False` sentinel, because the consumers are unattended digests that read an
  empty list as "nothing new". Lists return `[]` only when nothing matches; a fetch by an
  EntryID that doesn't exist, or whose item is the wrong type (appointment/task methods check
  Outlook's `Class`), returns `None` (not a failure). Only Outlook's not-found errors
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
  prefix and are cleaned up by prefix.
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
