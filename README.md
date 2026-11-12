# mailtool

A Python library and CLI for reading and acting on Outlook (classic) email, calendar, and
tasks through Windows COM automation. It runs on Windows, next to a running Outlook, and is
called from WSL2 through Windows interop (`mailtool.exe`).

Originally forked from [utsmok/mailtool](https://github.com/utsmok/mailtool); now maintained
independently. The MCP server the upstream project shipped has been removed; this repo is
the CLI and library only.

## Install

On Windows, with [uv](https://github.com/astral-sh/uv):

```powershell
uv tool install git+https://github.com/NodeJSmith/mailtool
```

Outlook must be installed, running, and signed in. From WSL2, call it as `mailtool.exe`.

## Usage

Every command prints JSON; exit code 1 means the operation failed.

```bash
mailtool emails --limit 5                 # recent inbox emails
mailtool email --id <entry_id>            # one email with body
mailtool calendar --days 7                # upcoming events (recurrences expanded, capped)
mailtool tasks                            # active tasks
mailtool send --to a@example.com --subject Hi --body Hello
mailtool --help                           # full command list
```

Commands: `emails`, `email`, `send`, `reply`, `forward`, `mark`, `move`, `delete-email`,
`calendar`, `edit-appt`, `delete-appt`, `freebusy`, `tasks`, `task`, `create-task`,
`edit-task`, `delete-task`.

As a library:

```python
from mailtool.bridge import OutlookBridge

bridge = OutlookBridge()
for email in bridge.list_emails(limit=5):
    print(email["subject"], email["sender"])
```

## Development

```bash
uv sync
uv run pytest            # live-Outlook tests skip unless on Windows with Outlook
uv run ruff check .
uv run ruff format .
```

See [CLAUDE.md](CLAUDE.md) for architecture notes and design decisions.
