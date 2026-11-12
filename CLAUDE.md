---
audience: personal tool
developers: solo
data-sensitivity: internal
---

# Mailtool: Outlook COM Automation Bridge

A WSL2-to-Windows bridge for Outlook automation via COM, optimized for AI agent integration.

**Version**: 2.3.0 | **Status**: Production/Stable

## Architecture

**Stack**: Python 3.13+ + pywin32 (COM) → Outlook (Windows)

**Entry Points**:
- `outlook.sh` (WSL2) → `outlook.bat` (Windows) → `src/mailtool/bridge.py`
- `run_tests.sh` (WSL2) → `run_tests.bat` (Windows) → `pytest`
- **MCP Server** → `src/mailtool/mcp/server.py` → Claude Code integration (26 tools, 7 resources)

**Dependency Management**: Uses `uv run --with pywin32` for zero-install Windows execution

**MCP Integration**: Model Context Protocol server using official MCP Python SDK v2 with FastMCP framework for Claude Code, Claude Desktop, and other MCP clients

**Development Tools**:
- `ruff` for linting and formatting (replaces Black, isort, Flake8, etc.)
- GitHub Actions CI/CD (windows-latest)
- Pre-commit hooks for code quality

## Key Design Decisions

### O(1) Access Pattern
All item lookups use `GetItemFromID(entry_id)` instead of iteration. This is critical for production use with large mailboxes.

### Recurrence Handling
Calendar events enable `IncludeRecurrences = True` + `Sort("[Start]")`, then apply COM-level `Restrict` filter **before** Python iteration to avoid the "Calendar Bomb" (infinite recurring meetings).

### Path Translation
WSL paths for attachments are auto-converted via `wslpath -w` in `outlook.sh` wrapper before being passed to Windows Python.

### Free/Busy API
Refactored to accept `email_address` directly, defaulting to current user. Legacy `entry_id` parameter supported but deprecated.

## File Structure

```
mailtool/
├── outlook.sh              # WSL2 entry point (translates paths)
├── outlook.bat             # Windows entry point (uv + pywin32)
├── run_tests.sh            # Test runner (WSL2)
├── run_tests.bat           # Test runner (Windows)
├── pytest.ini              # Pytest configuration
├── pyproject.toml          # Project config (uv, ruff, dependencies)
├── mcp_server.py           # Legacy MCP server (v2.2, kept for rollback)
├── test_mcp_server.py      # Legacy MCP server validation script
├── .claude-plugin/
│   └── plugin.json         # Claude Code plugin manifest
├── .github/
│   └── workflows/
│       ├── ci.yml          # Continuous Integration (tests + lint)
│       └── publish.yml     # PyPI publishing
├── src/
│   └── mailtool/
│       ├── __init__.py
│       ├── bridge.py       # Core COM automation (~1100 lines)
│       ├── cli.py          # CLI interface
│       └── mcp/            # MCP SDK v2 package (v2.3)
│           ├── __init__.py
│           ├── server.py   # FastMCP server with 26 tools
│           ├── models.py   # Pydantic models (10 models)
│           ├── resources.py # MCP resources (7 resources)
│           ├── lifespan.py # Outlook bridge lifecycle management
│           └── exceptions.py # Custom exception classes
├── tests/
│   ├── __init__.py
│   ├── conftest.py         # Session fixtures, warmup, cleanup
│   ├── test_bridge.py      # Core connectivity (6 tests)
│   ├── test_emails.py      # Email ops (12 tests)
│   ├── test_calendar.py    # Calendar ops (13 tests)
│   ├── test_tasks.py       # Task ops (13 tests)
│   └── mcp/                # MCP SDK v2 tests (166 tests)
│       ├── test_models.py      # Pydantic model validation (43 tests)
│       ├── test_tools.py       # MCP tool tests (44 tests)
│       ├── test_resources.py   # MCP resource tests (26 tests)
│       ├── test_integration.py # End-to-end workflows (34 tests)
│       └── test_exceptions.py  # Exception handling (19 tests)
└── docs/
    ├── README.md           # Main project README
    ├── CLAUDE.md           # This file - AI assistant guide
    ├── MCP_INTEGRATION.md  # MCP server documentation
    ├── MCP_SUMMARY.md      # MCP implementation summary (v2.2)
    ├── SUMMARY.md          # Proof of concept summary
    ├── FEATURES.md         # Feature list
    ├── COMMANDS.md         # CLI command reference
    ├── QUICKSTART.md       # Quick start guide
    └── PRODUCTION_UPGRADE.md  # v2.0 upgrade notes
```

## API Patterns

### Return Values
- **Draft emails**: Returns `EntryID` (string) for reference
- **Sent emails**: Returns `True` (boolean)
- **Failed ops**: Returns `False` (boolean)
- **Get ops**: Returns `dict` with full item data or `None`

### Test Isolation
All test-created items use `[TEST]` prefix for identification and auto-cleanup. Tests use real Outlook data - no mocking.

## MCP SDK v2 Architecture (v2.3)

### FastMCP Framework

The MCP server uses the official MCP Python SDK v2 with the FastMCP framework for type-safe, declarative tool and resource definitions.

**Key Components**:
- **FastMCP Server**: `src/mailtool/mcp/server.py` - Main server instance with 26 tools and 7 resources
- **Pydantic Models**: `src/mailtool/mcp/models.py` - 10 models for structured output (Email, Calendar, Task)
- **MCP Resources**: `src/mailtool/mcp/resources.py` - 7 resources for data access (Email, Calendar, Task)
- **Lifespan Management**: `src/mailtool/mcp/lifespan.py` - Async context manager for Outlook bridge lifecycle
- **Custom Exceptions**: `src/mailtool/mcp/exceptions.py` - 3 exception types for error handling

### FastMCP Decorator Pattern

All tools use the `@mcp.tool()` decorator for automatic registration and schema generation:

```python
from mcp.server import FastMCP
from mailtool.mcp.models import EmailSummary

mcp = FastMCP(name="mailtool-outlook-bridge", lifespan=outlook_lifespan)

@mcp.tool()
def list_emails(limit: int = 10, folder: str = "Inbox") -> list[EmailSummary]:
    """List emails from the specified folder.

    Args:
        limit: Maximum number of emails to return (default: 10)
        folder: Folder name to list emails from (default: "Inbox")

    Returns:
        list[EmailSummary]: List of email summaries with basic information
    """
    bridge = _get_bridge()
    emails = bridge.list_emails(limit=limit, folder=folder)
    return [EmailSummary(**email) for email in emails]
```

### Pydantic Models

All tools return structured Pydantic models for type safety and LLM understanding:

**Email Models**:
- `EmailSummary`: 7 fields (entry_id, subject, sender, sender_name, received_time, unread, has_attachments)
- `EmailDetails`: 8 fields (extends EmailSummary with body, html_body)
- `SendEmailResult`: 3 fields (success, entry_id, message)

**Calendar Models**:
- `AppointmentSummary`: 12 fields (entry_id, subject, start, end, location, organizer, all_day, required_attendees, optional_attendees, response_status, meeting_status, response_requested)
- `AppointmentDetails`: 13 fields (extends AppointmentSummary with body)
- `CreateAppointmentResult`: 3 fields (success, entry_id, message)
- `FreeBusyInfo`: 5 fields (email, start_date, end_date, free_busy, resolved, error)

**Task Models**:
- `TaskSummary`: 8 fields (entry_id, subject, body, due_date, status, priority, complete, percent_complete)
- `CreateTaskResult`: 3 fields (success, entry_id, message)

**Generic Models**:
- `OperationResult`: 2 fields (success, message) - Used for all boolean operations

### MCP Resources

Resources provide read-only data access using custom URI schemes:

**Email Resources**:
- `inbox://emails` - List recent emails (max 50)
- `inbox://unread` - List unread emails (max 50)
- `email://{entry_id}` - Get full email details (template resource)

**Calendar Resources**:
- `calendar://today` - List today's calendar events
- `calendar://week` - List calendar events for the next 7 days

**Task Resources**:
- `tasks://active` - List active (incomplete) tasks
- `tasks://all` - List all tasks (including completed)

Resources are registered using the `@mcp.resource()` decorator:

```python
@mcp.resource(uri="calendar://today")
def get_calendar_today() -> str:
    """Get today's calendar events as formatted text."""
    bridge = _get_bridge()
    events = bridge.list_calendar_events(days=1)
    return _format_appointments(events)
```

### Lifespan Management

The Outlook bridge lifecycle is managed by an async context manager:

```python
@asynccontextmanager
async def outlook_lifespan(app):
    """Async context manager for Outlook bridge lifecycle.

    1. Creates OutlookBridge instance on startup
    2. Warms up the connection with retry attempts (5 retries, 0.5s delay)
    3. Sets module-level bridge state for tool access
    4. Cleans up COM objects and forces garbage collection on shutdown
    """
    bridge = None
    try:
        # Create and warm up bridge
        bridge = await loop.run_in_executor(None, _create_bridge)
        await _warmup_bridge(bridge)
        app._bridge = bridge  # Set module-level state
        yield
    finally:
        # Cleanup COM objects
        if bridge:
            bridge.cleanup()
            gc.collect()
```

### Custom Exceptions

Three custom exception classes extend `McpError` for structured error handling:

```python
class OutlookNotFoundError(McpError):
    """Raised when an Outlook item is not found."""
    def __init__(self, entry_id: str):
        super().__init__(
            ErrorData(
                code=-32602,
                message=f"Outlook item not found: {entry_id}",
                data={"entry_id": entry_id}
            )
        )

class OutlookComError(McpError):
    """Raised when COM/bridge operations fail."""
    def __init__(self, details: str):
        super().__init__(
            ErrorData(
                code=-32603,
                message=f"Outlook COM error: {details}",
                data={"details": details}
            )
        )

class OutlookValidationError(McpError):
    """Raised when input validation fails."""
    def __init__(self, field: str, message: str):
        super().__init__(
            ErrorData(
                code=-32604,
                message=f"Validation error: {message}",
                data={"field": field}
            )
        )
```

## Recent Changes

### v2.3.0 (Current - MCP SDK v2 Migration)

1. **MCP SDK v2**: Migrated from hand-rolled MCP implementation to official MCP Python SDK v2
2. **FastMCP Framework**: Using FastMCP for type-safe, declarative tool and resource definitions
3. **Structured Output**: All 26 tools return Pydantic models (10 models: Email, Calendar, Task)
4. **MCP Resources**: Added 7 resources for read-only data access (Email, Calendar, Task)
5. **Async Lifespan**: Async context manager for Outlook bridge lifecycle (creation, warmup, cleanup)
6. **Custom Exceptions**: 3 exception types (OutlookNotFoundError, OutlookComError, OutlookValidationError)
7. **Enhanced Logging**: Comprehensive logging for debugging and monitoring (stderr)
8. **Test Coverage**: 170 MCP tests (models, tools, resources, integration, exceptions)
9. **Type Safety**: TYPE_CHECKING blocks for clean type hints without circular imports
10. **Module Organization**: 5-module MCP package (server, models, resources, lifespan, exceptions)

### v2.2.0 (Previous - MCP Integration)

1. **MCP Server**: Added Model Context Protocol server for Claude Code integration
2. **24 MCP Tools**: Email (10), Calendar (7), Tasks (7) operations exposed via JSON-RPC
3. **Plugin Manifest**: `.claude-plugin/plugin.json` for auto-loading in Claude Code
4. **Zero-Config MCP**: Uses `uv run --with pywin32` for dependency-free execution
5. **MCP Tools Integration**: 24 tools exposed via MCP for enhanced AI workflow
6. **Task Analysis**: List tasks, analyze by subject/deadline, recommend cleanup actions
7. **Pre-commit Hooks**: Automated code quality checks via pre-commit
8. **GitHub CI**: Automated testing and linting on Windows runners
9. **Ruff Integration**: Replaced multiple linters with unified ruff configuration

### v2.1.0 (Production Release)

1. **Calendar Bomb Fix**: Added `items.Restrict()` before iterating in `list_calendar_events()`
2. **WSL Path Translation**: Auto-convert attachment paths in `outlook.sh`
3. **Free/Busy Refactor**: Accepts `email_address` directly, defaults to current user
4. **Return Value Docs**: Clarified different return types in `send_email()` docstring
5. **Package Restructure**: Migrated from single-file to proper Python package structure

## Running Tests

```bash
# From WSL2
./run_tests.sh                 # All tests
./run_tests.sh -m email        # Email tests only
./run_tests.sh -m "not slow"   # Skip slow tests

# From Windows
run_tests.bat

# Test MCP server (requires Outlook running)
python test_mcp_server.py
```

## MCP Usage

### Installation

**IMPORTANT**: Due to [GitHub issue #16143](https://github.com/anthropics/claude-code/issues/16143), the `.claude-plugin/plugin.json` auto-loading is currently broken. Manual configuration is required.

#### Step 1: Install mailtool in editable mode

```bash
cd C:\dev\mailtool
uv pip install -e .
```

#### Step 2: Configure MCP server in user settings

Add the following to `C:\Users\Sam\.claude.json` in the `mcpServers` section:

**Basic configuration (no default account):**
```json
{
  "mcpServers": {
    "mailtool": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--with",
        "pywin32",
        "-m",
        "mailtool.mcp.server"
      ],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

**With default account specified:**
```json
{
  "mcpServers": {
    "mailtool": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--with",
        "pywin32",
        "-m",
        "mailtool.mcp.server",
        "--account",
        "s.mok@utwente.nl"
      ],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

#### Step 3: Restart Claude Code

The MCP server will auto-start on Claude Code launch. Ensure Outlook is running on Windows.

### Available MCP Tools

**Email (11 tools)**: `list_emails`, `list_unread_emails`, `get_email`, `send_email`, `reply_email`, `forward_email`, `mark_email`, `move_email`, `delete_email`, `search_emails`, `search_emails_by_sender`

**Calendar (7 tools)**: `list_calendar_events`, `create_appointment`, `get_appointment`, `edit_appointment`, `respond_to_meeting`, `delete_appointment`, `get_free_busy`

**Tasks (7 tools)**: `list_tasks`, `list_all_tasks`, `create_task`, `get_task`, `edit_task`, `complete_task`, `delete_task`

### Available MCP Resources

**Email Resources (3)**: `inbox://emails`, `inbox://unread`, `email://{entry_id}`

**Calendar Resources (2)**: `calendar://today`, `calendar://week`

**Task Resources (2)**: `tasks://active`, `tasks://all`

### Example Claude Code Interactions

```
You: Show me my last 5 unread emails

You: Create a task "Review Q1 report" due Friday with high priority

You: Schedule a team meeting for tomorrow at 2pm in Room 101

You: Accept the meeting invitation from John

You: What's on my calendar this week?
```

See [MCP_INTEGRATION.md](MCP_INTEGRATION.md) for complete documentation.

## Known Limitations

- **Date Format**: Outlook COM filters use locale-specific formats (currently MM/DD/YYYY HH:MM)
- **Parallel Execution**: COM is apartment-threaded; true parallel test execution not recommended
- **Sent Item ID**: Sent emails move to Sent Items with new EntryID (can't return original ID)

## Development Notes

- **COM Threading**: All COM calls must happen on same thread (session-scoped bridge fixture)
- **Warmup**: Tests include 2-5s warmup to ensure Outlook is responsive
- **Cleanup**: Test artifacts auto-cleaned via prefix-based deletion helpers

## Development Workflow

```bash
# Install dependencies (managed by uv)
uv sync --all-groups

# Run linter and formatter
uv run ruff check .           # Check code
uv run ruff check --fix .     # Auto-fix issues
uv run ruff format .          # Format code

# Run tests
./run_tests.sh                # All tests (WSL2)
run_tests.bat                 # All tests (Windows)
uv run pytest -v              # Direct pytest
uv run pytest -m email        # Run specific marker

# Run MCP server tests (requires Outlook running on Windows)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/ -v

# Run specific MCP test modules
uv run --with pytest --with pywin32 python -m pytest tests/mcp/test_models.py -v      # Pydantic model tests (43 tests)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/test_tools.py -v       # MCP tool tests (44 tests)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/test_resources.py -v   # MCP resource tests (26 tests)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/test_integration.py -v # Integration tests (34 tests)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/test_exceptions.py -v  # Exception tests (19 tests)

# Test MCP server manually (requires Outlook running)
uv run --with mcp --with pywin32 python -m mailtool.mcp.server

# Test MCP server with default account
uv run --with mcp --with pywin32 python -m mailtool.mcp.server --account "s.mok@utwente.nl"

# Add new dependency
uv add <package>

# Run pre-commit hooks manually
uv run pre-commit run --all-files
```

## Code Quality

- **Python Version**: Requires Python 3.13+
- **Linter/Formatter**: Ruff (unified tool replacing Black, isort, Flake8)
- **Line Length**: 88 characters (Black default)
- **CI/CD**: GitHub Actions runs tests and linting on Windows runners
- **Pre-commit Hooks**: Ensures code quality before commits

## Architecture Patterns

### Bridge Class (`src/mailtool/bridge.py`)
- **O(1) Lookups**: Uses `GetItemFromID(entry_id)` for all item access
- **Safe Attribute Access**: `_safe_get_attr()` wrapper for COM objects
- **Folder Access**: `get_inbox()`, `get_calendar()`, `get_tasks()`, `get_folder_by_name()`
- **Error Handling**: Returns `False` for failures, `None` for not found, EntryID string for drafts

### CLI Interface (`src/mailtool/cli.py`)
- **Entry Point**: `mailtool` command (installed via `uv add`)
- **Subcommands**: `emails`, `email`, `calendar`, `tasks`
- **Output**: JSON-formatted responses for easy parsing

### MCP Server (`src/mailtool/mcp/server.py`)
- **Framework**: FastMCP (MCP Python SDK v2)
- **Protocol**: JSON-RPC via stdio transport
- **Tool Registration**: Declarative `@mcp.tool()` decorator with Pydantic type hints
- **Resource Registration**: Declarative `@mcp.resource()` decorator with custom URI schemes
- **Lifespan Management**: Async context manager for Outlook bridge lifecycle
- **Structured Output**: All tools return Pydantic models (type-safe, self-documenting)
- **Error Handling**: Custom exception classes (OutlookNotFoundError, OutlookComError, OutlookValidationError)
- **Logging**: Comprehensive logging to stderr for debugging and monitoring
- **Initialization**: Single Outlook bridge instance per session (module-level state)
- **Module-Level Bridge Pattern**: `server._bridge` set by lifespan, accessed via `_get_bridge()` helper

### MCP Models (`src/mailtool/mcp/models.py`)
- **Purpose**: Pydantic models for structured output from MCP tools
- **Type Safety**: All models use Field() descriptions for LLM understanding
- **Validation**: Automatic validation and serialization via Pydantic
- **Email Models**: EmailSummary (7 fields), EmailDetails (8 fields), SendEmailResult (3 fields)
- **Calendar Models**: AppointmentSummary (12 fields), AppointmentDetails (13 fields), CreateAppointmentResult (3 fields), FreeBusyInfo (6 fields)
- **Task Models**: TaskSummary (8 fields), CreateTaskResult (3 fields)
- **Generic Models**: OperationResult (2 fields) - Used for all boolean operations

### MCP Resources (`src/mailtool/mcp/resources.py`)
- **Purpose**: Read-only data access via custom URI schemes
- **Email Resources**: inbox://emails, inbox://unread, email://{entry_id}
- **Calendar Resources**: calendar://today, calendar://week
- **Task Resources**: tasks://active, tasks://all
- **Format**: Resources return formatted text (str), not Pydantic models
- **Helper Functions**: Formatting functions for each resource type

### MCP Lifespan (`src/mailtool/mcp/lifespan.py`)
- **Purpose**: Async context manager for Outlook bridge lifecycle
- **Startup**: Create OutlookBridge instance via thread pool executor (COM is synchronous)
- **Warmup**: Retry attempts (5 retries, 0.5s delay) to ensure Outlook is responsive
- **State Management**: Set module-level bridge state in `server._bridge` and `resources._bridge`
- **Shutdown**: Release COM objects and force garbage collection
- **Logging**: Info level for lifecycle events, debug for operational details, error for failures

### MCP Exceptions (`src/mailtool/mcp/exceptions.py`)
- **OutlookNotFoundError**: Raised when an Outlook item is not found (entry_id attribute)
- **OutlookComError**: Raised when COM/bridge operations fail (details attribute)
- **OutlookValidationError**: Raised when input validation fails (field attribute)
- **Error Codes**: -32602 (not found), -32603 (COM error), -32604 (validation error)
- **Base Class**: All exceptions extend McpError with ErrorData for structured error information
