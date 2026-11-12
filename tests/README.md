# Test Suite for Mailtool Outlook Bridge

This directory contains integration tests for the mailtool Outlook COM bridge.

## Overview

These tests use **real Outlook data and COM automation** - there is no mocking. The tests are designed to be fast, isolated, and safe to run against your actual Outlook data.

## Key Design Principles

1. **Fast Execution**: Tests use a session-scoped bridge instance that is created once and reused
2. **Warmup Phase**: The bridge warms up the Outlook connection before tests run
3. **Test Isolation**: Test-created items use a `[TEST]` prefix for easy identification and cleanup
4. **Real Data**: Tests work with your actual Outlook folders and data
5. **Defensive**: Tests verify structure and behavior, not specific data content

## Running Tests

### From WSL2
```bash
# Run all tests with default settings
./run_tests.sh

# Run specific test file
./run_tests.sh tests/test_emails.py

# Run specific test
./run_tests.sh tests/test_emails.py::TestEmails::test_create_draft_email

# Run with verbose output
./run_tests.sh -v -s

# Run only quick tests (skip slow ones)
./run_tests.sh -m "not slow"

# Run only email tests
./run_tests.sh -m email
```

### From Windows
```cmd
REM Run all tests
run_tests.bat

REM Run with specific pytest arguments
run_tests.bat -v -s tests/test_bridge.py
```

### Direct with uv (Windows)
```cmd
cd /path/to/mailtool
uv run --with pywin32 --with pytest pytest
```

## Test Structure

### Fixtures (`conftest.py`)

- **`bridge`** (session-scoped): Single Outlook bridge instance, created with warmup
- **`test_timestamp`**: Unique timestamp for each test
- **`cleanup_helpers`**: Functions to delete test artifacts by prefix
- **`sample_email_data`**: Creates a test draft email
- **`sample_calendar_data`**: Creates a test calendar event
- **`sample_task_data`**: Creates a test task

### Test Files

- **`test_bridge.py`**: Core connectivity and basic operations
- **`test_emails.py`**: Email listing, creation, search, operations
- **`test_calendar.py`**: Calendar events, appointments, free/busy
- **`test_tasks.py`**: Task management, completion, status updates

## Test Markers

Tests are categorized with pytest markers:

```bash
# Run only specific categories
pytest -m email          # Email tests only
pytest -m calendar       # Calendar tests only
pytest -m tasks          # Task tests only
pytest -m slow           # Slower tests
pytest -m "not slow"     # Skip slow tests
```

## Test Isolation and Safety

### Test Prefix Convention

All test-created items use the `[TEST]` prefix in their subject/name:

```
[TEST] Sample Email 1704067200_abc123
[TEST] Sample Event 1704067200_def456
[TEST] Sample Task 1704067200_ghi789
```

### Automatic Cleanup

Tests that create items automatically clean them up using the `cleanup_helpers` fixture:

```python
def test_something(bridge, cleanup_helpers, test_timestamp):
    # Create test items...
    # Test passes or fails...
    # Cleanup runs automatically
    cleanup_helpers["delete_drafts_by_prefix"]("[TEST] ")
```

### Manual Cleanup

If tests are interrupted, you can manually clean up test artifacts:

```python
# In Python
bridge = OutlookBridge()
# Delete test drafts, calendar items, tasks
```

Or manually in Outlook:
- Search for `[TEST]` in subjects
- Delete all matching items

## Performance Characteristics

- **Warmup**: ~2-5 seconds for Outlook COM initialization
- **Individual tests**: ~0.1-0.5 seconds each
- **Full suite**: ~10-30 seconds (depending on mailbox size)

## Troubleshooting

### "Outlook not responding" during warmup

Increase the warmup timeout in `conftest.py`:

```python
WARMUP_TIMEOUT = 60  # Increase from 30
WARMUP_ATTEMPTS = 10  # Increase from 5
```

### Tests fail intermittently

- Outlook may be syncing or busy
- Try running tests again
- Close other Outlook add-ins that might interfere

### COM errors

- Ensure Outlook is installed
- Try repairing your Outlook installation
- Check Windows Event Viewer for COM errors

## Adding New Tests

1. Create a new test class with appropriate markers
2. Use the `bridge` fixture for access to OutlookBridge
3. Use `test_timestamp` for unique identifiers
4. Use `cleanup_helpers` for automatic cleanup
5. Follow the existing test structure and patterns

Example:

```python
@pytest.mark.integration
@pytest.mark.email
class TestMyNewFeature:
    def test_new_feature(self, bridge, test_timestamp, cleanup_helpers):
        subject = f"[TEST] My Test {test_timestamp}"

        # Create test data
        result = bridge.my_new_method(subject)

        # Assert
        assert result is not None

        # Cleanup
        cleanup_helpers["delete_drafts_by_prefix"]("[TEST] ")
```

## Continuous Integration

For CI/CD pipelines:

1. Ensure Outlook is installed on the CI agent
2. Use a dedicated test account or Outlook profile
3. Run tests with: `./run_tests.sh -x -v` (fail fast, verbose)
4. Consider increasing timeouts for slower CI environments
