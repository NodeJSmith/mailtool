# Test Suite for Mailtool Outlook Bridge

This directory holds two kinds of tests.

## Unit tests (run anywhere)

Pure-Python tests that drive `OutlookBridge` and the CLI through fake COM objects
(`OutlookBridge.__new__(OutlookBridge)` skips the COM connection, then folders, items and
namespaces are replaced by small fakes). They need no Outlook and no pywin32, and run on
Linux in CI. Examples: `test_filters.py`, `test_folder_resolution.py`,
`test_item_lookup_and_safe_get.py`, `test_fail_loudly.py`, `test_enhancements.py`.

New unit tests must not request the `bridge` fixture.

## Integration tests (Windows + Outlook)

`test_bridge.py`, `test_emails.py`, `test_calendar.py` and `test_tasks.py` use **real Outlook
data and COM automation**, with no mocking. They use a session-scoped bridge instance that
is created once and warmed up before tests run. Any test that requests the `bridge` fixture
(directly or through another fixture) is marked `windows` automatically by `conftest.py` and
skipped where pywin32 isn't available, so no manual marker is needed.

Key design principles:

1. **Fast execution**: one bridge for the whole session
2. **Warmup phase**: the connection to Outlook is warmed up before tests run
3. **Test isolation**: test-created items use a `[TEST]` prefix for identification and cleanup
4. **Real data**: tests work with your actual Outlook folders
5. **Defensive**: tests verify structure and behavior, not specific data content

## Running Tests

```bash
uv run pytest                      # everything; live-Outlook tests skip off Windows
uv run pytest -m unit              # only the pure-Python unit tests
uv run pytest -m windows           # only the live-Outlook tests (Windows + Outlook running)
uv run pytest -m "windows and not slow"
uv run pytest -m email             # one feature area (also: calendar, tasks)
uv run pytest tests/test_emails.py::TestEmails::test_create_draft_email
```

On Windows, run the live tests from a shell with Outlook open, with pywin32 available
(`uv run --with pywin32 pytest -m windows`). COM is apartment-threaded, so don't run the
integration tests in parallel.

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

The unit tests run in CI (`tests.yml`, ubuntu). The live-Outlook tests have no hosted CI; to run them on an agent:

1. Ensure Outlook is installed on the CI agent
2. Use a dedicated test account or Outlook profile
3. Run tests with: `uv run pytest -m windows -x` (fail fast)
4. Consider increasing timeouts for slower CI environments
