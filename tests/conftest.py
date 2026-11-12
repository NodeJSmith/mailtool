"""
Pytest configuration and fixtures for mailtool Outlook bridge tests

This module provides shared fixtures and test utilities for testing the Outlook bridge.
All tests use real Outlook data - no mocking.
"""

import time
import uuid
from datetime import datetime, timedelta

import pytest

from mailtool.bridge import OutlookBridge

# Prefix for all test-created items (enables easy identification and cleanup)
TEST_PREFIX = "[TEST] "

# How long to wait for Outlook warmup (in seconds)
WARMUP_TIMEOUT = 30
WARMUP_ATTEMPTS = 5
WARMUP_DELAY = 0.5


@pytest.fixture(scope="session")
def bridge():
    """
    Session-scoped Outlook bridge instance.
    Created once and reused across all tests for speed.
    Includes warmup period to ensure Outlook is responsive.
    """
    print("\n" + "=" * 70)
    print("INITIALIZING OUTLOOK BRIDGE (Session-scoped)")
    print("=" * 70)

    # Create bridge instance
    bridge_instance = OutlookBridge()
    print("✓ Bridge instance created")

    # Warmup: Ensure Outlook is fully responsive
    print("Warming up Outlook connection...")

    for attempt in range(WARMUP_ATTEMPTS):
        try:
            # Try a quick operation to verify connection
            inbox = bridge_instance.get_inbox()
            if inbox is not None:
                # Force a COM call by accessing Items
                _ = inbox.Items.Count
                print(f"✓ Outlook ready after {attempt + 1} attempt(s)")
                break
        except Exception as e:
            if attempt < WARMUP_ATTEMPTS - 1:
                print(f"  Attempt {attempt + 1} failed: {e}, retrying...")
                time.sleep(WARMUP_DELAY)
            else:
                print(f"✗ Outlook warmup failed after {WARMUP_ATTEMPTS} attempts")
                raise

    print("=" * 70 + "\n")
    yield bridge_instance

    # No explicit cleanup needed - bridge is just a COM reference
    print("\n" + "=" * 70)
    print("TEARDOWN: Bridge session complete")
    print("=" * 70)


@pytest.fixture
def test_timestamp():
    """
    Unique timestamp for each test.
    Useful for creating uniquely-named test items.
    """
    return f"{int(time.time())}_{uuid.uuid4().hex[:8]}"


@pytest.fixture
def cleanup_helpers(bridge):
    """
    Provides helper functions for test cleanup.

    Returns a dict with cleanup functions:
    - delete_drafts_by_prefix(prefix): Delete all draft emails with subject prefix
    - delete_calendar_by_prefix(prefix): Delete all calendar items with subject prefix
    - delete_tasks_by_prefix(prefix): Delete all tasks with subject prefix
    """
    helpers = {}

    def delete_drafts_by_prefix(prefix):
        """Delete all draft emails matching the prefix"""
        drafts = bridge.get_folder_by_name("Drafts") or bridge.get_inbox()
        items = drafts.Items
        items.Sort("[CreationTime]", False)  # Newest first

        deleted_count = 0
        for item in items:
            try:
                if (
                    hasattr(item, "Subject")
                    and item.Subject
                    and item.Subject.startswith(prefix)
                ):
                    item.Delete()
                    deleted_count += 1
            except Exception:
                pass

        if deleted_count > 0:
            print(f"  Cleanup: Deleted {deleted_count} draft(s)")
        return deleted_count

    def delete_calendar_by_prefix(prefix):
        """Delete all calendar items matching the prefix"""
        calendar = bridge.get_calendar()
        items = calendar.Items

        deleted_count = 0
        for item in items:
            try:
                if (
                    hasattr(item, "Subject")
                    and item.Subject
                    and item.Subject.startswith(prefix)
                ):
                    item.Delete()
                    deleted_count += 1
            except Exception:
                pass

        if deleted_count > 0:
            print(f"  Cleanup: Deleted {deleted_count} calendar event(s)")
        return deleted_count

    def delete_tasks_by_prefix(prefix):
        """Delete all tasks matching the prefix"""
        tasks = bridge.get_tasks()
        items = tasks.Items

        deleted_count = 0
        for item in items:
            try:
                if (
                    hasattr(item, "Subject")
                    and item.Subject
                    and item.Subject.startswith(prefix)
                ):
                    item.Delete()
                    deleted_count += 1
            except Exception:
                pass

        if deleted_count > 0:
            print(f"  Cleanup: Deleted {deleted_count} task(s)")
        return deleted_count

    helpers["delete_drafts_by_prefix"] = delete_drafts_by_prefix
    helpers["delete_calendar_by_prefix"] = delete_calendar_by_prefix
    helpers["delete_tasks_by_prefix"] = delete_tasks_by_prefix

    return helpers

    # Automatic cleanup of test artifacts is handled by individual tests
    # This fixture just provides the helpers


@pytest.fixture
def sample_email_data(bridge, cleanup_helpers):
    """
    Creates a sample draft email for testing.
    Automatically cleans up after the test.
    """
    test_unique = f"{int(time.time())}_{uuid.uuid4().hex[:8]}"
    subject = f"{TEST_PREFIX}Sample Email {test_unique}"
    body = f"This is a test email created at {time.ctime()}.\n\nTest ID: {test_unique}"

    entry_id = bridge.send_email(
        to="test@example.com", subject=subject, body=body, save_draft=True
    )

    yield {
        "entry_id": entry_id,
        "subject": subject,
        "body": body,
        "test_id": test_unique,
    }

    # Cleanup
    cleanup_helpers["delete_drafts_by_prefix"](TEST_PREFIX)


@pytest.fixture
def sample_calendar_data(bridge, cleanup_helpers):
    """
    Creates a sample calendar event for testing.
    Automatically cleans up after the test.
    """
    test_unique = f"{int(time.time())}_{uuid.uuid4().hex[:8]}"
    subject = f"{TEST_PREFIX}Sample Event {test_unique}"

    start = datetime.now() + timedelta(hours=1)
    end = start + timedelta(minutes=30)

    entry_id = bridge.create_appointment(
        subject=subject,
        start=start.strftime("%Y-%m-%d %H:%M:%S"),
        end=end.strftime("%Y-%m-%d %H:%M:%S"),
        location="Test Location",
        body="Test event body",
    )

    yield {"entry_id": entry_id, "subject": subject, "start": start, "end": end}

    # Cleanup
    cleanup_helpers["delete_calendar_by_prefix"](TEST_PREFIX)


@pytest.fixture
def sample_task_data(bridge, cleanup_helpers):
    """
    Creates a sample task for testing.
    Automatically cleans up after the test.
    """
    test_unique = f"{int(time.time())}_{uuid.uuid4().hex[:8]}"
    subject = f"{TEST_PREFIX}Sample Task {test_unique}"

    due_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")

    entry_id = bridge.create_task(
        subject=subject, body="Test task body", due_date=due_date
    )

    yield {"entry_id": entry_id, "subject": subject, "due_date": due_date}

    # Cleanup
    cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)


def assert_valid_entry_id(entry_id):
    """Helper to verify an EntryID looks valid"""
    assert entry_id is not None
    assert isinstance(entry_id, str)
    assert len(entry_id) > 10  # EntryIDs are typically long strings


def assert_email_structure(email):
    """Helper to verify email data has expected structure"""
    assert email is not None
    assert isinstance(email, dict)
    # May contain any of these fields depending on the call
    expected_fields = ["entry_id", "subject", "sender", "sender_name"]
    assert any(field in email for field in expected_fields)


def assert_calendar_structure(event):
    """Helper to verify calendar event has expected structure"""
    assert event is not None
    assert isinstance(event, dict)
    assert "entry_id" in event
    assert "subject" in event
    assert "start" in event


def assert_task_structure(task):
    """Helper to verify task has expected structure"""
    assert task is not None
    assert isinstance(task, dict)
    assert "entry_id" in task
    assert "subject" in task
