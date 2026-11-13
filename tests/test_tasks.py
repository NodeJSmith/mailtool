"""
Task functionality tests.

Tests for listing, creating, and managing Outlook tasks.
"""

import pytest

from mailtool.bridge import DATE_PATTERN

from .conftest import TEST_PREFIX, assert_valid_entry_id, from_now


@pytest.mark.integration
@pytest.mark.tasks
class TestTasks:
    """Test task-related functionality"""

    def test_list_tasks(self, bridge):
        """Test listing tasks (default: incomplete only)"""
        tasks = bridge.list_tasks()
        assert isinstance(tasks, list)

        for task in tasks:
            assert isinstance(task, dict)
            assert "entry_id" in task
            assert "subject" in task
            # Default behavior: should only return incomplete tasks
            assert task["complete"] is False

    def test_list_all_tasks(self, bridge):
        """Test listing all tasks including completed"""
        tasks = bridge.list_all_tasks()
        assert isinstance(tasks, list)

        for task in tasks:
            assert isinstance(task, dict)
            assert "entry_id" in task
            assert "subject" in task

    def test_list_tasks_include_completed(self, bridge):
        """Test listing tasks with include_completed flag"""
        # Default: incomplete only
        incomplete_tasks = bridge.list_tasks(include_completed=False)
        assert isinstance(incomplete_tasks, list)
        for task in incomplete_tasks:
            assert task["complete"] is False

        # All tasks
        all_tasks = bridge.list_tasks(include_completed=True)
        assert isinstance(all_tasks, list)
        # all_tasks should be >= incomplete_tasks
        assert len(all_tasks) >= len(incomplete_tasks)

    def test_create_task_basic(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a basic task"""
        subject = f"{TEST_PREFIX}Task Test {test_timestamp}"

        entry_id = bridge.create_task(subject=subject, body="Test task body")

        assert_valid_entry_id(entry_id)

        # Verify we can retrieve it
        task = bridge.get_task(entry_id)
        assert task is not None
        assert task["subject"] == subject

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_create_task_with_due_date(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating a task with a due date"""
        subject = f"{TEST_PREFIX}Due Date Test {test_timestamp}"
        due_date = from_now(days=7).date().format(DATE_PATTERN)

        entry_id = bridge.create_task(
            subject=subject, body="Task with due date", due_date=due_date
        )

        assert_valid_entry_id(entry_id)

        # Verify due date was set
        task = bridge.get_task(entry_id)
        assert task is not None
        assert task["due_date"] == due_date

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_create_task_with_priority(self, bridge, test_timestamp, cleanup_helpers):
        """Test creating tasks with different priority levels"""
        test_priorities = [(0, "Low"), (1, "Normal"), (2, "High")]

        for priority, priority_name in test_priorities:
            subject = f"{TEST_PREFIX}Priority {priority_name} Test {test_timestamp}"

            entry_id = bridge.create_task(
                subject=subject,
                due_date=from_now(days=1).date().format(DATE_PATTERN),
                importance=priority,
            )

            assert_valid_entry_id(entry_id)

            # Verify priority was set
            task = bridge.get_task(entry_id)
            assert task is not None
            assert task["priority"] == priority

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_get_task(self, bridge, sample_task_data):
        """Test retrieving task by EntryID"""
        task = bridge.get_task(sample_task_data["entry_id"])

        assert task is not None
        assert isinstance(task, dict)
        assert task["entry_id"] == sample_task_data["entry_id"]
        assert task["subject"] == sample_task_data["subject"]

        # Verify all expected fields are present
        expected_fields = [
            "entry_id",
            "subject",
            "body",
            "due_date",
            "status",
            "priority",
            "complete",
            "percent_complete",
        ]
        for field in expected_fields:
            assert field in task

    def test_edit_task_subject(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing a task's subject"""
        # Create task
        original_subject = f"{TEST_PREFIX}Edit Subject Test {test_timestamp}"
        entry_id = bridge.create_task(subject=original_subject, body="Test body")

        # Edit subject
        new_subject = f"{TEST_PREFIX}Updated Subject {test_timestamp}"
        result = bridge.edit_task(entry_id, subject=new_subject)

        assert result is True

        # Verify change
        task = bridge.get_task(entry_id)
        assert task["subject"] == new_subject

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_edit_task_body(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing a task's body"""
        # Create task
        subject = f"{TEST_PREFIX}Edit Body Test {test_timestamp}"
        entry_id = bridge.create_task(subject=subject, body="Original body")

        # Edit body
        new_body = "Updated body content"
        result = bridge.edit_task(entry_id, body=new_body)

        assert result is True

        # Verify change
        task = bridge.get_task(entry_id)
        assert task["body"] == new_body

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_edit_task_due_date(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing a task's due date"""
        # Create task
        subject = f"{TEST_PREFIX}Edit Due Date Test {test_timestamp}"
        original_due = from_now(days=5).date().format(DATE_PATTERN)
        entry_id = bridge.create_task(subject=subject, due_date=original_due)

        # Edit due date
        new_due = from_now(days=14).date().format(DATE_PATTERN)
        result = bridge.edit_task(entry_id, due_date=new_due)

        assert result is True

        # Verify change
        task = bridge.get_task(entry_id)
        assert task["due_date"] == new_due

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_edit_task_priority(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing a task's priority"""
        # Create task with normal priority
        subject = f"{TEST_PREFIX}Edit Priority Test {test_timestamp}"
        entry_id = bridge.create_task(
            subject=subject,
            importance=1,  # Normal
        )

        # Edit to high priority
        result = bridge.edit_task(entry_id, importance=2)

        assert result is True

        # Verify change
        task = bridge.get_task(entry_id)
        assert task["priority"] == 2

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_edit_task_percent_complete(self, bridge, test_timestamp, cleanup_helpers):
        """Test editing a task's percent complete"""
        # Create task
        subject = f"{TEST_PREFIX}Percent Complete Test {test_timestamp}"
        entry_id = bridge.create_task(subject=subject)

        # Set to 50% complete
        result = bridge.edit_task(entry_id, percent_complete=50)

        assert result is True

        # Verify change
        task = bridge.get_task(entry_id)
        assert task["percent_complete"] == 50
        assert task["complete"] is False  # Not fully complete

        # Set to 100% complete
        bridge.edit_task(entry_id, percent_complete=100)

        task = bridge.get_task(entry_id)
        assert task["percent_complete"] == 100
        # Note: complete and status may auto-update

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_complete_task(self, bridge, test_timestamp, cleanup_helpers):
        """Test marking a task as complete"""
        # Create task
        subject = f"{TEST_PREFIX}Complete Test {test_timestamp}"
        entry_id = bridge.create_task(subject=subject)

        # Verify not complete initially
        task = bridge.get_task(entry_id)
        assert task["complete"] is False

        # Mark as complete
        result = bridge.complete_task(entry_id)

        assert result is True

        # Verify completion
        task = bridge.get_task(entry_id)
        assert task["complete"] is True
        assert task["percent_complete"] == 100

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)

    def test_delete_task(self, bridge, test_timestamp):
        """Test deleting a task"""
        # Create task
        subject = f"{TEST_PREFIX}Delete Task Test {test_timestamp}"
        entry_id = bridge.create_task(subject=subject)

        # Verify it exists
        task = bridge.get_task(entry_id)
        assert task is not None

        # Delete it
        result = bridge.delete_task(entry_id)
        assert result is True

        # Verify it's gone
        task = bridge.get_task(entry_id)
        assert task is None

    def test_task_status_updates(self, bridge, test_timestamp, cleanup_helpers):
        """Test that task status updates correctly with different percent_complete values"""
        subject = f"{TEST_PREFIX}Status Test {test_timestamp}"
        entry_id = bridge.create_task(subject=subject)

        # 0% should be "Not started" (status=0)
        bridge.edit_task(entry_id, percent_complete=0)
        task = bridge.get_task(entry_id)
        assert task["percent_complete"] == 0
        assert task["status"] == 0  # Not started

        # 50% should be "In progress" (status=1)
        bridge.edit_task(entry_id, percent_complete=50)
        task = bridge.get_task(entry_id)
        assert task["percent_complete"] == 50
        assert task["status"] == 1  # In progress

        # 100% should be "Complete" (status=2)
        bridge.edit_task(entry_id, percent_complete=100)
        task = bridge.get_task(entry_id)
        assert task["percent_complete"] == 100
        assert task["status"] == 2  # Complete

        # Cleanup
        cleanup_helpers["delete_tasks_by_prefix"](TEST_PREFIX)
