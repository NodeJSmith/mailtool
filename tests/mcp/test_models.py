"""
Model Validation Tests

Tests all Pydantic models to ensure they validate data correctly.
This test suite runs without requiring Outlook to be running.
"""

import pytest
from pydantic import ValidationError

from mailtool.mcp.models import (
    AppointmentDetails,
    AppointmentSummary,
    CreateAppointmentResult,
    CreateTaskResult,
    EmailDetails,
    EmailSummary,
    FreeBusyInfo,
    OperationResult,
    SendEmailResult,
    TaskSummary,
)


class TestEmailSummary:
    """Test EmailSummary model validation"""

    def test_valid_email_summary(self):
        """Test EmailSummary accepts valid data"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": "2025-01-19 10:30:00",
            "unread": True,
            "has_attachments": False,
        }
        email = EmailSummary(**data)
        assert email.entry_id == "test-entry-id-123"
        assert email.subject == "Test Subject"
        assert email.sender == "sender@example.com"
        assert email.sender_name == "Test Sender"
        assert email.received_time == "2025-01-19 10:30:00"
        assert email.unread is True
        assert email.has_attachments is False

    def test_email_summary_with_none_received_time(self):
        """Test EmailSummary accepts None for received_time"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": None,
            "unread": False,
            "has_attachments": True,
        }
        email = EmailSummary(**data)
        assert email.received_time is None

    def test_email_summary_missing_required_fields(self):
        """Test EmailSummary raises ValidationError for missing required fields"""
        data = {
            "entry_id": "test-entry-id-123",
            # Missing: subject, sender, sender_name, unread, has_attachments
        }
        with pytest.raises(ValidationError):
            EmailSummary(**data)


class TestEmailDetails:
    """Test EmailDetails model validation"""

    def test_valid_email_details(self):
        """Test EmailDetails accepts valid data"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": "2025-01-19 10:30:00",
            "has_attachments": False,
        }
        email = EmailDetails(**data)
        assert email.entry_id == "test-entry-id-123"
        assert email.subject == "Test Subject"
        assert email.body == "Plain text body"
        assert email.html_body == "<html><body>HTML body</body></html>"
        assert email.received_time == "2025-01-19 10:30:00"
        assert email.has_attachments is False

    def test_email_details_with_none_received_time(self):
        """Test EmailDetails accepts None for received_time"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": None,
            "has_attachments": False,
        }
        email = EmailDetails(**data)
        assert email.received_time is None

    def test_email_details_missing_required_fields(self):
        """Test EmailDetails raises ValidationError for missing required fields"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            # Missing: sender, sender_name, body, html_body, has_attachments
        }
        with pytest.raises(ValidationError):
            EmailDetails(**data)


class TestSendEmailResult:
    """Test SendEmailResult model validation"""

    def test_sent_email_result(self):
        """Test SendEmailResult for successfully sent email"""
        data = {
            "success": True,
            "entry_id": None,
            "message": "Email sent successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id is None
        assert result.message == "Email sent successfully"

    def test_draft_saved_result(self):
        """Test SendEmailResult for saved draft"""
        data = {
            "success": True,
            "entry_id": "draft-entry-id-456",
            "message": "Draft saved successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id == "draft-entry-id-456"
        assert result.message == "Draft saved successfully"

    def test_failed_send_result(self):
        """Test SendEmailResult for failed email"""
        data = {
            "success": False,
            "entry_id": None,
            "message": "Failed to send email",
        }
        result = SendEmailResult(**data)
        assert result.success is False
        assert result.entry_id is None
        assert result.message == "Failed to send email"

    def test_send_result_default_entry_id(self):
        """Test SendEmailResult entry_id defaults to None"""
        data = {
            "success": True,
            "message": "Email sent successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id is None
        assert result.message == "Email sent successfully"

    def test_send_result_missing_required_fields(self):
        """Test SendEmailResult raises ValidationError for missing required fields"""
        data = {
            "success": True,
            # Missing: message
        }
        with pytest.raises(ValidationError):
            SendEmailResult(**data)


class TestModelSerialization:
    """Test model serialization and deserialization"""

    def test_email_summary_serialization(self):
        """Test EmailSummary can be serialized to dict and JSON"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": "2025-01-19 10:30:00",
            "unread": True,
            "has_attachments": False,
            "message_class": "IPM.Note",
            "to": "",
            "cc": "",
            "sent_time": None,
            "conversation_id": None,
            "conversation_topic": None,
        }
        email = EmailSummary(**data)
        # Test model_dump
        dumped = email.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = email.model_dump_json()
        assert "test-entry-id-123" in json_str
        assert "Test Subject" in json_str

    def test_email_details_serialization(self):
        """Test EmailDetails can be serialized to dict and JSON"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": "2025-01-19 10:30:00",
            "has_attachments": False,
            "message_class": "IPM.Note",
            "to": "",
            "cc": "",
            "bcc": "",
            "sent_time": None,
            "conversation_id": None,
            "conversation_topic": None,
            "attachments": [],
            "body_top": "",
        }
        email = EmailDetails(**data)
        # Test model_dump
        dumped = email.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = email.model_dump_json()
        assert "Plain text body" in json_str

    def test_send_email_result_serialization(self):
        """Test SendEmailResult can be serialized to dict and JSON"""
        data = {
            "success": True,
            "entry_id": "draft-entry-id-456",
            "message": "Draft saved successfully",
        }
        result = SendEmailResult(**data)
        # Test model_dump
        dumped = result.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = result.model_dump_json()
        assert "draft-entry-id-456" in json_str


class TestAppointmentSummary:
    """Test AppointmentSummary model validation"""

    def test_valid_appointment_summary(self):
        """Test AppointmentSummary accepts valid data"""
        data = {
            "entry_id": "apt-entry-id-123",
            "subject": "Team Meeting",
            "start": "2025-01-20 14:00:00",
            "end": "2025-01-20 15:00:00",
            "location": "Conference Room A",
            "organizer": "organizer@example.com",
            "all_day": False,
            "required_attendees": "attendee1@example.com; attendee2@example.com",
            "optional_attendees": "optional@example.com",
            "response_status": "Accepted",
            "meeting_status": "Meeting",
            "response_requested": True,
        }
        appointment = AppointmentSummary(**data)
        assert appointment.entry_id == "apt-entry-id-123"
        assert appointment.subject == "Team Meeting"
        assert appointment.start == "2025-01-20 14:00:00"
        assert appointment.end == "2025-01-20 15:00:00"
        assert appointment.location == "Conference Room A"
        assert appointment.organizer == "organizer@example.com"
        assert appointment.all_day is False
        assert appointment.response_status == "Accepted"
        assert appointment.meeting_status == "Meeting"

    def test_appointment_summary_with_none_times(self):
        """Test AppointmentSummary accepts None for start/end/organizer"""
        data = {
            "entry_id": "apt-entry-id-456",
            "subject": "All Day Event",
            "start": None,
            "end": None,
            "location": "",
            "organizer": None,
            "all_day": True,
            "required_attendees": "",
            "optional_attendees": "",
            "response_status": "Organizer",
            "meeting_status": "NonMeeting",
            "response_requested": False,
        }
        appointment = AppointmentSummary(**data)
        assert appointment.start is None
        assert appointment.end is None
        assert appointment.organizer is None

    def test_appointment_summary_missing_required_fields(self):
        """Test AppointmentSummary raises ValidationError for missing required fields"""
        data = {
            "entry_id": "apt-entry-id-789",
            # Missing: subject, response_status, meeting_status, response_requested, all_day
        }
        with pytest.raises(ValidationError):
            AppointmentSummary(**data)


class TestAppointmentDetails:
    """Test AppointmentDetails model validation"""

    def test_valid_appointment_details(self):
        """Test AppointmentDetails accepts valid data with body"""
        data = {
            "entry_id": "apt-entry-id-123",
            "subject": "Team Meeting",
            "start": "2025-01-20 14:00:00",
            "end": "2025-01-20 15:00:00",
            "location": "Conference Room A",
            "organizer": "organizer@example.com",
            "all_day": False,
            "required_attendees": "attendee1@example.com; attendee2@example.com",
            "optional_attendees": "optional@example.com",
            "response_status": "Accepted",
            "meeting_status": "Meeting",
            "response_requested": True,
            "body": "Agenda: Q1 planning and review",
        }
        appointment = AppointmentDetails(**data)
        assert appointment.body == "Agenda: Q1 planning and review"
        assert appointment.subject == "Team Meeting"

    def test_appointment_details_default_body(self):
        """Test AppointmentDetails body defaults to empty string"""
        data = {
            "entry_id": "apt-entry-id-456",
            "subject": "Quick Sync",
            "start": "2025-01-20 10:00:00",
            "end": "2025-01-20 10:15:00",
            "location": "",
            "organizer": None,
            "all_day": False,
            "required_attendees": "",
            "optional_attendees": "",
            "response_status": "Organizer",
            "meeting_status": "NonMeeting",
            "response_requested": False,
        }
        appointment = AppointmentDetails(**data)
        assert appointment.body == ""


class TestCreateAppointmentResult:
    """Test CreateAppointmentResult model validation"""

    def test_successful_appointment_creation(self):
        """Test CreateAppointmentResult for successful appointment creation"""
        data = {
            "success": True,
            "entry_id": "new-apt-entry-id-123",
            "message": "Appointment created successfully",
        }
        result = CreateAppointmentResult(**data)
        assert result.success is True
        assert result.entry_id == "new-apt-entry-id-123"
        assert result.message == "Appointment created successfully"

    def test_failed_appointment_creation(self):
        """Test CreateAppointmentResult for failed appointment creation"""
        data = {
            "success": False,
            "entry_id": None,
            "message": "Failed to create appointment",
        }
        result = CreateAppointmentResult(**data)
        assert result.success is False
        assert result.entry_id is None
        assert result.message == "Failed to create appointment"

    def test_appointment_result_default_entry_id(self):
        """Test CreateAppointmentResult entry_id defaults to None"""
        data = {
            "success": True,
            "message": "Appointment created successfully",
        }
        result = CreateAppointmentResult(**data)
        assert result.success is True
        assert result.entry_id is None


class TestFreeBusyInfo:
    """Test FreeBusyInfo model validation"""

    def test_successful_free_busy_query(self):
        """Test FreeBusyInfo for successful query"""
        data = {
            "email": "user@example.com",
            "start_date": "2025-01-20",
            "end_date": "2025-01-21",
            "free_busy": "0000111222000",
            "resolved": True,
            "error": None,
        }
        info = FreeBusyInfo(**data)
        assert info.email == "user@example.com"
        assert info.start_date == "2025-01-20"
        assert info.end_date == "2025-01-21"
        assert info.free_busy == "0000111222000"
        assert info.resolved is True
        assert info.error is None

    def test_failed_free_busy_query(self):
        """Test FreeBusyInfo for failed query (unresolved email)"""
        data = {
            "email": "invalid@example.com",
            "start_date": "2025-01-20",
            "end_date": "2025-01-21",
            "free_busy": None,
            "resolved": False,
            "error": "Could not resolve email address",
        }
        info = FreeBusyInfo(**data)
        assert info.email == "invalid@example.com"
        assert info.free_busy is None
        assert info.resolved is False
        assert info.error == "Could not resolve email address"

    def test_free_busy_error_without_dates(self):
        """Test FreeBusyInfo error case without start/end dates"""
        data = {
            "email": "unknown@example.com",
            "start_date": None,
            "end_date": None,
            "free_busy": None,
            "resolved": False,
            "error": "Connection error",
        }
        info = FreeBusyInfo(**data)
        assert info.start_date is None
        assert info.end_date is None
        assert info.resolved is False

    def test_free_busy_defaults(self):
        """Test FreeBusyInfo optional fields default to None"""
        data = {
            "email": "user@example.com",
            "resolved": True,
        }
        info = FreeBusyInfo(**data)
        assert info.start_date is None
        assert info.end_date is None
        assert info.free_busy is None
        assert info.error is None


class TestCalendarModelSerialization:
    """Test calendar model serialization and deserialization"""

    def test_appointment_summary_serialization(self):
        """Test AppointmentSummary can be serialized to dict and JSON"""
        data = {
            "entry_id": "apt-entry-id-123",
            "subject": "Team Meeting",
            "start": "2025-01-20 14:00:00",
            "end": "2025-01-20 15:00:00",
            "location": "Conference Room A",
            "organizer": "organizer@example.com",
            "all_day": False,
            "required_attendees": "attendee1@example.com; attendee2@example.com",
            "optional_attendees": "optional@example.com",
            "response_status": "Accepted",
            "meeting_status": "Meeting",
            "response_requested": True,
        }
        appointment = AppointmentSummary(**data)
        # Test model_dump
        dumped = appointment.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = appointment.model_dump_json()
        assert "Team Meeting" in json_str
        assert "Conference Room A" in json_str

    def test_appointment_details_serialization(self):
        """Test AppointmentDetails can be serialized to dict and JSON"""
        data = {
            "entry_id": "apt-entry-id-123",
            "subject": "Team Meeting",
            "start": "2025-01-20 14:00:00",
            "end": "2025-01-20 15:00:00",
            "location": "Conference Room A",
            "organizer": "organizer@example.com",
            "all_day": False,
            "required_attendees": "attendee1@example.com; attendee2@example.com",
            "optional_attendees": "optional@example.com",
            "response_status": "Accepted",
            "meeting_status": "Meeting",
            "response_requested": True,
            "body": "Agenda: Q1 planning",
        }
        appointment = AppointmentDetails(**data)
        # Test model_dump
        dumped = appointment.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = appointment.model_dump_json()
        assert "Agenda: Q1 planning" in json_str

    def test_free_busy_info_serialization(self):
        """Test FreeBusyInfo can be serialized to dict and JSON"""
        data = {
            "email": "user@example.com",
            "start_date": "2025-01-20",
            "end_date": "2025-01-21",
            "free_busy": "0000111222000",
            "resolved": True,
            "error": None,
        }
        info = FreeBusyInfo(**data)
        # Test model_dump
        dumped = info.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = info.model_dump_json()
        assert "user@example.com" in json_str


class TestTaskSummary:
    """Test TaskSummary model validation"""

    def test_valid_task_summary(self):
        """Test TaskSummary accepts valid data"""
        data = {
            "entry_id": "task-entry-id-123",
            "subject": "Review Q1 Report",
            "body": "Complete review of Q1 financial report",
            "due_date": "2025-01-25",
            "status": 0,
            "priority": 2,
            "complete": False,
            "percent_complete": 0.0,
        }
        task = TaskSummary(**data)
        assert task.entry_id == "task-entry-id-123"
        assert task.subject == "Review Q1 Report"
        assert task.body == "Complete review of Q1 financial report"
        assert task.due_date == "2025-01-25"
        assert task.status == 0
        assert task.priority == 2
        assert task.complete is False
        assert task.percent_complete == 0.0

    def test_task_summary_with_none_optional_fields(self):
        """Test TaskSummary accepts None for optional fields"""
        data = {
            "entry_id": "task-entry-id-456",
            "subject": "Quick Task",
            "body": "",
            "due_date": None,
            "status": None,
            "priority": None,
            "complete": False,
            "percent_complete": 0.0,
        }
        task = TaskSummary(**data)
        assert task.due_date is None
        assert task.status is None
        assert task.priority is None
        assert task.body == ""

    def test_task_summary_default_values(self):
        """Test TaskSummary default values for body"""
        data = {
            "entry_id": "task-entry-id-789",
            "subject": "Task with default body",
            "due_date": "2025-01-30",
            "status": 1,
            "priority": 1,
            "complete": False,
            "percent_complete": 50.0,
        }
        task = TaskSummary(**data)
        assert task.body == ""

    def test_task_summary_missing_required_fields(self):
        """Test TaskSummary raises ValidationError for missing required fields"""
        data = {
            "entry_id": "task-entry-id-999",
            # Missing: subject, complete, percent_complete
        }
        with pytest.raises(ValidationError):
            TaskSummary(**data)

    def test_task_summary_completed(self):
        """Test TaskSummary with completed task"""
        data = {
            "entry_id": "task-entry-id-111",
            "subject": "Completed Task",
            "body": "This task is done",
            "due_date": "2025-01-20",
            "status": 2,
            "priority": 1,
            "complete": True,
            "percent_complete": 100.0,
        }
        task = TaskSummary(**data)
        assert task.complete is True
        assert task.percent_complete == 100.0
        assert task.status == 2


class TestCreateTaskResult:
    """Test CreateTaskResult model validation"""

    def test_successful_task_creation(self):
        """Test CreateTaskResult for successful task creation"""
        data = {
            "success": True,
            "entry_id": "new-task-entry-id-123",
            "message": "Task created successfully",
        }
        result = CreateTaskResult(**data)
        assert result.success is True
        assert result.entry_id == "new-task-entry-id-123"
        assert result.message == "Task created successfully"

    def test_failed_task_creation(self):
        """Test CreateTaskResult for failed task creation"""
        data = {
            "success": False,
            "entry_id": None,
            "message": "Failed to create task",
        }
        result = CreateTaskResult(**data)
        assert result.success is False
        assert result.entry_id is None
        assert result.message == "Failed to create task"

    def test_task_result_default_entry_id(self):
        """Test CreateTaskResult entry_id defaults to None"""
        data = {
            "success": True,
            "message": "Task created successfully",
        }
        result = CreateTaskResult(**data)
        assert result.success is True
        assert result.entry_id is None


class TestOperationResult:
    """Test OperationResult model validation"""

    def test_successful_operation(self):
        """Test OperationResult for successful operation"""
        data = {
            "success": True,
            "message": "Operation completed successfully",
        }
        result = OperationResult(**data)
        assert result.success is True
        assert result.message == "Operation completed successfully"

    def test_failed_operation(self):
        """Test OperationResult for failed operation"""
        data = {
            "success": False,
            "message": "Operation failed",
        }
        result = OperationResult(**data)
        assert result.success is False
        assert result.message == "Operation failed"

    def test_operation_result_missing_required_fields(self):
        """Test OperationResult raises ValidationError for missing required fields"""
        data = {
            "success": True,
            # Missing: message
        }
        with pytest.raises(ValidationError):
            OperationResult(**data)


class TestTaskModelSerialization:
    """Test task model serialization and deserialization"""

    def test_task_summary_serialization(self):
        """Test TaskSummary can be serialized to dict and JSON"""
        data = {
            "entry_id": "task-entry-id-123",
            "subject": "Review Q1 Report",
            "body": "Complete review of Q1 financial report",
            "due_date": "2025-01-25",
            "status": 0,
            "priority": 2,
            "complete": False,
            "percent_complete": 0.0,
        }
        task = TaskSummary(**data)
        # Test model_dump
        dumped = task.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = task.model_dump_json()
        assert "Review Q1 Report" in json_str
        assert "Complete review of Q1 financial report" in json_str

    def test_create_task_result_serialization(self):
        """Test CreateTaskResult can be serialized to dict and JSON"""
        data = {
            "success": True,
            "entry_id": "new-task-entry-id-456",
            "message": "Task created successfully",
        }
        result = CreateTaskResult(**data)
        # Test model_dump
        dumped = result.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = result.model_dump_json()
        assert "new-task-entry-id-456" in json_str

    def test_operation_result_serialization(self):
        """Test OperationResult can be serialized to dict and JSON"""
        data = {
            "success": True,
            "message": "Task marked as complete",
        }
        result = OperationResult(**data)
        # Test model_dump
        dumped = result.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = result.model_dump_json()
        assert "Task marked as complete" in json_str
