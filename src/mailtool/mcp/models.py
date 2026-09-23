"""
Pydantic Models for MCP Tools

This module contains all Pydantic models for structured output from MCP tools.
Models provide type safety, automatic schema generation, and descriptive field metadata.
"""

from pydantic import BaseModel, Field

# ============================================================================
# Email Models (US-004)
# ============================================================================


class EmailSummary(BaseModel):
    """Summary representation of an email for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    unread: bool = Field(description="Whether the email is unread")
    has_attachments: bool = Field(description="Whether the email has attachments")
    message_class: str = Field(
        default="IPM.Note",
        description="Outlook MessageClass (IPM.Note for mail; IPM.Schedule.Meeting.* for meeting items)",
    )
    to: str = Field(default="", description="Semicolon-separated To recipients")
    cc: str = Field(default="", description="Semicolon-separated CC recipients")
    sent_time: str | None = Field(
        default=None,
        description="Sent timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    conversation_id: str | None = Field(
        default=None,
        description="Outlook ConversationID for thread grouping",
    )
    conversation_topic: str | None = Field(
        default=None,
        description="Conversation topic (normalized subject) for thread grouping",
    )


class AttachmentInfo(BaseModel):
    """Metadata for a single email attachment (bytes not included)."""

    filename: str = Field(default="", description="Attachment filename")
    size: int = Field(default=0, description="Attachment size in bytes")
    display_name: str = Field(default="", description="Outlook display name")
    content_type: str | None = Field(
        default=None, description="MIME content type if known"
    )
    is_inline: bool = Field(
        default=False,
        description="Whether the attachment is inline (e.g. embedded image)",
    )


class EmailDetails(BaseModel):
    """Full email details including body content."""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    body: str = Field(description="Plain text email body")
    html_body: str = Field(description="HTML email body")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    has_attachments: bool = Field(description="Whether the email has attachments")
    message_class: str = Field(
        default="IPM.Note",
        description="Outlook MessageClass (IPM.Note for mail; IPM.Schedule.Meeting.* for meeting items)",
    )
    to: str = Field(default="", description="Semicolon-separated To recipients")
    cc: str = Field(default="", description="Semicolon-separated CC recipients")
    bcc: str = Field(
        default="", description="Semicolon-separated BCC recipients (sent items only)"
    )
    sent_time: str | None = Field(
        default=None,
        description="Sent timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    conversation_id: str | None = Field(
        default=None,
        description="Outlook ConversationID for thread grouping",
    )
    conversation_topic: str | None = Field(
        default=None,
        description="Conversation topic (normalized subject) for thread grouping",
    )
    attachments: list[AttachmentInfo] = Field(
        default_factory=list,
        description="Metadata for each attachment (filenames, sizes, types)",
    )
    body_top: str = Field(
        default="",
        description="Cleaned top message: the new content only, with quoted reply chains and signatures stripped",
    )


class SendEmailResult(BaseModel):
    """Result of sending an email or saving a draft"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of saved draft (None if sent or failed)"
    )
    message: str = Field(description="Human-readable result message")


# ============================================================================
# Calendar Models (US-005)
# ============================================================================


class AppointmentSummary(BaseModel):
    """Summary representation of a calendar event for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Appointment subject line")
    start: str | None = Field(
        default=None,
        description="Start timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    end: str | None = Field(
        default=None,
        description="End timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    location: str = Field(description="Appointment location", default="")
    organizer: str | None = Field(
        default=None, description="Email address of the meeting organizer"
    )
    all_day: bool = Field(description="Whether this is an all-day event")
    required_attendees: str = Field(
        default="", description="Semicolon-separated list of required attendees"
    )
    optional_attendees: str = Field(
        default="", description="Semicolon-separated list of optional attendees"
    )
    response_status: str = Field(
        description="Meeting response status: None, Organizer, Tentative, Accepted, Declined, NotResponded, Unknown"
    )
    meeting_status: str = Field(
        description="Meeting status: NonMeeting, Meeting, Received, Canceled, Unknown"
    )
    response_requested: bool = Field(description="Whether response was requested")


class AppointmentDetails(AppointmentSummary):
    """Full appointment details including body content"""

    body: str = Field(default="", description="Appointment body/description text")


class CreateAppointmentResult(BaseModel):
    """Result of creating an appointment"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of created appointment (None if failed)"
    )
    message: str = Field(description="Human-readable result message")


class FreeBusyInfo(BaseModel):
    """Free/busy information for an email address"""

    email: str = Field(description="Email address that was checked")
    start_date: str | None = Field(
        default=None, description="Start date in 'YYYY-MM-DD' format or None if error"
    )
    end_date: str | None = Field(
        default=None, description="End date in 'YYYY-MM-DD' format or None if error"
    )
    free_busy: str | None = Field(
        default=None,
        description="Free/busy string with time slot status codes (0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=Working Elsewhere) or None if error",
    )
    resolved: bool = Field(
        description="Whether the email address was successfully resolved"
    )
    error: str | None = Field(
        default=None, description="Error message if the operation failed"
    )


# ============================================================================
# Task Models (US-006)
# ============================================================================


class TaskSummary(BaseModel):
    """Summary representation of a task for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Task subject line")
    body: str = Field(default="", description="Task description/body text")
    due_date: str | None = Field(
        default=None,
        description="Due date in 'YYYY-MM-DD' format or None if not set",
    )
    status: int | None = Field(
        default=None,
        description="Outlook task status code (0=NotStarted, 1=InProgress, 2=Complete, 3=Waiting, 4=Deferred, 5=Other)",
    )
    priority: int | None = Field(
        default=None,
        description="Task priority/importance (0=Low, 1=Normal, 2=High)",
    )
    complete: bool = Field(description="Whether the task is marked complete")
    percent_complete: float = Field(
        description="Task completion percentage (0.0 to 100.0)"
    )


class CreateTaskResult(BaseModel):
    """Result of creating a task"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of created task (None if failed)"
    )
    message: str = Field(description="Human-readable result message")


# ============================================================================
# Common Result Models (US-007)
# ============================================================================


class OperationResult(BaseModel):
    """Generic result for boolean operations (mark, delete, move, complete, etc.)"""

    success: bool = Field(description="Whether the operation succeeded")
    message: str = Field(description="Human-readable result message")


# ============================================================================
# Mailbox Statistics Models
# ============================================================================


class InboxStats(BaseModel):
    """Folder statistics for cheap pagination/monitoring without fetching items."""

    folder: str = Field(description="Folder name the stats were collected from")
    total: int = Field(default=0, description="Total item count in the folder")
    unread: int = Field(default=0, description="Unread item count in the folder")
