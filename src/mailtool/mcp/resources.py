"""MCP Resources

This module provides MCP resource implementations for read-only data access.
Resources are registered with the FastMCP server and return formatted text/JSON content.

Email resources:
- inbox://emails - List recent emails from inbox
- inbox://unread - List unread emails from inbox
- email://{entry_id} - Get full email details by entry ID (template resource)

Calendar resources:
- calendar://today - List today's calendar events
- calendar://week - List calendar events for the next 7 days

Task resources:
- tasks://active - List active (incomplete) tasks
- tasks://all - List all tasks (including completed)
"""

import logging
from typing import TYPE_CHECKING

from mcp.server import FastMCP

from mailtool.mcp.com_state import ensure_com_initialized
from mailtool.mcp.exceptions import OutlookComError
from mailtool.mcp.models import (
    AppointmentDetails,
    AppointmentSummary,
    EmailDetails,
    EmailSummary,
    TaskSummary,
)

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

# Configure logging
logger = logging.getLogger(__name__)

# Module-level bridge instance (set by server.py)
_bridge: "OutlookBridge | None" = None


def _get_bridge() -> "OutlookBridge":
    """Get the current bridge instance.

    Returns:
        OutlookBridge: The bridge instance

    Raises:
        OutlookComError: If bridge is not initialized
    """
    global _bridge

    # Ensure COM is initialized for the current thread before accessing bridge
    ensure_com_initialized()

    if _bridge is None:
        raise OutlookComError("Outlook bridge not initialized")
    return _bridge


def _set_bridge(bridge: "OutlookBridge") -> None:
    """Set the bridge instance (called by server.py lifespan).

    Args:
        bridge: The bridge instance
    """
    global _bridge
    _bridge = bridge


def _format_email_summary(email: EmailSummary) -> str:
    """Format an email summary as readable text.

    Args:
        email: EmailSummary model

    Returns:
        Formatted text representation
    """
    return f"""Subject: {email.subject}
From: {email.sender_name} <{email.sender}>
To: {email.to or "(none)"}
CC: {email.cc or "(none)"}
Received: {email.received_time}
Sent: {email.sent_time}
Unread: {"Yes" if email.unread else "No"}
Attachments: {"Yes" if email.has_attachments else "No"}
MessageClass: {email.message_class}
Conversation: {email.conversation_topic or "(none)"} [{email.conversation_id or "-"}]
Entry ID: {email.entry_id}
"""


def _format_email_details(email: EmailDetails) -> str:
    """Format full email details as readable text.

    Args:
        email: EmailDetails model

    Returns:
        Formatted text representation with body
    """
    attachment_lines = (
        "\n".join(
            f"  - {a.filename} ({a.size} bytes{', inline' if a.is_inline else ''})"
            for a in email.attachments
        )
        if email.attachments
        else "  (none)"
    )
    return f"""Subject: {email.subject}
From: {email.sender_name} <{email.sender}>
To: {email.to or "(none)"}
CC: {email.cc or "(none)"}
Received: {email.received_time}
Sent: {email.sent_time}
Attachments: {"Yes" if email.has_attachments else "No"}
MessageClass: {email.message_class}
Conversation: {email.conversation_topic or "(none)"} [{email.conversation_id or "-"}]
Entry ID: {email.entry_id}

Attachments:
{attachment_lines}

Body (top message):
{email.body_top}

Body (full):
{email.body}
"""


def _email_summary_to_dict(email: EmailSummary) -> dict:
    """Convert EmailSummary to dict for JSON serialization.

    Args:
        email: EmailSummary model

    Returns:
        Dictionary representation
    """
    return {
        "entry_id": email.entry_id,
        "subject": email.subject,
        "sender": email.sender,
        "sender_name": email.sender_name,
        "received_time": email.received_time,
        "unread": email.unread,
        "has_attachments": email.has_attachments,
        "message_class": email.message_class,
        "to": email.to,
        "cc": email.cc,
        "sent_time": email.sent_time,
        "conversation_id": email.conversation_id,
        "conversation_topic": email.conversation_topic,
    }


def _email_details_to_dict(email: EmailDetails) -> dict:
    """Convert EmailDetails to dict for JSON serialization.

    Args:
        email: EmailDetails model

    Returns:
        Dictionary representation
    """
    return {
        "entry_id": email.entry_id,
        "subject": email.subject,
        "sender": email.sender,
        "sender_name": email.sender_name,
        "body": email.body,
        "html_body": email.html_body,
        "body_top": email.body_top,
        "received_time": email.received_time,
        "sent_time": email.sent_time,
        "has_attachments": email.has_attachments,
        "message_class": email.message_class,
        "to": email.to,
        "cc": email.cc,
        "bcc": email.bcc,
        "conversation_id": email.conversation_id,
        "conversation_topic": email.conversation_topic,
        "attachments": [a.model_dump() for a in email.attachments],
    }


def register_email_resources(mcp: FastMCP) -> None:
    """Register all email resources with the FastMCP server.

    Args:
        mcp: FastMCP server instance

    This function registers three email resources:
    1. inbox://emails - Lists recent emails (limit 50)
    2. inbox://unread - Lists unread emails (limit 50)
    3. email://{entry_id} - Gets full email details (template resource)
    """

    @mcp.resource(
        uri="inbox://emails",
        name="inbox_emails",
        title="Inbox Emails",
        description="List recent emails from the inbox (max 50)",
    )
    def inbox_emails() -> str:
        """Get recent emails from inbox.

        Returns:
            Formatted text with email summaries
        """
        bridge = _get_bridge()

        # Get emails from bridge
        emails_data = bridge.list_emails(limit=50, folder="Inbox")

        # Convert to EmailSummary models
        emails = [
            EmailSummary(
                entry_id=email["entry_id"],
                subject=email["subject"],
                sender=email["sender"],
                sender_name=email["sender_name"],
                received_time=email["received_time"],
                unread=email["unread"],
                has_attachments=email["has_attachments"],
            )
            for email in emails_data
        ]

        # Format as text
        if not emails:
            return "No emails found in inbox"

        lines = [f"Inbox Emails ({len(emails)} items)", ""]
        for email in emails:
            lines.append(_format_email_summary(email))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="inbox://unread",
        name="inbox_unread",
        title="Unread Inbox Emails",
        description="List unread emails from the inbox (max 50)",
    )
    def inbox_unread() -> str:
        """Get unread emails from inbox.

        Returns:
            Formatted text with unread email summaries
        """
        bridge = _get_bridge()

        # Get all emails from bridge
        emails_data = bridge.list_emails(limit=50, folder="Inbox")

        # Filter to only unread emails
        unread_emails_data = [
            email for email in emails_data if email.get("unread", False)
        ]

        # Convert to EmailSummary models
        emails = [
            EmailSummary(
                entry_id=email["entry_id"],
                subject=email["subject"],
                sender=email["sender"],
                sender_name=email["sender_name"],
                received_time=email["received_time"],
                unread=email["unread"],
                has_attachments=email["has_attachments"],
            )
            for email in unread_emails_data
        ]

        # Format as text
        if not emails:
            return "No unread emails in inbox"

        lines = [f"Unread Emails ({len(emails)} items)", ""]
        for email in emails:
            lines.append(_format_email_summary(email))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="email://{entry_id}",
        name="email_details",
        title="Email Details",
        description="Get full email details by entry ID",
    )
    def email_details(entry_id: str) -> str:
        """Get full email details by entry ID.

        Args:
            entry_id: Outlook EntryID of the email

        Returns:
            Formatted text with full email details including body

        Raises:
            OutlookComError: If bridge is not initialized
        """
        bridge = _get_bridge()

        # Get email details from bridge
        email_data = bridge.get_email_body(entry_id)

        if email_data is None:
            return f"Email not found: {entry_id}"

        # Convert to EmailDetails model
        email = EmailDetails(
            entry_id=email_data["entry_id"],
            subject=email_data["subject"],
            sender=email_data["sender"],
            sender_name=email_data["sender_name"],
            body=email_data["body"],
            html_body=email_data["html_body"],
            received_time=email_data["received_time"],
            has_attachments=email_data["has_attachments"],
        )

        # Format as text
        return _format_email_details(email)


def _format_appointment_summary(appt: AppointmentSummary) -> str:
    """Format an appointment summary as readable text.

    Args:
        appt: AppointmentSummary model

    Returns:
        Formatted text representation
    """
    return f"""Subject: {appt.subject}
Start: {appt.start}
End: {appt.end}
Location: {appt.location or "No location"}
Organizer: {appt.organizer or "Unknown"}
All Day: {"Yes" if appt.all_day else "No"}
Required Attendees: {appt.required_attendees or "None"}
Optional Attendees: {appt.optional_attendees or "None"}
Response Status: {appt.response_status or "N/A"}
Meeting Status: {appt.meeting_status or "N/A"}
Entry ID: {appt.entry_id}
"""


def _format_appointment_details(appt: AppointmentDetails) -> str:
    """Format full appointment details as readable text.

    Args:
        appt: AppointmentDetails model

    Returns:
        Formatted text representation with body
    """
    return f"""Subject: {appt.subject}
Start: {appt.start}
End: {appt.end}
Location: {appt.location or "No location"}
Organizer: {appt.organizer or "Unknown"}
All Day: {"Yes" if appt.all_day else "No"}
Required Attendees: {appt.required_attendees or "None"}
Optional Attendees: {appt.optional_attendees or "None"}
Response Status: {appt.response_status or "N/A"}
Meeting Status: {appt.meeting_status or "N/A"}
Entry ID: {appt.entry_id}

Body:
{appt.body or "No body text"}
"""


def register_calendar_resources(mcp: FastMCP) -> None:
    """Register all calendar resources with the FastMCP server.

    Args:
        mcp: FastMCP server instance

    This function registers two calendar resources:
    1. calendar://today - Lists today's calendar events
    2. calendar://week - Lists calendar events for the next 7 days
    """

    @mcp.resource(
        uri="calendar://today",
        name="calendar_today",
        title="Today's Calendar",
        description="List calendar events for today",
    )
    def calendar_today() -> str:
        """Get today's calendar events.

        Returns:
            Formatted text with appointment summaries
        """
        bridge = _get_bridge()

        # Get today's events from bridge (days=1)
        events_data = bridge.list_calendar_events(days=1)

        # Convert to AppointmentSummary models
        events = [
            AppointmentSummary(
                entry_id=event["entry_id"],
                subject=event["subject"],
                start=event["start"],
                end=event["end"],
                location=event["location"],
                organizer=event["organizer"],
                all_day=event["all_day"],
                required_attendees=event["required_attendees"],
                optional_attendees=event["optional_attendees"],
                response_status=event["response_status"],
                meeting_status=event["meeting_status"],
                response_requested=event["response_requested"],
            )
            for event in events_data
        ]

        # Format as text
        if not events:
            return "No calendar events for today"

        lines = [f"Today's Calendar ({len(events)} events)", ""]
        for event in events:
            lines.append(_format_appointment_summary(event))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="calendar://week",
        name="calendar_week",
        title="Week's Calendar",
        description="List calendar events for the next 7 days",
    )
    def calendar_week() -> str:
        """Get calendar events for the next 7 days.

        Returns:
            Formatted text with appointment summaries
        """
        bridge = _get_bridge()

        # Get this week's events from bridge (days=7)
        events_data = bridge.list_calendar_events(days=7)

        # Convert to AppointmentSummary models
        events = [
            AppointmentSummary(
                entry_id=event["entry_id"],
                subject=event["subject"],
                start=event["start"],
                end=event["end"],
                location=event["location"],
                organizer=event["organizer"],
                all_day=event["all_day"],
                required_attendees=event["required_attendees"],
                optional_attendees=event["optional_attendees"],
                response_status=event["response_status"],
                meeting_status=event["meeting_status"],
                response_requested=event["response_requested"],
            )
            for event in events_data
        ]

        # Format as text
        if not events:
            return "No calendar events for the next 7 days"

        lines = [f"Week's Calendar ({len(events)} events)", ""]
        for event in events:
            lines.append(_format_appointment_summary(event))
            lines.append("-" * 60)

        return "\n".join(lines)


def _format_task_summary(task: TaskSummary) -> str:
    """Format a task summary as readable text.

    Args:
        task: TaskSummary model

    Returns:
        Formatted text representation
    """
    # Map status codes to human-readable names
    status_map = {
        0: "Not Started",
        1: "In Progress",
        2: "Complete",
        3: "Waiting",
        4: "Deferred",
        5: "Other",
    }
    status_str = (
        status_map.get(task.status, "Unknown") if task.status is not None else "N/A"
    )

    # Map priority codes to human-readable names
    priority_map = {0: "Low", 1: "Normal", 2: "High"}
    priority_str = (
        priority_map.get(task.priority, "Unknown")
        if task.priority is not None
        else "N/A"
    )

    return f"""Subject: {task.subject}
Due Date: {task.due_date or "No due date"}
Status: {status_str}
Priority: {priority_str}
Complete: {"Yes" if task.complete else "No"}
Percent Complete: {task.percent_complete:.1f}%
Entry ID: {task.entry_id}
"""


def register_task_resources(mcp: FastMCP) -> None:
    """Register all task resources with the FastMCP server.

    Args:
        mcp: FastMCP server instance

    This function registers two task resources:
    1. tasks://active - Lists active (incomplete) tasks
    2. tasks://all - Lists all tasks (including completed)
    """

    @mcp.resource(
        uri="tasks://active",
        name="tasks_active",
        title="Active Tasks",
        description="List active (incomplete) tasks",
    )
    def tasks_active() -> str:
        """Get active (incomplete) tasks.

        Returns:
            Formatted text with task summaries
        """
        bridge = _get_bridge()

        # Get active tasks from bridge (include_completed=False)
        tasks_data = bridge.list_tasks(include_completed=False)

        # Convert to TaskSummary models
        tasks = [
            TaskSummary(
                entry_id=task["entry_id"],
                subject=task["subject"],
                body=task["body"],
                due_date=task["due_date"],
                status=task["status"],
                priority=task["priority"],
                complete=task["complete"],
                percent_complete=task["percent_complete"],
            )
            for task in tasks_data
        ]

        # Format as text
        if not tasks:
            return "No active tasks"

        lines = [f"Active Tasks ({len(tasks)} items)", ""]
        for task in tasks:
            lines.append(_format_task_summary(task))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="tasks://all",
        name="tasks_all",
        title="All Tasks",
        description="List all tasks (including completed)",
    )
    def tasks_all() -> str:
        """Get all tasks (including completed).

        Returns:
            Formatted text with task summaries
        """
        bridge = _get_bridge()

        # Get all tasks from bridge (include_completed=True)
        tasks_data = bridge.list_tasks(include_completed=True)

        # Convert to TaskSummary models
        tasks = [
            TaskSummary(
                entry_id=task["entry_id"],
                subject=task["subject"],
                body=task["body"],
                due_date=task["due_date"],
                status=task["status"],
                priority=task["priority"],
                complete=task["complete"],
                percent_complete=task["percent_complete"],
            )
            for task in tasks_data
        ]

        # Format as text
        if not tasks:
            return "No tasks found"

        lines = [f"All Tasks ({len(tasks)} items)", ""]
        for task in tasks:
            lines.append(_format_task_summary(task))
            lines.append("-" * 60)

        return "\n".join(lines)
