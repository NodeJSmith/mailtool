"""Mailtool MCP Server

This module provides the main FastMCP server instance for Outlook automation.
It implements the Model Context Protocol (MCP) using the official MCP Python SDK v2
with the FastMCP framework.

The server provides 25 tools and 7 resources for Outlook email, calendar, and task management.
All tools return structured Pydantic models for type safety and LLM understanding.
"""

import argparse
import logging
from typing import TYPE_CHECKING

from mcp.server import FastMCP

from mailtool.mcp.com_state import ensure_com_initialized
from mailtool.mcp.exceptions import OutlookComError, OutlookNotFoundError
from mailtool.mcp.lifespan import outlook_lifespan
from mailtool.mcp.models import (
    AppointmentDetails,
    AppointmentSummary,
    AttachmentInfo,
    CreateAppointmentResult,
    CreateTaskResult,
    EmailDetails,
    EmailSummary,
    FreeBusyInfo,
    InboxStats,
    OperationResult,
    SendEmailResult,
    TaskSummary,
)
from mailtool.mcp.resources import (
    register_calendar_resources,
    register_email_resources,
    register_task_resources,
)

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

# Configure logging for the MCP server
# Logs are written to stderr for debugging and monitoring
logger = logging.getLogger(__name__)

# Global variable to store default account from CLI args
# This is set in main() before the server runs, and read by the lifespan
_default_account: str | None = None


# Create FastMCP server instance
# The lifespan parameter manages Outlook COM bridge lifecycle (creation, warmup, cleanup)
# Note: The lifespan reads from the global _default_account variable
mcp = FastMCP(
    name="mailtool-outlook-bridge",
    lifespan=outlook_lifespan,
)

# Register email resources (US-022), calendar resources (US-028), and task resources (US-033)
register_email_resources(mcp)
register_calendar_resources(mcp)
register_task_resources(mcp)

# Module-level bridge instance (set by lifespan, accessed by tools)
_bridge: "OutlookBridge | None" = None


def _get_bridge():
    """Get the current bridge instance

    Returns:
        OutlookBridge: The bridge instance

    Raises:
        OutlookComError: If bridge is not initialized (server not running)
    """
    global _bridge

    # Ensure COM is initialized for the current thread before accessing bridge
    ensure_com_initialized()

    if _bridge is None:
        logger.error("Outlook bridge not initialized. Is the server running?")
        raise OutlookComError("Outlook bridge not initialized. Is the server running?")
    logger.debug("Retrieved Outlook bridge instance")
    return _bridge


def _email_summary_from_dict(email: dict) -> EmailSummary:
    """Build an EmailSummary from a bridge result dict (defaults for missing keys)."""
    return EmailSummary(
        entry_id=email.get("entry_id", ""),
        subject=email.get("subject", ""),
        sender=email.get("sender", ""),
        sender_name=email.get("sender_name", ""),
        received_time=email.get("received_time"),
        unread=email.get("unread", False),
        has_attachments=email.get("has_attachments", False),
        message_class=email.get("message_class", "IPM.Note"),
        to=email.get("to", ""),
        cc=email.get("cc", ""),
        sent_time=email.get("sent_time"),
        conversation_id=email.get("conversation_id"),
        conversation_topic=email.get("conversation_topic"),
    )


def _email_details_from_dict(email: dict) -> EmailDetails:
    """Build an EmailDetails from a bridge result dict (defaults for missing keys)."""
    attachments = [
        AttachmentInfo(
            filename=a.get("filename", ""),
            size=a.get("size", 0),
            display_name=a.get("display_name", ""),
            content_type=a.get("content_type"),
            is_inline=a.get("is_inline", False),
        )
        for a in email.get("attachments", [])
    ]
    return EmailDetails(
        entry_id=email.get("entry_id", ""),
        subject=email.get("subject", ""),
        sender=email.get("sender", ""),
        sender_name=email.get("sender_name", ""),
        body=email.get("body", ""),
        html_body=email.get("html_body", ""),
        received_time=email.get("received_time"),
        has_attachments=email.get("has_attachments", False),
        message_class=email.get("message_class", "IPM.Note"),
        to=email.get("to", ""),
        cc=email.get("cc", ""),
        bcc=email.get("bcc", ""),
        sent_time=email.get("sent_time"),
        conversation_id=email.get("conversation_id"),
        conversation_topic=email.get("conversation_topic"),
        attachments=attachments,
        body_top=email.get("body_top", ""),
    )


# ============================================================================
# Email Tools (US-008: get_email, US-011: mark_email, US-013: delete_email, US-016: list_emails, US-017: send_email, US-018: reply_email, US-019: forward_email, US-020: move_email, US-021: search_emails)
# ============================================================================


@mcp.tool()
def list_emails(
    limit: int = 10, folder: str = "Inbox", include_non_mail: bool = False
) -> list[EmailSummary]:
    """
    List emails from the specified folder.

    Retrieves a list of email summaries from the specified folder, sorted by
    received time (most recent first). Uses O(1) direct access for each email.

    By default only real emails (MessageClass IPM.Note) are returned; meeting
    notifications, meeting cancellations, and other non-mail inbox items are
    excluded. Set include_non_mail=True to include them.

    Args:
        limit: Maximum number of emails to return (default: 10)
        folder: Folder name to list emails from (default: "Inbox")
        include_non_mail: If True, also return non-mail items (meeting
            notifications, post items, etc.) — default False

    Returns:
        list[EmailSummary]: List of email summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized or folder cannot be accessed
    """
    bridge = _get_bridge()
    result = bridge.list_emails(
        limit=limit, folder=folder, include_non_mail=include_non_mail
    )
    return [_email_summary_from_dict(email) for email in result]


@mcp.tool()
def list_unread_emails(limit: int = 10) -> list[EmailSummary]:
    """
    List unread emails from the Inbox.

    Retrieves the most recent unread emails from the Inbox, sorted by received
    time (most recent first). Uses Outlook Restrict filter for efficient querying
    (O(1) search at COM level).

    Only real emails (MessageClass IPM.Note) are returned; unread meeting
    notifications and other non-mail items are filtered out.

    Args:
        limit: Maximum number of unread emails to return (default: 10)

    Returns:
        list[EmailSummary]: List of unread email summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized

    Note:
        This function uses the Outlook Restrict filter with '[Unread] = TRUE'
        for efficient querying at the COM level, avoiding unnecessary iteration.
    """
    bridge = _get_bridge()
    result = bridge.search_emails(filter_query="[Unread] = TRUE", limit=limit)
    return [_email_summary_from_dict(email) for email in result]


@mcp.tool()
def get_email(entry_id: str) -> EmailDetails:
    """
    Get full email body and details by entry ID.

    Retrieves complete email information including body content (both plain text
    and HTML), recipients, sent time, conversation/thread identifiers, and
    attachment metadata, using O(1) direct access via EntryID.

    Non-mail items (e.g. meeting notifications that share an Inbox EntryID) are
    returned with their message_class set (e.g. "IPM.Schedule.Meeting.Request")
    rather than raising, so callers can branch. OutlookNotFoundError is raised
    only when no item at all matches the ID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)

    Returns:
        EmailDetails: Complete email details including body content

    Raises:
        OutlookNotFoundError: If no item matches the entry_id
        OutlookComError: If bridge is not initialized
    """
    bridge = _get_bridge()
    result = bridge.get_email_body(entry_id)
    if result is None:
        logger.error(f"Email not found: {entry_id}")
        raise OutlookNotFoundError("Email not found", entry_id=entry_id)
    logger.debug(f"Retrieved email: {entry_id}")
    return _email_details_from_dict(result)


@mcp.tool()
def get_emails(entry_ids: list[str], include_body: bool = True) -> list[EmailDetails]:
    """
    Fetch full details for many emails in a single call (bulk get_email).

    Avoids the N+1 round-trip of calling get_email once per item. Items that
    cannot be found are silently omitted from the result; the result order
    follows the input order for items that were found.

    Args:
        entry_ids: List of Outlook EntryIDs (O(1) direct access each)
        include_body: If False, return only summary fields (faster) — default True

    Returns:
        list[EmailDetails]: Details for each item that was found

    Raises:
        OutlookComError: If bridge is not initialized
    """
    bridge = _get_bridge()
    result = bridge.get_email_bodies(entry_ids, include_body=include_body)
    return [_email_details_from_dict(email) for email in result]


@mcp.tool()
def mark_email(entry_id: str, unread: bool = False) -> OperationResult:
    """
    Mark an email as read or unread.

    Changes the read/unread status of an email using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        unread: True to mark as unread, False to mark as read (default: False)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Mark email as read/unread via bridge
    result = bridge.mark_email_read(entry_id, unread=unread)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email marked as {'unread' if unread else 'read'}",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to mark email as {'unread' if unread else 'read'}",
        )


@mcp.tool()
def delete_email(entry_id: str) -> OperationResult:
    """
    Delete an email.

    Permanently deletes an email using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete email via bridge
    result = bridge.delete_email(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Email deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete email",
        )


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    bcc: str | None = None,
    html_body: str | None = None,
    file_paths: list[str] | None = None,
    save_draft: bool = False,
) -> SendEmailResult:
    """
    Send an email or save it as a draft.

    Creates and sends an email, or saves it to the Drafts folder.
    Supports attachments, CC/BCC recipients, and HTML body.

    Args:
        to: Primary recipient email address
        subject: Email subject line
        body: Plain text email body
        cc: CC recipients (optional)
        bcc: BCC recipients (optional)
        html_body: HTML email body (optional, overrides plain text body if provided)
        file_paths: List of file paths to attach (optional)
        save_draft: If True, save to Drafts instead of sending (default: False)

    Returns:
        SendEmailResult: Result with success status, draft entry ID (if saved), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Send email via bridge
    result = bridge.send_email(
        to=to,
        subject=subject,
        body=body,
        cc=cc,
        bcc=bcc,
        html_body=html_body,
        file_paths=file_paths,
        save_draft=save_draft,
    )

    # Convert bridge result to SendEmailResult
    # Bridge returns: False (failed), True (sent), str (draft EntryID)
    if result is False:
        return SendEmailResult(
            success=False,
            entry_id=None,
            message="Failed to send email",
        )
    elif result is True:
        return SendEmailResult(
            success=True,
            entry_id=None,
            message="Email sent successfully",
        )
    else:  # str - draft EntryID
        return SendEmailResult(
            success=True,
            entry_id=result,
            message=f"Email saved as draft: {result}",
        )


@mcp.tool()
def reply_email(entry_id: str, body: str, reply_all: bool = False) -> OperationResult:
    """
    Reply to an email.

    Replies to an email using O(1) direct access via EntryID.
    Can reply to sender only or reply to all recipients.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        body: Reply body text
        reply_all: True to reply to all recipients, False to reply to sender only (default: False)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Reply to email via bridge
    result = bridge.reply_email(entry_id, body=body, reply_all=reply_all)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email {'replied to all' if reply_all else 'replied'} successfully",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to {'reply to all' if reply_all else 'reply'}",
        )


@mcp.tool()
def forward_email(entry_id: str, to: str, body: str = "") -> OperationResult:
    """
    Forward an email.

    Forwards an email to a recipient using O(1) direct access via EntryID.
    Optionally adds additional body text.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        to: Recipient email address to forward to
        body: Optional additional body text (default: "")

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Forward email via bridge
    result = bridge.forward_email(entry_id, to=to, body=body)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Email forwarded successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to forward email",
        )


@mcp.tool()
def move_email(entry_id: str, folder: str) -> OperationResult:
    """
    Move an email to a different folder.

    Moves an email to a specified folder using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        folder: Target folder name (e.g., "Archive", "Drafts", "Sent Items")

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Move email via bridge
    result = bridge.move_email(entry_id, folder_name=folder)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email moved to {folder}",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to move email to {folder}",
        )


@mcp.tool()
def search_emails(
    filter_query: str, limit: int = 100, include_non_mail: bool = False
) -> list[EmailSummary]:
    """
    Search emails using Outlook filter query.

    Searches emails in the Inbox using Outlook Restriction filter (O(1) search).
    Supports SQL-like filter syntax for advanced queries.

    SQL-style LIKE patterns are translated automatically to a DASL (@SQL=)
    query, because Items.Restrict's default Jet syntax has no LIKE. Wildcard
    semantics are SQL LIKE (case-insensitive): '%x%' contains, 'x%' starts
    with, '%x' ends with. Invalid filters raise a tool error instead of
    silently returning [].

    By default the search is scoped to real emails (MessageClass IPM.Note); set
    include_non_mail=True to also match meeting items and other non-mail items.

    NOTE: For searching by sender email address, especially for internal/Exchange
    users, use search_emails_by_sender() instead. The SenderEmailAddress filter
    does not work for Exchange addresses (internal emails).

    Args:
        filter_query: SQL-like filter query string (examples below)
        limit: Maximum number of results to return (default: 100)
        include_non_mail: If True, do not scope to IPM.Note items — default False

    Returns:
        list[EmailSummary]: List of matching email summaries

    Raises:
        OutlookComError: If bridge is not initialized

    Examples:
        search_emails("[Subject] LIKE '%project%'")  # Search by subject
        search_emails("[Unread] = TRUE")  # Find unread emails
        search_emails("[SenderName] LIKE '%John%'")  # By sender name (better than email)
        search_emails("[ReceivedTime] >= '07/01/2026 00:00' AND "
                      "[ReceivedTime] <= '07/31/2026 23:59'")  # Date range
        search_emails("[HasAttachments] = TRUE")  # Only emails with attachments
    """
    bridge = _get_bridge()
    result = bridge.search_emails(
        filter_query=filter_query, limit=limit, include_non_mail=include_non_mail
    )
    return [_email_summary_from_dict(email) for email in result]


@mcp.tool()
def search_emails_by_sender(
    sender_email: str,
    limit: int = 100,
    folder: str = "Inbox",
    include_non_mail: bool = False,
) -> list[EmailSummary]:
    """
    Search emails by sender email address (handles Exchange addresses).

    This method properly handles both SMTP and Exchange email addresses.
    For Exchange users (internal emails), it resolves the Exchange address
    to SMTP address before matching.

    This is the recommended way to search for emails from a specific sender,
    especially for internal/Exchange email addresses where SenderEmailAddress
    filter would not work.

    Args:
        sender_email: Email address to search for (e.g., "f.muijzer@utwente.nl")
        limit: Maximum number of results to return (default: 100)
        folder: Folder name to search in (default: "Inbox")
        include_non_mail: If True, also consider non-mail items — default False

    Returns:
        list[EmailSummary]: List of matching email summaries

    Raises:
        OutlookComError: If bridge is not initialized

    Examples:
        search_emails_by_sender("john@example.com")  # Search by sender
        search_emails_by_sender("f.muijzer@utwente.nl", limit=50)  # Internal Exchange user
    """
    bridge = _get_bridge()
    result = bridge.search_by_sender(
        sender_email=sender_email,
        limit=limit,
        folder=folder,
        include_non_mail=include_non_mail,
    )
    return [_email_summary_from_dict(email) for email in result]


@mcp.tool()
def get_inbox_stats(folder: str = "Inbox") -> InboxStats:
    """
    Return cheap total/unread counts for a folder without fetching items.

    Uses Restrict+Count at the COM level so it is fast regardless of folder
    size. Useful for deciding whether list_unread_emails(limit=N) has more
    results to page through, or for inbox monitoring.

    Args:
        folder: Folder name (default: "Inbox")

    Returns:
        InboxStats: {folder, total, unread}

    Raises:
        OutlookComError: If bridge is not initialized
    """
    bridge = _get_bridge()
    result = bridge.get_inbox_stats(folder=folder)
    return InboxStats(
        folder=result.get("folder", folder),
        total=result.get("total", 0),
        unread=result.get("unread", 0),
    )


# ============================================================================
# Calendar Tools (US-009: get_appointment, US-014: delete_appointment, US-023: list_calendar_events, US-024: create_appointment, US-025: edit_appointment, US-026: respond_to_meeting, US-027: get_free_busy)
# ============================================================================


@mcp.tool()
def list_calendar_events(
    days: int = 7, all_events: bool = False
) -> list[AppointmentSummary]:
    """
    List calendar events for the next N days.

    Retrieves a list of calendar events/appointments from the Outlook calendar.
    Supports date filtering and includes recurring meetings.
    Uses O(1) direct access and COM-level filtering for performance.

    Args:
        days: Number of days ahead to look (default: 7)
        all_events: If True, scan from now through +365 days instead of `days`.
            Unbounded scans (the old behaviour) can expand recurring series
            forever and wedge Outlook's COM apartment, so the horizon is capped.

    Returns:
        list[AppointmentSummary]: List of appointment summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized

    Note:
        This function handles the "Calendar Bomb" issue by applying COM-level
        Restrict filters before Python iteration to prevent infinite recurring meetings.
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # List calendar events via bridge
    result = bridge.list_calendar_events(days=days, all_events=all_events)

    # Convert bridge result to list of AppointmentSummary models
    return [
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
        for event in result
    ]


@mcp.tool()
def get_appointment(entry_id: str) -> AppointmentDetails:
    """
    Get full appointment details and body by entry ID.

    Retrieves complete appointment information including body content and all
    meeting metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)

    Returns:
        AppointmentDetails: Complete appointment details including body content

    Raises:
        OutlookNotFoundError: If appointment not found
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get appointment details from bridge
    result = bridge.get_appointment(entry_id)

    # Check if appointment was found
    if result is None:
        logger.error(f"Appointment not found: {entry_id}")
        raise OutlookNotFoundError("Appointment not found", entry_id=entry_id)

    logger.debug(f"Retrieved appointment: {entry_id}")

    # Convert bridge result to AppointmentDetails model
    # Note: AppointmentDetails extends AppointmentSummary, adding 'body' field
    return AppointmentDetails(
        entry_id=result["entry_id"],
        subject=result["subject"],
        start=result["start"],
        end=result["end"],
        location=result["location"],
        organizer=result["organizer"],
        body=result["body"],
        all_day=result["all_day"],
        required_attendees=result["required_attendees"],
        optional_attendees=result["optional_attendees"],
        response_status=result["response_status"],
        meeting_status=result["meeting_status"],
        response_requested=result["response_requested"],
    )


@mcp.tool()
def delete_appointment(entry_id: str) -> OperationResult:
    """
    Delete an appointment.

    Permanently deletes an appointment using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete appointment via bridge
    result = bridge.delete_appointment(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Appointment deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete appointment",
        )


@mcp.tool()
def create_appointment(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    required_attendees: str | None = None,
    optional_attendees: str | None = None,
) -> CreateAppointmentResult:
    """
    Create a calendar appointment.

    Creates a new appointment or meeting in the Outlook calendar.
    Supports all-day events, location, body/description, and attendees.

    Args:
        subject: Appointment subject line
        start: Start timestamp in 'YYYY-MM-DD HH:MM:SS' format
        end: End timestamp in 'YYYY-MM-DD HH:MM:SS' format
        location: Appointment location (default: "")
        body: Appointment body/description text (default: "")
        all_day: True for all-day event, False for timed event (default: False)
        required_attendees: Semicolon-separated list of required attendees (optional)
        optional_attendees: Semicolon-separated list of optional attendees (optional)

    Returns:
        CreateAppointmentResult: Result with success status, appointment entry ID (if created), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Create appointment via bridge
    result = bridge.create_appointment(
        subject=subject,
        start=start,
        end=end,
        location=location,
        body=body,
        all_day=all_day,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees,
    )

    # Convert bridge result to CreateAppointmentResult
    # Bridge returns: str (EntryID) if successful, None if failed
    if result:
        return CreateAppointmentResult(
            success=True,
            entry_id=result,
            message=f"Appointment created successfully: {result}",
        )
    else:
        return CreateAppointmentResult(
            success=False,
            entry_id=None,
            message="Failed to create appointment",
        )


@mcp.tool()
def edit_appointment(
    entry_id: str,
    required_attendees: str | None = None,
    optional_attendees: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
) -> OperationResult:
    """
    Edit an existing appointment.

    Updates an existing appointment's fields using O(1) direct access via EntryID.
    Only updates fields that are provided (non-None parameters).

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)
        required_attendees: Semicolon-separated list of required attendees (optional)
        optional_attendees: Semicolon-separated list of optional attendees (optional)
        subject: New subject (optional)
        start: New start timestamp in 'YYYY-MM-DD HH:MM:SS' format (optional)
        end: New end timestamp in 'YYYY-MM-DD HH:MM:SS' format (optional)
        location: New location (optional)
        body: New body/description text (optional)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Edit appointment via bridge
    result = bridge.edit_appointment(
        entry_id=entry_id,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees,
        subject=subject,
        start=start,
        end=end,
        location=location,
        body=body,
    )

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Appointment updated successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to update appointment",
        )


@mcp.tool()
def respond_to_meeting(entry_id: str, response: str) -> OperationResult:
    """
    Respond to a meeting invitation.

    Responds to a meeting invitation request using O(1) direct access via EntryID.
    Supports accept, decline, and tentative responses.

    Args:
        entry_id: Outlook EntryID of the meeting invitation (O(1) direct access)
        response: Response type - "accept", "decline", or "tentative"

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Respond to meeting via bridge
    result = bridge.respond_to_meeting(entry_id, response)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Meeting {response.lower()}ed successfully",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to {response.lower()} meeting",
        )


@mcp.tool()
def get_free_busy(
    email_address: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> FreeBusyInfo:
    """
    Get free/busy status for an email address.

    Retrieves free/busy information for a specified email address using the
    Outlook FreeBusy method. Defaults to current user and today's date if not specified.

    Args:
        email_address: Email address to check (optional, defaults to current user)
        start_date: Start date in 'YYYY-MM-DD' format (optional, defaults to today)
        end_date: End date in 'YYYY-MM-DD' format (optional, defaults to start + 1 day)

    Returns:
        FreeBusyInfo: Free/busy information with email, dates, status string, and resolution status

    Raises:
        OutlookComError: If bridge is not initialized

    Note:
        Free/busy status codes: 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get free/busy info via bridge
    result = bridge.get_free_busy(
        email_address=email_address,
        start_date=start_date,
        end_date=end_date,
    )

    # Convert bridge result to FreeBusyInfo model
    # Bridge returns dict with different fields depending on success/error
    return FreeBusyInfo(
        email=result.get("email", email_address or "unknown"),
        start_date=result.get("start_date"),
        end_date=result.get("end_date"),
        free_busy=result.get("free_busy"),
        resolved=result.get("resolved", False),
        error=result.get("error"),
    )


# ============================================================================
# Task Tools (US-010: get_task, US-012: complete_task, US-015: delete_task, US-029: list_tasks, US-030: list_all_tasks, US-031: create_task, US-032: edit_task)
# ============================================================================


@mcp.tool()
def list_tasks(include_completed: bool = False) -> list[TaskSummary]:
    """
    List tasks from the Outlook Tasks folder.

    Retrieves a list of tasks from the Outlook Tasks folder.
    By default, returns only incomplete tasks. Optionally includes completed tasks.

    Args:
        include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

    Returns:
        list[TaskSummary]: List of task summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # List tasks via bridge
    result = bridge.list_tasks(include_completed=include_completed)

    # Convert bridge result to list of TaskSummary models
    return [
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
        for task in result
    ]


@mcp.tool()
def list_all_tasks() -> list[TaskSummary]:
    """
    List all tasks from the Outlook Tasks folder (including completed).

    Retrieves a complete list of all tasks from the Outlook Tasks folder,
    including both incomplete and completed tasks. This is a convenience
    function that calls list_tasks with include_completed=True.

    Returns:
        list[TaskSummary]: List of all task summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # List all tasks via bridge (hardcoded include_completed=True)
    result = bridge.list_tasks(include_completed=True)

    # Convert bridge result to list of TaskSummary models
    return [
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
        for task in result
    ]


@mcp.tool()
def get_task(entry_id: str) -> TaskSummary:
    """
    Get full task details and body by entry ID.

    Retrieves complete task information including body content and all
    task metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        TaskSummary: Complete task details including body content

    Raises:
        OutlookNotFoundError: If task not found
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get task details from bridge
    result = bridge.get_task(entry_id)

    # Check if task was found
    if result is None:
        logger.error(f"Task not found: {entry_id}")
        raise OutlookNotFoundError("Task not found", entry_id=entry_id)

    logger.debug(f"Retrieved task: {entry_id}")

    # Convert bridge result to TaskSummary model
    # Note: TaskSummary includes all fields from bridge.get_task()
    return TaskSummary(
        entry_id=result["entry_id"],
        subject=result["subject"],
        body=result["body"],
        due_date=result["due_date"],
        status=result["status"],
        priority=result["priority"],
        complete=result["complete"],
        percent_complete=result["percent_complete"],
    )


@mcp.tool()
def complete_task(entry_id: str) -> OperationResult:
    """
    Mark a task as complete.

    Marks a task as complete with 100% percent complete status using O(1)
    direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Mark task as complete via bridge
    result = bridge.complete_task(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Task marked as complete",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to mark task as complete",
        )


@mcp.tool()
def delete_task(entry_id: str) -> OperationResult:
    """
    Delete a task.

    Permanently deletes a task using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete task via bridge
    result = bridge.delete_task(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Task deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete task",
        )


@mcp.tool()
def create_task(
    subject: str,
    body: str = "",
    due_date: str | None = None,
    priority: int = 1,
) -> CreateTaskResult:
    """
    Create a new task.

    Creates a new task in the Outlook Tasks folder.
    Supports task description, due date, and priority level.

    Args:
        subject: Task subject/title
        body: Task description or body text (default: "")
        due_date: Due date in 'YYYY-MM-DD' format (optional)
        priority: Task priority - 0=Low, 1=Normal (default), 2=High

    Returns:
        CreateTaskResult: Result with success status, task entry ID (if created), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Create task via bridge
    # Note: bridge parameter is 'importance', not 'priority'
    result = bridge.create_task(
        subject=subject,
        body=body,
        due_date=due_date,
        importance=priority,
    )

    # Convert bridge result to CreateTaskResult
    # Bridge returns: str (EntryID) if successful, None if failed
    if result:
        return CreateTaskResult(
            success=True,
            entry_id=result,
            message=f"Task created successfully: {result}",
        )
    else:
        return CreateTaskResult(
            success=False,
            entry_id=None,
            message="Failed to create task",
        )


@mcp.tool()
def edit_task(
    entry_id: str,
    subject: str | None = None,
    body: str | None = None,
    due_date: str | None = None,
    priority: int | None = None,
    percent_complete: float | None = None,
    complete: bool | None = None,
) -> OperationResult:
    """
    Edit an existing task.

    Updates an existing task in the Outlook Tasks folder.
    Only updates fields that are provided (non-None parameters).
    Supports updating subject, description, due date, priority, completion status, and percent complete.

    Args:
        entry_id: Task entry ID to edit
        subject: New task subject/title (optional)
        body: New task description or body text (optional)
        due_date: New due date in 'YYYY-MM-DD' format (optional)
        priority: New task priority - 0=Low, 1=Normal, 2=High (optional)
        percent_complete: New percent complete value 0-100 (optional)
        complete: Mark task as complete or incomplete (optional)

    Returns:
        OperationResult: Result with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Edit task via bridge
    # Note: bridge parameter is 'importance', not 'priority'
    result = bridge.edit_task(
        entry_id=entry_id,
        subject=subject,
        body=body,
        due_date=due_date,
        importance=priority,
        percent_complete=percent_complete,
        complete=complete,
    )

    # Convert bridge result to OperationResult
    # Bridge returns: True if successful, False if failed
    if result:
        return OperationResult(success=True, message="Task edited successfully")
    else:
        return OperationResult(success=False, message="Failed to edit task")


def main(default_account: str | None = None):
    """Entry point for the MCP server.

    Args:
        default_account: Optional default account name/email (bypasses argparse if provided)
    """
    global _default_account

    # If default_account is provided directly, use it
    # Otherwise, parse CLI arguments
    if default_account is None:
        # Parse CLI arguments for default account
        parser = argparse.ArgumentParser(
            description="Mailtool MCP Server - Outlook automation via MCP",
            epilog=(
                "Examples:\n"
                "  uv run --with pywin32 -m mailtool.mcp.server\n"
                "  uv run --with pywin32 -m mailtool.mcp.server --account 'john@example.com'\n"
                "  uv run --with pywin32 -m mailtool.mcp.server --acc 'john@example.com'\n"
            ),
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        parser.add_argument(
            "--account",
            "--acc",
            dest="account",
            help="Default account name or email address for Outlook operations",
        )

        args = parser.parse_args()
        default_account = args.account

    # Set the global default account that the lifespan will read
    _default_account = default_account

    # Run the MCP server with stdio transport
    # Uses the global 'mcp' instance which has all tools registered
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
