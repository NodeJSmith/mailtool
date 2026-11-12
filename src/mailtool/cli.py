"""
Mailtool CLI Entry Point
Provides platform-specific error handling for Windows-only COM automation.

The bridge module imports without pywin32, so this module checks the platform and
pywin32 up front and gives a helpful error instead of a COM failure.
"""

import argparse
import importlib.util
import json
import sys
import traceback
from typing import NoReturn

from mailtool.bridge import MailtoolError, OutlookBridge


def _check_platform() -> None:
    """
    Verify that the tool is running on Windows.

    Raises:
        SystemExit: With error code 1 and helpful message if not on Windows.
    """
    if sys.platform != "win32":
        print(
            "Error: mailtool requires Windows with Microsoft Outlook installed.\n",
            file=sys.stderr,
        )
        print(
            "This tool uses COM automation to communicate with Outlook "
            "and is only supported on Windows.\n",
            file=sys.stderr,
        )
        print("For WSL2/Linux users:", file=sys.stderr)
        print(
            "  - Install mailtool on the Windows side and call it as mailtool.exe\n",
            file=sys.stderr,
        )
        print("For direct Windows access:", file=sys.stderr)
        print("  - Run from Windows PowerShell or Command Prompt", file=sys.stderr)
        print(
            "  - Or use: uv run --with mailtool --no-project mailtool <command>",
            file=sys.stderr,
        )
        fail("mailtool requires Windows with Microsoft Outlook installed")


def _check_pywin32() -> None:
    """
    Verify that pywin32 is available on Windows.

    Raises:
        SystemExit: With error code 1 and helpful message if pywin32 is missing.
    """
    try:
        if importlib.util.find_spec("win32com.client") is None:
            raise ImportError()
    except (ImportError, ValueError):
        print(
            "Error: pywin32 is required but not installed.\n",
            file=sys.stderr,
        )
        print(
            "This package provides COM bindings for Outlook automation.\n",
            file=sys.stderr,
        )
        print("To fix:", file=sys.stderr)
        print(
            "  uv run --with pywin32 mailtool <command>",
            file=sys.stderr,
        )
        print("\nOr install pywin32 in your environment:", file=sys.stderr)
        print("  uv add pywin32", file=sys.stderr)
        fail("pywin32 is required but not installed")


def parse_bool(value: str) -> bool:
    """argparse type for true/false flags (type=bool would make "false" True)."""
    lowered = value.strip().lower()
    if lowered in ("true", "yes", "1"):
        return True
    if lowered in ("false", "no", "0"):
        return False
    raise argparse.ArgumentTypeError(f"expected true or false, got {value!r}")


def main() -> None:
    """
    Main CLI entry point for mailtool.

    Checks the platform and pywin32, then dispatches commands to the
    OutlookBridge class.

    All commands return JSON output for machine readability.
    Exit code 1 indicates an error.
    """
    _check_platform()
    _check_pywin32()

    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Outlook COM Bridge - Email and Calendar Automation",
        epilog=(
            "Examples:\n"
            "  mailtool emails --limit 10\n"
            "  mailtool calendar --days 7\n"
            "  mailtool send --to user@example.com --subject 'Hello' --body 'World'"
        ),
    )
    subparsers = parser.add_subparsers(
        dest="command", help="Command to run", required=False
    )

    # Emails command
    email_parser = subparsers.add_parser("emails", help="List emails")
    email_parser.add_argument(
        "--limit", type=int, default=10, help="Max emails to return"
    )
    email_parser.add_argument(
        "--folder", default="Inbox", help="Folder name (default: Inbox)"
    )

    # Calendar command
    cal_parser = subparsers.add_parser("calendar", help="List calendar events")
    cal_parser.add_argument("--days", type=int, default=7, help="Days ahead to look")
    cal_parser.add_argument(
        "--all", action="store_true", help="Show all events without date filtering"
    )

    # Get email command
    get_parser = subparsers.add_parser("email", help="Get email details")
    get_parser.add_argument("--id", required=True, help="Email entry ID")

    # Send email command
    send_parser = subparsers.add_parser("send", help="Send an email")
    send_parser.add_argument("--to", required=True, help="Recipient email address")
    send_parser.add_argument("--subject", required=True, help="Email subject")
    send_parser.add_argument("--body", required=True, help="Email body")
    send_parser.add_argument("--cc", help="CC recipients")
    send_parser.add_argument("--bcc", help="BCC recipients")
    send_parser.add_argument("--html", help="HTML body (rich text)")
    send_parser.add_argument("--attach", nargs="+", help="File paths to attach")
    send_parser.add_argument(
        "--draft", action="store_true", help="Save as draft instead of sending"
    )

    # Download attachments command
    attach_parser = subparsers.add_parser(
        "attachments", help="Download email attachments"
    )
    attach_parser.add_argument("--id", required=True, help="Email entry ID")
    attach_parser.add_argument(
        "--dir", required=True, help="Directory to save attachments"
    )

    # Reply email command
    reply_parser = subparsers.add_parser("reply", help="Reply to an email")
    reply_parser.add_argument("--id", required=True, help="Email entry ID")
    reply_parser.add_argument("--body", required=True, help="Reply body")
    reply_parser.add_argument(
        "--all", action="store_true", help="Reply all instead of just sender"
    )

    # Forward email command
    forward_parser = subparsers.add_parser("forward", help="Forward an email")
    forward_parser.add_argument("--id", required=True, help="Email entry ID")
    forward_parser.add_argument("--to", required=True, help="Recipient to forward to")
    forward_parser.add_argument("--body", default="", help="Additional body text")

    # Search emails command
    search_parser = subparsers.add_parser(
        "search", help="Search emails using Restriction"
    )
    search_parser.add_argument(
        "--query",
        required=True,
        help="SQL filter query (e.g., urn:schemas:httpmail:subject LIKE '%%keyword%%')",
    )
    search_parser.add_argument(
        "--limit", type=int, default=100, help="Max results to return"
    )

    # Mark email command
    mark_parser = subparsers.add_parser("mark", help="Mark email as read/unread")
    mark_parser.add_argument("--id", required=True, help="Email entry ID")
    mark_parser.add_argument(
        "--unread", action="store_true", help="Mark as unread (default: read)"
    )

    # Move email command
    move_parser = subparsers.add_parser("move", help="Move email to folder")
    move_parser.add_argument("--id", required=True, help="Email entry ID")
    move_parser.add_argument("--folder", required=True, help="Target folder name")

    # Delete email command
    del_email_parser = subparsers.add_parser("delete-email", help="Delete an email")
    del_email_parser.add_argument("--id", required=True, help="Email entry ID")

    # Create appointment command
    create_appt_parser = subparsers.add_parser(
        "create-appt", help="Create calendar appointment"
    )
    create_appt_parser.add_argument(
        "--subject", required=True, help="Appointment subject"
    )
    create_appt_parser.add_argument(
        "--start", required=True, help="Start time (YYYY-MM-DD HH:MM:SS)"
    )
    create_appt_parser.add_argument(
        "--end", required=True, help="End time (YYYY-MM-DD HH:MM:SS)"
    )
    create_appt_parser.add_argument("--location", default="", help="Location")
    create_appt_parser.add_argument(
        "--body", default="", help="Appointment description"
    )
    create_appt_parser.add_argument(
        "--all-day", action="store_true", help="All-day event"
    )
    create_appt_parser.add_argument(
        "--required", help="Required attendees (semicolon-separated)"
    )
    create_appt_parser.add_argument(
        "--optional", help="Optional attendees (semicolon-separated)"
    )

    # Get appointment command
    get_appt_parser = subparsers.add_parser(
        "appointment", help="Get appointment details"
    )
    get_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")

    # Delete appointment command
    del_appt_parser = subparsers.add_parser("delete-appt", help="Delete an appointment")
    del_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")

    # Respond to meeting command
    respond_parser = subparsers.add_parser(
        "respond", help="Respond to meeting invitation"
    )
    respond_parser.add_argument("--id", required=True, help="Appointment entry ID")
    respond_parser.add_argument(
        "--response",
        required=True,
        choices=["accept", "decline", "tentative"],
        help="Meeting response",
    )

    # Free/busy command
    freebusy_parser = subparsers.add_parser("freebusy", help="Get free/busy status")
    freebusy_parser.add_argument(
        "--email", help="Email address to check (defaults to current user)"
    )
    freebusy_parser.add_argument(
        "--start", help="Start date (YYYY-MM-DD, defaults to today)"
    )
    freebusy_parser.add_argument(
        "--end", help="End date (YYYY-MM-DD, defaults to tomorrow)"
    )
    freebusy_parser.add_argument(
        "--id", help="DEPRECATED: Appointment entry ID (use --email instead)"
    )

    # Edit appointment command
    edit_appt_parser = subparsers.add_parser("edit-appt", help="Edit an appointment")
    edit_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")
    edit_appt_parser.add_argument(
        "--required", help="Required attendees (comma-separated)"
    )
    edit_appt_parser.add_argument(
        "--optional", help="Optional attendees (comma-separated)"
    )
    edit_appt_parser.add_argument("--subject", help="New subject")
    edit_appt_parser.add_argument(
        "--start", help="New start time (YYYY-MM-DD HH:MM:SS)"
    )
    edit_appt_parser.add_argument("--end", help="New end time (YYYY-MM-DD HH:MM:SS)")
    edit_appt_parser.add_argument("--location", help="New location")
    edit_appt_parser.add_argument("--body", help="New body/description")

    # Tasks command
    subparsers.add_parser("tasks", help="List all tasks")

    # List folders command (new)
    folders_parser = subparsers.add_parser(
        "folders", help="List folders for an account or all accounts"
    )
    folders_parser.add_argument(
        "--account", help="Account/display name or email to filter folders for"
    )

    # Set default account command (new)
    setacc_parser = subparsers.add_parser(
        "set-account", help="Set the default account/store to use"
    )
    setacc_parser.add_argument("--name", required=True, help="Account name or email")

    # Get task command
    get_task_parser = subparsers.add_parser("task", help="Get task details")
    get_task_parser.add_argument("--id", required=True, help="Task entry ID")

    # Create task command
    create_task_parser = subparsers.add_parser("create-task", help="Create a new task")
    create_task_parser.add_argument("--subject", required=True, help="Task subject")
    create_task_parser.add_argument("--body", default="", help="Task description")
    create_task_parser.add_argument("--due", help="Due date (YYYY-MM-DD)")
    create_task_parser.add_argument(
        "--priority",
        type=int,
        choices=[0, 1, 2],
        default=1,
        help="Priority: 0=Low, 1=Normal, 2=High",
    )

    # Edit task command
    edit_task_parser = subparsers.add_parser("edit-task", help="Edit a task")
    edit_task_parser.add_argument("--id", required=True, help="Task entry ID")
    edit_task_parser.add_argument("--subject", help="New subject")
    edit_task_parser.add_argument("--body", help="New description")
    edit_task_parser.add_argument("--due", help="New due date (YYYY-MM-DD)")
    edit_task_parser.add_argument(
        "--priority", type=int, choices=[0, 1, 2], help="New priority"
    )
    edit_task_parser.add_argument(
        "--percent", type=int, choices=range(0, 101), help="Percent complete (0-100)"
    )
    edit_task_parser.add_argument(
        "--complete", type=parse_bool, help="Mark complete/incomplete (true/false)"
    )

    # Complete task command
    complete_task_parser = subparsers.add_parser(
        "complete-task", help="Mark task as complete"
    )
    complete_task_parser.add_argument("--id", required=True, help="Task entry ID")

    # Delete task command
    del_task_parser = subparsers.add_parser("delete-task", help="Delete a task")
    del_task_parser.add_argument("--id", required=True, help="Task entry ID")

    args = parser.parse_args()

    if not args.command:
        parser.print_help(sys.stderr)
        fail("No command given")

    try:
        bridge = OutlookBridge()
        run_command(bridge, args)
    except MailtoolError as e:
        fail(str(e))
    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        fail(f"Unexpected error: {type(e).__name__}: {e}")


def fail(message: str) -> NoReturn:
    """Print the error JSON on stdout (where consumers parse) and exit 1."""
    # main() relies on SystemExit not being caught by its `except Exception`.
    print(json.dumps({"status": "error", "message": message}))
    sys.exit(1)


def run_command(bridge: OutlookBridge, args: argparse.Namespace) -> None:
    """
    Run the parsed command against the bridge and print its JSON result.

    Bridge failures raise MailtoolError; main() turns that into the error
    JSON. A fetch that finds nothing (None) is reported here as not found.
    """
    if args.command == "emails":
        emails = bridge.list_emails(limit=args.limit, folder=args.folder)
        print(json.dumps(emails, indent=2))

    elif args.command == "calendar":
        events = bridge.list_calendar_events(days=args.days, all_events=args.all)
        print(json.dumps(events, indent=2))

    elif args.command == "email":
        email = bridge.get_email_body(entry_id=args.id)
        if email is None:
            fail("Email not found")
        print(json.dumps(email, indent=2))

    elif args.command == "send":
        result = bridge.send_email(
            args.to,
            args.subject,
            args.body,
            args.cc,
            args.bcc,
            html_body=args.html,
            file_paths=args.attach,
            save_draft=args.draft,
        )
        if args.draft:
            print(
                json.dumps(
                    {
                        "status": "success",
                        "entry_id": result,
                        "message": "Draft saved",
                    }
                )
            )
        else:
            print(json.dumps({"status": "success", "message": "Email sent"}))

    elif args.command == "attachments":
        downloaded = bridge.download_attachments(args.id, args.dir)
        print(json.dumps({"status": "success", "attachments": downloaded}))

    elif args.command == "reply":
        bridge.reply_email(args.id, args.body, reply_all=args.all)
        print(json.dumps({"status": "success", "message": "Reply sent"}))

    elif args.command == "forward":
        bridge.forward_email(args.id, args.to, args.body)
        print(json.dumps({"status": "success", "message": "Email forwarded"}))

    elif args.command == "search":
        emails = bridge.search_emails(args.query, limit=args.limit)
        print(json.dumps(emails, indent=2))

    elif args.command == "folders":
        folders = bridge.list_folders(getattr(args, "account", None))
        print(json.dumps(folders, indent=2))

    elif args.command == "set-account":
        bridge.set_default_account(args.name)
        print(json.dumps({"status": "success", "account": args.name}))

    elif args.command == "mark":
        bridge.mark_email_read(args.id, unread=args.unread)
        status = "unread" if args.unread else "read"
        print(json.dumps({"status": "success", "message": f"Email marked as {status}"}))

    elif args.command == "move":
        bridge.move_email(args.id, args.folder)
        print(
            json.dumps(
                {"status": "success", "message": f"Email moved to {args.folder}"}
            )
        )

    elif args.command == "delete-email":
        bridge.delete_email(args.id)
        print(json.dumps({"status": "success", "message": "Email deleted"}))

    elif args.command == "create-appt":
        entry_id = bridge.create_appointment(
            args.subject,
            args.start,
            args.end,
            args.location,
            args.body,
            args.all_day,
            args.required,
            args.optional,
        )
        print(
            json.dumps(
                {
                    "status": "success",
                    "entry_id": entry_id,
                    "message": "Appointment created",
                }
            )
        )

    elif args.command == "appointment":
        appointment = bridge.get_appointment(args.id)
        if appointment is None:
            fail("Appointment not found")
        print(json.dumps(appointment, indent=2))

    elif args.command == "delete-appt":
        bridge.delete_appointment(args.id)
        print(json.dumps({"status": "success", "message": "Appointment deleted"}))

    elif args.command == "edit-appt":
        bridge.edit_appointment(
            args.id,
            required_attendees=args.required,
            optional_attendees=args.optional,
            subject=args.subject,
            start=args.start,
            end=args.end,
            location=args.location,
            body=args.body,
        )
        print(json.dumps({"status": "success", "message": "Appointment updated"}))

    elif args.command == "respond":
        bridge.respond_to_meeting(args.id, args.response)
        outcome = {
            "accept": "accepted",
            "decline": "declined",
            "tentative": "tentatively accepted",
        }[args.response]
        print(json.dumps({"status": "success", "message": f"Meeting {outcome}"}))

    elif args.command == "freebusy":
        freebusy = bridge.get_free_busy(
            email_address=getattr(args, "email", None),
            start_date=getattr(args, "start", None),
            end_date=getattr(args, "end", None),
            entry_id=getattr(args, "id", None),
        )
        if not freebusy.get("resolved"):
            fail(f"Could not resolve email address: {freebusy.get('email')}")
        print(json.dumps(freebusy, indent=2))

    elif args.command == "tasks":
        tasks = bridge.list_tasks()
        print(json.dumps(tasks, indent=2))

    elif args.command == "task":
        task = bridge.get_task(args.id)
        if task is None:
            fail("Task not found")
        print(json.dumps(task, indent=2))

    elif args.command == "create-task":
        entry_id = bridge.create_task(args.subject, args.body, args.due, args.priority)
        print(
            json.dumps(
                {
                    "status": "success",
                    "entry_id": entry_id,
                    "message": "Task created",
                }
            )
        )

    elif args.command == "edit-task":
        bridge.edit_task(
            args.id,
            args.subject,
            args.body,
            args.due,
            args.priority,
            args.percent,
            args.complete,
        )
        print(json.dumps({"status": "success", "message": "Task updated"}))

    elif args.command == "complete-task":
        bridge.complete_task(args.id)
        print(json.dumps({"status": "success", "message": "Task marked as complete"}))

    elif args.command == "delete-task":
        bridge.delete_task(args.id)
        print(json.dumps({"status": "success", "message": "Task deleted"}))


if __name__ == "__main__":
    main()
