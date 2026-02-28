#!/usr/bin/env python3
"""
Poll Gmail for new emails with PDF attachments, extract serial numbers,
and upload to Google Sheets.

Designed to run periodically via Windows Task Scheduler.

Usage:
    python gmail_poll.py                    # Process new unread PDFs
    python gmail_poll.py --label "LoadOut"  # Use custom Gmail label
    python gmail_poll.py --dry-run          # Check without processing
"""

import os
import sys
import base64
import json
import argparse
import logging
from datetime import datetime
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Add parent directory so we can import sibling modules
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(SCRIPT_DIR))

from extract_serials import process_pdf
from upload_to_sheet import upload_to_sheet, get_credentials as get_sheets_credentials

# Gmail + Sheets scopes
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets",
]

DEFAULT_SHEET_ID = "1wK92FpXq-07LdYYPCwZi7-C2vruLPs59JM14w4nAggs"
PDF_DIR = PROJECT_DIR / ".tmp" / "pdfs"
LOG_FILE = PROJECT_DIR / "output" / "gmail_poll.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


def get_credentials(creds_path=None, token_path=None):
    """Get OAuth credentials with Gmail + Sheets scopes."""
    if not creds_path:
        # Auto-detect credentials file
        for f in os.listdir(PROJECT_DIR):
            if f.startswith("client_secret") and f.endswith(".json"):
                creds_path = PROJECT_DIR / f
                break
        if not creds_path:
            creds_path = PROJECT_DIR / "credentials.json"

    if not token_path:
        token_path = PROJECT_DIR / "token_gmail.json"

    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                log.error(f"Credentials not found: {creds_path}")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, "w") as f:
            f.write(creds.to_json())
        log.info(f"Token saved to: {token_path}")

    return creds


def find_unread_pdf_emails(gmail, label_name="PDFs"):
    """Find unread emails with PDF attachments in the given label."""
    query = f"is:unread has:attachment filename:pdf"
    if label_name:
        query = f"label:{label_name} " + query

    results = gmail.users().messages().list(
        userId="me", q=query, maxResults=20
    ).execute()

    messages = results.get("messages", [])
    return messages


def download_pdf_attachments(gmail, message_id):
    """Download PDF attachments from a Gmail message. Returns list of file paths."""
    msg = gmail.users().messages().get(
        userId="me", id=message_id, format="full"
    ).execute()

    subject = ""
    for header in msg.get("payload", {}).get("headers", []):
        if header["name"].lower() == "subject":
            subject = header["value"]
            break

    pdf_files = []
    parts = msg.get("payload", {}).get("parts", [])

    for part in parts:
        filename = part.get("filename", "")
        if not filename.lower().endswith(".pdf"):
            continue

        attachment_id = part.get("body", {}).get("attachmentId")
        if not attachment_id:
            continue

        attachment = gmail.users().messages().attachments().get(
            userId="me", messageId=message_id, id=attachment_id
        ).execute()

        data = base64.urlsafe_b64decode(attachment["data"])

        # Save to temp directory
        PDF_DIR.mkdir(parents=True, exist_ok=True)
        safe_filename = "".join(
            c if c.isalnum() or c in ".-_ " else "_" for c in filename
        )
        filepath = PDF_DIR / safe_filename

        with open(filepath, "wb") as f:
            f.write(data)

        pdf_files.append(str(filepath))
        log.info(f"  Downloaded: {filename} ({len(data)} bytes)")

    return pdf_files, subject


def mark_as_read(gmail, message_id):
    """Remove UNREAD label from a message."""
    gmail.users().messages().modify(
        userId="me",
        id=message_id,
        body={"removeLabelIds": ["UNREAD"]},
    ).execute()


def main():
    parser = argparse.ArgumentParser(
        description="Poll Gmail for PDFs, extract serials, upload to Sheets"
    )
    parser.add_argument(
        "--label", default="LoadOut",
        help="Gmail label to filter on (default: LoadOut). Use --no-label to skip.",
    )
    parser.add_argument(
        "--no-label", action="store_true",
        help="Process ALL unread PDF emails (no label filter)",
    )
    parser.add_argument(
        "--sheet-id", default=DEFAULT_SHEET_ID,
        help="Google Sheet ID",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Check for emails without processing",
    )
    args = parser.parse_args()

    # Ensure output/log directory exists
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

    log.info("=" * 50)
    log.info(f"Gmail poll started at {datetime.now().isoformat(timespec='seconds')}")

    # Authenticate
    creds = get_credentials()
    gmail = build("gmail", "v1", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)

    # Create label if it doesn't exist
    label_name = "" if args.no_label else args.label
    if label_name:
        existing_labels = gmail.users().labels().list(userId="me").execute()
        label_exists = any(
            lbl["name"].lower() == label_name.lower()
            for lbl in existing_labels.get("labels", [])
        )
        if not label_exists:
            gmail.users().labels().create(
                userId="me",
                body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
            ).execute()
            log.info(f"Created Gmail label: {label_name}")

    # Find unread emails with PDFs
    messages = find_unread_pdf_emails(gmail, label_name)

    if not messages:
        log.info("No new PDF emails found.")
        return

    log.info(f"Found {len(messages)} unread email(s) with PDFs")

    if args.dry_run:
        for msg in messages:
            msg_detail = gmail.users().messages().get(
                userId="me", id=msg["id"], format="metadata",
                metadataHeaders=["Subject", "From"],
            ).execute()
            headers = {
                h["name"]: h["value"]
                for h in msg_detail.get("payload", {}).get("headers", [])
            }
            log.info(f"  From: {headers.get('From', '?')} | Subject: {headers.get('Subject', '?')}")
        log.info("Dry run - no processing done.")
        return

    # Process each email
    total_serials = 0
    all_records = []

    for msg in messages:
        msg_id = msg["id"]
        log.info(f"Processing email {msg_id}...")

        # Download PDF attachments
        pdf_files, subject = download_pdf_attachments(gmail, msg_id)

        if not pdf_files:
            log.warning(f"  No PDF attachments found in email: {subject}")
            mark_as_read(gmail, msg_id)
            continue

        # Extract serials from each PDF
        for pdf_path in pdf_files:
            log.info(f"  Extracting from: {os.path.basename(pdf_path)}")
            records = process_pdf(pdf_path)
            log.info(f"  Found {len(records)} serial number(s)")
            all_records.extend(records)
            total_serials += len(records)

        # Mark email as read
        mark_as_read(gmail, msg_id)
        log.info(f"  Email marked as read: {subject}")

    # Upload all records to Google Sheet
    if all_records:
        log.info(f"Uploading {total_serials} serial numbers to Google Sheet...")
        upload_to_sheet(all_records, args.sheet_id, creds, append=True)
        log.info("Upload complete!")
    else:
        log.info("No serial numbers found in any PDFs.")

    log.info(f"Done. Processed {len(messages)} email(s), {total_serials} serial(s).")


if __name__ == "__main__":
    main()
