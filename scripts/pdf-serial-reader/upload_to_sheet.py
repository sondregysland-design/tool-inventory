#!/usr/bin/env python3
"""
Upload extracted serial numbers to Google Sheets using OAuth.

First run: opens browser for Google login, saves token.json.
Subsequent runs: reuses token automatically.
"""

import os
import sys
import json
import argparse

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

DEFAULT_SHEET_ID = "1wK92FpXq-07LdYYPCwZi7-C2vruLPs59JM14w4nAggs"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = [
    "Filename",
    "Customer",
    "Well Name",
    "Load Out Date",
    "Material #",
    "Description",
    "Serial / Trace #",
    "Qty",
    "Extracted At",
]

FIELD_MAP = [
    "filename",
    "customer",
    "well_name",
    "load_out_date",
    "material_num",
    "description",
    "serial_trace",
    "qty",
    "extracted_at",
]


def get_credentials(creds_path, token_path):
    """Load or create OAuth credentials."""
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                print(f"Error: credentials not found at: {creds_path}", file=sys.stderr)
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, "w") as f:
            f.write(creds.to_json())
        print(f"Token saved to: {token_path}")

    return creds


def _get_existing_keys(sheets, sheet_id):
    """Read existing rows and return set of (filename, material#, serial#) keys."""
    result = sheets.values().get(
        spreadsheetId=sheet_id, range="Out!A:G"
    ).execute()
    existing = set()
    for row in result.get("values", [])[1:]:  # skip header
        if len(row) >= 7:
            key = (row[0], row[4], row[6])  # Filename, Material #, Serial / Trace #
            existing.add(key)
    return existing


def upload_to_sheet(records, sheet_id, creds, append=False):
    """Upload records to Google Sheet."""
    service = build("sheets", "v4", credentials=creds)
    sheets = service.spreadsheets()

    rows = [HEADERS]
    for record in records:
        row = [str(record.get(field, "")) for field in FIELD_MAP]
        rows.append(row)

    if append:
        # Deduplicate: skip records already in the sheet
        existing_keys = _get_existing_keys(sheets, sheet_id)
        new_rows = []
        for row in rows[1:]:
            key = (row[0], row[4], row[6])  # Filename, Material #, Serial / Trace #
            if key not in existing_keys:
                new_rows.append(row)

        if not new_rows:
            print(f"All {len(rows) - 1} rows already exist in sheet - nothing to append.")
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}"

        skipped = len(rows) - 1 - len(new_rows)
        if skipped:
            print(f"Skipped {skipped} duplicate(s)")

        sheets.values().append(
            spreadsheetId=sheet_id,
            range="Out!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": new_rows},
        ).execute()
        print(f"Appended {len(new_rows)} new row(s)")
    else:
        sheets.values().clear(spreadsheetId=sheet_id, range="Out!A:Z").execute()
        sheets.values().update(
            spreadsheetId=sheet_id,
            range="Out!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        print(f"Uploaded {len(rows) - 1} data rows (plus header)")

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"Sheet URL: {url}")
    return url


def main():
    parser = argparse.ArgumentParser(
        description="Upload serial data to Google Sheets"
    )
    parser.add_argument(
        "--input", required=True, help="Path to JSON file with extracted serials"
    )
    parser.add_argument(
        "--sheet-id", default=DEFAULT_SHEET_ID, help="Google Sheet ID"
    )
    parser.add_argument(
        "--credentials", default=None, help="Path to OAuth credentials JSON"
    )
    parser.add_argument(
        "--token", default="token.json", help="Path to token file"
    )
    parser.add_argument(
        "--append", action="store_true", help="Append to existing data"
    )

    args = parser.parse_args()

    # Auto-detect credentials file if not specified
    creds_path = args.credentials
    if not creds_path:
        for f in os.listdir("."):
            if f.startswith("client_secret") and f.endswith(".json"):
                creds_path = f
                break
        if not creds_path:
            creds_path = "credentials.json"

    with open(args.input, "r", encoding="utf-8") as f:
        records = json.load(f)

    if not records:
        print("No records to upload.")
        sys.exit(0)

    print(f"Uploading {len(records)} records to Google Sheet...")
    creds = get_credentials(creds_path, args.token)
    upload_to_sheet(records, args.sheet_id, creds, args.append)


if __name__ == "__main__":
    main()
