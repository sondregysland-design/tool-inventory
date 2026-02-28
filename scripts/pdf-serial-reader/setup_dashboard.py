#!/usr/bin/env python3
"""
Set up Inventory + Dashboard sheets in the Google Sheet.

- Uses existing 'Out' sheet as data source
- Creates Inventory sheet (user fills in Total Stock & Redress)
- Creates Dashboard sheet with formulas that auto-calculate Load Out & Ready
"""

import os
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

DEFAULT_SHEET_ID = "1wK92FpXq-07LdYYPCwZi7-C2vruLPs59JM14w4nAggs"
DATA_SHEET_NAME = "Out"  # Existing sheet with load out data
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_credentials(token_path="token_gmail.json"):
    """Load OAuth credentials."""
    if not os.path.exists(token_path):
        token_path = "token.json"
    if not os.path.exists(token_path):
        print("Error: no token file found. Run upload_to_sheet.py first.", file=sys.stderr)
        sys.exit(1)

    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(token_path, "w") as f:
                f.write(creds.to_json())
        else:
            print("Error: token expired and cannot refresh.", file=sys.stderr)
            sys.exit(1)
    return creds


def get_sheet_id_by_title(spreadsheet, title):
    """Find a sheet's ID by its title, or return None."""
    for sheet in spreadsheet.get("sheets", []):
        if sheet["properties"]["title"] == title:
            return sheet["properties"]["sheetId"]
    return None


def setup_dashboard(sheet_id=DEFAULT_SHEET_ID):
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheets = service.spreadsheets()

    # Get current spreadsheet info
    spreadsheet = sheets.get(spreadsheetId=sheet_id).execute()
    existing_titles = [s["properties"]["title"] for s in spreadsheet["sheets"]]
    print(f"Existing sheets: {existing_titles}")

    if DATA_SHEET_NAME not in existing_titles:
        print(f"Error: '{DATA_SHEET_NAME}' sheet not found!", file=sys.stderr)
        sys.exit(1)

    requests = []

    # --- Step 1: Create Inventory sheet if it doesn't exist ---
    if "Inventory" not in existing_titles:
        requests.append({
            "addSheet": {
                "properties": {
                    "title": "Inventory",
                    "index": 1,
                }
            }
        })
        print("  Creating 'Inventory' sheet")

    # --- Step 2: Create Dashboard sheet if it doesn't exist ---
    if "Dashboard" not in existing_titles:
        requests.append({
            "addSheet": {
                "properties": {
                    "title": "Dashboard",
                    "index": 0,  # First tab = Dashboard
                }
            }
        })
        print("  Creating 'Dashboard' sheet")

    if requests:
        sheets.batchUpdate(spreadsheetId=sheet_id, body={"requests": requests}).execute()
        print("  Sheets created.")
        spreadsheet = sheets.get(spreadsheetId=sheet_id).execute()

    # --- Step 3: Get unique descriptions from Out sheet ---
    result = sheets.values().get(
        spreadsheetId=sheet_id, range=f"'{DATA_SHEET_NAME}'!F:F"
    ).execute()
    all_descriptions = []
    for row in result.get("values", [])[1:]:  # skip header
        if row and row[0]:
            all_descriptions.append(row[0])

    unique_descriptions = sorted(set(all_descriptions))
    print(f"  Found {len(unique_descriptions)} unique tool descriptions")

    # --- Step 4: Populate Inventory sheet ---
    inv_result = sheets.values().get(
        spreadsheetId=sheet_id, range="Inventory!A:C"
    ).execute()
    existing_inv = inv_result.get("values", [])

    if len(existing_inv) <= 1:
        inv_rows = [["Tool", "Total Stock", "Redress"]]
        for desc in unique_descriptions:
            inv_rows.append([desc, 0, 0])

        sheets.values().update(
            spreadsheetId=sheet_id,
            range="Inventory!A1",
            valueInputOption="RAW",
            body={"values": inv_rows},
        ).execute()
        print(f"  Populated Inventory with {len(unique_descriptions)} tools")
    else:
        existing_tools = {row[0] for row in existing_inv[1:] if row}
        new_tools = [d for d in unique_descriptions if d not in existing_tools]
        if new_tools:
            new_rows = [[t, 0, 0] for t in new_tools]
            sheets.values().append(
                spreadsheetId=sheet_id,
                range="Inventory!A1",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": new_rows},
            ).execute()
            print(f"  Added {len(new_tools)} new tools to Inventory")
        else:
            print("  Inventory already up to date")

    # --- Step 5: Build Dashboard with formulas ---
    inv_result = sheets.values().get(
        spreadsheetId=sheet_id, range="Inventory!A:A"
    ).execute()
    inv_rows_count = len(inv_result.get("values", [])) - 1  # minus header

    dashboard_rows = [["Tool", "Redress", "Ready", "Load Out", "Total Stock"]]
    for i in range(2, inv_rows_count + 2):  # row 2, 3, ... in Inventory
        dash_row = len(dashboard_rows) + 1  # current dashboard row number
        # Use semicolon as separator for no_NO locale
        row = [
            f"=Inventory!A{i}",
            f"=Inventory!C{i}",
            f"=Inventory!B{i}-Inventory!C{i}-COUNTIF('{DATA_SHEET_NAME}'!F:F;A{dash_row})",
            f"=COUNTIF('{DATA_SHEET_NAME}'!F:F;A{dash_row})",
            f"=Inventory!B{i}",
        ]
        dashboard_rows.append(row)

    sheets.values().clear(
        spreadsheetId=sheet_id, range="Dashboard!A:Z"
    ).execute()
    sheets.values().update(
        spreadsheetId=sheet_id,
        range="Dashboard!A1",
        valueInputOption="USER_ENTERED",
        body={"values": dashboard_rows},
    ).execute()
    print(f"  Dashboard built with {inv_rows_count} tool rows + formulas")

    # --- Step 6: Format Dashboard ---
    dashboard_sheet_id = get_sheet_id_by_title(spreadsheet, "Dashboard")
    if dashboard_sheet_id is not None:
        fmt_requests = []

        # Header: dark background, white bold text
        fmt_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_sheet_id,
                    "startRowIndex": 0, "endRowIndex": 1,
                    "startColumnIndex": 0, "endColumnIndex": 5,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.2, "green": 0.3, "blue": 0.4},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        })

        # Center-align number columns (B-E)
        fmt_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_sheet_id,
                    "startRowIndex": 1, "endRowIndex": inv_rows_count + 1,
                    "startColumnIndex": 1, "endColumnIndex": 5,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment)",
            }
        })

        # Column widths
        col_widths = [300, 100, 100, 100, 100]
        for idx, width in enumerate(col_widths):
            fmt_requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": dashboard_sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": idx, "endIndex": idx + 1,
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            })

        # Alternating row colors
        for row_idx in range(1, inv_rows_count + 1):
            if row_idx % 2 == 0:
                fmt_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": dashboard_sheet_id,
                            "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                            "startColumnIndex": 0, "endColumnIndex": 5,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.93, "green": 0.95, "blue": 0.97},
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)",
                    }
                })

        # Freeze header row
        fmt_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": dashboard_sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        })

        sheets.batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": fmt_requests},
        ).execute()
        print("  Dashboard formatted")

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    print(f"\nDone! Open your sheet: {url}")
    return url


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Set up dashboard in Google Sheet")
    parser.add_argument("--sheet-id", default=DEFAULT_SHEET_ID)
    args = parser.parse_args()
    setup_dashboard(args.sheet_id)
