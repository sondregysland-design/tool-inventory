---
name: pdf-serial-reader
description: >
  Extract Serial / Trace numbers from Halliburton Load Out List PDFs and upload to Google Sheets.
  Use when user asks to read PDFs, extract serial numbers, process load out lists, monitor Gmail for PDFs, or export equipment data.
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# PDF Serial / Trace # Reader

## Goal
Read Halliburton Load Out List PDFs, extract Serial / Trace # values from the LOAD OUT LIST table, and upload to Google Sheets. Supports both manual processing and automatic Gmail polling.

## Scripts
- `scripts/pdf-serial-reader/extract_serials.py` - Extract serial numbers from PDFs
- `scripts/pdf-serial-reader/upload_to_sheet.py` - Upload to Google Sheets
- `scripts/pdf-serial-reader/gmail_poll.py` - Auto-poll Gmail for PDF emails

## Method 1: Manual (from local PDFs)

### Extract + upload
```bash
cd "C:\Users\sondr\PDF reading"
python scripts/pdf-serial-reader/extract_serials.py --dir "." --json --output .tmp/serials.json
python scripts/pdf-serial-reader/upload_to_sheet.py --input .tmp/serials.json --append
```

### Extract to CSV (for Power BI)
```bash
python scripts/pdf-serial-reader/extract_serials.py --dir "."
```

## Method 2: Automatic (Gmail polling)

### Run once
```bash
python scripts/pdf-serial-reader/gmail_poll.py
```

### Dry run (check without processing)
```bash
python scripts/pdf-serial-reader/gmail_poll.py --dry-run
```

### How it works
1. Checks Gmail for unread emails with PDF attachments (label: "PDFs")
2. Downloads PDFs, extracts serial numbers
3. Appends data to Google Sheet
4. Marks email as read

### Gmail setup (one-time)
1. Create Gmail label "PDFs"
2. Create filter: emails with PDF attachments → add label "PDFs"
3. Gmail API must be enabled in Google Cloud Console

### Schedule with Task Scheduler
Run every 5 minutes automatically via Windows Task Scheduler.

## Google Sheet Columns

| Column | Description |
|--------|-------------|
| Filename | PDF file name |
| Customer | From GENERAL INFO |
| Well Name | From GENERAL INFO |
| Load Out Date | From GENERAL INFO |
| Material # | From LOAD OUT LIST |
| Description | Equipment description |
| Serial / Trace # | The serial/trace number |
| Qty | Quantity |
| Extracted At | Extraction timestamp |

## Target Sheet
https://docs.google.com/spreadsheets/d/1wK92FpXq-07LdYYPCwZi7-C2vruLPs59JM14w4nAggs

## Error Handling
- **No PDFs found**: Prints warning, exits cleanly
- **PDF read error**: Logs error, skips file, continues with others
- **Table extraction fails**: Falls back to text-based regex extraction
- **No new emails**: Exits silently (normal for polling)
