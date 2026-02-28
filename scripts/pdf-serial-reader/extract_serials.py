#!/usr/bin/env python3
"""
Extract Serial / Trace # values from Halliburton Load Out List PDFs.

Reads PDFs from a directory, finds the LOAD OUT LIST table on page 1,
and extracts all Serial/Trace numbers along with metadata.
"""

import os
import sys
import json
import csv
import glob
import argparse
import re
from datetime import datetime

import pdfplumber


def extract_general_info(page):
    """Extract metadata from the GENERAL INFO section using layout text."""
    info = {
        "customer": "",
        "well_name": "",
        "load_out_date": "",
    }

    # Use layout=True to preserve column alignment
    text = page.extract_text(layout=True) or ""
    lines = text.split("\n")

    for i, line in enumerate(lines):
        stripped = line.strip()

        # The data row after "Sales Order/Planning Order Number Customer ..."
        # contains customer and well name on the same line, column-aligned.
        # Typical: "                       VÅR ENERGIASA-EBUS  EESSA_NO_SPT_VAAR ..."
        # We detect it by looking for the line after the header row.
        if "Sales Order/Planning Order Number" in stripped:
            for next_line in lines[i + 1 : i + 3]:
                next_stripped = next_line.strip()
                if not next_stripped or "Shipping" in next_stripped:
                    continue
                # Parse the data line: customer is first chunk, well name has underscores
                # Use regex to split known patterns
                cust_match = re.search(
                    r"(V[ÅA]R\s*ENERGI[\w\s\-]*?)\s{2,}(\S+)", next_stripped
                )
                if cust_match:
                    info["customer"] = re.sub(r"\s+", " ", cust_match.group(1)).strip()
                    # Well name follows customer, before coordinator
                    remainder = next_stripped[cust_match.end(1) :].strip()
                    well_match = re.search(
                        r"(EESSA[\w_\-]+\s*[\w_\-\s]*\d+[\w\-\s]*H?)", remainder
                    )
                    if well_match:
                        info["well_name"] = well_match.group(1).strip()
                break

        # Load Out Date
        if "Load Out Date" in stripped:
            for next_line in lines[i + 1 : i + 3]:
                date_match = re.search(r"\d{1,2}\s+\w+\s+\d{4}", next_line)
                if date_match:
                    info["load_out_date"] = date_match.group()
                    break

    # Fallback for customer from raw text
    if not info["customer"]:
        raw = page.extract_text() or ""
        for line in raw.split("\n"):
            if "ENERGI" in line and ("ASA" in line or "EBUS" in line):
                m = re.search(r"(V[ÅA]R\s*ENERGI[\w\s\-]*?(?:EBUS|ASA\S*))", line)
                if m:
                    info["customer"] = re.sub(r"\s+", " ", m.group(1)).strip()
                    break

    # Normalize customer name (fix missing spaces from PDF extraction)
    if info["customer"]:
        info["customer"] = re.sub(r"ENERGI(?=ASA)", "ENERGI ", info["customer"])

    # Fallback for well name from raw text
    if not info["well_name"]:
        raw = page.extract_text() or ""
        for line in raw.split("\n"):
            well_match = re.search(
                r"(EESSA[\w_\-]+\s*[\w_\-]*\s*\d+[\w\-\s]*H?)", line
            )
            if well_match:
                info["well_name"] = well_match.group(1).strip()
                break

    # Clean well name: remove coordinator name if appended
    # Well names contain underscores and end with something like "H" or a number
    if info["well_name"]:
        # Stop at first word that looks like a person name (capitalized, no underscores)
        cleaned = re.match(
            r"([\w_\-]+(?:\s+[\d_\-]+[\w\-]*)*(?:\s+H)?)", info["well_name"]
        )
        if cleaned:
            info["well_name"] = cleaned.group(1).strip()

    return info


def extract_loadout_serials_from_table(page):
    """
    Extract Serial/Trace # values from the LOAD OUT LIST table using
    pdfplumber's table extraction.

    Returns list of dicts with: material_num, description, serial, qty
    """
    results = []

    tables = page.extract_tables()

    for table in tables:
        if not table or len(table) < 2:
            continue

        # Check if this is the LOAD OUT LIST table
        header_row = table[0]
        header_text = " ".join(str(cell or "") for cell in header_row)

        if "Serial" not in header_text and "Trace" not in header_text:
            continue

        # Find column indices
        serial_col = None
        material_col = None
        desc_col = None
        qty_col = None

        for idx, cell in enumerate(header_row):
            cell_text = str(cell or "").strip()
            if "Serial" in cell_text or "Trace" in cell_text:
                serial_col = idx
            elif "Material" in cell_text:
                material_col = idx
            elif "Description" in cell_text:
                desc_col = idx
            elif "Qty" in cell_text:
                qty_col = idx

        if serial_col is None:
            continue

        # Process data rows
        for row in table[1:]:
            if not row or all(
                cell is None or str(cell).strip() == "" for cell in row
            ):
                continue

            serial_cell = str(row[serial_col] or "").strip()
            material_cell = (
                str(row[material_col] or "").strip()
                if material_col is not None
                else ""
            )
            desc_cell = (
                str(row[desc_col] or "").strip() if desc_col is not None else ""
            )
            qty_cell = (
                str(row[qty_col] or "").strip() if qty_col is not None else ""
            )

            # Split multi-line values (common in these PDFs)
            serials = [s.strip() for s in serial_cell.split("\n") if s.strip()]
            materials = [
                m.strip() for m in material_cell.split("\n") if m.strip()
            ]
            descriptions = [
                d.strip() for d in desc_cell.split("\n") if d.strip()
            ]

            # Pair serials with their corresponding material/description
            for i, serial in enumerate(serials):
                if not serial or serial.lower() in ("none", "", "n/a"):
                    continue
                results.append(
                    {
                        "material_num": (
                            materials[i]
                            if i < len(materials)
                            else (materials[0] if materials else "")
                        ),
                        "description": (
                            descriptions[i]
                            if i < len(descriptions)
                            else (descriptions[0] if descriptions else "")
                        ),
                        "serial": serial,
                        "qty": qty_cell,
                    }
                )

    return results


def extract_loadout_serials_from_text(page_text):
    """
    Extract Serial/Trace # values from raw text using line-by-line parsing.
    Each data row in the Load Out List has the format:
      [Material#] Description Serial/Trace# [Qty]
    on a single line.
    """
    results = []

    # Find the LOAD OUT LIST section
    loadout_match = re.search(
        r"LOAD OUT LIST\s*\n(.+?)(?:Load Out Verif|$)",
        page_text,
        re.DOTALL,
    )
    if not loadout_match:
        return results

    section = loadout_match.group(1)

    # Serial number patterns we recognize
    serial_re = re.compile(r"(OWS-[\w\-]+|\b\d{8}\b)")

    for line in section.split("\n"):
        line = line.strip()
        if not line or "Material #" in line or "Description" in line:
            continue

        serial_match = serial_re.search(line)
        if not serial_match:
            continue

        serial = serial_match.group(1)
        before_serial = line[: serial_match.start()].strip()
        after_serial = line[serial_match.end() :].strip()

        # Parse what's before the serial: [Material#] [Description]
        # Material# is first token if it matches alphanumeric pattern
        material_num = ""
        description = before_serial

        mat_match = re.match(r"^(\d{9}|\d{6,}|[A-Z]{2,5})\s+(.+)", before_serial)
        if mat_match:
            material_num = mat_match.group(1)
            description = mat_match.group(2).strip()

        # Qty is the number after the serial (if any)
        qty = ""
        qty_match = re.match(r"^(\d+)", after_serial)
        if qty_match:
            qty = qty_match.group(1)

        results.append(
            {
                "material_num": material_num,
                "description": description,
                "serial": serial,
                "qty": qty,
            }
        )

    return results


def process_pdf(pdf_path):
    """Process a single PDF and return all extracted serial data."""
    filename = os.path.basename(pdf_path)
    all_records = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                print(f"  WARNING: {filename} has no pages", file=sys.stderr)
                return []

            page1 = pdf.pages[0]
            page_text = page1.extract_text() or ""

            # Extract metadata (uses layout=True internally)
            info = extract_general_info(page1)

            # Try table extraction first, fall back to text-based
            loadout_serials = extract_loadout_serials_from_table(page1)

            if not loadout_serials:
                print(
                    f"  Table extraction found nothing, trying text fallback...",
                    file=sys.stderr,
                )
                loadout_serials = extract_loadout_serials_from_text(page_text)

            for item in loadout_serials:
                all_records.append(
                    {
                        "filename": filename,
                        "customer": info["customer"],
                        "well_name": info["well_name"],
                        "load_out_date": info["load_out_date"],
                        "material_num": item["material_num"],
                        "description": item["description"],
                        "serial_trace": item["serial"],
                        "qty": item["qty"],
                        "extracted_at": datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    }
                )

    except Exception as e:
        print(f"  ERROR processing {filename}: {e}", file=sys.stderr)
        return []

    return all_records


CSV_HEADERS = [
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

CSV_FIELDS = [
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


def save_csv(records, output_path, append=False):
    """Save records to CSV file. Appends if file exists and --append is set."""
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    file_exists = os.path.exists(output_path) and os.path.getsize(output_path) > 0
    mode = "a" if append and file_exists else "w"

    with open(output_path, mode, newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if mode == "w" or not file_exists:
            writer.writerow(CSV_HEADERS)
        for record in records:
            writer.writerow([record.get(field, "") for field in CSV_FIELDS])

    action = "Appended to" if mode == "a" and file_exists else "Saved to"
    print(f"{action}: {output_path}")


def save_json(records, output_path):
    """Save records to JSON file."""
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    print(f"Saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Extract Serial/Trace # from Halliburton Load Out PDFs"
    )
    parser.add_argument(
        "--dir", default=".", help="Directory containing PDF files"
    )
    parser.add_argument(
        "--output",
        default="output/serials.csv",
        help="Output file path (default: output/serials.csv)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output as JSON instead of CSV",
    )
    parser.add_argument(
        "--append",
        action="store_true",
        help="Append to existing CSV instead of overwriting",
    )

    args = parser.parse_args()

    # Find all PDFs (both .PDF and .pdf)
    pdf_files = glob.glob(os.path.join(args.dir, "*.PDF"))
    pdf_files += glob.glob(os.path.join(args.dir, "*.pdf"))
    pdf_files = list(set(pdf_files))

    if not pdf_files:
        print(f"No PDF files found in {args.dir}")
        sys.exit(0)

    print(f"Found {len(pdf_files)} PDF file(s)")

    all_records = []
    for pdf_path in sorted(pdf_files):
        print(f"Processing: {os.path.basename(pdf_path)}")
        records = process_pdf(pdf_path)
        print(f"  Extracted {len(records)} serial number(s)")
        all_records.extend(records)

    print(f"\nTotal: {len(all_records)} serial numbers extracted")

    if args.json:
        output = args.output if args.output != "output/serials.csv" else ".tmp/serials.json"
        save_json(all_records, output)
    else:
        save_csv(all_records, args.output, append=args.append)


if __name__ == "__main__":
    main()
