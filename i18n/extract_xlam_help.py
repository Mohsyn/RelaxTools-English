#!/usr/bin/env python3
"""
RelaxTools XLAM HELP Sheet Extractor
======================================
Treats RelaxTools.xlam as a ZIP file, locates the HELP worksheet,
and extracts all Japanese cell values into strings_raw.csv (appended).

The HELP sheet provides:
  - Ribbon button labels        (col C_COL_LABEL = 4)
  - Menu category               (col C_COL_CATEGORY = 2)
  - Supertip / tooltip text     (col C_COL_HELP = 6)
  - Description                 (col C_COL_DESCRIPTION = 7)

Usage:
    python i18n/extract_xlam_help.py [--xlam RelaxTools.xlam] [--out i18n/strings_raw.csv]

Rows added to CSV have file = "XLAM:HELP" and context = "cell".
"""

import argparse
import csv
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# Excel XML namespace
NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF\uFF65-\uFF9F]')


def has_jp(text: str) -> bool:
    return bool(JP_RE.search(text))


FIELDNAMES = ['key', 'file', 'line', 'context', 'attr_name', 'ja_text', 'en_text', 'status']


# ---------------------------------------------------------------------------
# Parse workbook.xml to find sheet name -> rId -> target mapping
# ---------------------------------------------------------------------------
def get_sheet_targets(zf: zipfile.ZipFile) -> dict:
    """Returns dict: sheet_name -> xl/worksheets/sheetN.xml path."""
    # Read workbook.xml
    wb_xml = zf.read('xl/workbook.xml')
    wb_tree = ET.fromstring(wb_xml)

    # Read workbook _rels
    rels_xml = zf.read('xl/_rels/workbook.xml.rels')
    rels_tree = ET.fromstring(rels_xml)

    # Build rId -> target map
    rid_to_target = {}
    for rel in rels_tree:
        rid = rel.get('Id', '')
        target = rel.get('Target', '')
        rid_to_target[rid] = target

    # Build sheet name -> path
    sheets = {}
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    for sheet in wb_tree.findall('.//main:sheet', ns):
        name = sheet.get('name', '')
        rid  = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        if rid in rid_to_target:
            target = rid_to_target[rid]
            # Target is relative to xl/, e.g. "worksheets/sheet13.xml"
            sheets[name] = f"xl/{target}"
    return sheets


# ---------------------------------------------------------------------------
# Parse shared strings table
# ---------------------------------------------------------------------------
def get_shared_strings(zf: zipfile.ZipFile) -> list:
    """Returns list of string values from xl/sharedStrings.xml."""
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    xml = zf.read('xl/sharedStrings.xml')
    tree = ET.fromstring(xml)
    strings = []
    ns_url = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns = {'main': ns_url}
    for si in tree.findall('main:si', ns):
        # Concatenate all <t> text nodes within the <si>
        parts = [t.text or '' for t in si.findall('.//main:t', ns)]
        strings.append(''.join(parts))
    return strings


# ---------------------------------------------------------------------------
# Parse sheet XML to extract cell values
# ---------------------------------------------------------------------------
def extract_sheet_cells(zf: zipfile.ZipFile, sheet_path: str, shared: list) -> list:
    """
    Returns list of (row_num: int, col_num: int, value: str) for all cells
    that have string values containing Japanese text.
    col_num is 1-based (A=1, B=2, ...).
    """
    xml = zf.read(sheet_path)
    tree = ET.fromstring(xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    cells = []
    for row_el in tree.findall('.//main:row', ns):
        row_num_str = row_el.get('r', '0')
        try:
            row_num = int(row_num_str)
        except ValueError:
            continue

        for c_el in row_el.findall('main:c', ns):
            ref  = c_el.get('r', '')   # e.g. "D25"
            type_ = c_el.get('t', '')  # 's' = shared string, '' = inline/number

            # Parse column from ref
            col_letters = re.match(r'([A-Z]+)', ref)
            if not col_letters:
                continue
            col_str = col_letters.group(1)
            col_num = 0
            for ch in col_str:
                col_num = col_num * 26 + (ord(ch) - ord('A') + 1)

            # Get value
            v_el = c_el.find('main:v', ns)
            if v_el is None:
                # Try inline string
                is_el = c_el.find('main:is', ns)
                if is_el is not None:
                    parts = [t.text or '' for t in is_el.findall('.//main:t', ns)]
                    value = ''.join(parts)
                else:
                    continue
            elif type_ == 's':
                # Shared string index
                try:
                    idx = int(v_el.text or '0')
                    value = shared[idx] if idx < len(shared) else ''
                except (ValueError, IndexError):
                    continue
            else:
                value = v_el.text or ''

            if value and has_jp(value):
                cells.append((row_num, col_num, value))

    return cells


# ---------------------------------------------------------------------------
# Column number to context label
# Based on basRibbon.bas constants:
#   C_COL_NO=1, C_COL_CATEGORY=2, C_COL_MACRO=3, C_COL_LABEL=4,
#   C_COL_DIVISION=5, C_COL_HELP=6, C_COL_DESCRIPTION=7
# ---------------------------------------------------------------------------
COL_CONTEXT = {
    1: 'no',
    2: 'category',
    3: 'macro_name',
    4: 'label',
    5: 'division',
    6: 'supertip',
    7: 'description',
}


def col_to_letter(n: int) -> str:
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord('A') + rem) + result
    return result


# ---------------------------------------------------------------------------
# Load existing CSV to avoid duplicates
# ---------------------------------------------------------------------------
def load_existing(csv_path: Path) -> tuple:
    """Returns (rows: list, existing_keys: set)."""
    if not csv_path.exists():
        return [], set()
    with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    keys = {r['key'] for r in rows}
    return rows, keys


def write_csv(rows: list, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, '') for k in FIELDNAMES})


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description='Extract HELP sheet strings from XLAM')
    parser.add_argument('--xlam',   default='RelaxTools.xlam',       help='Path to XLAM file')
    parser.add_argument('--out',    default='i18n/strings_raw.csv',  help='Output/append CSV')
    parser.add_argument('--sheet',  default='HELP',                   help='Worksheet name to extract')
    parser.add_argument('--root',   default='.',                       help='Repo root')
    args = parser.parse_args()

    root      = Path(args.root).resolve()
    xlam_path = (root / args.xlam).resolve()
    out_path  = (root / args.out).resolve()

    if not xlam_path.exists():
        print(f"ERROR: XLAM not found: {xlam_path}", file=sys.stderr)
        sys.exit(1)

    print(f"Opening {xlam_path} as ZIP ...")
    try:
        with zipfile.ZipFile(xlam_path, 'r') as zf:
            sheets = get_sheet_targets(zf)
            print(f"  Found {len(sheets)} worksheets: {list(sheets.keys())}")

            if args.sheet not in sheets:
                print(f"ERROR: Sheet '{args.sheet}' not found.", file=sys.stderr)
                print(f"Available sheets: {list(sheets.keys())}", file=sys.stderr)
                sys.exit(1)

            sheet_path = sheets[args.sheet]
            print(f"  Extracting from: {sheet_path}")

            shared = get_shared_strings(zf)
            print(f"  Shared strings count: {len(shared)}")

            cells = extract_sheet_cells(zf, sheet_path, shared)
            print(f"  Japanese cells found: {len(cells)}")

    except zipfile.BadZipFile:
        print(f"ERROR: {xlam_path} is not a valid ZIP/XLAM file.", file=sys.stderr)
        sys.exit(1)

    # Load existing CSV
    existing_rows, existing_keys = load_existing(out_path)
    print(f"\nExisting CSV rows: {len(existing_rows)}")

    new_rows = []
    for row_num, col_num, value in cells:
        ctx_label = COL_CONTEXT.get(col_num, f'col{col_num}')
        col_letter = col_to_letter(col_num)
        key = f"XLAM_HELP_R{row_num}_{col_letter}_{ctx_label}"

        if key in existing_keys:
            continue  # already extracted

        new_rows.append({
            'key':      key,
            'file':     f"XLAM:{args.sheet}",
            'line':     row_num,
            'context':  f'cell:{ctx_label}',
            'attr_name': col_letter,
            'ja_text':  value,
            'en_text':  '',
            'status':   'pending',
        })

    print(f"New HELP sheet rows to add: {len(new_rows)}")

    if new_rows:
        all_rows = existing_rows + new_rows
        write_csv(all_rows, out_path)
        print(f"Updated CSV: {out_path} ({len(all_rows)} total rows)")
    else:
        print("No new rows to add. CSV unchanged.")

    print("\nDone. Translate XLAM:HELP rows in the CSV, then run apply_xlam_help.py")


if __name__ == '__main__':
    main()
