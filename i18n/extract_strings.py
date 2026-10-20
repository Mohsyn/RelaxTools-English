#!/usr/bin/env python3
"""
RelaxTools String Extractor
============================
Scans all source files under Source/ (and optionally the .xlam HELP sheet)
for Japanese text and writes i18n/strings_raw.csv.

Usage:
    python i18n/extract_strings.py [--source SOURCE_DIR] [--out OUTPUT_CSV]

Output CSV columns:
    key         - Stable identifier: {file_slug}_L{line}_{context}_{instance}
    file        - Relative path from repo root
    line        - Line number in the source file (1-based)
    context     - Type: property | comment | string_literal | xml_attr | readme
    attr_name   - Attribute/property name if context is property/xml_attr
    ja_text     - Original Japanese string
    en_text     - English translation (fill this in)
    status      - pending | approved | skip
"""

import argparse
import csv
import hashlib
import os
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Japanese character detection
# Covers Hiragana, Katakana, CJK Unified Ideographs, half-width Katakana
# ---------------------------------------------------------------------------
JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF\uFF65-\uFF9F]')


def has_jp(text: str) -> bool:
    return bool(JP_RE.search(text))


# ---------------------------------------------------------------------------
# File reading with encoding fallback
# ---------------------------------------------------------------------------
def read_file(path: Path) -> str:
    for enc in ('utf-8-sig', 'cp932', 'utf-8', 'latin-1'):
        try:
            return path.read_text(encoding=enc, errors='strict')
        except (UnicodeDecodeError, LookupError):
            continue
    return path.read_text(encoding='latin-1', errors='replace')


# ---------------------------------------------------------------------------
# Key generation
# ---------------------------------------------------------------------------
def slugify(text: str) -> str:
    """Convert a file path or label to a safe identifier slug."""
    text = re.sub(r'[\\\/]', '_', text)
    text = re.sub(r'[^a-zA-Z0-9_]', '', text)
    return text.strip('_')


def make_key(file_rel: str, line: int, context: str, instance: int = 0) -> str:
    slug = slugify(file_rel)
    inst = f"_{instance}" if instance > 0 else ""
    ctx = re.sub(r'[^a-zA-Z0-9]', '', context)[:12]
    return f"{slug}_L{line}_{ctx}{inst}"


def row(key, file_rel, line, context, attr_name, ja_text):
    return {
        'key':      key,
        'file':     str(file_rel),
        'line':     line,
        'context':  context,
        'attr_name': attr_name,
        'ja_text':  ja_text,
        'en_text':  '',
        'status':   'pending',
    }


# ---------------------------------------------------------------------------
# VBA / FRM parser
# ---------------------------------------------------------------------------
# VBA strings can contain "" as escaped double-quote.
# We parse them character by character to handle this correctly.
def parse_vba_strings_on_line(line: str):
    """
    Return list of (start_col, string_value) for each "..." literal on the line,
    stopping when an unquoted ' (comment) is encountered.
    """
    results = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == "'":
            # Everything from here is a comment - stop scanning for strings
            break
        if ch == '"':
            # Start of a VBA string literal
            j = i + 1
            buf = []
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        buf.append('"')   # escaped ""
                        j += 2
                    else:
                        break             # end of string
                else:
                    buf.append(line[j])
                    j += 1
            results.append((i, ''.join(buf)))
            i = j + 1
        else:
            i += 1
    return results


def get_inline_comment(line: str) -> str:
    """
    Return the comment text after the first unquoted ' on the line, or ''.
    """
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_str:
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
            else:
                in_str = True
        elif ch == "'" and not in_str:
            return line[i + 1:].strip()
        i += 1
    return ''


def extract_vba(content: str, file_rel: str) -> list:
    """Extract Japanese strings from .bas, .cls, and .frm files."""
    rows = []
    lines = content.splitlines()

    # .frm files have a Begin...End header section before the VBA code.
    # We track nesting depth to know when we're in that section.
    in_header = False
    header_done = False
    depth = 0
    # Counter per line to make keys unique for multiple strings on same line
    line_inst: dict = {}

    def add(line_num, context, attr_name, ja_text):
        inst = line_inst.get(line_num, 0)
        line_inst[line_num] = inst + 1
        k = make_key(file_rel, line_num, context, inst)
        rows.append(row(k, file_rel, line_num, context, attr_name, ja_text))

    is_frm = str(file_rel).endswith('.frm')

    for idx, line in enumerate(lines, 1):
        stripped = line.strip()

        # --- FRM HEADER SECTION (Begin...End blocks) ---
        if is_frm and not header_done:
            if re.match(r'^Begin\b', stripped):
                depth += 1
                in_header = True
                continue

            if stripped == 'End' and in_header:
                depth -= 1
                if depth == 0:
                    header_done = True
                continue

            if in_header and depth > 0:
                # Property = "value"
                m = re.match(r'^\s*(\w+)\s*=\s*"([^"]*)"', line)
                if m and has_jp(m.group(2)):
                    add(idx, 'property', m.group(1), m.group(2))

                # NumProp = value  'inline comment
                m2 = re.match(r'^\s*(\w+)\s*=\s*\S+\s+\'(.+)$', line)
                if m2 and has_jp(m2.group(2)):
                    add(idx, 'comment', '', m2.group(2).strip())
                continue

        # --- VBA CODE SECTION ---
        # Skip Attribute lines (VB metadata, not translatable)
        if re.match(r'^\s*Attribute\s+VB_', stripped):
            continue

        # Full-line comment
        if re.match(r"^\s*'", stripped):
            comment_text = re.sub(r"^\s*'+\s*", '', stripped)
            if has_jp(comment_text):
                add(idx, 'comment', '', comment_text)
            continue

        # Inline comment on a code line
        ic = get_inline_comment(line)
        if ic and has_jp(ic):
            add(idx, 'comment', '', ic)

        # String literals in code
        for _col, sval in parse_vba_strings_on_line(line):
            if has_jp(sval):
                add(idx, 'string_literal', '', sval)

    return rows


# ---------------------------------------------------------------------------
# XML parser (customUI.xml)
# ---------------------------------------------------------------------------
# Attributes that can contain user-visible Japanese text
XML_UI_ATTRS = {'label', 'title', 'supertip', 'description', 'screentip', 'keytip'}


def extract_xml(content: str, file_rel: str) -> list:
    rows = []
    lines = content.splitlines()
    line_inst: dict = {}

    def add(line_num, attr_name, ja_text):
        inst = line_inst.get(line_num, 0)
        line_inst[line_num] = inst + 1
        k = make_key(file_rel, line_num, f'xml_{attr_name}', inst)
        rows.append(row(k, file_rel, line_num, 'xml_attr', attr_name, ja_text))

    # Match attribute="value" pairs
    attr_re = re.compile(
        r'\b(' + '|'.join(XML_UI_ATTRS) + r')\s*=\s*"([^"]*)"',
        re.IGNORECASE
    )
    # Also match title= in menuSeparator
    sep_re = re.compile(r'\btitle\s*=\s*"([^"]*)"', re.IGNORECASE)

    for idx, line in enumerate(lines, 1):
        for m in attr_re.finditer(line):
            val = m.group(2)
            if has_jp(val):
                add(idx, m.group(1).lower(), val)

    return rows


# ---------------------------------------------------------------------------
# VBS parser
# ---------------------------------------------------------------------------
def extract_vbs(content: str, file_rel: str) -> list:
    """VBScript is similar to VBA - same string/comment rules."""
    return extract_vba(content, file_rel)


# ---------------------------------------------------------------------------
# Markdown / plain text parser
# ---------------------------------------------------------------------------
def extract_text(content: str, file_rel: str) -> list:
    rows = []
    lines = content.splitlines()
    line_inst: dict = {}

    for idx, line in enumerate(lines, 1):
        stripped = line.strip()
        if has_jp(stripped) and stripped:
            inst = line_inst.get(idx, 0)
            line_inst[idx] = inst + 1
            k = make_key(file_rel, idx, 'readme', inst)
            rows.append(row(k, file_rel, idx, 'readme', '', stripped))

    return rows


# ---------------------------------------------------------------------------
# Dispatch by extension
# ---------------------------------------------------------------------------
EXTRACTORS = {
    '.bas': extract_vba,
    '.cls': extract_vba,
    '.frm': extract_vba,
    '.vbs': extract_vbs,
    '.xml': extract_xml,
    '.md':  extract_text,
    '.txt': extract_text,
}

# Files/dirs to skip entirely
SKIP_DIRS = {'Source_EN', '.git', 'i18n', '__pycache__'}
SKIP_FILES = {'strings_raw.csv', 'strings_prev.csv', 'new_strings.csv'}


# ---------------------------------------------------------------------------
# Main scan
# ---------------------------------------------------------------------------
def scan_source(source_dir: Path, repo_root: Path) -> list:
    all_rows = []
    seen_keys: set = set()

    for fpath in sorted(source_dir.rglob('*')):
        # Skip directories
        if fpath.is_dir():
            continue

        # Skip excluded dirs
        if any(p in SKIP_DIRS for p in fpath.parts):
            continue

        # Skip excluded files
        if fpath.name in SKIP_FILES:
            continue

        ext = fpath.suffix.lower()
        if ext not in EXTRACTORS:
            continue

        # Skip .frx binary form files (not text)
        if ext == '.frx':
            continue

        try:
            content = read_file(fpath)
        except Exception as e:
            print(f"  [WARN] Cannot read {fpath}: {e}", file=sys.stderr)
            continue

        file_rel = fpath.relative_to(repo_root)
        extractor = EXTRACTORS[ext]

        try:
            file_rows = extractor(content, str(file_rel).replace('\\', '/'))
        except Exception as e:
            print(f"  [WARN] Error parsing {fpath}: {e}", file=sys.stderr)
            continue

        # Deduplicate keys (can happen with multiple strings same line)
        for r in file_rows:
            orig_key = r['key']
            counter = 0
            while r['key'] in seen_keys:
                counter += 1
                r['key'] = f"{orig_key}_x{counter}"
            seen_keys.add(r['key'])
            all_rows.append(r)

    return all_rows


# ---------------------------------------------------------------------------
# Also scan root-level README and text files
# ---------------------------------------------------------------------------
def scan_root_docs(repo_root: Path) -> list:
    rows = []
    for name in ('README.md', 'README.txt', 'Version.txt'):
        fpath = repo_root / name
        if fpath.exists():
            try:
                content = read_file(fpath)
                file_rel = name
                rows.extend(extract_text(content, file_rel))
            except Exception as e:
                print(f"  [WARN] Cannot read {fpath}: {e}", file=sys.stderr)
    return rows


# ---------------------------------------------------------------------------
# Merge with existing CSV (preserve en_text and status for known keys)
# ---------------------------------------------------------------------------
def merge_with_existing(new_rows: list, existing_csv: Path) -> list:
    if not existing_csv.exists():
        return new_rows

    existing: dict = {}
    try:
        with open(existing_csv, 'r', encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f):
                existing[r['key']] = r
    except Exception as e:
        print(f"  [WARN] Cannot read existing CSV: {e}", file=sys.stderr)
        return new_rows

    # Also build lookup by ja_text for fuzzy matching when keys shift
    existing_by_text: dict = {}
    for k, r in existing.items():
        txt = r.get('ja_text', '')
        if txt and r.get('en_text', ''):
            existing_by_text.setdefault(txt, r)

    merged = []
    for r in new_rows:
        key = r['key']
        ja  = r['ja_text']
        if key in existing:
            # Exact key match — preserve translation
            r['en_text'] = existing[key].get('en_text', '')
            r['status']  = existing[key].get('status', 'pending')
        elif ja in existing_by_text:
            # Same Japanese text, key shifted (e.g., line number changed)
            old = existing_by_text[ja]
            r['en_text'] = old.get('en_text', '')
            r['status']  = old.get('status', 'pending') if old.get('en_text') else 'pending'
        merged.append(r)

    return merged


# ---------------------------------------------------------------------------
# Write CSV
# ---------------------------------------------------------------------------
FIELDNAMES = ['key', 'file', 'line', 'context', 'attr_name', 'ja_text', 'en_text', 'status']


def write_csv(rows: list, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  Written {len(rows)} rows -> {path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description='Extract Japanese strings from RelaxTools source')
    parser.add_argument('--source',  default='Source',         help='Source directory (relative to repo root)')
    parser.add_argument('--out',     default='i18n/strings_raw.csv', help='Output CSV path')
    parser.add_argument('--root',    default='.',              help='Repo root directory')
    args = parser.parse_args()

    repo_root  = Path(args.root).resolve()
    source_dir = (repo_root / args.source).resolve()
    out_csv    = (repo_root / args.out).resolve()

    if not source_dir.exists():
        print(f"ERROR: Source directory not found: {source_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Scanning {source_dir} ...")
    rows = scan_source(source_dir, repo_root)
    print(f"  Found {len(rows)} Japanese strings in Source/")

    print("Scanning root docs (README, Version.txt) ...")
    rows += scan_root_docs(repo_root)
    print(f"  Total after docs: {len(rows)} rows")

    print("Merging with existing translations ...")
    rows = merge_with_existing(rows, out_csv)

    # Sort: by file then line number for readability
    rows.sort(key=lambda r: (r['file'], int(r['line'])))

    write_csv(rows, out_csv)

    pending   = sum(1 for r in rows if r['status'] == 'pending')
    approved  = sum(1 for r in rows if r['status'] == 'approved')
    skipped   = sum(1 for r in rows if r['status'] == 'skip')
    print(f"\nSummary: {len(rows)} total | {approved} approved | {pending} pending | {skipped} skip")
    print("Done. Open i18n/strings_raw.csv in Excel to add translations.")


if __name__ == '__main__':
    main()
