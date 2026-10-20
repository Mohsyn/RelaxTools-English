#!/usr/bin/env python3
"""
RelaxTools Translation Applier
================================
Reads strings_raw.csv (rows with status=approved) and patches
each source file, writing translated output to Source_EN/.

Usage:
    python i18n/apply_translations.py [--approved i18n/strings_raw.csv] [--source Source] [--output Source_EN]

Strategy:
  - For each approved translation, find the source file and line.
  - Replace the exact ja_text with en_text on that line.
  - Process lines in descending order so line numbers stay valid.
  - Write the patched file to Source_EN/ (mirroring Source/ structure).
  - Files without any approved translations are copied as-is.
  - Original Source/ files are NEVER modified.

Notes on VBA strings:
  - String literals: replace "ja_text" with "en_text"
  - Comments: replace the Japanese comment text
  - Properties in .frm Begin blocks: replace the value in quotes
"""

import argparse
import csv
import os
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path


# ---------------------------------------------------------------------------
# File reading / writing with encoding detection
# ---------------------------------------------------------------------------
def read_file_bytes(path: Path) -> bytes:
    return path.read_bytes()


def detect_encoding(path: Path) -> str:
    raw = path.read_bytes()
    # UTF-8 BOM
    if raw.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig'
    # Try cp932 (Shift-JIS, common for VBA on Japanese Windows)
    for enc in ('utf-8', 'cp932', 'latin-1'):
        try:
            raw.decode(enc, errors='strict')
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return 'latin-1'


def read_file(path: Path) -> tuple:
    """Returns (lines: list[str], encoding: str, line_ending: str)."""
    enc = detect_encoding(path)
    text = path.read_text(encoding=enc, errors='replace')

    # Detect line ending
    if '\r\n' in text:
        le = '\r\n'
    elif '\r' in text:
        le = '\r'
    else:
        le = '\n'

    lines = text.splitlines()
    return lines, enc, le


def write_file(path: Path, lines: list, encoding: str, line_ending: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    text = line_ending.join(lines)
    if line_ending in ('\r\n', '\r'):
        text += line_ending  # preserve trailing newline
    path.write_text(text, encoding=encoding)


# ---------------------------------------------------------------------------
# Patch logic
# ---------------------------------------------------------------------------
def patch_string_literal(line: str, ja_text: str, en_text: str) -> str:
    """Replace "ja_text" with "en_text" in VBA string literal."""
    # Escape any regex special chars in ja_text
    escaped = re.escape(ja_text)
    # Match the string inside double-quotes (exact)
    pattern = f'"{escaped}"'
    replacement = f'"{en_text}"'
    # Replace first occurrence only on this line
    new_line, count = re.subn(pattern, replacement, line, count=1)
    if count == 0:
        # Try without quotes (might be a bare property value)
        new_line = line.replace(ja_text, en_text, 1)
    return new_line


def patch_comment(line: str, ja_text: str, en_text: str) -> str:
    """Replace Japanese comment text, preserving comment structure."""
    return line.replace(ja_text, en_text, 1)


def patch_property(line: str, ja_text: str, en_text: str) -> str:
    """Replace property value in Begin block: PropName = "ja_text" -> "en_text"."""
    escaped = re.escape(ja_text)
    pattern = f'"{escaped}"'
    replacement = f'"{en_text}"'
    new_line, count = re.subn(pattern, replacement, line, count=1)
    if count == 0:
        new_line = line.replace(ja_text, en_text, 1)
    return new_line


def patch_xml_attr(line: str, ja_text: str, en_text: str, attr_name: str) -> str:
    """Replace XML attribute value: attr="ja" -> attr="en"."""
    escaped_attr = re.escape(attr_name)
    escaped_val  = re.escape(ja_text)
    pattern = rf'({escaped_attr}\s*=\s*"){escaped_val}(")'
    replacement = rf'\g<1>{en_text}\g<2>'
    new_line, count = re.subn(pattern, replacement, line, count=1, flags=re.IGNORECASE)
    if count == 0:
        new_line = line.replace(ja_text, en_text, 1)
    return new_line


PATCH_FNS = {
    'string_literal': patch_string_literal,
    'comment':        patch_comment,
    'property':       patch_property,
    'xml_attr':       patch_xml_attr,
    'readme':         patch_comment,   # plain text replacement
}


def patch_line(line: str, context: str, attr_name: str, ja_text: str, en_text: str) -> str:
    """Apply the appropriate patch for a given context."""
    fn = PATCH_FNS.get(context, patch_comment)
    if context == 'xml_attr' and attr_name:
        return fn(line, ja_text, en_text, attr_name)
    return fn(line, ja_text, en_text)


# ---------------------------------------------------------------------------
# Load approved translations
# ---------------------------------------------------------------------------
def load_approved(csv_path: Path) -> dict:
    """
    Returns dict: file_rel -> sorted list of (line_num, context, attr_name, ja_text, en_text)
    Only includes rows with status=approved and non-empty en_text.
    """
    by_file = defaultdict(list)
    if not csv_path.exists():
        print(f"ERROR: CSV not found: {csv_path}", file=sys.stderr)
        sys.exit(1)

    with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
        for r in csv.DictReader(f):
            if r.get('status', '').lower() != 'approved':
                continue
            en = r.get('en_text', '').strip()
            if not en:
                continue
            # Normalise file path separators
            file_rel = r['file'].replace('\\', '/')
            try:
                line_num = int(r['line'])
            except (ValueError, KeyError):
                continue

            # Skip XLAM:HELP — handled by apply_xlam_help.py
            if file_rel.startswith('XLAM:'):
                continue

            by_file[file_rel].append({
                'line':     line_num,
                'context':  r.get('context', 'string_literal'),
                'attr_name': r.get('attr_name', ''),
                'ja_text':  r.get('ja_text', ''),
                'en_text':  en,
            })

    # Sort each file's patches by line descending (patch bottom-up to preserve line numbers)
    for k in by_file:
        by_file[k].sort(key=lambda x: x['line'], reverse=True)

    return dict(by_file)


# ---------------------------------------------------------------------------
# Mirror a file to Source_EN/ (no patches needed)
# ---------------------------------------------------------------------------
def mirror_file(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


# ---------------------------------------------------------------------------
# Apply patches to a file and write to dest
# ---------------------------------------------------------------------------
def apply_file_patches(src: Path, dst: Path, patches: list) -> tuple:
    """
    Returns (patched_count, failed_count).
    """
    lines, enc, le = read_file(src)
    patched = 0
    failed  = 0

    for p in patches:
        line_idx = p['line'] - 1  # 0-based
        if line_idx < 0 or line_idx >= len(lines):
            print(f"    [WARN] Line {p['line']} out of range in {src.name}", file=sys.stderr)
            failed += 1
            continue

        original = lines[line_idx]
        if p['ja_text'] not in original:
            # Text not found - may have already been patched or line shifted
            failed += 1
            continue

        new_line = patch_line(original, p['context'], p['attr_name'], p['ja_text'], p['en_text'])
        if new_line == original:
            failed += 1
        else:
            lines[line_idx] = new_line
            patched += 1

    write_file(dst, lines, enc, le)
    return patched, failed


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description='Apply approved translations to Source_EN/')
    parser.add_argument('--approved', default='i18n/strings_raw.csv', help='Translation CSV')
    parser.add_argument('--source',   default='Source',               help='Source directory')
    parser.add_argument('--output',   default='Source_EN',            help='Output directory')
    parser.add_argument('--root',     default='.',                     help='Repo root')
    args = parser.parse_args()

    root       = Path(args.root).resolve()
    src_dir    = (root / args.source).resolve()
    out_dir    = (root / args.output).resolve()
    csv_path   = (root / args.approved).resolve()

    if not src_dir.exists():
        print(f"ERROR: Source directory not found: {src_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Loading approved translations from {csv_path} ...")
    by_file = load_approved(csv_path)
    print(f"  {sum(len(v) for v in by_file.values())} approved patches across {len(by_file)} files")

    total_files   = 0
    total_patched = 0
    total_failed  = 0
    total_copied  = 0

    print(f"\nApplying patches: {src_dir} -> {out_dir}")

    for src_file in sorted(src_dir.rglob('*')):
        if src_file.is_dir():
            continue

        rel = src_file.relative_to(src_dir)
        rel_str = str(rel).replace('\\', '/')
        # Normalise to match CSV keys which use Source/... prefix
        csv_key = f"Source/{rel_str}"

        dst_file = out_dir / rel
        total_files += 1

        patches = by_file.get(csv_key, [])

        if not patches:
            # No patches — just mirror the file
            mirror_file(src_file, dst_file)
            total_copied += 1
            continue

        patched, failed = apply_file_patches(src_file, dst_file, patches)
        total_patched += patched
        total_failed  += failed
        status = f"[{patched} patched, {failed} failed]"
        print(f"  {csv_key:<60} {status}")

    print(f"\nSummary:")
    print(f"  Files processed : {total_files}")
    print(f"  Files with patches: {len(by_file)}")
    print(f"  Successful patches: {total_patched}")
    print(f"  Failed patches    : {total_failed}")
    print(f"  Files copied (no patches): {total_copied}")
    print(f"\nOutput: {out_dir}")


if __name__ == '__main__':
    main()
