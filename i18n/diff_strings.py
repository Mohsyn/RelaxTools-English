#!/usr/bin/env python3
"""
RelaxTools String Differ
=========================
Compares two versions of strings_raw.csv to identify:
  - NEW strings     (in curr but not in prev, by ja_text+file)
  - CHANGED strings (ja_text changed for same key)
  - REMOVED strings (in prev but not in curr)

Usage:
    python i18n/diff_strings.py
    python i18n/diff_strings.py --prev i18n/strings_prev.csv --curr i18n/strings_raw.csv --out i18n/new_strings.csv

Intended workflow:
    1. Before syncing upstream: copy strings_raw.csv -> strings_prev.csv
    2. Run extract_strings.py to refresh strings_raw.csv from new source
    3. Run diff_strings.py to see what's new/changed
    4. Translate new_strings.csv, merge back into strings_raw.csv
"""

import argparse
import csv
import sys
from pathlib import Path


FIELDNAMES = ['key', 'file', 'line', 'context', 'attr_name', 'ja_text', 'en_text', 'status', 'diff_type']


def load_csv(path: Path) -> list:
    if not path.exists():
        return []
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def write_csv(rows: list, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, '') for k in FIELDNAMES})
    print(f"  Written {len(rows)} rows -> {path}")


def diff(prev_rows: list, curr_rows: list) -> tuple:
    """
    Returns (new_rows, changed_rows, removed_rows).
    Matching strategy:
      1. Exact key match first
      2. Fall back to (file, line) match
      3. Fall back to ja_text match within same file
    """
    prev_by_key  = {r['key']: r for r in prev_rows}
    curr_by_key  = {r['key']: r for r in curr_rows}

    # Build file+line index for prev
    prev_by_loc  = {(r['file'], r['line']): r for r in prev_rows}
    # Build ja_text index per file for prev
    prev_by_text = {}
    for r in prev_rows:
        prev_by_text.setdefault(r['file'], {})[r['ja_text']] = r

    new_rows     = []
    changed_rows = []
    matched_prev_keys = set()

    for r in curr_rows:
        key = r['key']
        file_ = r['file']
        line  = r['line']
        ja    = r['ja_text']

        # --- Strategy 1: exact key match ---
        if key in prev_by_key:
            prev_r = prev_by_key[key]
            matched_prev_keys.add(key)
            if prev_r['ja_text'] != ja:
                # String was changed under the same key (unusual but possible)
                r_out = dict(r)
                r_out['diff_type'] = 'changed'
                r_out['en_text'] = ''
                r_out['status'] = 'pending'
                changed_rows.append(r_out)
            continue

        # --- Strategy 2: same file+line (key changed but same location) ---
        if (file_, line) in prev_by_loc:
            prev_r = prev_by_loc[(file_, line)]
            matched_prev_keys.add(prev_r['key'])
            if prev_r['ja_text'] != ja:
                r_out = dict(r)
                r_out['diff_type'] = 'changed'
                r_out['en_text'] = ''
                r_out['status'] = 'pending'
                changed_rows.append(r_out)
            continue

        # --- Strategy 3: same file, same Japanese text (line shifted) ---
        file_prev_texts = prev_by_text.get(file_, {})
        if ja in file_prev_texts:
            prev_r = file_prev_texts[ja]
            matched_prev_keys.add(prev_r['key'])
            # Not new - just moved
            continue

        # --- No match: genuinely new string ---
        r_out = dict(r)
        r_out['diff_type'] = 'new'
        new_rows.append(r_out)

    # Removed strings: in prev but never matched
    removed_rows = []
    for r in prev_rows:
        if r['key'] not in matched_prev_keys and r.get('en_text', ''):
            r_out = dict(r)
            r_out['diff_type'] = 'removed'
            removed_rows.append(r_out)

    return new_rows, changed_rows, removed_rows


def main():
    parser = argparse.ArgumentParser(description='Diff two strings CSV files after upstream sync')
    parser.add_argument('--prev', default='i18n/strings_prev.csv',  help='Previous snapshot CSV')
    parser.add_argument('--curr', default='i18n/strings_raw.csv',   help='Current (new) CSV')
    parser.add_argument('--out',  default='i18n/new_strings.csv',   help='Output: new/changed strings')
    parser.add_argument('--root', default='.',                        help='Repo root')
    args = parser.parse_args()

    root = Path(args.root).resolve()
    prev_path = (root / args.prev).resolve()
    curr_path = (root / args.curr).resolve()
    out_path  = (root / args.out).resolve()

    if not curr_path.exists():
        print(f"ERROR: Current CSV not found: {curr_path}", file=sys.stderr)
        print("Run extract_strings.py first.", file=sys.stderr)
        sys.exit(1)

    if not prev_path.exists():
        print(f"NOTE: No previous snapshot found at {prev_path}")
        print("Treating all rows as 'new'. Run this script after a sync to see actual diffs.")
        prev_rows = []
    else:
        prev_rows = load_csv(prev_path)
        print(f"Loaded {len(prev_rows)} rows from previous snapshot")

    curr_rows = load_csv(curr_path)
    print(f"Loaded {len(curr_rows)} rows from current extraction")

    new_rows, changed_rows, removed_rows = diff(prev_rows, curr_rows)

    print(f"\nDiff results:")
    print(f"  NEW     strings: {len(new_rows)}")
    print(f"  CHANGED strings: {len(changed_rows)}")
    print(f"  REMOVED strings: {len(removed_rows)} (had translations, no longer in source)")

    all_action = new_rows + changed_rows
    if all_action:
        write_csv(all_action, out_path)
        print(f"\n-> Review {out_path}")
        print("   Add en_text for each row, then run:")
        print("   python i18n/apply_translations.py")
    else:
        print("\nNo new or changed strings. Nothing to translate!")

    if removed_rows:
        removed_path = out_path.parent / 'removed_strings.csv'
        write_csv(removed_rows, removed_path)
        print(f"\n-> Removed strings (previously translated, now gone): {removed_path}")

    print("\nDone.")


if __name__ == '__main__':
    main()
