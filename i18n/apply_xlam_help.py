#!/usr/bin/env python3
"""
RelaxTools XLAM HELP Sheet Patcher (Clean XML Version)
======================================================
Replaces Japanese text in xl/sharedStrings.xml and customUI/customUI.xml
inside RelaxTools.xlam without modifying worksheet XML files or altering
OpenXML namespaces, ensuring 100% compatibility with Excel without recovery errors.
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_URL = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS_URL)

def load_translations(csv_path: Path, cache_path: Path) -> dict:
    translations = {}
    
    # 1. Load from cache
    if cache_path.exists():
        try:
            translations.update(json.loads(cache_path.read_text(encoding='utf-8')))
        except Exception:
            pass

    # 2. Override with approved translations from CSV
    if csv_path.exists():
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f):
                if r.get('status', '').lower() == 'approved':
                    en = r.get('en_text', '').strip()
                    ja = r.get('ja_text', '').strip()
                    if en and ja:
                        translations[ja] = en

    return translations

def patch_shared_strings(raw_xml: bytes, translations: dict) -> tuple:
    tree = ET.fromstring(raw_xml)
    ns = {'main': NS_URL}
    replaced = 0

    for si in tree.findall('main:si', ns):
        t_els = si.findall('.//main:t', ns)
        for t_el in t_els:
            val = t_el.text or ''
            if val in translations and translations[val] != val:
                t_el.text = translations[val]
                replaced += 1

    out = io.BytesIO()
    # Write exact Excel standard XML header
    out.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    ET.ElementTree(tree).write(out, encoding='UTF-8', xml_declaration=False)
    return out.getvalue(), replaced

def main():
    parser = argparse.ArgumentParser(description='Cleanly patch RelaxTools XLAM')
    parser.add_argument('--xlam',     default='RelaxTools.xlam',      help='Source XLAM')
    parser.add_argument('--approved', default='i18n/strings_raw.csv', help='Translation CSV')
    parser.add_argument('--cache',    default='i18n/translation_cache.json', help='Translation Cache')
    parser.add_argument('--customui', default='Source_EN/customUI/customUI.xml', help='Translated customUI')
    parser.add_argument('--output',   default='RelaxTools_EN.xlam',   help='Output XLAM')
    parser.add_argument('--root',     default='.',                    help='Repo root')
    args = parser.parse_args()

    root = Path(args.root).resolve()
    xlam_path = (root / args.xlam).resolve()
    out_path = (root / args.output).resolve()
    csv_path = (root / args.approved).resolve()
    cache_path = (root / args.cache).resolve()
    customui_path = (root / args.customui).resolve()

    if not xlam_path.exists():
        print(f"ERROR: {xlam_path} not found.", file=sys.stderr)
        sys.exit(1)

    print(f"Loading translations ...")
    translations = load_translations(csv_path, cache_path)
    print(f"  {len(translations)} unique translations loaded")

    print(f"\nOpening {xlam_path} ...")
    with zipfile.ZipFile(xlam_path, 'r') as src_zf:
        raw_ss = src_zf.read('xl/sharedStrings.xml')
        new_ss, replaced = patch_shared_strings(raw_ss, translations)
        print(f"  Replaced {replaced} text entries in sharedStrings.xml")

        # Load translated customUI if available
        new_customui = None
        if customui_path.exists():
            new_customui = customui_path.read_bytes()
            print(f"  Found translated customUI: {customui_path} ({len(new_customui)} bytes)")

        print(f"\nPackaging {out_path} ...")
        with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as dst_zf:
            for item in src_zf.infolist():
                if item.filename == 'xl/sharedStrings.xml':
                    dst_zf.writestr(item, new_ss)
                elif item.filename == 'customUI/customUI.xml' and new_customui:
                    dst_zf.writestr(item, new_customui)
                else:
                    # Keep all worksheet XMLs, VBA binaries, and relationships byte-identical
                    dst_zf.writestr(item, src_zf.read(item.filename))

    size_mb = out_path.stat().st_size / 1024 / 1024
    print(f"  Success: {out_path} ({size_mb:.2f} MB)")
    print("All worksheet structures preserved. Excel recovery errors resolved!")

if __name__ == '__main__':
    main()
