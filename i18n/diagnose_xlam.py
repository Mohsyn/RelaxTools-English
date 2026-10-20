#!/usr/bin/env python3
"""
Diagnose remaining Japanese text inside RelaxTools_EN.xlam
"""
import json
import re
import zipfile
from xml.etree import ElementTree as ET

JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')

def main():
    cache = json.load(open('i18n/translation_cache.json', encoding='utf-8'))
    zf = zipfile.ZipFile('RelaxTools_EN.xlam', 'r')

    # --- 1. customUI.xml remaining Japanese ---
    raw_custom = zf.read('customUI/customUI.xml').decode('utf-8')
    attr_re = re.compile(r'\b(label|title|supertip|description|screentip)="([^"]*)"', re.IGNORECASE)
    custom_jp = []
    for m in attr_re.finditer(raw_custom):
        val = m.group(2)
        if JP_RE.search(val):
            custom_jp.append((m.group(1), val))

    print(f"=== customUI.xml: {len(custom_jp)} attributes still have Japanese ===")
    for attr, val in custom_jp[:20]:
        cleaned = val.replace('&#xD;&#xA;', ' ').strip()
        in_cache = cleaned in cache or val in cache
        print(f"  [{attr}] {val[:60]!r}  (in_cache={in_cache})")

    # --- 2. sharedStrings.xml remaining Japanese ---
    raw_ss = zf.read('xl/sharedStrings.xml')
    tree = ET.fromstring(raw_ss)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    ss_jp = []
    for si in tree.findall('main:si', ns):
        # Get ALL text including from <rPh> phonetic tags
        main_t = si.find('main:t', ns)
        main_text = main_t.text if main_t is not None else ''
        if not main_text:
            # Rich text: multiple <r><t>...</t></r> runs
            r_els = si.findall('main:r', ns)
            if r_els:
                main_text = ''.join([(t.text or '') for r in r_els for t in r.findall('main:t', ns)])
        if main_text and JP_RE.search(main_text):
            ss_jp.append(main_text)

    print(f"\n=== sharedStrings.xml: {len(ss_jp)} entries still have Japanese ===")
    for txt in ss_jp[:30]:
        cleaned = txt.strip()
        in_cache = cleaned in cache
        print(f"  {txt[:70]!r}  (in_cache={in_cache})")

    # --- 3. Count what's NOT in cache at all ---
    all_missing = set()
    for txt in ss_jp:
        cleaned = txt.strip()
        if cleaned not in cache:
            all_missing.add(cleaned)
    for _, val in custom_jp:
        cleaned = val.replace('&#xD;&#xA;', '\n').strip()
        if cleaned not in cache:
            all_missing.add(cleaned)

    print(f"\n=== TOTAL unique Japanese phrases NOT in cache: {len(all_missing)} ===")
    for t in sorted(all_missing, key=len)[:15]:
        print(f"  {t[:80]!r}")

    # --- 4. Count what IS in cache but wasn't applied ---
    in_cache_not_applied = 0
    for txt in ss_jp:
        cleaned = txt.strip()
        if cleaned in cache and cache[cleaned] != cleaned:
            in_cache_not_applied += 1
    print(f"\n=== sharedStrings: in cache but NOT applied: {in_cache_not_applied} ===")

    # Dump all missing to a file for batch translation
    with open('i18n/missing_phrases.json', 'w', encoding='utf-8') as f:
        json.dump(list(all_missing), f, ensure_ascii=False, indent=2)
    print(f"\nDumped {len(all_missing)} missing phrases to i18n/missing_phrases.json")

if __name__ == '__main__':
    main()
