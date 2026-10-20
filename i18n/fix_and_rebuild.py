#!/usr/bin/env python3
"""
Translate missing XLAM phrases and rebuild RelaxTools_EN.xlam.
1. Reads missing_phrases.json
2. Translates each via online API
3. Merges into translation_cache.json
4. Runs build_xlam.py logic to repackage
"""
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.sax.saxutils as saxutils
import zipfile
from xml.etree import ElementTree as ET

JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')
CACHE_FILE = 'i18n/translation_cache.json'
MISSING_FILE = 'i18n/missing_phrases.json'

def translate_phrase(text):
    clean = text.strip()
    if not clean or not JP_RE.search(clean):
        return clean
    encoded = urllib.parse.quote(clean)
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=ja&tl=en&dt=t&q={encoded}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                parts = [p[0] for p in data[0] if p[0]]
                return ''.join(parts).strip()
        except Exception:
            time.sleep(1 + attempt)
    return ''

def main():
    # Load cache
    cache = {}
    if os.path.exists(CACHE_FILE):
        cache = json.load(open(CACHE_FILE, encoding='utf-8'))
    print(f"Existing cache: {len(cache)} entries")

    # Load missing phrases
    missing = json.load(open(MISSING_FILE, encoding='utf-8'))
    to_translate = [p for p in missing if p.strip() not in cache]
    print(f"Missing phrases: {len(missing)}, need API calls: {len(to_translate)}")

    # Translate
    count = 0
    try:
        for idx, phrase in enumerate(to_translate, 1):
            clean = phrase.strip()
            trans = translate_phrase(clean)
            if trans:
                cache[clean] = trans
            else:
                cache[clean] = clean  # fallback
            count += 1
            if idx % 50 == 0 or idx == len(to_translate):
                elapsed_pct = idx / len(to_translate) * 100
                print(f"  Translated {idx}/{len(to_translate)} ({elapsed_pct:.0f}%)...")
                json.dump(cache, open(CACHE_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
                time.sleep(0.2)
            else:
                time.sleep(0.05)
    except KeyboardInterrupt:
        print("Interrupted! Saving cache...")
    finally:
        json.dump(cache, open(CACHE_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print(f"Cache saved: {len(cache)} total entries")

    # Now rebuild XLAM with full cache
    print("\n--- Rebuilding RelaxTools_EN.xlam ---")

    src_xlam = 'RelaxTools.xlam'
    dst_xlam = 'RelaxTools_EN.xlam'

    src_zf = zipfile.ZipFile(src_xlam, 'r')

    # 1. Translate customUI
    raw_custom = src_zf.read('customUI/customUI.xml').decode('utf-8')
    # Update callbacks
    en_custom = raw_custom.replace('RelaxTools.xlam!', 'RelaxTools_EN.xlam!')
    # Translate attributes
    attr_re = re.compile(r'\b(label|title|supertip|description|screentip)="([^"]*)"', re.IGNORECASE)
    def replace_attr(m):
        attr_name, val = m.group(1), m.group(2)
        if not JP_RE.search(val):
            return m.group(0)
        cleaned = val.replace('&#xD;&#xA;', '\n').strip()
        trans = cache.get(cleaned) or cache.get(val.strip())
        if trans and trans != cleaned:
            escaped = saxutils.escape(trans, {'"': '&quot;'})
            if '&#xD;&#xA;' in val and '\n' in trans:
                escaped = escaped.replace('\n', '&#xD;&#xA;')
            return f'{attr_name}="{escaped}"'
        return m.group(0)
    en_custom = attr_re.sub(replace_attr, en_custom)
    ET.fromstring(en_custom)
    print("  customUI.xml translated and valid.")

    # 2. Translate sharedStrings - handle ALL <t> elements including inside <r> runs
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    raw_ss = src_zf.read('xl/sharedStrings.xml')
    tree = ET.fromstring(raw_ss)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    replaced = 0
    for si in tree.findall('main:si', ns):
        # Case 1: Simple <si><t>text</t>...</si>
        main_t = si.find('main:t', ns)
        if main_t is not None and main_t.text and JP_RE.search(main_t.text):
            clean_t = main_t.text.strip()
            trans = cache.get(clean_t) or cache.get(main_t.text) or cache.get(clean_t.replace('\r\n', '\n'))
            if trans and trans != clean_t:
                main_t.text = trans
                replaced += 1
                # Remove phonetic ruby tags
                for rph in list(si.findall('main:rPh', ns)):
                    si.remove(rph)
            continue

        # Case 2: Rich text <si><r><t>text</t></r>...</si>
        r_els = si.findall('main:r', ns)
        if r_els:
            full_text = ''.join([(t.text or '') for r in r_els for t in r.findall('main:t', ns)])
            if JP_RE.search(full_text):
                clean_t = full_text.strip()
                trans = cache.get(clean_t) or cache.get(full_text)
                if trans and trans != clean_t:
                    # Replace rich text with simple text
                    for r_el in list(r_els):
                        si.remove(r_el)
                    for rph in list(si.findall('main:rPh', ns)):
                        si.remove(rph)
                    new_t = ET.SubElement(si, f'{{{ns["main"]}}}t')
                    new_t.text = trans
                    replaced += 1

    print(f"  Replaced {replaced} entries in sharedStrings.xml")

    out = io.BytesIO()
    out.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    ET.ElementTree(tree).write(out, encoding='UTF-8', xml_declaration=False)
    new_ss = out.getvalue()

    # 3. Package
    with zipfile.ZipFile(dst_xlam, 'w', compression=zipfile.ZIP_DEFLATED) as dst_zf:
        for item in src_zf.infolist():
            if item.filename == 'customUI/customUI.xml':
                dst_zf.writestr(item, en_custom.encode('utf-8'))
            elif item.filename == 'xl/sharedStrings.xml':
                dst_zf.writestr(item, new_ss)
            else:
                dst_zf.writestr(item, src_zf.read(item.filename))

    size_mb = os.path.getsize(dst_xlam) / 1024 / 1024
    print(f"  Built {dst_xlam} ({size_mb:.2f} MB)")

    # Verify remaining Japanese
    zf2 = zipfile.ZipFile(dst_xlam, 'r')
    remaining_custom = len(re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]', zf2.read('customUI/customUI.xml').decode('utf-8')))
    remaining_ss = len([1 for si in ET.fromstring(zf2.read('xl/sharedStrings.xml')).findall('main:si', ns)
                       if JP_RE.search(''.join([(t.text or '') for t in si.findall('.//main:t', ns)]))])
    print(f"\n  Remaining Japanese chars in customUI: {remaining_custom}")
    print(f"  Remaining Japanese shared strings: {remaining_ss}")

if __name__ == '__main__':
    main()
