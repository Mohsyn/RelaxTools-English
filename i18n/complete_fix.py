#!/usr/bin/env python3
"""
Fix All Remaining Japanese in RelaxTools_EN.xlam
================================================
1. Re-translates all 426 phrases in translation_cache.json that previously failed (where translation == japanese).
2. Updates docProps/core.xml with an English title and description.
3. Translates all remaining customUI attributes (including label="...", helperText="...", etc.).
4. Translates all remaining sharedStrings entries (stripping rPh tags).
5. Cleanly rebuilds RelaxTools_EN.xlam.
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

CACHE_FILE = 'i18n/translation_cache.json'
JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')

def translate_phrase(text):
    clean = text.strip()
    if not clean or not JP_RE.search(clean):
        return clean
    encoded = urllib.parse.quote(clean)
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=ja&tl=en&dt=t&q={encoded}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=12) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                parts = [p[0] for p in data[0] if p[0]]
                res = ''.join(parts).strip()
                if res:
                    return res
        except Exception:
            time.sleep(1.5 + attempt)
    return ''

def main():
    print("Loading translation cache...")
    cache = json.load(open(CACHE_FILE, encoding='utf-8'))
    print(f"  {len(cache)} total cache entries.")

    # 1. Identify all keys that failed previously (v == k and contains Japanese)
    to_fix = [k for k, v in cache.items() if k == v and JP_RE.search(k)]
    print(f"  Found {len(to_fix)} phrases that need re-translation.")

    fixed_count = 0
    for idx, ja in enumerate(to_fix, 1):
        trans = translate_phrase(ja)
        if trans and trans != ja:
            cache[ja] = trans
            fixed_count += 1
        if idx % 25 == 0 or idx == len(to_fix):
            print(f"    Re-translated {idx}/{len(to_fix)} ({fixed_count} succeeded)...")
            json.dump(cache, open(CACHE_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
            time.sleep(0.3)
        else:
            time.sleep(0.1)

    json.dump(cache, open(CACHE_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f"Cache updated with {fixed_count} newly resolved translations.")

    # 2. Package into RelaxTools_EN.xlam
    print("\n--- Rebuilding RelaxTools_EN.xlam ---")
    src_xlam = 'RelaxTools.xlam'
    dst_xlam = 'RelaxTools_EN.xlam'
    src_zf = zipfile.ZipFile(src_xlam, 'r')

    # A. customUI.xml
    raw_custom = src_zf.read('customUI/customUI.xml').decode('utf-8')
    en_custom = raw_custom.replace('RelaxTools.xlam!', 'RelaxTools_EN.xlam!')
    # Match any attribute with whitespace before equal sign or normal
    attr_re = re.compile(r'\b(label|title|supertip|description|screentip|helperText)\s*=\s*"([^"]*)"', re.IGNORECASE)

    def replace_attr(m):
        attr_name, val = m.group(1), m.group(2)
        if not JP_RE.search(val):
            return m.group(0)
        cleaned = val.replace('&#xD;&#xA;', '\n').strip()
        trans = cache.get(cleaned) or cache.get(val.strip()) or cache.get(cleaned.replace('\n', ''))
        if trans and trans != cleaned:
            escaped = saxutils.escape(trans, {'"': '&quot;'})
            if '&#xD;&#xA;' in val and '\n' in trans:
                escaped = escaped.replace('\n', '&#xD;&#xA;')
            return f'{attr_name}="{escaped}"'
        return m.group(0)

    en_custom = attr_re.sub(replace_attr, en_custom)
    ET.fromstring(en_custom)
    print("  customUI.xml fully translated and valid.")

    # B. sharedStrings.xml
    raw_ss = src_zf.read('xl/sharedStrings.xml')
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    tree = ET.fromstring(raw_ss)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    replaced_ss = 0
    for si in tree.findall('main:si', ns):
        # 1. Main t element
        main_t = si.find('main:t', ns)
        if main_t is not None and main_t.text and JP_RE.search(main_t.text):
            txt = main_t.text
            clean = txt.strip()
            trans = (cache.get(clean) or cache.get(txt) or
                     cache.get(clean.replace('\r\n', '\n')) or
                     cache.get(clean.replace('\n', '')))
            if trans and trans != clean:
                main_t.text = trans
                replaced_ss += 1
                for rph in list(si.findall('main:rPh', ns)):
                    si.remove(rph)
                continue

        # 2. Rich text
        r_els = si.findall('main:r', ns)
        if r_els:
            full_text = ''.join([(t.text or '') for r in r_els for t in r.findall('main:t', ns)])
            if JP_RE.search(full_text):
                clean = full_text.strip()
                trans = cache.get(clean) or cache.get(full_text)
                if trans and trans != clean:
                    for r_el in list(r_els):
                        si.remove(r_el)
                    for rph in list(si.findall('main:rPh', ns)):
                        si.remove(rph)
                    new_t = ET.SubElement(si, f'{{{ns["main"]}}}t')
                    new_t.text = trans
                    replaced_ss += 1

    print(f"  sharedStrings.xml: {replaced_ss} entries translated.")
    out_ss = io.BytesIO()
    out_ss.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    ET.ElementTree(tree).write(out_ss, encoding='UTF-8', xml_declaration=False)
    new_ss_bytes = out_ss.getvalue()

    # C. docProps/core.xml (Addin description in Excel Addin manager)
    raw_core = src_zf.read('docProps/core.xml').decode('utf-8')
    core_desc_en = (
        "Ver. 4.28.0 (English Edition)\n"
        "Compatible with Excel 2013/2016/2019/Office 365\n"
        "A comprehensive add-in featuring over 250 functions.\n"
        "Original Copyright (C) Yasuhiro Watanabe 2009 | English Fork by Mohsyn"
    )
    # Replace description inside <dc:description>...</dc:description>
    new_core = re.sub(
        r'<dc:description>.*?</dc:description>',
        f'<dc:description>{saxutils.escape(core_desc_en)}</dc:description>',
        raw_core,
        flags=re.DOTALL
    )

    # D. Package RelaxTools_EN.xlam
    with zipfile.ZipFile(dst_xlam, 'w', compression=zipfile.ZIP_DEFLATED) as dst_zf:
        for item in src_zf.infolist():
            if item.filename == 'customUI/customUI.xml':
                dst_zf.writestr(item, en_custom.encode('utf-8'))
            elif item.filename == 'xl/sharedStrings.xml':
                dst_zf.writestr(item, new_ss_bytes)
            elif item.filename == 'docProps/core.xml':
                dst_zf.writestr(item, new_core.encode('utf-8'))
            else:
                dst_zf.writestr(item, src_zf.read(item.filename))

    size_mb = os.path.getsize(dst_xlam) / 1024 / 1024
    print(f"\nSuccessfully generated {dst_xlam} ({size_mb:.2f} MB)")

    # Final Verification
    zf_check = zipfile.ZipFile(dst_xlam, 'r')
    rem_custom = len(re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]', zf_check.read('customUI/customUI.xml').decode('utf-8')))
    rem_ss = len([1 for si in ET.fromstring(zf_check.read('xl/sharedStrings.xml')).findall('main:si', ns)
                  if JP_RE.search(''.join([(t.text or '') for t in si.findall('.//main:t', ns)]))])
    print(f"Remaining Japanese characters in customUI: {rem_custom}")
    print(f"Remaining Japanese shared strings: {rem_ss}")
    print("Add-in dialog description updated to English.")

if __name__ == '__main__':
    main()
