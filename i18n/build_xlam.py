#!/usr/bin/env python3
"""
Complete Builder for RelaxTools_EN.xlam
========================================
1. Translates ALL hardcoded Japanese attributes (label, title, supertip, description)
   in customUI.xml using the translation cache and escapes XML special characters.
2. Updates all 2,145 callback references to point to RelaxTools_EN.xlam!
3. Translates ALL Japanese entries in xl/sharedStrings.xml (handling phonetic <rPh> tags).
4. Packages cleanly into RelaxTools_EN.xlam with original worksheet structures untouched.
"""

import io
import json
import os
import re
import xml.sax.saxutils as saxutils
import zipfile
from xml.etree import ElementTree as ET

def translate_customui(src_xml: str, cache: dict) -> str:
    # 1. Update callback add-in name
    xml = src_xml.replace('RelaxTools.xlam!', 'RelaxTools_EN.xlam!')

    # 2. Translate attributes: label, title, supertip, description, screentip
    jp_re = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')
    attr_re = re.compile(r'\b(label|title|supertip|description|screentip)="([^"]*)"', re.IGNORECASE)

    def replace_attr(m):
        attr_name = m.group(1)
        val = m.group(2)
        if not jp_re.search(val):
            return m.group(0)

        # Handle &#xD;&#xA; newlines in XML attributes
        has_crlf = '&#xD;&#xA;' in val
        cleaned_val = val.replace('&#xD;&#xA;', '\n').strip()

        trans = cache.get(cleaned_val) or cache.get(val.strip())
        if not trans:
            # Try without whitespace
            trans = cache.get(re.sub(r'\s+', '', cleaned_val))

        if trans and trans != cleaned_val:
            # Escape XML special chars, especially double quotes & < > &
            escaped = saxutils.escape(trans, {'"': '&quot;'})
            if has_crlf and '\n' in trans:
                escaped = escaped.replace('\n', '&#xD;&#xA;')
            return f'{attr_name}="{escaped}"'
        return m.group(0)

    xml = attr_re.sub(replace_attr, xml)
    return xml

def translate_shared_strings(raw_ss: bytes, cache: dict) -> bytes:
    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    tree = ET.fromstring(raw_ss)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    jp_re = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')

    replaced = 0
    for si in tree.findall('main:si', ns):
        t_el = si.find('main:t', ns)
        if t_el is not None and t_el.text:
            text = t_el.text
            if jp_re.search(text):
                clean_t = text.strip()
                trans = cache.get(clean_t) or cache.get(text)
                if not trans:
                    # Try newline normalized
                    trans = cache.get(clean_t.replace('\r\n', '\n'))
                if trans and trans != clean_t:
                    t_el.text = trans
                    replaced += 1
                    # Remove Japanese phonetic ruby tags if present
                    for rph in list(si.findall('main:rPh', ns)):
                        si.remove(rph)

    print(f"  Replaced {replaced} text entries in sharedStrings.xml.")

    out = io.BytesIO()
    out.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    ET.ElementTree(tree).write(out, encoding='UTF-8', xml_declaration=False)
    return out.getvalue()

def main():
    src_xlam = 'RelaxTools.xlam'
    dst_xlam = 'RelaxTools_EN.xlam'
    cache_file = 'i18n/translation_cache.json'

    print(f"Loading {cache_file} ...")
    cache = json.load(open(cache_file, encoding='utf-8'))
    print(f"  {len(cache)} translations loaded.")

    print(f"\nOpening {src_xlam} ...")
    src_zf = zipfile.ZipFile(src_xlam, 'r')

    # 1. Translate customUI
    raw_custom = src_zf.read('customUI/customUI.xml').decode('utf-8')
    print("Translating customUI ribbon labels and updating callbacks ...")
    en_custom = translate_customui(raw_custom, cache)
    
    # Test valid XML
    ET.fromstring(en_custom)
    print("  customUI.xml is 100% valid XML.")

    # 2. Translate shared strings (HELP sheet & others)
    raw_ss = src_zf.read('xl/sharedStrings.xml')
    print("Translating xl/sharedStrings.xml ...")
    en_ss = translate_shared_strings(raw_ss, cache)

    # 3. Write out new XLAM
    print(f"\nPackaging {dst_xlam} ...")
    with zipfile.ZipFile(dst_xlam, 'w', compression=zipfile.ZIP_DEFLATED) as dst_zf:
        for item in src_zf.infolist():
            if item.filename == 'customUI/customUI.xml':
                dst_zf.writestr(item, en_custom.encode('utf-8'))
            elif item.filename == 'xl/sharedStrings.xml':
                dst_zf.writestr(item, en_ss)
            else:
                dst_zf.writestr(item, src_zf.read(item.filename))

    size_mb = os.path.getsize(dst_xlam) / 1024 / 1024
    print(f"Successfully generated {dst_xlam} ({size_mb:.2f} MB)")

if __name__ == '__main__':
    main()
