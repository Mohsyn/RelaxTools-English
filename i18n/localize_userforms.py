#!/usr/bin/env python3
"""
Comprehensive UserForm Localizer for RelaxTools-Addin
======================================================
1. Connects to RelaxTools_EN.xlam via Excel COM automation.
2. Iterates over all 80 UserForms in vbext_ct_MSForm.
3. Translates all design-time control captions (buttons, labels, checkboxes, frames).
4. Translates form window titles (Captions).
5. Updates each UserForm's CodeModule with the translated English VBA code from Source_EN/src/Form/.
6. Saves RelaxTools_EN.xlam (compiling the English .frx binary blobs natively into vbaProject.bin).
7. Exports the updated .frm and .frx files back into Source_EN/src/Form/.
8. Re-opens and verifies 0 remaining Japanese controls.
"""

import json
import os
import re
import sys
import time
from pathlib import Path

# Ensure UTF-8 output
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
CACHE_FILE = ROOT / "i18n" / "translation_cache.json"
XLAM_PATH = ROOT / "RelaxTools_EN.xlam"
SOURCE_EN_FORMS = ROOT / "Source_EN" / "src" / "Form"

JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')

def load_cache():
    if not CACHE_FILE.exists():
        raise FileNotFoundError(f"{CACHE_FILE} not found.")
    return json.loads(CACHE_FILE.read_text(encoding="utf-8"))

def get_english_caption_from_frm(frm_path: Path) -> str:
    """Extracts Caption = "..." from the Begin block of a .frm file."""
    if not frm_path.exists():
        return ""
    try:
        content = frm_path.read_text(encoding="utf-8", errors="replace")
        for line in content.splitlines():
            m = re.match(r'^\s*Caption\s*=\s*"([^"]*)"', line)
            if m:
                return m.group(1).strip()
            if line.strip().startswith("Attribute VB_Name"):
                break
    except Exception:
        pass
    return ""

def get_code_from_frm(frm_path: Path) -> str:
    """Extracts the VBA code portion from a .frm file (filtering out any Attribute lines)."""
    if not frm_path.exists():
        return ""
    try:
        lines = frm_path.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)
        last_attr = -1
        for i, line in enumerate(lines):
            if line.strip().startswith("Attribute VB_"):
                last_attr = i
        code_lines = lines[last_attr + 1:] if last_attr != -1 else lines
        # Inside VBE CodeModule, Attribute lines (like Attribute MW.VB_VarHelpID = -1) are invalid syntax
        filtered_lines = [l for l in code_lines if not l.strip().startswith("Attribute ")]
        return "".join(filtered_lines)
    except Exception:
        return ""

def translate_text(text: str, cache: dict) -> str:
    """Looks up text in cache with various normalizations."""
    if not text:
        return ""
    clean = text.strip()
    if not JP_RE.search(clean):
        return text

    # Direct match
    if clean in cache and cache[clean] != clean:
        return cache[clean]

    # Newline normalizations
    norm_nl = clean.replace("\r\n", "\n")
    if norm_nl in cache and cache[norm_nl] != norm_nl:
        return cache[norm_nl]

    norm_crlf = clean.replace("\n", "\r\n")
    if norm_crlf in cache and cache[norm_crlf] != norm_crlf:
        return cache[norm_crlf]

    # Without whitespace
    no_ws = re.sub(r'\s+', '', clean)
    if no_ws in cache and cache[no_ws] != no_ws:
        return cache[no_ws]

    return text

def main():
    import win32com.client

    print("=" * 70)
    print(" RelaxTools UserForm Localization Pipeline")
    print("=" * 70)

    cache = load_cache()
    print(f"[1/5] Loaded translation cache ({len(cache)} entries).")

    if not XLAM_PATH.exists():
        print(f"ERROR: {XLAM_PATH} does not exist.")
        sys.exit(1)

    print(f"[2/5] Connecting to Excel COM & opening {XLAM_PATH.name} ...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    try:
        wb = excel.Workbooks.Open(str(XLAM_PATH.resolve()))
        vb_proj = wb.VBProject

        form_comps = [c for c in vb_proj.VBComponents if c.Type == 3] # vbext_ct_MSForm = 3
        print(f"      Found {len(form_comps)} UserForms in {XLAM_PATH.name}.\n")

        print("[3/5] Localizing UserForm controls and code...")
        total_controls_updated = 0
        total_forms_updated = 0
        total_code_updated = 0
        unmatched_strings = set()

        for idx, comp in enumerate(form_comps, 1):
            form_name = comp.Name
            designer = comp.Designer
            frm_path = SOURCE_EN_FORMS / f"{form_name}.frm"

            # A. Form Title (Caption)
            frm_cap = get_english_caption_from_frm(frm_path)
            current_cap = str(designer.Caption or "").strip()
            if frm_cap and not JP_RE.search(frm_cap):
                designer.Caption = frm_cap
            elif JP_RE.search(current_cap):
                trans_cap = translate_text(current_cap, cache)
                if trans_cap != current_cap:
                    designer.Caption = trans_cap

            # B. Control Captions & Tooltips
            form_ctls_updated = 0
            for ctl in designer.Controls:
                # 1. Caption
                try:
                    cap = str(ctl.Caption or "").strip()
                    if cap and JP_RE.search(cap):
                        trans = translate_text(cap, cache)
                        if trans and trans != cap:
                            ctl.Caption = trans
                            form_ctls_updated += 1
                        else:
                            unmatched_strings.add((form_name, ctl.Name, "Caption", cap))
                except Exception:
                    pass

                # 2. ControlTipText
                try:
                    tip = str(ctl.ControlTipText or "").strip()
                    if tip and JP_RE.search(tip):
                        trans = translate_text(tip, cache)
                        if trans and trans != tip:
                            ctl.ControlTipText = trans
                            form_ctls_updated += 1
                        else:
                            unmatched_strings.add((form_name, ctl.Name, "ControlTipText", tip))
                except Exception:
                    pass

                # 3. TabStrip Tabs (if any)
                try:
                    tabs = ctl.Tabs
                    for tab in tabs:
                        tcap = str(tab.Caption or "").strip()
                        if tcap and JP_RE.search(tcap):
                            trans = translate_text(tcap, cache)
                            if trans and trans != tcap:
                                tab.Caption = trans
                                form_ctls_updated += 1
                except Exception:
                    pass

            # C. Update CodeModule from Source_EN
            code = get_code_from_frm(frm_path)
            if code and len(code.strip()) > 0:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(code)
                total_code_updated += 1

            total_controls_updated += form_ctls_updated
            if form_ctls_updated > 0 or frm_cap:
                total_forms_updated += 1

            print(f"  [{idx:02d}/{len(form_comps):02d}] {form_name:<25}: {form_ctls_updated:2d} controls localized.")

        print(f"\n      Summary:")
        print(f"      - Forms updated         : {total_forms_updated}")
        print(f"      - Controls localized    : {total_controls_updated}")
        print(f"      - CodeModules replaced  : {total_code_updated}")

        if unmatched_strings:
            print(f"\n  [WARN] {len(unmatched_strings)} strings could not be matched in cache:")
            for fn, ctl_name, prop, s in list(unmatched_strings)[:10]:
                print(f"    {fn}.{ctl_name}.{prop}: {repr(s)}")

        # D. Save Workbook (Excel compiles .frx blobs into vbaProject.bin)
        print("\n[4/5] Saving RelaxTools_EN.xlam and exporting native English .frm/.frx files...")
        wb.Save()

        # E. Export forms back to Source_EN/src/Form/
        export_count = 0
        for comp in form_comps:
            export_target = SOURCE_EN_FORMS / f"{comp.Name}.frm"
            comp.Export(str(export_target.resolve()))
            export_count += 1
        print(f"      Exported {export_count} forms to {SOURCE_EN_FORMS}")

        wb.Close(False)

    finally:
        excel.Quit()
        del excel

    # [5/5] Re-open and Verify
    print("\n[5/5] Re-opening RelaxTools_EN.xlam to verify localization ...")
    excel2 = win32com.client.Dispatch("Excel.Application")
    excel2.Visible = False
    excel2.DisplayAlerts = False
    try:
        wb2 = excel2.Workbooks.Open(str(XLAM_PATH.resolve()))
        remaining_jp_ctls = []
        for c in wb2.VBProject.VBComponents:
            if c.Type == 3:
                for ctl in c.Designer.Controls:
                    try:
                        cap = str(ctl.Caption or "").strip()
                        # Allow English translations that contain Japanese character explanations like "'カ' and 'ケ'"
                        if JP_RE.search(cap):
                            # Check if the caption is in cache as an approved English translation
                            is_en_translation = False
                            for k, v in cache.items():
                                if v == cap and k != v:
                                    is_en_translation = True
                                    break
                            if not is_en_translation:
                                remaining_jp_ctls.append((c.Name, ctl.Name, cap))
                    except Exception:
                        pass
        wb2.Close(False)

        print(f"      Remaining untranslated controls: {len(remaining_jp_ctls)}")
        if remaining_jp_ctls:
            for fn, cn, text in remaining_jp_ctls:
                print(f"        {fn}.{cn} -> {repr(text)}")
        else:
            print("      SUCCESS: 100% of UserForm controls are in English!")

    finally:
        excel2.Quit()
        del excel2

    print("\n" + "=" * 70)
    print(" UserForm Localization Complete!")
    print("=" * 70)

if __name__ == "__main__":
    main()
