#!/usr/bin/env python3
"""
UserForm Caption Translator for RelaxTools i18n
===============================================
Translates missing UserForm control captions from i18n/userform_missing_captions.json,
applies Excel UI terminology overrides, and updates translation_cache.json & strings_raw.csv.
"""

import csv
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CACHE_FILE = Path("i18n/translation_cache.json")
CSV_PATH = Path("i18n/strings_raw.csv")
MISSING_FILE = Path("i18n/userform_missing_captions.json")

# Domain-specific UI glossary for standard Windows/Excel dialog terms
GLOSSARY = {
    "保存": "Save",
    "キャンセル": "Cancel",
    "閉じる": "Close",
    "終了": "Exit",
    "設定": "Settings",
    "実行": "Execute",
    "全選択": "Select All",
    "全解除": "Deselect All",
    "有効": "Enabled",
    "無効": "Disabled",
    "追加": "Add",
    "削除": "Delete",
    "編集": "Edit",
    "更新": "Update",
    "参照": "Browse...",
    "取得": "Get",
    "フォルダを選択": "Select Folder",
    "フォルダ選択": "Select Folder",
    "ヘルプ": "Help",
    "クリア": "Clear",
    "リセット": "Reset",
    "デフォルト": "Default",
    "初期値": "Default Value",
    "OK": "OK",
    "はい": "Yes",
    "いいえ": "No",
    "行": "Row",
    "列": "Column",
    "シート": "Sheet",
    "ブック": "Workbook",
    "セル": "Cell",
    "数値": "Numeric",
    "条件": "Condition",
    "オプション": "Options",
    "フォント": "Font",
    "詳細": "Details",
    "履歴": "History",
    "置換": "Replace",
    "検索": "Search",
    "インポート": "Import",
    "エクスポート": "Export",
    "上へ": "Move Up",
    "下へ": "Move Down",
    "プレビュー": "Preview",
    "適用": "Apply",
    "コピー": "Copy",
    "貼り付け": "Paste",
    "切り取り": "Cut",
}

JP_RE = re.compile(r'[\u3040-\u30FF\u4E00-\u9FFF]')

def translate_phrase(text: str) -> str:
    clean = text.strip()
    if not clean or not JP_RE.search(clean):
        return clean
    if clean in GLOSSARY:
        return GLOSSARY[clean]

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
            time.sleep(1.0 + attempt * 1.5)
    return ""

def main():
    if not MISSING_FILE.exists():
        print(f"Error: {MISSING_FILE} not found.")
        sys.exit(1)

    missing_dict = json.loads(MISSING_FILE.read_text(encoding="utf-8"))
    cache = json.loads(CACHE_FILE.read_text(encoding="utf-8")) if CACHE_FILE.exists() else {}

    print(f"Loaded {len(cache)} existing cache items.")
    print(f"Loaded {len(missing_dict)} missing captions to translate.")

    to_translate = [k for k in missing_dict.keys() if k not in cache or cache[k] == k or not cache[k]]
    print(f"Phrases needing translation: {len(to_translate)}")

    translated_count = 0
    for idx, ja in enumerate(to_translate, 1):
        en = translate_phrase(ja)
        if en:
            cache[ja] = en
            missing_dict[ja] = en
            translated_count += 1
        else:
            print(f"  [WARN] Failed to translate: {repr(ja)}")

        if idx % 20 == 0 or idx == len(to_translate):
            print(f"  Progress: {idx}/{len(to_translate)} ({translated_count} translated)...")
            CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
            MISSING_FILE.write_text(json.dumps(missing_dict, ensure_ascii=False, indent=2), encoding="utf-8")
            time.sleep(0.2)
        else:
            time.sleep(0.08)

    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    MISSING_FILE.write_text(json.dumps(missing_dict, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nDone! Translated {translated_count}/{len(to_translate)} captions.")

if __name__ == "__main__":
    main()
