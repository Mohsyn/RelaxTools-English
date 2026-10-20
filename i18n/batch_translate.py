#!/usr/bin/env python3
"""
Batch Translator for RelaxTools i18n
====================================
Translates unique Japanese phrases using online translation API,
caches translations in a local JSON file so it can be resumed anytime,
and updates strings_raw.csv with en_text and status=approved.
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

CACHE_FILE = Path("i18n/translation_cache.json")
CSV_PATH = Path("i18n/strings_raw.csv")

def load_cache():
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_cache(cache):
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

def translate_phrase(text):
    clean_text = text.strip()
    if not clean_text:
        return ""
    
    # Check if text is mostly punctuation or numbers
    if not re.search(r'[\u3040-\u30FF\u4E00-\u9FFF]', clean_text):
        return clean_text

    encoded = urllib.parse.quote(clean_text)
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=ja&tl=en&dt=t&q={encoded}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=10) as response:
                res_data = json.loads(response.read().decode("utf-8"))
                # Combine parts if multi-sentence
                translated_parts = [part[0] for part in res_data[0] if part[0]]
                return "".join(translated_parts).strip()
        except Exception as e:
            time.sleep(1 + attempt)
    return ""

def main():
    cache = load_cache()
    print(f"Loaded {len(cache)} cached translations.")

    if not CSV_PATH.exists():
        print(f"Error: {CSV_PATH} not found.")
        sys.exit(1)

    with open(CSV_PATH, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    print(f"Total rows in CSV: {len(rows)}")

    # We will prioritize UI contexts:
    # Ribbon labels, supertips, descriptions, categories, properties, xml attributes, and string literals
    ui_contexts = {
        'cell:label', 'cell:category', 'cell:supertip', 'cell:description',
        'xml_attr', 'property', 'string_literal', 'cell:no', 'cell:division'
    }

    # Extract all distinct texts needing translation
    to_translate = set()
    for r in rows:
        ja = r["ja_text"].strip()
        if not ja:
            continue
        # If not translated yet, consider for translation
        if not r.get("en_text"):
            to_translate.add(ja)

    print(f"Unique phrases needing translation: {len(to_translate)}")

    # Translate missing items
    needed = [t for t in to_translate if t not in cache]
    print(f"Unique phrases to query from API: {len(needed)}")

    count = 0
    start_time = time.time()
    try:
        for idx, ja in enumerate(needed, 1):
            translated = translate_phrase(ja)
            if translated:
                cache[ja] = translated
            else:
                cache[ja] = ja  # fallback
            count += 1

            if idx % 50 == 0 or idx == len(needed):
                elapsed = time.time() - start_time
                print(f"Translated {idx}/{len(needed)} phrases ({elapsed:.1f}s)...")
                save_cache(cache)
                time.sleep(0.2)
            else:
                time.sleep(0.05)  # slight throttle to avoid rate limits
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. Saving progress...")
    finally:
        save_cache(cache)

    print(f"\nCache has {len(cache)} translations. Applying to CSV rows...")

    applied = 0
    for r in rows:
        ja = r["ja_text"].strip()
        if not r.get("en_text") and ja in cache:
            trans = cache[ja]
            if trans and trans != ja:
                r["en_text"] = trans
                r["status"] = "approved"
                applied += 1
            elif trans == ja and r["context"] in ui_contexts:
                # Text didn't change (e.g. ASCII / numbers)
                r["en_text"] = trans
                r["status"] = "approved"
                applied += 1

    fieldnames = ['key', 'file', 'line', 'context', 'attr_name', 'ja_text', 'en_text', 'status']
    with open(CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    approved = sum(1 for r in rows if r["status"] == "approved")
    pending = sum(1 for r in rows if r["status"] == "pending")
    print(f"Updated CSV: {applied} newly filled. Current totals: {approved} approved, {pending} pending.")

if __name__ == "__main__":
    main()
