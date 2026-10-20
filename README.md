<p align="center">
  <img src="https://raw.githubusercontent.com/wiki/RelaxTools/RelaxTools-Addin/image/RelaxChan.png" />  
  <h2 align="center">RelaxTools Addin For Excel 2013/2016/2019/Office365 (English Version)</h2>
  <p align="center">A versatile Excel add-in for conversion, editing, searching, and advanced formatting utilities.</p>
  <p align="center">Original Copyright (C) Yasuhiro Watanabe 2009 | English Fork by Mohsyn</p>
</p>

---

* **Supported OS**: Windows 10 / Windows 11 (macOS is not supported)
* **Supported Excel**: 2013, 2016, 2019, Office365 (Desktop)

This add-in systematically aggregates and refines macros developed across several years into a comprehensive toolkit containing **over 250 functions**.

### Key Features
* **Excel Grep & Replace**: Fast search and replacement within workbooks and external files.
* **HTML Table Export**: Convert Excel tables directly to HTML (ideal for using Excel as a UI/layout designer).
* **Regex Search & Replace**: Regular expression search and replace within cells or shape text.
* **2007-compatible print preview** screen support.
* **Open Duplicate Books**: Open workbooks with identical filenames (read-only mode).
* **Data Verification & Checking**:
  * Check for duplicate cells, character counts, and byte length.
  * Validate Half-width / Full-width / Numeric / Alphabetic / Alphanumeric / Date formatting and check digits (Luhn, Modulus 10/11, Corporate number, etc.).
* **Text Manipulation**:
  * Trim leading/trailing whitespace, insert/delete characters, remove line breaks.
  * Case/character conversion: Uppercase, Lowercase, Hiragana, Katakana, Full-width, Half-width.
* **Table & Structure Tools**:
  * 1-click row/column transpose, vertical merge, and advanced grid generation.
  * Convert Excel tables into ASCII/Unicode text border representations.
* **Shapes & Visual Aids**:
  * Highlight shapes (rectangles, ellipses, arrows, callouts).
  * Shape alignment utilities (straighten arrows, equalize dimensions).
  * Cross-line cursor for large data navigation.
  * Sticky notes (Fusen) creation.
* **Workbook & Sheet Management**:
  * Copy book path/name to clipboard, open current workbook folder in Explorer.
  * Sheet manager: sort sheet names, batch select, show/hide, bulk delete.
  * Split workbook by worksheet, merge multiple workbooks into one.
  * Auto-reset cursor to **A1 (Home position)** upon saving.
  * Lightweight workbook version/generation management.
* **Developer & System Tools**:
  * SQL formatter & External editor integration for cell content.
  * Export embedded images to files.
  * Export file list / file tree structure into spreadsheets.
  * Generate folder structures from cell text lists.
  * Version control integration: quick Commit/Revert with TortoiseSVN and TortoiseGit.
  * Page counter for Excel/Word/PowerPoint documents; VBA line/step counter.
  * Assign almost any RelaxTools function to custom shortcut keys; disable the F1 key.
  * Auto paragraph numbering (e.g. 1.1.1) and table of contents generation.
  * Digital stamp / Seal utilities (data stamp, approval stamp, business stamp).

---

## File Contents
* `README.md` — This English guide.
* `README_JP.md` — Original Japanese guide.
* `RelaxTools.xlam` — Original RelaxTools add-in package.
* `RelaxTools_EN.xlam` — Generated English version of the add-in.
* `install.vbs` — Installation script.
* `Section_No_Manual.xlsx` — User guide for paragraph numbering features.
* `relaxtools.png` — Icon asset.
* `ExcelReadOnly.vbs` — Adds "Open as Read-Only" to Windows Explorer context menu.
* `ExcelSetHomePosition.vbs` — Utility script to set cursor home positions.
* `keyConfig/` — Recommended shortcut key configuration files.
* `Source/` — Original Japanese source tree (VBA modules, forms, classes, XML).
* `Source_EN/` — Generated English source tree.
* `i18n/` — Automated localization and string tracking pipeline.

---

## Installation / Update Instructions
1. Make sure Microsoft Excel is not running (close Excel if open).
2. Right-click `RelaxTools.xlam` (or `RelaxTools_EN.xlam`), open **Properties**, and check **Unblock** if Windows flagged it as downloaded from the internet.
3. Run `install.vbs`.
4. Launch Excel and verify that the `RelaxTools` tab appears in the ribbon.

### Resetting Settings
* RelaxTools Help → Click **Initialize RelaxTools** (RelaxToolsの初期化).

### Uninstallation
* RelaxTools Help → Click **Uninstall RelaxTools** (RelaxToolsのアンインストール).

---

## Localization & Upstream Sync Pipeline
This repository includes a future-proof, automated localization pipeline in `i18n/`:

* **`i18n/strings_raw.csv`**: Central translation table containing all extracted strings (code comments, string literals, form captions, customUI XML attributes, and XLAM HELP sheet cells).
* **`i18n/extract_strings.py`**: Scans the source tree for Japanese text and merges with existing translations.
* **`i18n/diff_strings.py`**: Compares versions during upstream updates to flag new strings (`i18n/new_strings.csv`).
* **`i18n/apply_translations.py`**: Injects approved translations into `Source_EN/`.
* **`i18n/apply_xlam_help.py`**: Patches the HELP sheet inside the XLAM package.
* **`sync_upstream.ps1`**: Syncs from a fresh upstream release ZIP and flags newly introduced strings.
* **`build_english.ps1`**: Applies all approved translations to generate `Source_EN/` and `RelaxTools_EN.xlam`.

Refer to [`i18n/README.md`](file:///d:/aaa/relax_ag/RelaxTools-Addin-master/i18n/README.md) for full instructions.

---

## License & Credits
* Original Author: **Yasuhiro Watanabe** ([RelaxTools/RelaxTools-Addin](https://github.com/RelaxTools/RelaxTools-Addin))
* English Localization Fork: **Mohsyn** ([Mohsyn/RelaxTools-Addin](https://github.com/Mohsyn/RelaxTools-Addin))
* License: **MIT License** (refer to `LICENSE` file)
