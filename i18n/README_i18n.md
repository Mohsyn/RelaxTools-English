# RelaxTools-Addin: repeatable Japanese -> English localisation kit

Put this folder in the root of your fork as `i18n/` (next to `Source/`, `README.md`, ...).

## The idea

Upstream files are **never edited**. All translations live in one table, `translations.csv`, keyed by the
exact Japanese text. `rlx_i18n.py apply` writes a translated copy of the tree into `build_en/`.
When upstream changes: merge it as usual (no conflicts, you never touched those files), run `extract`
(it lists only new/changed strings), translate those, `apply`. Nothing you translated is lost.

## Files

| file | purpose |
|---|---|
| `rlx_i18n.py` | the pipeline (Python 3, standard library only) |
| `translations.csv` | translation memory. Columns: id, scope, hint, status, source, translation, count, first_location, context, note |
| `line_overrides.csv` | whole-line replacements for the few lines where English needs a different word order or the Japanese must stay (keyed by the exact original line text, so line-number shifts don't matter) |
| `identifiers.csv` | Japanese identifiers in code (e.g. a parameter name) -> English |
| `rlx_build.vbs` | **untested** Excel automation: imports modules into a copy of the xlam, dumps/applies form-control texts |
| `frx_strings.tsv` | (created later) form-control texts dumped from Excel |

Statuses: `new` = needs translating, `done`, `review`, `keep` = functional Japanese, do not translate,
`skip` = not translated by default (code comments), `obsolete` = no longer in the source (kept for reuse).

## Everyday commands (run from the repo root; global options go before the command)

```
python i18n/rlx_i18n.py extract        # after every upstream merge: finds new/changed/obsolete strings
python i18n/rlx_i18n.py status         # progress
python i18n/rlx_i18n.py todo --limit 50 [--scope vba|xml|text] [--hint ui|const|frx|sheet|...]
python i18n/rlx_i18n.py apply          # -> build_en/   (default --encoding ascii-safe)
python i18n/rlx_i18n.py check          # validates translations + proves only strings/comments changed
python i18n/rlx_i18n.py sheetnames     # worksheet tabs inside the xlam that must be renamed
```
Edit `translations.csv` in any editor/Excel (UTF-8). Rules the tool applies for you:
* Quotes are un-escaped in the table; the tool re-escapes (`"` -> `""` in VBA, `&quot;` in XML).
* Leading/trailing spaces: if your translation has its own, they are kept (English needs spaces where Japanese had
  none, e.g. `" item(s) deleted."`); otherwise the source's edge spaces are copied.
* Many messages are built as `"text" & value & "text"`. Each fragment is its own row, so translate the fragments so that
  they read correctly when joined. If the word order cannot work, use `line_overrides.csv`.
* Two-line ribbon labels: write `\n` for the line break (CRLF is restored automatically).

`ascii-safe` mode makes every generated VBA file pure ASCII: the Japanese that must stay is written as `ChrW()`
expressions, constants that contain Japanese become parameterless Functions (module level) or local variables,
Japanese comments become `?`. Result: the files import correctly on any Windows language/locale.

## Building the English xlam

1. `python i18n/rlx_i18n.py apply` -> `build_en/`
2. Copy the original `RelaxTools.xlam` (from the Japanese release) somewhere. In Excel enable
   *Trust access to the VBA project object model*, close all Excel windows.
3. `cscript //nologo i18n\rlx_build.vbs import RelaxTools.xlam RelaxTools_stage.xlam build_en\Source\src`
4. **Form controls (button/label/checkbox texts live inside the binary .frx files):**
   `cscript //nologo i18n\rlx_build.vbs dump RelaxTools_stage.xlam frx_dump.txt` then
   `python i18n/rlx_i18n.py import-frx frx_dump.txt`, translate the new rows (`todo --hint frx`),
   `python i18n/rlx_i18n.py apply`, then
   `cscript //nologo i18n\rlx_build.vbs apply-forms RelaxTools_stage.xlam build_en\frx_translations.txt`
   (run the dump on the *Japanese* original if you want Japanese sources; the table is value-based, so either works).
5. Ribbon: `python i18n/rlx_i18n.py patch-xlam RelaxTools_stage.xlam RelaxTools_en.xlam`
6. Rename the template worksheets listed by `sheetnames` inside the xlam (Excel: unhide the add-in sheets via
   VBE properties -> Visible = -1, rename tabs, hide again).

## Known limits (be aware)

* `rlx_build.vbs` and the form-text step were written without Excel available: untested. Work on copies.
* Cell contents of sheets inside the xlam (menu definition tables, category names, templates) are not in `Source/`
  and are not translated. Several functions compare against those Japanese category names (kept Japanese on purpose).
* `keyConfig/*.key`: only the `#` comment lines are translated; data lines stay (their category column must match the
  xlam data sheets).
* Code comments (1,609 unique) are not translated (`skip`). Set status to `new`, translate, and they are used.
* Locale-dependent behaviour is not changed (e.g. Japanese date format `ggge`, half/full-width conversion, font
  "MS Gothic" JP name checks) - these are features of the add-in, not UI text.
* Strings whose length the code relies on were kept length-compatible (`(Ref)` = 5 characters, see `C_REF_TEXT`).
  If you change translations of `hint=const`/`func?` rows, check how the code uses them first.
