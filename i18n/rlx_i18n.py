#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rlx_i18n.py - repeatable Japanese -> English localisation pipeline for RelaxTools-Addin.

IDEA
    The upstream files are NEVER edited.  All translations live in one table
    (translations.csv, keyed by the exact Japanese text).  `apply` generates a
    translated copy of the tree in build_en/.  When upstream changes you merge,
    run `extract` (finds new / changed strings), translate only those, `apply`.

COMMANDS
    extract   scan source, add new strings to the table, mark vanished ones obsolete
    status    progress summary
    todo      print untranslated rows (to translate / review)
    merge F   import translations from a CSV (columns: id,translation[,status])
    apply     write the translated tree to --out
    check     validate the table and the generated tree
    import-frx F  fold the form-control text dump (from rlx_build.vbs) into the table
    patch-xlam    put the translated ribbon XML into a built .xlam
    sheetnames  worksheet names the code looks up (template sheets in the xlam must match)

Only the Python 3 standard library is used.
"""
from __future__ import annotations

import argparse
import collections
import csv
import difflib
import hashlib
import re
import shutil
import sys
import xml.dom.minidom
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_SRC = HERE.parent
DEFAULT_TM = HERE / "translations.csv"
DEFAULT_OVR = HERE / "line_overrides.csv"
DEFAULT_IDS = HERE / "identifiers.csv"
DEFAULT_FRX = HERE / "frx_strings.tsv"
DEFAULT_OUT = HERE.parent / "build_en"

# ---------------------------------------------------------------- what to process
VBA_EXT = {".bas", ".cls", ".frm", ".vbs"}
TEXT_EXT = {".md", ".txt", ".key"}
XML_NAMES = {"customUI.xml"}
XML_UI_ATTRS = {"label", "screentip", "supertip", "description", "title",
                "helperText", "keytip", "alt"}
SKIP_FILES = {"Version.txt"}                   # change-log, not user facing
SKIP_TOP_DIRS = {".git", "i18n", "build_en", "__pycache__"}
CP932_ALWAYS = {".vbs", ".key"}                # read by Windows tools as ANSI

JP = re.compile("[\u3000-\u303f\u3040-\u30ff\u31f0-\u31ff\u3200-\u33ff"
                "\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]")
NONASCII = re.compile(r"[^\x00-\x7f]")
JP_LETTERS = re.compile("[\u3040-\u30fa\u30fc-\u30ff\u3400-\u9fff]")
WS = " \t\u3000"
PUNCT = str.maketrans({"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
                       "\u2026": "...", "\u2013": "-", "\u2014": "-",
                       "\u00a0": " "})
PLACEHOLDER = re.compile(r"%[0-9.]*[sdfx]|\{\d+\}|\$\d|\\[nrt]|%\w+%|<[A-Za-z/][^>]{0,20}>")
FIELDS = ["id", "scope", "hint", "status", "source", "translation",
          "count", "first_location", "context", "note"]
STATUSES = ("new", "done", "review", "keep", "skip", "obsolete")
VBA_MAX_LINE = 1023


# ---------------------------------------------------------------- small helpers
def key_of(scope: str, text: str) -> str:
    return hashlib.sha1((scope + "\0" + text).encode("utf-8")).hexdigest()[:10]


def read_text(path: Path):
    b = path.read_bytes()
    if b.startswith(b"\xef\xbb\xbf"):
        enc = "utf-8-sig"
    else:
        try:
            b.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            enc = "cp932"
    return b.decode(enc), enc


def enc_vba(t: str) -> str:
    return t.replace('"', '""')


def dec_vba(r: str) -> str:
    return r.replace('""', '"')


def dec_xml(r: str) -> str:
    r = re.sub(r"&#[xX]([0-9A-Fa-f]+);", lambda m: chr(int(m.group(1), 16)), r)
    r = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), r)
    return (r.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


def enc_xml(t: str) -> str:
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;").replace("\r", "&#xD;").replace("\n", "&#xA;"))


def utf16_units(ch: str):
    b = ch.encode("utf-16-le")
    return [int.from_bytes(b[i:i + 2], "little") for i in range(0, len(b), 2)]


CONST_RE = re.compile(r"^(\s*)((?:Public|Private|Global)\s+)?Const\s+(\w+)(?:\s+As\s+String)?\s*=\s*"
                      r'"((?:[^"]|"")*)"\s*(\'.*)?$')


def literal_parts(text: str):
    parts, buf = [], ""
    for ch in text:
        if ord(ch) < 128:
            buf += ch
        else:
            if buf:
                parts.append('"' + enc_vba(buf) + '"')
                buf = ""
            parts += ["ChrW(&H%X&)" % u for u in utf16_units(ch)]
    if buf:
        parts.append('"' + enc_vba(buf) + '"')
    return parts


def render_vba_literal(text: str, literal_only: bool):
    """Return an ASCII-only VBA expression for text, or None when impossible."""
    if text.isascii():
        return '"' + enc_vba(text) + '"'
    if literal_only:
        return None
    return "(" + " & ".join(literal_parts(text)) + ")"


# ---------------------------------------------------------------- tokenizer
def tokenize_vba(line: str):
    """yield ('str'|'com', start, end) -- spans exclude the quotes / the apostrophe."""
    n = len(line) - (1 if line.endswith("\r") else 0)
    m = re.match(r"\s*Rem\b", line, re.I)
    if m:
        yield ("com", m.end(), n)
        return
    i = 0
    while i < n:
        c = line[i]
        if c == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            yield ("str", i + 1, j)
            i = j + 1
        elif c == "'":
            yield ("com", i + 1, n)
            return
        else:
            i += 1


UI_CALL = re.compile(r"\b(MsgBox|InputBox|StatusBar|Err\.Raise|AddItem|Caption|ControlTipText|"
                     r"\.Text|\.Title|\.Label|C_TITLE|Application\.Dialogs|GetOpenFilename|"
                     r"GetSaveAsFilename|FileDialog|Description)\b", re.I)
FUNC_CALL = re.compile(r"\b(Replace|InStr|InStrRev|StrConv|Split|Join|Format|Like|Case|Pattern|"
                       r"NumberFormat|Evaluate|CreateObject|GetObject|Environ|Styles|Worksheets|"
                       r"Sheets|Names|Font\.Name|Dir|Chr|Asc|Left|Right|Mid|Trim|Tag|Key)\b|"
                       r"[=<>]\s*$", re.I)


def hint_for(line: str, start: int, header: bool) -> str:
    if header:
        return "form"
    s = line.strip()
    if re.match(r"Attribute\s", s):
        return "attr"
    if re.match(r"(Public\s+|Private\s+|Global\s+)?Const\b", s, re.I):
        return "const"
    before = line[:max(start - 1, 0)]
    if re.search(r"\b(Worksheets|Sheets|Charts)\s*\(\s*$", before, re.I):
        return "sheet"                       # worksheet-name lookup: template sheet lives in the .xlam
    if UI_CALL.search(before):
        return "ui"
    if FUNC_CALL.search(before):
        return "func?"
    return "other"


class Unit:
    __slots__ = ("scope", "hint", "li", "s", "e", "text", "lineno", "ctx",
                 "literal_only", "has_jp")


class FileScan:
    def __init__(self, rel, enc, kind, lines):
        self.rel, self.enc, self.kind, self.lines = rel, enc, kind, lines
        self.units: list[Unit] = []
        self.code_jp: list[tuple[int, str]] = []


def scan_file(root: Path, path: Path) -> FileScan:
    rel = path.relative_to(root).as_posix()
    ext = path.suffix.lower()
    text, enc = read_text(path)
    if path.name in XML_NAMES:
        fs = FileScan(rel, enc, "xml", [text])
        for m in re.finditer(r'\b([\w:]+)\s*=\s*"([^"]*)"', text):
            if m.group(1) in XML_UI_ATTRS and JP.search(m.group(2)):
                u = Unit()
                u.scope, u.hint, u.li = "xml", m.group(1), 0
                u.s, u.e = m.start(2), m.end(2)
                u.text = dec_xml(m.group(2))
                u.lineno = text.count("\n", 0, u.s) + 1
                u.ctx = text[max(0, text.rfind("\n", 0, u.s) + 1):u.s].strip()[-100:] + m.group(1)
                u.literal_only, u.has_jp = True, True
                fs.units.append(u)
            elif m.group(1) not in XML_UI_ATTRS and m.group(1) not in ("id", "onAction") \
                    and JP.search(m.group(2)):
                fs.code_jp.append((text.count("\n", 0, m.start()) + 1, m.group(0)[:80]))
        for m in re.finditer(r"<!--(.*?)-->", text, re.S):
            if JP.search(m.group(1)):
                u = Unit()
                u.scope, u.hint, u.li = "comment", "xml", 0
                u.s, u.e = m.start(1), m.end(1)
                u.text = m.group(1)
                u.lineno = text.count("\n", 0, u.s) + 1
                u.ctx = m.group(1).strip()[:100]
                u.literal_only, u.has_jp = True, True
                fs.units.append(u)
        return fs

    lines = text.split("\n")
    if ext in TEXT_EXT:
        fs = FileScan(rel, enc, "text", lines)
        for li, line in enumerate(lines):
            if JP.search(line):
                core = line.rstrip("\r")
                u = Unit()
                u.scope, u.hint, u.li = "text", "doc", li
                u.s = len(core) - len(core.lstrip(WS))
                u.e = len(core.rstrip(WS))
                u.text = core[u.s:u.e]
                u.lineno, u.ctx = li + 1, ""
                u.literal_only, u.has_jp = True, True
                fs.units.append(u)
        return fs

    fs = FileScan(rel, enc, "vba", lines)
    header = ext == ".frm"
    last_begin = ""
    for li, line in enumerate(lines):
        if header and line.startswith("Attribute VB_Name"):
            header = False
        if header:
            m = re.match(r"\s*Begin\s+\S+\s+(\w+)", line)
            if m:
                last_begin = m.group(1)
        covered = []
        for kind, s, e in tokenize_vba(line):
            covered.append((s, e))
            raw = line[s:e]
            if not NONASCII.search(raw):
                continue
            u = Unit()
            u.li, u.s, u.e, u.lineno = li, s, e, li + 1
            u.has_jp = bool(JP.search(raw))
            if kind == "str":
                u.scope, u.text = "vba", dec_vba(raw)
                u.hint = hint_for(line, s, header)
                u.literal_only = (header or u.hint in ("const", "attr")
                                  or bool(re.search(r"\bOptional\b", line)))
            else:
                u.scope, u.text, u.hint, u.literal_only = "comment", raw, "vba", True
            u.ctx = ((last_begin + ": ") if header and last_begin else "") + line.strip()[:140]
            fs.units.append(u)
        if not header:
            rest, prev = [], 0
            for s, e in sorted(covered):
                rest.append(line[prev:max(s - 1, 0)])
                prev = e + 1
            rest.append(line[prev:])
            if JP.search("".join(rest)):
                fs.code_jp.append((li + 1, line.strip()[:100]))
    return fs


def iter_files(root: Path, out: Path | None = None):
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        parts = p.relative_to(root).parts
        if parts[0] in SKIP_TOP_DIRS or p.name in SKIP_FILES:
            continue
        if out is not None and out in p.parents:
            continue
        if p.suffix.lower() in VBA_EXT | TEXT_EXT or p.name in XML_NAMES:
            yield p


# ---------------------------------------------------------------- translation memory
class TM:
    def __init__(self, path: Path):
        self.path, self.rows = path, {}
        if path.exists():
            with path.open(encoding="utf-8-sig", newline="") as f:
                for r in csv.DictReader(f):
                    r = {k: (r.get(k) or "") for k in FIELDS}
                    self.rows[r["id"]] = r

    def save(self):
        rows = sorted(self.rows.values(), key=lambda r: (r["scope"], r["first_location"] or "~", r["id"]))
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, FIELDS, lineterminator="\n")
            w.writeheader()
            w.writerows(rows)

    def get(self, scope, text):
        return self.rows.get(key_of(scope, text))


def tm_key_text(u: Unit) -> str:
    return u.text if u.scope in ("vba", "xml") else u.text.strip(WS)


def load_overrides(path: Path):
    ovr = {}
    if path.exists():
        with path.open(encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if r.get("file") and r.get("source_line") is not None:
                    ovr[(r["file"], r["source_line"].rstrip("\r"))] = r.get("replacement_line", "")
    return ovr


def load_identifiers(path: Path):
    ids = {}
    if path.exists():
        with path.open(encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if r.get("source") and r.get("replacement"):
                    ids[r["source"]] = r["replacement"]
    return ids


def rename_code(line: str, ids: dict) -> str:
    """Rename Japanese identifiers, touching code only (never strings or comments)."""
    if not ids or not JP.search(line):
        return line
    out, prev = [], 0
    spans = [(s - (1 if k == "str" else 1), e + (1 if k == "str" else 0)) for k, s, e in tokenize_vba(line)]
    for s, e in sorted(spans):
        out.append(rename_chunk(line[prev:s], ids))
        out.append(line[s:e])
        prev = e
    out.append(rename_chunk(line[prev:], ids))
    return "".join(out)


def rename_chunk(chunk: str, ids: dict) -> str:
    for src in sorted(ids, key=len, reverse=True):
        chunk = chunk.replace(src, ids[src])
    return chunk


# ---------------------------------------------------------------- commands
def cmd_extract(a):
    root = Path(a.src).resolve()
    tm = TM(Path(a.tm))
    found = {}
    code_jp = []
    for p in iter_files(root, Path(a.out).resolve()):
        fs = scan_file(root, p)
        code_jp += [(fs.rel, n, t) for n, t in fs.code_jp]
        for u in fs.units:
            if not u.has_jp:
                continue
            k = tm_key_text(u)
            if not k:
                continue
            d = found.setdefault((u.scope, k), {"hint": u.hint, "count": 0,
                                                 "loc": "%s:%d" % (fs.rel, u.lineno), "ctx": u.ctx})
            d["count"] += 1
    for form, ctl, prop, text in load_frx(Path(a.frx)):
        if JP.search(text):
            d = found.setdefault(("vba", text), {"hint": "frx", "count": 0,
                                                 "loc": "frx:%s.%s.%s" % (form, ctl, prop),
                                                 "ctx": "form control text: %s / %s / %s" % (form, ctl, prop)})
            d["count"] += 1
    added = changed_back = obsoleted = 0
    seen = set()
    for (scope, text), d in found.items():
        i = key_of(scope, text)
        seen.add(i)
        r = tm.rows.get(i)
        if r is None:
            tm.rows[i] = {"id": i, "scope": scope, "hint": d["hint"],
                          "status": "skip" if scope == "comment" else "new",
                          "source": text, "translation": "", "count": str(d["count"]),
                          "first_location": d["loc"], "context": d["ctx"], "note": ""}
            added += 1
        else:
            if r["status"] == "obsolete":
                r["status"] = "done" if r["translation"] else "new"
                changed_back += 1
            r.update(hint=d["hint"], count=str(d["count"]), first_location=d["loc"], context=d["ctx"])
    for i, r in tm.rows.items():
        if i not in seen and r["status"] != "obsolete":
            r["status"], r["count"] = "obsolete", "0"
            obsoleted += 1
    # suggestions for changed strings: nearest obsolete row of the same scope
    obs = [r for r in tm.rows.values() if r["status"] == "obsolete" and r["translation"]]
    if obs:
        for r in tm.rows.values():
            if r["status"] == "new" and not r["note"]:
                best = max((difflib.SequenceMatcher(None, r["source"], o["source"]).ratio(), o)
                           for o in obs if o["scope"] == r["scope"]) if any(
                    o["scope"] == r["scope"] for o in obs) else (0, None)
                if best[0] >= 0.7:
                    r["note"] = "similar to old (%.2f): %s => %s" % (
                        best[0], best[1]["source"][:60], best[1]["translation"][:60])
    tm.save()
    print("extract: %d unique strings | +%d new | %d re-appeared | %d now obsolete"
          % (len(found), added, changed_back, obsoleted))
    if code_jp:
        print("\nNOTE: Japanese outside strings/comments (identifiers or unhandled XML attributes) - "
              "cannot be translated by this tool:")
        for f, n, t in code_jp[:20]:
            print("   %s:%d  %s" % (f, n, t))
        if len(code_jp) > 20:
            print("   ... %d more" % (len(code_jp) - 20))
    print_status(tm)


def print_status(tm: TM):
    c = collections.Counter((r["scope"], r["status"]) for r in tm.rows.values())
    scopes = sorted({s for s, _ in c})
    print("\n%-9s" % "scope" + "".join("%9s" % s for s in STATUSES))
    for sc in scopes:
        print("%-9s" % sc + "".join("%9d" % c.get((sc, s), 0) for s in STATUSES))
    new = [r for r in tm.rows.values() if r["status"] == "new"]
    h = collections.Counter((r["scope"], r["hint"]) for r in new)
    if new:
        print("\nuntranslated by scope/hint:", ", ".join("%s/%s=%d" % (k[0], k[1], v) for k, v in sorted(h.items())))
        print("untranslated Japanese characters: %d" % sum(len(JP.findall(r["source"])) for r in new))


# ---------------------------------------------------------------- form control texts (.frx)
def esc_tsv(t: str) -> str:
    return t.replace("\\", "\\\\").replace("\t", "\\t").replace("\r", "\\r").replace("\n", "\\n")


def unesc_tsv(t: str) -> str:
    out, i = [], 0
    while i < len(t):
        c = t[i]
        if c == "\\" and i + 1 < len(t):
            out.append({"t": "\t", "r": "\r", "n": "\n", "\\": "\\"}.get(t[i + 1], t[i + 1]))
            i += 2
        else:
            out.append(c)
            i += 1
    return "".join(out)


def read_any_text(path: Path) -> str:
    b = path.read_bytes()
    if b[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return b.decode("utf-16")
    return b.decode("utf-8-sig")


def load_frx(path: Path):
    """rows (form, control, property, text) dumped from Excel by rlx_build.vbs"""
    rows = []
    if path.exists():
        for ln in read_any_text(path).splitlines():
            parts = ln.split("\t")
            if len(parts) >= 4 and parts[3]:
                rows.append((parts[0], parts[1], parts[2], unesc_tsv("\t".join(parts[3:]))))
    return rows


def cmd_import_frx(a):
    """store the dump produced by `rlx_build.vbs dump` and fold it into the translation table"""
    rows = []
    text = read_any_text(Path(a.file))
    Path(a.frx).write_text(text, encoding="utf-8", newline="\n")
    n = len([1 for ln in text.splitlines() if ln.strip()])
    print("import-frx: stored %d control strings in %s" % (n, a.frx))
    cmd_extract(a)


def write_frx_translations(tm: TM, frx_path: Path, outdir: Path):
    rows = load_frx(frx_path)
    if not rows:
        return
    pairs, missing = {}, 0
    for form, ctl, prop, text in rows:
        if not JP.search(text):
            continue
        r = tm.get("vba", text)
        if r and r["status"] in ("done", "review") and r["translation"].strip():
            u = Unit()
            u.scope, u.text = "vba", text
            pairs[text] = fix_translation(u, r["translation"])
        else:
            missing += 1
    dst = outdir / "frx_translations.txt"
    dst.parent.mkdir(parents=True, exist_ok=True)
    body = "\r\n".join("%s\t%s" % (esc_tsv(k), esc_tsv(v)) for k, v in pairs.items()) + "\r\n"
    dst.write_bytes(b"\xff\xfe" + body.encode("utf-16-le"))
    print("  form control texts       : %d translations written to %s (%d control strings still untranslated)"
          % (len(pairs), dst.name, missing))


def cmd_patch_xlam(a):
    """replace the ribbon XML inside a (built) .xlam with the translated customUI.xml"""
    src, dst = Path(a.xlam), Path(a.xlam_out)
    xml_path = Path(a.out) / "Source" / "customUI" / "customUI.xml"
    if not xml_path.exists():
        sys.exit("run `apply` first: %s not found" % xml_path)
    data_new = xml_path.read_bytes()
    done = []
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.lower().endswith(("customui/customui.xml", "customui/customui14.xml")):
                data = data_new
                done.append(item.filename)
            zout.writestr(item, data)
    if not done:
        dst.unlink()
        sys.exit("no customUI part found inside %s" % src)
    print("patch-xlam: replaced %s in %s -> %s" % (", ".join(done), src.name, dst))


def cmd_status(a):
    print_status(TM(Path(a.tm)))


def cmd_todo(a):
    tm = TM(Path(a.tm))
    rows = [r for r in tm.rows.values() if r["status"] == a.status
            and (not a.scope or r["scope"] == a.scope) and (not a.hint or r["hint"] == a.hint)]
    rows.sort(key=lambda r: (r["scope"], r["first_location"]))
    for r in rows[a.offset:a.offset + a.limit]:
        print("%s|%s/%s|x%s|%s\n    SRC: %s\n    CTX: %s%s" % (
            r["id"], r["scope"], r["hint"], r["count"], r["first_location"],
            r["source"].replace("\n", "\\n"), r["context"],
            ("\n    NOTE: " + r["note"]) if r["note"] else ""))
    print("-- %d rows matched, showing %d" % (len(rows), min(a.limit, max(0, len(rows) - a.offset))))


def merge_rows(tm: TM, items, force=False):
    """items: iterable of (id, translation, status_or_None)"""
    ok = bad = 0
    for i, tr, st in items:
        r = tm.rows.get(i)
        if r is None:
            print("  unknown id", i)
            bad += 1
            continue
        if r["status"] == "keep" and not force and not st:
            continue
        r["translation"] = tr
        r["status"] = st or ("done" if tr.strip() else r["status"])
        ok += 1
    return ok, bad


def cmd_merge(a):
    tm = TM(Path(a.tm))
    with open(a.file, encoding="utf-8-sig", newline="") as f:
        items = [(r["id"], r.get("translation", ""), r.get("status") or None) for r in csv.DictReader(f)]
    ok, bad = merge_rows(tm, items, a.force)
    tm.save()
    print("merge: %d updated, %d unknown ids" % (ok, bad))


def fix_translation(u: Unit, tr: str) -> str:
    """Normalise a translation for insertion.
    Edge whitespace: if the translation has its own leading/trailing spaces they are kept
    (English needs spaces where Japanese had none); otherwise the source's edge whitespace
    is copied (ideographic spaces become ASCII spaces)."""
    tr = tr.translate(PUNCT)
    if "\n" in tr and "\n" in u.text:            # keep the source's newline style (usually CRLF)
        tr = tr.replace("\r\n", "\n")
        if "\r\n" in u.text:
            tr = tr.replace("\n", "\r\n")
    core = tr.strip(WS)
    t_lead = tr[:len(tr) - len(tr.lstrip(WS))]
    t_trail = tr[len(tr.rstrip(WS)):]
    s_lead = u.text[:len(u.text) - len(u.text.lstrip(WS))].replace("\u3000", " ")
    s_trail = u.text[len(u.text.rstrip(WS)):].replace("\u3000", " ")
    return (t_lead or s_lead) + core + (t_trail or s_trail)


def asciify_line(line: str) -> str:
    """ascii-safe: turn non-ASCII string literals (and comments) of an override line into ChrW()/'?'"""
    core = line.rstrip("\r")
    cr = line[len(core):]
    out, prev = [], 0
    for kind, s_, e_ in tokenize_vba(core):
        raw = core[s_:e_]
        if not NONASCII.search(raw):
            continue
        if kind == "str":
            out.append(core[prev:s_ - 1])
            out.append("(" + " & ".join(literal_parts(dec_vba(raw))) + ")")
            prev = e_ + 1
        else:
            out.append(core[prev:s_])
            out.append(re.sub(r"\?+", "?", NONASCII.sub("?", raw)))
            prev = e_
    out.append(core[prev:])
    return "".join(out) + cr


def apply_file(fs: FileScan, tm: TM, ovr, mode: str, st: collections.Counter, ids=None):
    ext = Path(fs.rel).suffix.lower()
    out = list(fs.lines)
    done_lines = set()
    for li, l in enumerate(fs.lines):
        k = (fs.rel, l.rstrip("\r"))
        if k in ovr:
            out[li] = ovr[k] + ("\r" if l.endswith("\r") else "")
            if mode == "ascii-safe" and fs.kind == "vba" and NONASCII.search(out[li]):
                out[li] = asciify_line(out[li])
            done_lines.add(li)
            st["override_used:" + fs.rel + "\0" + k[1]] += 1
    by_line = collections.defaultdict(list)
    for u in fs.units:
        if u.li not in done_lines:
            by_line[u.li].append(u)
    for li, us in by_line.items():
        line = fs.lines[li]
        for u in sorted(us, key=lambda u: u.s, reverse=True):
            row = tm.get(u.scope, tm_key_text(u)) if u.has_jp else None
            tr = None
            if row and row["status"] in ("done", "review") and row["translation"].strip():
                tr = fix_translation(u, row["translation"])
                st["translated"] += 1
            elif u.has_jp:
                st["untranslated_" + u.scope] += 1
            s, e = u.s, u.e
            if u.scope == "vba":
                text = tr if tr is not None else u.text
                if mode == "ascii-safe":
                    expr = render_vba_literal(text, u.literal_only)
                    if expr is None:
                        st["unsafe_literal"] += 1
                        new = enc_vba(text)
                    else:
                        new = expr[1:-1] if expr.startswith('"') and expr.endswith('"') and text.isascii() else expr
                        if not text.isascii():
                            s, e = s - 1, e + 1
                            st["chrw"] += 1
                else:
                    new = enc_vba(text)
                if tr is None and new == line[u.s:u.e]:
                    continue
            elif u.scope == "xml":
                if tr is None:
                    continue
                new = enc_xml(tr)
            elif u.scope == "comment":
                if tr is not None:
                    new = tr
                elif mode == "ascii-safe" and NONASCII.search(u.text) and fs.kind == "vba":
                    new = NONASCII.sub("?", u.text)
                    new = re.sub(r"\?+", "?", new)
                    st["comment_scrubbed"] += 1
                else:
                    continue
            else:  # text
                if tr is None:
                    continue
                new = tr
            line = line[:s] + new + line[e:]
        out[li] = line
    if mode == "ascii-safe" and fs.kind == "vba":
        cr = "\r" if "\r\n" in "\n".join(fs.lines[:50]) or (fs.lines and fs.lines[0].endswith("\r")) else ""
        extra = []
        for li, line in enumerate(out):
            if li in done_lines or not NONASCII.search(line):
                continue
            core = line.rstrip("\r")
            m = CONST_RE.match(core)
            if not m:
                continue
            ind, scope, name, raw, com = m.groups()
            text_ = dec_vba(raw)
            if not NONASCII.search(text_):
                continue
            parts = literal_parts(text_)
            com = com if com and not NONASCII.search(com) else ""
            if ind:                                    # local constant -> local variable
                out[li] = "%sDim %s As String: %s = (%s)%s%s" % (
                    ind, name, name, " & ".join(parts), (" " + com) if com else "", cr)
            else:                                      # module constant -> parameterless Function
                out[li] = "' %s: defined as a Function at the end of this module%s" % (name, cr)
                body = (" & _" + cr + "\n        ").join(" & ".join(parts[i:i + 6]) for i in range(0, len(parts), 6))
                vis = "Public " if scope and scope.strip() in ("Public", "Global") else "Private "
                extra += ["%sFunction %s() As String%s" % (vis, name, cr), "    %s = %s%s" % (name, body, cr),
                          "End Function" + cr]
            st["const_converted"] += 1
            st["unsafe_literal"] = max(0, st["unsafe_literal"] - 1)
        if extra:
            if out and out[-1] == "":
                out[-1:-1] = extra
            else:
                out.extend(extra)
    if ids and fs.kind == "vba":
        for li, line in enumerate(out):
            if li not in done_lines and JP.search(line):
                out[li] = rename_code(line, ids)
    text = "\n".join(out)
    if fs.kind == "vba" and mode == "ascii-safe" and not text.isascii():
        st["not_ascii:" + fs.rel] += 1
    return text


def cmd_apply(a):
    root, outdir = Path(a.src).resolve(), Path(a.out).resolve()
    tm = TM(Path(a.tm))
    if not tm.rows:
        sys.exit("translation table is empty - run `extract` first")
    ovr = load_overrides(Path(a.overrides))
    ids = load_identifiers(Path(a.identifiers))
    if outdir.exists():
        shutil.rmtree(outdir)
    # copy everything that is not ours (frx binaries, rels, ...) so the tree is import-ready
    for p in root.rglob("*"):
        parts = p.relative_to(root).parts
        if not p.is_file() or parts[0] in SKIP_TOP_DIRS or outdir in p.parents:
            continue
        dst = outdir / p.relative_to(root)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, dst)
    st = collections.Counter()
    nfiles = 0
    for p in iter_files(root, outdir):
        fs = scan_file(root, p)
        text = apply_file(fs, tm, ovr, a.encoding, st, ids)
        ext = p.suffix.lower()
        enc = fs.enc
        if ext in CP932_ALWAYS:
            enc = "cp932"
        elif fs.kind == "vba" and a.encoding == "cp932":
            enc = "cp932"
        elif fs.kind == "vba" and a.encoding == "ascii-safe" and not text.isascii():
            enc = fs.enc if fs.enc != "utf-8-sig" else "utf-8-sig"
        try:
            data = text.encode(enc)
        except UnicodeEncodeError as ex:
            print("  !! %s: cannot encode as %s (%s) - written as utf-8" % (fs.rel, enc, ex.reason))
            data = text.encode("utf-8")
        dst = outdir / fs.rel
        dst.write_bytes(data)
        nfiles += 1
    write_frx_translations(tm, Path(a.frx), outdir)
    used = {k.split("\0", 1)[1] for k in st if k.startswith("override_used:")}
    stale = [k for k in ovr if k[1] not in used]
    print("apply: %d files written to %s  (mode=%s)" % (nfiles, outdir, a.encoding))
    print("  strings/lines translated : %d" % st["translated"])
    for sc in ("vba", "xml", "text", "comment"):
        if st["untranslated_" + sc]:
            print("  left in Japanese (%-7s): %d" % (sc, st["untranslated_" + sc]))
    if st["chrw"]:
        print("  converted to ChrW()      : %d string literals" % st["chrw"])
    if st["const_converted"]:
        print("  Const -> Function/Dim conversions (non-ASCII constants): %d" % st["const_converted"])
    if st["comment_scrubbed"]:
        print("  untranslated comments made ASCII: %d" % st["comment_scrubbed"])
    if st["unsafe_literal"]:
        print("  !! %d Const/Optional/form literals still contain non-ASCII (cannot use ChrW)" % st["unsafe_literal"])
    bad = [k[len("not_ascii:"):] for k in st if k.startswith("not_ascii:")]
    if bad:
        print("  files that are NOT pure ASCII (need a Japanese system locale to import safely):")
        for b in bad[:15]:
            print("     ", b)
    if stale:
        print("  stale line overrides (upstream line changed): %d" % len(stale))
        for f, l in stale[:10]:
            print("     %s: %s" % (f, l[:80]))


def cmd_sheetnames(a):
    """Worksheet names the code looks up.  The matching template sheets inside the .xlam must be
    renamed to the English names (Excel limit: 31 characters, no  : \\ / ? * [ ] )."""
    tm = TM(Path(a.tm))
    print("%-34s %-34s %s" % ("Japanese sheet name (in xlam)", "English name (code now expects)", "status"))
    for r in sorted(tm.rows.values(), key=lambda r: r["source"]):
        if r["hint"] == "sheet" and r["status"] != "obsolete":
            en = r["translation"] or "(untranslated)"
            warn = "  !! >31 chars" if len(en) > 31 else ""
            print("%-34s %-34s %s%s" % (r["source"], en, r["status"], warn))


def skeleton(line: str) -> str:
    """line with the *contents* of strings and comments removed"""
    out, prev = [], 0
    for kind, s, e in tokenize_vba(line):
        out.append(line[prev:s])
        prev = e
    out.append(line[prev:])
    return "".join(out)


def has_unterminated(line: str) -> bool:
    n = len(line) - (1 if line.endswith("\r") else 0)
    return any(k == "str" and e >= n and (e == n and (e == s or line[e - 1:e] != '"' or True)) and
               not (e < len(line) and line[e:e + 1] == '"') for k, s, e in tokenize_vba(line))


def selftest(a) -> int:
    """Apply in utf-8 mode and prove that only string/comment contents changed (code is untouched)."""
    root, tm, ovr = Path(a.src).resolve(), TM(Path(a.tm)), load_overrides(Path(a.overrides))
    bad = 0
    for p in iter_files(root, Path(a.out).resolve()):
        fs = scan_file(root, p)
        if fs.kind != "vba":
            continue
        new = apply_file(fs, tm, ovr, "utf-8", collections.Counter(), None).split("\n")
        for i, (o, n) in enumerate(zip(fs.lines, new)):
            if (fs.rel, o.rstrip("\r")) in ovr:
                continue
            if skeleton(o) != skeleton(n):
                print("ERROR  code changed outside strings/comments: %s:%d" % (fs.rel, i + 1))
                bad += 1
            elif has_unterminated(n) and not has_unterminated(o):
                print("ERROR  unbalanced quote introduced: %s:%d" % (fs.rel, i + 1))
                bad += 1
        if len(new) != len(fs.lines):
            print("ERROR  line count changed: %s" % fs.rel)
            bad += 1
    print("selftest: %s" % ("OK - only string/comment contents differ from upstream" if not bad else "%d problem(s)" % bad))
    return bad


def cmd_check(a):
    tm = TM(Path(a.tm))
    warn = err = 0
    for r in tm.rows.values():
        if r["status"] not in ("done", "review") or not r["translation"].strip():
            continue
        tr, src, where = r["translation"], r["source"], "%s %s" % (r["id"], r["first_location"])
        if r["scope"] in ("vba", "text") and ("\n" in tr or "\r" in tr) and "\n" not in src:
            print("ERROR  newline in translation (%s): %s" % (where, tr[:60]))
            err += 1
        if JP_LETTERS.search(tr):
            print("WARN   Japanese left in translation (%s): %s" % (where, tr[:60]))
            warn += 1
        if collections.Counter(PLACEHOLDER.findall(src)) != collections.Counter(PLACEHOLDER.findall(tr)):
            print("WARN   placeholder mismatch (%s): %s  =>  %s" % (where, src[:50], tr[:50]))
            warn += 1
        if r["scope"] == "vba" and len(tr) > 40 and len(tr) > 8 * max(len(src), 8):
            print("WARN   very long translation (%s): %s" % (where, tr[:60]))
            warn += 1
    err += selftest(a)
    outdir = Path(a.out)
    if outdir.exists():
        for p in sorted(outdir.rglob("*")):
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            if ext in (".bas", ".cls", ".frm", ".vbs"):
                text, _ = read_text(p)
                for n, l in enumerate(text.split("\n"), 1):
                    if len(l.rstrip("\r")) > VBA_MAX_LINE:
                        print("ERROR  line > %d chars (VBA limit): %s:%d" % (VBA_MAX_LINE, p.relative_to(outdir), n))
                        err += 1
            if p.name in XML_NAMES:
                try:
                    xml.dom.minidom.parseString(p.read_bytes())
                except Exception as ex:  # noqa
                    print("ERROR  %s is not well-formed XML: %s" % (p.name, ex))
                    err += 1
    print("check: %d error(s), %d warning(s)" % (err, warn))
    return 1 if err else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", default=str(DEFAULT_SRC), help="repo root (default: parent of this folder)")
    ap.add_argument("--tm", default=str(DEFAULT_TM))
    ap.add_argument("--overrides", default=str(DEFAULT_OVR))
    ap.add_argument("--identifiers", default=str(DEFAULT_IDS))
    ap.add_argument("--frx", default=str(DEFAULT_FRX), help="form control strings dumped from Excel")
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("extract").set_defaults(fn=cmd_extract)
    sub.add_parser("status").set_defaults(fn=cmd_status)
    t = sub.add_parser("todo")
    t.add_argument("--status", default="new")
    t.add_argument("--scope")
    t.add_argument("--hint")
    t.add_argument("--limit", type=int, default=50)
    t.add_argument("--offset", type=int, default=0)
    t.set_defaults(fn=cmd_todo)
    m = sub.add_parser("merge")
    m.add_argument("file")
    m.add_argument("--force", action="store_true")
    m.set_defaults(fn=cmd_merge)
    p = sub.add_parser("apply")
    p.add_argument("--encoding", choices=["ascii-safe", "utf-8", "cp932"], default="ascii-safe",
                   help="ascii-safe (default): pure-ASCII VBA files, import works on any Windows locale")
    p.set_defaults(fn=cmd_apply)
    sub.add_parser("check").set_defaults(fn=cmd_check)
    sub.add_parser("sheetnames").set_defaults(fn=cmd_sheetnames)
    f = sub.add_parser("import-frx")
    f.add_argument("file", help="text file written by `cscript rlx_build.vbs dump ...`")
    f.set_defaults(fn=cmd_import_frx)
    x = sub.add_parser("patch-xlam")
    x.add_argument("xlam", help="built .xlam (after importing the modules)")
    x.add_argument("xlam_out", help="output .xlam with the English ribbon")
    x.set_defaults(fn=cmd_patch_xlam)
    a = ap.parse_args()
    sys.exit(a.fn(a) or 0)


if __name__ == "__main__":
    main()
