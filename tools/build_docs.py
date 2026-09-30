"""Fill the generated sections of the Markdown files from the workbooks.

Run after tools/build_excel.py and after the workbooks have been recalculated
(open and save in Excel, or LibreOffice). If the workbooks have no cached values,
this script recalculates temporary copies with LibreOffice (soffice).

Markers in the Markdown files:
  <!-- BEGIN GENERATED: questions ch08 --> ... <!-- END GENERATED: questions ch08 -->
  <!-- BEGIN GENERATED: answers ch08 -->   ... <!-- END GENERATED: answers ch08 -->
  <!-- BEGIN GENERATED: ch08.ex1.journal --> ... (a tagged worked-example block)
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
from itertools import zip_longest
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
LEVELS = {1: "Level 1: Easy", 2: "Level 2: Easy to medium", 3: "Level 3: Medium", 4: "Level 4: Hard",
          5: "Level 5: Challenge"}
PCT, INT, NUM2 = "0.0%", "0", '#,##0.00;(#,##0.00);"-"'


def load_values(fname):
    path = ROOT / "excel" / fname
    wb = load_workbook(path, data_only=True)
    formulas = load_workbook(path)
    cells = [(ws.title, c.coordinate) for ws in formulas.worksheets for row in ws.iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith("=")]
    empty = sum(1 for t, a in cells if wb[t][a].value is None)
    if empty < len(cells) / 2:   # blank-by-design cells (e.g. unanswered checks) read as None
        return wb
    if False:
        return wb
    print(f"  {fname}: no cached values, recalculating a copy with LibreOffice")
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "in" / fname
        src.parent.mkdir()
        shutil.copy(path, src)
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(src)],
                       check=True, capture_output=True)
        return load_workbook(Path(tmp) / fname, data_only=True)


def money(v, fmt=None, zero="-"):
    if v is None or v == "":
        return ""
    if isinstance(v, str):
        return v
    if fmt == PCT:
        return f"{v * 100:g}%"
    if fmt == INT:
        return f"{v:g}"
    if abs(v) < 1e-9:
        return zero
    if fmt == NUM2 or abs(v - round(v)) > 1e-6:
        s = f"{abs(v):,.2f}"
    else:
        s = f"{abs(round(v)):,}"
    return f"({s})" if v < 0 else s


def cell(get, a, fmt=None):
    return money(get(a), fmt)


def esc(t):
    return str(t).replace("|", "\\|")


# ------------------------------------------------------------------ block renderers
def r_data(b):
    out = ["| Item | Figure |", "|---|---|"]
    for label, value, fmt in b["rows"]:
        if label == "#":
            out.append(f"| **{esc(value)}** | |")
        else:
            out.append(f"| {esc(label)} | {money(value, fmt, zero='0')} |")
    return "\n".join(out)


def r_tb(b):
    out = [f"**{b['title']}**", "", "| Particulars | Debit (RM) | Credit (RM) |", "|---|---:|---:|"]
    for label, dr, cr in b["rows"]:
        out.append(f"| {esc(label)} | {money(dr) if dr is not None else ''} | {money(cr) if cr is not None else ''} |")
    d, c = b["totals"]
    out.append(f"| **Total** | **{money(d)}** | **{money(c)}** |")
    return "\n".join(out)


def r_calc(b, get):
    out = ["| Item | Answer |", "|---|---:|"]
    for label, a, fmt in b["rows"]:
        out.append(f"| {esc(label)} | **{cell(get, a, fmt)}** |")
    return "\n".join(out)


def r_choice(b, get):
    out = ["| Item | Answer |", "|---|---|"]
    for label, a, _acc in b["rows"]:
        out.append(f"| {esc(label)} | **{esc(get(a))}** |")
    return "\n".join(out)


def r_written(b):
    return "\n\n".join(line.strip() for line in b["text"].split("\n") if line.strip())


def r_table(b, get):
    heads = b["headers"]
    out = ["| " + " | ".join(esc(h) for h in heads) + " |", "|---" + "|---:" * (len(heads) - 1) + "|"]
    for label, cells in b["rows"]:
        vals = []
        values = []
        for c in cells:
            if c is None:
                vals.append("")
            elif c[0] == "v":
                v = get(c[1])
                values.append(v)
                vals.append(money(v, c[2]))
            elif c[0] == "d":
                vals.append(money(c[1], c[2], zero="0"))
            else:
                vals.append(esc(c[1]))
        if values and all(v in (None, "") for v in values):
            continue
        out.append(f"| {esc(label)} | " + " | ".join(vals) + " |")
    return "\n".join(out)


def r_journal(b, get):
    out = ["| Date | Particulars | Debit (RM) | Credit (RM) |", "|---|---|---:|---:|"]
    for row in b["rows"]:
        if row[0] == "n":
            out.append(f"| | *({esc(row[1])})* | | |")
        elif row[0] == "d":
            out.append(f"| {row[1]} | {esc(row[2])} | {cell(get, row[3])} | |")
        else:
            out.append(f"| {row[1]} | &emsp;&emsp;{esc(row[2])} | | {cell(get, row[3])} |")
    return "\n".join(out)


def r_ledger(b, get):
    out = [f"**Dr** &emsp;&emsp; **{b['name']}** &emsp;&emsp; **Cr**", "",
           "| Date | Particulars | Folio | Amount (RM) | Date | Particulars | Folio | Amount (RM) |",
           "|---|---|---|---:|---|---|---|---:|"]
    sides = {"dr": [], "cr": []}

    def flush(total=None):
        for d, c in zip_longest(sides["dr"], sides["cr"]):
            left = f"{d[0]} | {esc(d[1])} | {d[2]} | {money(get(d[3]))}" if d else " | | | "
            right = f"{c[0]} | {esc(c[1])} | {c[2]} | {money(get(c[3]))}" if c else " | | | "
            out.append(f"| {left} | {right} |")
        if total:
            out.append(f"| | | | **{money(get(total['dr']))}** | | | | **{money(get(total['cr']))}** |")
        sides["dr"].clear()
        sides["cr"].clear()

    year = {"dr": "", "cr": ""}
    for row in b["rows"]:
        if row.get("total"):
            flush(row)
            continue
        for side in ("dr", "cr"):
            e = row.get(side)
            if not e:
                continue
            m = re.match(r"^(\d{4})\b", e[0])
            if get(e[3]) in (None, "", 0, 0.0):
                if m:
                    year[side] = m.group(1)   # keep the year of a dropped line for the next line shown
                continue
            if not m and not sides[side] and year[side] and e[0]:
                e = [f"{year[side]} {e[0]}"] + list(e[1:])
            year[side] = ""
            sides[side].append(e)
    flush()
    return "\n".join(out)


def r_statement(b, get):
    out = ["  \n".join(f"**{h}**" for h in b["heading"]), "",
           "| | Particulars | RM | RM | RM |", "|---|---|---:|---:|---:|"]
    if b["kind"] == "sofp":
        out.append("| | | **Cost** | **Accumulated Depreciation** | **Carrying Amount** |")
    for m, p, vals, flags in b["rows"]:
        name = f"**{esc(p)}**" if p and ("h" in flags or "b" in flags) else esc(p)
        cols = []
        for c in ("1", "2", "3"):
            if c in vals:
                a, line = vals[c]
                s = cell(get, a)
                cols.append(f"**{s}**" if line == "total" or "b" in flags else s)
            else:
                cols.append("")
        out.append(f"| {m} | {name} | " + " | ".join(cols) + " |")
    return "\n".join(out)


def render_answer_block(b, get):
    k = b["k"]
    if k == "part":
        return f"**{b['text']}**"
    if k == "calc":
        return r_calc(b, get)
    if k == "choice":
        return r_choice(b, get)
    if k == "written":
        return r_written(b)
    if k == "table":
        return r_table(b, get) if any(c and c[0] == "v" for _, cells in b["rows"] for c in cells) else None
    if k == "journal":
        return r_journal(b, get)
    if k == "ledger":
        return r_ledger(b, get)
    if k == "statement":
        return r_statement(b, get)
    return None


def render_question(sheet):
    parts = [f"**{sheet['qid']} ({LEVELS[sheet['level']]})**"]
    bullets = []

    def flush_bullets():
        if bullets:
            parts.append("\n".join(f"* {x}" for x in bullets))
            bullets.clear()

    for b in sheet["blocks"]:
        k = b["k"]
        if k == "text" and b["style"] == "bullet":
            bullets.append(b["text"])
            continue
        flush_bullets()
        if k == "text":
            parts.append(b["text"])
        elif k == "data" and b["md"]:
            parts.append(r_data(b))
        elif k == "tb":
            parts.append(r_tb(b))
        elif k == "table" and not any(c and c[0] == "v" for _, cells in b["rows"] for c in cells):
            parts.append(r_table(b, lambda a: None))
        elif k == "part":
            parts.append(b["text"])
        elif k == "choice":
            subs = [label for label, _a, _acc in b["rows"] if re.match(r"^[a-z]\) ", label)]
            if subs:
                parts.append("\n\n".join(subs))
    flush_bullets()
    return "\n\n".join(parts)


def render_answers(sheet, get):
    parts = [f"### {sheet['qid']} (Level {sheet['level']})"]
    for b in sheet["blocks"]:
        s = render_answer_block(b, get)
        if s:
            parts.append(s)
    return "\n\n".join(parts)


def replace_section(text, name, body, path):
    pat = re.compile(rf"(<!-- BEGIN GENERATED: {re.escape(name)} -->).*?(<!-- END GENERATED: {re.escape(name)} -->)",
                     re.S)
    if not pat.search(text):
        return text, False
    return pat.sub(lambda m: m.group(1) + "\n" + body + "\n" + m.group(2), text), True


def main():
    manifest = json.loads((ROOT / "tools" / "manifest.json").read_text())
    sections = {}
    for fname, book in manifest.items():
        wb = load_values(fname)
        chap = book["chapter"]
        qs, ans = [], []
        for sheet in book["sheets"]:
            ws = wb[sheet["sheet"]]

            def get(a, ws=ws):
                return ws[a].value

            if sheet["mode"] == "q" and sheet["qid"] != "quiz":
                qs.append(render_question(sheet))
            elif sheet["mode"] == "a" and sheet["qid"] != "quiz":
                ans.append(render_answers(sheet, get))
            elif sheet["mode"] == "w":
                for b in sheet["blocks"]:
                    if "tag" in b:
                        sections[b["tag"]] = render_answer_block(b, get) if b["k"] != "tb" else r_tb(b)
        sections[f"questions {chap}"] = "\n\n".join(qs)
        sections[f"answers {chap}"] = "\n\n".join(ans)
    used = set()
    files = sorted((ROOT / "notes").glob("*.md")) + sorted((ROOT / "model-questions").glob("*.md"))
    for path in files:
        text = path.read_text(encoding="utf-8")
        for name, body in sections.items():
            text, hit = replace_section(text, name, body, path)
            if hit:
                used.add(name)
        path.write_text(text, encoding="utf-8")
    unused = sorted(set(sections) - used)
    missing = []
    for path in files:
        for name in re.findall(r"<!-- BEGIN GENERATED: (.*?) -->", path.read_text(encoding="utf-8")):
            if name not in sections:
                missing.append(f"{path.name}: {name}")
    print("filled", len(used), "sections")
    if unused:
        print("not placed:", ", ".join(unused))
    if missing:
        print("MISSING:", "; ".join(missing))
        sys.exit(1)


if __name__ == "__main__":
    main()
