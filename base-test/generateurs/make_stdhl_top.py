"""Test STDHL documents FCT1 .. FCT10 (stdhl/perTome/2026-01-01_TEST_STDHL_FCTk_rev1.docx), from base-test.

Only the requirements delegated to TOP are specified in a STDHL: requirements of release A whose row in the acceptance
plan of their tome has the delegation TOP (written or empty cell: default platform), not N/A. Each one goes into the
STDHL of its column STDHL (the function), 5 to 10 % left uncovered for the tests; about a third of them appear on
several test lines. The Word package of the STDHL template is reused: only the body changes.
Usage (from nexus-v/programme): python make_stdhl_top.py <base-test> <template.docx>
"""
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, ".")
from configuration import Configuration
from plans_db import PlansDatabase

BASE, TEMPLATE = Path(sys.argv[1]), Path(sys.argv[2])
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only: how plans are read
DEFAULT = "TOP"
TYPES = ("IHM", "COM", "PERF", "SECU")
ACTIONS = ("Lancer le scénario nominal", "Provoquer le cas dégradé", "Rejouer la séquence opérateur",
           "Injecter le message de test")


def p(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{ppr}<w:r><w:t xml:space=\"preserve\">{escape(text)}</w:t></w:r></w:p>"


def cell(paragraphs, width, span=None):
    extra = f'<w:gridSpan w:val="{span}"/>' if span else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{extra}</w:tcPr>{"".join(p(t) for t in paragraphs)}</w:tc>'


def table(lines):
    """lines: [(description, expected, [IDs])], under the expected header and a phase row."""
    rows = ["<w:tr>" + cell(["Description du test"], 3020) + cell(["Résultat attendu"], 3021)
            + cell(["Exigences associées"], 3021) + "</w:tr>",
            "<w:tr>" + cell(["Phase 1"], 9062, span=3) + "</w:tr>"]
    rows += ["<w:tr>" + cell([d], 3020) + cell([e], 3021) + cell(ids, 3021) + "</w:tr>" for d, e, ids in lines]
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="Grilledutableau"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3020"/><w:gridCol w:w="3021"/><w:gridCol w:w="3021"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")


def build(name, reqs, rng):
    rate = rng.uniform(0.05, 0.10)
    uncovered = set(rng.sample([r for r, _ in reqs], round(rate * len(reqs))))
    titles = dict(reqs)
    covered = [r for r, _ in reqs if r not in uncovered]
    # Test lines: 1 to 3 requirements each; about a third of them come back on a later line.
    queue = list(covered)
    repeated = rng.sample(queue, len(queue) // 3)
    for rid in repeated:
        queue.insert(rng.randrange(len(queue) + 1), rid)
    lines, i = [], 0
    while i < len(queue):
        n = rng.randint(1, 3)
        ids = list(dict.fromkeys(queue[i:i + n]))
        i += n
        lines.append((f"{rng.choice(ACTIONS)} : {titles[ids[0]]}", "Comportement conforme à l'exigence", ids))
    body, sheet, j = [p("Description des tests", "Titre1")], 0, 0
    while j < len(lines):
        sheet += 1
        body += [p(f"SYS_{rng.choice(TYPES)}_{name}_{sheet:03d}", "Titre2"), p("Procédure de test", "Titre3")]
        for case in range(1, rng.randint(1, 2) + 1):
            part = lines[j:j + rng.randint(3, 6)]
            j += len(part)
            if part:
                body += [p(f"TEST CASE #{case} : fiche {sheet:03d} de {name}"), table(part), p("")]
    return "".join(body), uncovered, repeated


# Requirements of release A with their title and function (column STDHL).
requirements = {}
for k in range(1, 11):
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    for rid, title, rel, stdhl in ws.iter_rows(min_row=2, values_only=True):
        if "A" in (rel or "").split(", "):
            requirements[rid] = (title, stdhl)
# Requirements delegated to TOP in the plans (not N/A).
top = set()
for k in range(1, 11):
    plan = PlansDatabase.read_docx((BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx").read_bytes(), cfg)
    for row in plan["rows"]:
        if row["iadt"].strip().upper() == "N/A":
            continue
        if (row["delegation"].strip() or DEFAULT) == DEFAULT:
            top.update(row["requirements"])

with zipfile.ZipFile(TEMPLATE) as z:
    parts = {i: z.read(i) for i in z.infolist()}
document = next(v for k, v in parts.items() if k.filename == "word/document.xml").decode("utf-8")
start, end = document.index("<w:body>") + 8, document.index("<w:sectPr")
total = 0
for k in range(1, 11):
    name = f"FCT{k}"
    reqs = sorted((rid, title) for rid, (title, stdhl) in requirements.items() if stdhl == name and rid in top)
    body, uncovered, repeated = build(name, reqs, random.Random(206 + k))
    target = BASE / "stdhl" / "perTome" / f"2026-01-01_TEST_STDHL_{name}_rev1.docx"
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in parts.items():
            out.writestr(info, (document[:start] + body + document[end:]).encode("utf-8")
                         if info.filename == "word/document.xml" else data)
    total += len(reqs)
    print(f"{target.name} : {len(reqs)} exigences TOP, {len(uncovered)} non couvertes, {len(repeated)} sur plusieurs lignes")
print("exigences TOP :", len(top), "| dans les STDHL :", total)
