"""Test STDHL of format 2 for the function FCT1, built on the template 2026-11-11-STDHL-template2.docx.

Every requirement of FCT1 delegated to TOP (release A, plan row TOP, not N/A) is covered, spread over 3 to 8 tests.
Exact proportions of the 40 phases: 20 without requirement (50 %), 10 with one (25 %), 4 with two (10 %), 4 with three
(10 %), 2 with four to six (5 %). Phases beyond the number of requirements cite again a requirement of the same test
(in another phase), so that the list « Exigences associées » of a test is exactly the requirements of its phases.
Test names: <type of test, 30 characters at most>_<n°> : <title>, the number unique in the document. The document
has tables (history of versions, test means under each test), none with the header of a STDHL of format 1.
Usage (from nexus-v/programme): python make_stdhl_format2.py <base-test> <output.docx>
"""
import random
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, ".")
from configuration import Configuration
from plans_db import PlansDatabase

BASE, OUT = Path(sys.argv[1]), Path(sys.argv[2])
TEMPLATE = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
rng = random.Random(20261111)
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only: how plans are read

# Requirements of FCT1 delegated to TOP.
titles = {r[0]: r[1] for r in openpyxl.load_workbook(BASE / "exigences" / "2026-01-01-fct1.xlsx").active.iter_rows(min_row=2, values_only=True)
          if "A" in (r[2] or "").split(", ")}
plan = PlansDatabase.read_docx((BASE / "plans" / "2026-01-01_PA_FCT1_A-plan.docx").read_bytes(), cfg)
top = list(dict.fromkeys(i for row in plan["rows"] if row["iadt"].strip().upper() != "N/A"
                         and (row["delegation"].strip() or "TOP") == "TOP" for i in row["requirements"] if i in titles))
rng.shuffle(top)

# Exact sizes of the 40 phases, in random order.
sizes = [0] * 20 + [1] * 10 + [2] * 4 + [3] * 4 + [rng.randint(4, 6) for _ in range(2)]
rng.shuffle(sizes)
assert sum(sizes) >= len(top), "not enough places for every requirement"
# Phases into test cases (2 to 6 phases each), test cases into 3 to 8 tests.
cases, at = [], 0
while at < len(sizes):
    n = rng.randint(2, 6)
    cases.append(sizes[at:at + n])
    at += n
n_tests = max(3, min(8, len(cases), rng.randint(3, 8)))
tests = [[] for _ in range(n_tests)]
for k, case in enumerate(cases):
    tests[k if k < n_tests else rng.randrange(n_tests)].append(case)

# Requirements into the places: a new one while there are as many places left as new requirements, otherwise a
# requirement already cited by the same test (in another phase), so that every requirement is cited.
places = sum(sizes)
queue, filled = list(top), 0
content = []   # per test, per test case, per phase: requirement numbers
for test_cases in tests:
    own, test_content = [], []
    for case in test_cases:
        case_content = []
        for size in case:
            phase = []
            for _ in range(size):
                left = places - filled
                again = [i for i in own if i not in phase]
                if queue and (left <= len(queue) or not again):
                    phase.append(queue.pop(0))
                    own.append(phase[-1])
                else:
                    phase.append(rng.choice(again))
                filled += 1
            case_content.append(phase)
        test_content.append(case_content)
    content.append(test_content)
assert not queue, "requirements left without place"

TYPES = ["HORODATAGE_EVENEMENTS", "GESTION_ALARMES", "SUPERVISION_EQUIPEMENTS", "MODES_DEGRADES", "ARCHIVAGE_JOURNAUX",
         "MESSAGERIE_OPERATEUR", "TELECOMMANDES", "SECURITE_ACCES"]
assert all(len(t) <= 30 for t in TYPES)
ACTIONS = ["Lancer le scénario nominal", "Provoquer le cas dégradé", "Rejouer la séquence opérateur",
           "Vérifier l'affichage de l'IHM", "Contrôler les journaux", "Couper puis rétablir la liaison"]
MEANS = [("Banc de test V&V", "Configuration nominale"), ("Simulateur réseau", "Liaison dégradée"),
         ("Poste opérateur", "Profil administrateur"), ("Enregistreur", "Journalisation complète")]


def p(text, style=""):
    """A paragraph of the document, with its style."""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    """A Word table (first row = header), with borders."""
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    cell = lambda text: f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>{p(text)}</w:tc>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            + "".join(f"<w:tr>{''.join(cell(c) for c in row)}</w:tr>" for row in rows) + "</w:tbl>")


types = rng.sample(TYPES, n_tests)
names = [f"{types[k]}_{k + 1:03d}" for k in range(n_tests)]
heads = [f"vérification de {types[k].lower().replace('_', ' ')} de FCT1" for k in range(n_tests)]
body = [p("Contents", "En-ttedetabledesmatires"), p("1 Historique 1", "TM1"), p("2 Description des tests 2", "TM1")]
body += [p(f"2.{k + 1} {names[k]} : {heads[k]} {k + 2}", "TM3") for k in range(n_tests)]
body += [p("Historique des versions", "Titre1"),
         table([["Version", "Date", "Auteur", "Modification"], ["1.0", "2026-10-15", "V&V", "Création"],
                ["1.1", "2026-11-11", "V&V", "Tests de la fonction FCT1"]]),
         p("Domaine d'application", "Titre1"), p("Généralités sur le système", "Titre2"), p("blabla"),
         p("Description des tests", "Titre1"), p("Tests de la fonction FCT1", "Titre2")]
# Lists « Exigences associées »: the requirements of the phases, then a few inconsistencies on purpose (at most
# INCONSISTENCY of the requirements): a requirement of the phases left out of the list, or a requirement of another
# test added to the list while no phase of this test cites it.
INCONSISTENCY = 0.10
lists = [list(dict.fromkeys(i for case in t for phase in case for i in phase)) for t in content]
budget = int(INCONSISTENCY * len(top))
inconsistent = []
for n in range(budget):
    k = rng.choice([j for j in range(n_tests) if len(lists[j]) > 1])
    if n % 3 < 2:   # two out of three: cited in a phase, missing from the list
        i = rng.choice(lists[k])
        lists[k].remove(i)
        inconsistent.append(f"{i} cité dans les phases de {k + 1} mais absent de sa liste")
    else:           # one out of three: in the list, cited by no phase of the test
        i = rng.choice([r for j, l in enumerate(lists) if j != k for r in l if r not in lists[k]])
        lists[k].append(i)
        inconsistent.append(f"{i} dans la liste de {k + 1} mais cité dans aucune de ses phases")
for k, test_content in enumerate(content):
    listed = lists[k]
    body +=[p(f" {names[k]} : {heads[k]}", "Titre3"), p("État : Revue"), p(f"Objectif : {heads[k]}"),
             p("Applicabilité : RELEASE A"), p("Exigences associées :")]
    body += [p(f"{i} : {titles[i]}", "Paragraphedeliste") for i in listed]
    body += [p("Conditions et configuration initiales :"), p("Banc de test en configuration nominale", "Paragraphedeliste"),
             p("Moyens d'essai :"), table([["Moyen", "Configuration"], *rng.sample(MEANS, 2)]),
             p("DESCRIPTION DE LA FICHE DE TEST")]
    for c, case in enumerate(test_content, 1):
        body.append(p(f"TEST CASE#{c} : {rng.choice(ACTIONS).lower()}"))
        for f, phase in enumerate(case, 1):
            body.append(p(f"Phase {f} : {rng.choice(ACTIONS)}" + (" : " + ", ".join(phase) if phase else ""), "Paragraphedeliste"))

with zipfile.ZipFile(TEMPLATE) as z, zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as w:
    for info in z.infolist():
        data = z.read(info.filename)
        if info.filename == "word/document.xml":
            x = data.decode("utf-8")
            # The body is replaced up to the last section properties (the first sectPr may sit inside a paragraph).
            a, b = x.index("<w:body>") + len("<w:body>"), x.rindex("<w:sectPr")
            # Every section in A4 portrait (the template ends with a landscape section).
            data = re.sub(r"<w:pgSz [^>]*/>", '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>', x[:a] + "".join(body) + x[b:]).encode("utf-8")
        w.writestr(info, data)

phases = [len(ph) for t in content for case in t for ph in case]
count = {label: sum(1 for s in phases if ok(s)) for label, ok in
         (("0", lambda s: s == 0), ("1", lambda s: s == 1), ("2", lambda s: s == 2), ("3", lambda s: s == 3), ("4-6", lambda s: s >= 4))}
print(f"{OUT.name} : {len(top)} exigences TOP de FCT1, {n_tests} tests, {len(cases)} test cases, {len(phases)} phases, "
      f"{sum(phases)} citations")
print("phases :", ", ".join(f"{k} exigence(s) : {v} ({100 * v / len(phases):.0f} %)" for k, v in count.items()))
for k, t in enumerate(content):
    listed = {i for case in t for ph in case for i in ph}
    print(f"  {names[k]} ({len(names[k])} car.) : {len(t)} test cases, {len(listed)} exigences")
print(f"incohérences voulues ({len(inconsistent)}) :")
for line in inconsistent:
    k = int(line.split(" de ")[1].split()[0]) - 1
    print("  " + line.replace(f" de {k + 1} ", f" de {names[k]} "))
