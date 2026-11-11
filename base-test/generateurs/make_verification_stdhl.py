"""STDHL files of format 1 (one per tome) to try the check before import (loupe of the STDHL box).

They start from the test STDHL stdhl/perTome/2026-01-01_TEST_STDHL_FCTk_rev1.docx, which must be loaded in the base
with the requirements and the plans of the tests (tomes fct1 to fct10, release A). Written into verification-import/stdhl/:
- one file per case, with LISEZMOI.html giving the expected result of each one;
- complement/2026-03-15_STDHL_FCTk-complement_rev1.docx: 5 % of the requirements of the STDHL with another test, plus
  the new requirements of the requirement complement files (SSS_FCTk_REQ_500...);
- remplacement/2026-03-15_TEST_STDHL_FCTk_rev2.docx (same STDHL name as the one of the base): the whole STDHL, 5 % of
  its requirements removed, 5 % with another test, plus the new requirements of the requirement replacement files
  (SSS_FCTk_REQ_600...).
The new requirements are unknown of the base until those requirement files (and their plans) are imported: the check
reports them as unknown, then as newly specified. A test is a paragraph « SYS_... » followed by its tables (Description
du test, Résultat attendu, Exigences associées). Word documents in A4 portrait, on the styles of
stdhl/2026-11-11-STDHL-template2.docx.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_verification_stdhl.py ../../base-test
"""
import math
import random
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, ".")
sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, next to this script
from lisezmoi import write_html
from configuration import Configuration
from plans_db import PlansDatabase
from stdhl_db import StdhlDatabase

BASE = Path(sys.argv[1])
OUT = BASE / "verification-import" / "stdhl"
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only: how plans and STDHL are read
SHARE = 0.05
HEADER = ["Description du test", "Résultat attendu", "Exigences associées"]


def p(text, style=""):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="3000"/>' for _ in rows[0])
    body = "".join("<w:tr>" + "".join(f"<w:tc>{p(c)}</w:tc>" for c in r) + "</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>")


def write_stdhl(path, title, tests, header=HEADER, orphan=None):
    """A STDHL of format 1: a title, then for each (test, [requirements]) the paragraph of the test and its table, one
    row per requirement. `orphan`: requirements of a table put before the first test (no SYS_ paragraph)."""
    body = [p(title, "Titre"), p("Description des tests", "Titre1")]
    if orphan:
        body += [p("TEST CASE #1 : tableau placé avant tout test"),
                 table([header] + [["Vérifier l'exigence", "Conforme", i] for i in orphan])]
    for test, reqs in tests:
        body += [p(test, "Titre2"), p("Procédure de test"), p("TEST CASE #1 : cas nominal"),
                 table([header] + [[f"Vérifier {i}", "Conforme à l'exigence", i] for i in reqs])]
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SKELETON) as z, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "word/document.xml":
                x = data.decode("utf-8")
                a, b = x.index("<w:body>") + len("<w:body>"), x.rindex("<w:sectPr")
                x = x[:a] + "".join(body) + x[b:]
                data = re.sub(r"<w:pgSz [^>]*/>", '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>', x).encode("utf-8")
            w.writestr(info, data)


def stdhl_of(k):
    """Tests of the test STDHL of fctk: [(test, [requirements])], in document order."""
    name = f"2026-01-01_TEST_STDHL_FCT{k}_rev1.docx"
    data = StdhlDatabase.read_docx((BASE / "stdhl" / "perTome" / name).read_bytes(), name, cfg)
    tests = {}
    for r in data["rows"]:
        for i in r["requirements"]:
            if i not in tests.setdefault(r["fiche"], []):
                tests[r["fiche"]].append(i)
    return list(tests.items())


def plan_of(k):
    data = PlansDatabase.read_docx((BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx").read_bytes(), cfg)
    return PlansDatabase.entries(data["rows"])


cases = []


def case(name, what, expected, title, tests, **kw):
    write_stdhl(OUT / name, title, tests, **kw)
    cases.append((name, what, expected))


case("2026-03-01_TEST_STDHL_FCT1_rev2.docx", "copie exacte du STDHL fct1 de la base",
     "Chargeable dans les deux cas : aucune exigence nouvelle ni modifiée ; un STDHL de même nom (FCT1) sera remplacé "
     "(Information) ; les exigences de TOP déjà sans STDHL dans la base sont rappelées (À vérifier)",
     "STDHL FCT1", stdhl_of(1))
case("2026-03-01_TEST_STDHL_FCT2-entetes_rev1.docx", "en-tête « Exigences associées » renommé « Exigences »",
     "Refusé dans les deux cas : aucun tableau STDHL valide (Bloquant)",
     "STDHL FCT2", stdhl_of(2), header=["Description du test", "Résultat attendu", "Exigences"])
case("stdhl-fct3.docx", "STDHL fct3, nom de fichier hors format (ni « STDHL_ » ni « _rev »)",
     "Chargeable dans les deux cas : nom de fichier hors format (À vérifier) ; les exigences sont reconnues dans le "
     "tome fct3", "STDHL FCT3", stdhl_of(3))
t4 = stdhl_of(4)
case("2026-03-01_TEST_STDHL_FCT4-inconnues_rev1.docx", "3 exigences absentes de la base (SSS_FCT4_REQ_990 à 992)",
     "Chargeable avec des choix à faire dans les deux cas : 3 exigences inconnues de la base, non enregistrées (À régler)",
     "STDHL FCT4", t4[:-1] + [(t4[-1][0], t4[-1][1] + [f"SSS_FCT4_REQ_{n}" for n in range(990, 993)])])
na5 = [e["requirement"] for e in plan_of(5) if e["iadt"] == "N/A"][:2]
t5 = stdhl_of(5)
case("2026-03-01_TEST_STDHL_FCT5-na_rev1.docx", f"{', '.join(na5)} ajoutées : N/A dans le plan",
     "Chargeable dans les deux cas : 2 exigences N/A dans les plans, ignorées (Information)",
     "STDHL FCT5", t5[:-1] + [(t5[-1][0], t5[-1][1] + na5)])
other6 = [e["requirement"] for e in plan_of(6) if e["iadt"] != "N/A" and e["delegation"] not in ("", "TOP")][:2]
t6 = stdhl_of(6)
case("2026-03-01_TEST_STDHL_FCT6-plateforme_rev1.docx", f"{', '.join(other6)} ajoutées : déléguées à une autre plateforme",
     "Chargeable, à vérifier, dans les deux cas : 2 exigences déléguées à une autre plateforme que TOP (À vérifier)",
     "STDHL FCT6", t6[:-1] + [(t6[-1][0], t6[-1][1] + other6)])
t7 = stdhl_of(7)
case("2026-03-01_TEST_STDHL_FCT7-sans-test_rev1.docx",
     f"les exigences du premier test ({t7[0][0]}) dans un tableau placé avant tout paragraphe SYS_",
     "Chargeable, à vérifier, dans les deux cas : tableau sans fiche SYS_ avant lui et exigences sans test SYS_ (À "
     "vérifier) ; ces exigences ont un test modifié (Information)",
     "STDHL FCT7", t7[1:], orphan=t7[0][1])
t8, t9 = stdhl_of(8), stdhl_of(9)
case("2026-03-01_TEST_STDHL_FCT8-deux-tomes_rev1.docx", f"STDHL fct8 et le premier test du STDHL fct9 ({t9[0][0]})",
     "Chargeable, à vérifier, dans les deux cas : exigences de 2 tomes, fct8 et fct9 (À vérifier) ; s'il écrase la "
     "base, les autres exigences de TOP du tome fct9 perdent leur STDHL (À vérifier)",
     "STDHL FCT8", t8 + t9[:1])
t10 = stdhl_of(10)
reqs10 = list(dict.fromkeys(i for _, rs in t10 for i in rs))
half = len(reqs10) // 2
share = reqs10[half:half + 3]
test10 = t10[0][0]
write_stdhl(OUT / "2026-03-01_STDHL_FCT10-partie1_rev1.docx", "STDHL FCT10 — partie 1", [(test10, reqs10[:half + 3])])
write_stdhl(OUT / "2026-03-01_STDHL_FCT10-partie2_rev1.docx", "STDHL FCT10 — partie 2", [(test10, reqs10[half:])])
cases.append(("2026-03-01_STDHL_FCT10-partie1_rev1.docx et 2026-03-01_STDHL_FCT10-partie2_rev1.docx",
              f"les exigences du STDHL fct10 coupées en deux fichiers, sous le même test, 3 exigences ({', '.join(share)}) "
              "dans les deux",
              "Vérifiés ensemble : 3 exigences dans plusieurs STDHL (À régler) ; s'ils écrasent la base, le tome fct10 "
              "concerné par 2 STDHL (À régler) ; seul, chacun fait perdre leur STDHL aux exigences de l'autre partie s'il "
              "écrase la base (À vérifier)"))

lines = ["# STDHL pour essayer la vérification avant import", "",
         "Prérequis : les exigences, les plans et les STDHL de test fct1 à fct10 sont chargés dans la base. Vérifier ces "
         "fichiers avec la loupe de la boîte STDHL ; rien n'est importé. Seul le format 1 (tableaux, un STDHL par tome) "
         "est vérifié.", "",
         "La vérification ne se sert pas du nom du fichier pour le tome : chaque exigence est cherchée dans toute la "
         "base, et les tests demandés pour elle comparés à ceux des STDHL de la base. Le rapport donne pour chaque "
         "fichier ce qui se passe s'il écrase la base (les STDHL des tomes concernés sont remplacés) et s'il la modifie "
         "(ses lignes remplacent celles des mêmes exigences) ; les résultats ci-dessous valent pour les deux cas, sauf "
         "mention.", "",
         "Dans la base de test, 16 des 258 exigences de TOP n'ont pas de STDHL : chaque rapport rappelle celles des "
         "tomes concernés, « déjà sans STDHL dans la base » (À vérifier). Ce rappel n'est pas répété ci-dessous.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_stdhl.py ../../base-test` depuis "
         "`nexus-v/programme`.", "", "| Fichier | Contenu | Résultat attendu |", "| --- | --- | --- |"]
lines += [f"| {f} | {what} | {expected} |" for f, what, expected in cases]
write_html(OUT / "LISEZMOI.html", lines)


def moved(tests, chosen, rng):
    """The tests with the requirements `chosen` asked for another test of the STDHL (a new one when there is one
    test only); returns (tests, {requirement: new test})."""
    names = [t for t, _ in tests]
    changes = {}
    for i in chosen:
        own = next(t for t, rs in tests if i in rs)
        others = [t for t in names if t != own]
        changes[i] = rng.choice(others) if others else re.sub(r"\d+$", "999", own)
    out = {t: [i for i in rs if i not in changes] for t, rs in tests}
    for i, t in changes.items():
        out.setdefault(t, []).append(i)
    return [(t, rs) for t, rs in out.items() if rs], changes


report = []
for k in range(1, 11):
    rng = random.Random(3500 + k)
    tests = stdhl_of(k)
    reqs = list(dict.fromkeys(i for _, rs in tests for i in rs))   # a requirement may be under several tests
    n = max(1, math.ceil(SHARE * len(reqs)))
    new_test = f"SYS_NOUVEAU_FCT{k}_001"

    # Complement: the requirements whose test changes, and the new ones; nothing else.
    chosen = rng.sample(reqs, n)
    full, changes = moved(tests, chosen, rng)
    only = [(t, [i for i in rs if i in changes]) for t, rs in full]
    added = [f"SSS_FCT{k}_REQ_{m}" for m in range(500, 500 + n)]
    write_stdhl(OUT / "complement" / f"2026-03-15_STDHL_FCT{k}-complement_rev1.docx", f"STDHL FCT{k} — complément",
                [(t, rs) for t, rs in only if rs] + [(new_test, added)])
    report.append((f"2026-03-15_STDHL_FCT{k}-complement_rev1.docx", "complément",
                   {"nouvelles": added, "test modifié": sorted(changes)}))

    # Replacement: the whole STDHL, some requirements removed, some with another test, some new.
    picked = rng.sample(reqs, 2 * n)
    removed, chosen = set(picked[:n]), picked[n:]
    kept = [(t, [i for i in rs if i not in removed]) for t, rs in tests]
    full, changes = moved([(t, rs) for t, rs in kept if rs], chosen, rng)
    added = [f"SSS_FCT{k}_REQ_{m}" for m in range(600, 600 + n)]
    write_stdhl(OUT / "remplacement" / f"2026-03-15_TEST_STDHL_FCT{k}_rev2.docx", f"STDHL FCT{k} — remplacement",
                full + [(new_test, added)])
    report.append((f"2026-03-15_TEST_STDHL_FCT{k}_rev2.docx", "remplacement",
                   {"retirées": sorted(removed), "nouvelles": added, "test modifié": sorted(changes)}))

kinds = ("retirées", "nouvelles", "test modifié")
lines = ["# STDHL de complément et de remplacement", "",
         "Prérequis : les exigences, les plans et les STDHL de test fct1 à fct10 sont chargés dans la base.", "",
         "- `complement/` : STDHL à importer en complément : 5 % des exigences du STDHL demandées pour un autre test, et "
         "les exigences nouvelles des fichiers d'exigences de complément (`SSS_FCTk_REQ_500…`), sous un test "
         "`SYS_NOUVEAU_FCTk_001`. Dans le rapport, regarder la partie « Si le fichier modifie la base ».",
         "- `remplacement/` : STDHL à importer en remplacement, de même nom que celui de la base : tout le STDHL, moins 5 % "
         "d'exigences retirées, 5 % demandées pour un autre test, et les exigences nouvelles des fichiers d'exigences de "
         "remplacement (`SSS_FCTk_REQ_600…`). Dans le rapport, regarder la partie « Si le fichier écrase la base ».", "",
         "Tant que les exigences nouvelles ne sont pas importées (avec leur ligne de plan), elles sont signalées "
         "« exigences inconnues de la base » (À régler). Les exigences retirées « perdent leur STDHL avec l'import » "
         "si le fichier écrase la base.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_stdhl.py ../../base-test` depuis "
         "`nexus-v/programme`.", "",
         "| Fichier | Mode | " + " | ".join(k.capitalize() for k in kinds) + " |",
         "| --- | --- | " + " | ".join("---" for _ in kinds) + " |"]
for f, mode, found in report:
    lines.append(f"| {f} | {mode} | " + " | ".join(str(len(found[k])) if k in found else "—" for k in kinds) + " |")
lines += ["", "## Numéros par fichier", ""]
for f, mode, found in report:
    lines += [f"**{f}** ({mode})", ""] + [f"- {c} : {', '.join(v)}" for c, v in found.items() if v] + [""]
write_html(OUT / "LISEZMOI-complement-remplacement.html", lines)
print(f"{len(list(OUT.glob('*.docx')))} STDHL de cas, {len(report)} STDHL de complément et de remplacement écrits dans {OUT}")
