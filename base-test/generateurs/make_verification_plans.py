"""Acceptance plan files to try the check before import (loupe of the Plan d'acceptation box).

They start from the test plans plans/2026-01-01_PA_FCTk_A-plan.docx, which must be loaded in the base with their
requirements (tomes fct1 to fct10, release A). Written into verification-import/plans/:
- one file per case (cases), with LISEZMOI.html giving the expected result of each one;
- complement/2026-03-15_PA_FCTk_A-complement.docx: 5 % of the rows of the plan modified, plus rows for the new
  requirements of the requirement complement files (SSS_FCTk_REQ_500...);
- remplacement/2026-03-15_PA_FCTk_A-remplacement.docx: the whole plan, 5 % of its rows removed, 5 % modified, plus
  rows for the new requirements of the requirement replacement files (SSS_FCTk_REQ_600...).
The new requirements are unknown of the base until those requirement files are imported: the check reports them as
unknown, then, once imported, as new rows. A modified row changes, in turn, its IADT (D and T swapped, MOC4 kept), its
MOC (MOC2 and MOC3 swapped, IADT A) or its delegation; the rules IADT / MOC stay respected. Word documents in A4
portrait, on the styles of stdhl/2026-11-11-STDHL-template2.docx.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_verification_plans.py ../../base-test
"""
import math
import random
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, ".")
sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, next to this script
from lisezmoi import write_html
from configuration import Configuration
from plans_db import PlansDatabase

BASE = Path(sys.argv[1])
OUT = BASE / "verification-import" / "plans"
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only: how plans are read
SHARE = 0.05
HEADER = ["Exigence", "Méthodes de test V&V", "IADT", "MOC", "Délégation"]


def p(text, style=""):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="2000"/>' for _ in rows[0])
    body = "".join("<w:tr>" + "".join(f"<w:tc>{p(c)}</w:tc>" for c in r) + "</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>")


def write_plan(path, title, tables):
    """A plan document: a title, then for each (heading, rows) a heading and a plan table; A4 portrait."""
    body = [p(title, "Titre")]
    for heading, rows, header in tables:
        body += [p(heading, "Titre2"), table([header] + rows)]
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


def titles_of(k):
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    return {r[0]: r[1] for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}


def plan_of(k):
    """Entries of the test plan of fctk: [{requirement, iadt, moc, delegation}], in document order."""
    data = PlansDatabase.read_docx((BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx").read_bytes(), cfg)
    return [{f: e[f] for f in ("requirement", "iadt", "moc", "delegation")} for e in PlansDatabase.entries(data["rows"])]


def row(e, titles):
    """A plan table row of an entry."""
    title = titles.get(e["requirement"], "nouvelle exigence")
    return [f"{e['requirement']} {title}", f"V&V [{e['delegation'] or 'TOP'}] : méthode de test et moyens", e["iadt"],
            e["moc"], e["delegation"]]


def modify(e, field):
    """The entry with one field changed, the rules IADT / MOC kept."""
    e = dict(e)
    if field == "IADT":
        e["iadt"] = "T" if e["iadt"] == "D" else "D"
    elif field == "MOC":
        e["moc"] = "MOC3" if e["moc"] == "MOC2" else "MOC2"
    else:
        e["delegation"] = "PLT2_N-1" if not e["delegation"] or e["delegation"] == "TOP" else ""
    return e


def eligible(e, field):
    """Whether an entry can take this modification: IADT on D or T rows, MOC on A rows, delegation on tested rows."""
    if e["iadt"] == "N/A":
        return False
    return {"IADT": e["iadt"] in ("D", "T"), "MOC": e["iadt"] == "A" and e["moc"] in ("MOC2", "MOC3"),
            "délégation": True}[field]


def pick_modified(entries, count, rng, avoid=()):
    """{index: field} of `count` entries, the fields in turn, each entry eligible for its field."""
    chosen, fields = {}, ("IADT", "MOC", "délégation")
    for j in range(count):
        field = fields[j % 3]
        pool = [i for i, e in enumerate(entries) if i not in chosen and i not in avoid and eligible(e, field)]
        if not pool:
            pool = [i for i, e in enumerate(entries) if i not in chosen and i not in avoid and eligible(e, "délégation")]
            field = "délégation"
        chosen[rng.choice(pool)] = field
    return chosen


def new_entries(k, first, count):
    return [{"requirement": f"SSS_FCT{k}_REQ_{n}", "iadt": "T", "moc": "MOC4", "delegation": ""}
            for n in range(first, first + count)]


# ------------------------------------------------------------------ cases
cases = []


def case(name, what, expected, title, tables):
    write_plan(OUT / name, title, tables)
    cases.append((name, what, expected))


t1 = titles_of(1)
case("2026-03-01_PA_FCT1_A-identique.docx", "copie exacte du plan fct1 de la base",
     "Chargeable dans les deux cas, aucun problème : 0 ligne nouvelle, 0 modifiée",
     "Plan d'acceptation FCT1", [("FCT1", [row(e, t1) for e in plan_of(1)], HEADER)])
t2 = titles_of(2)
case("2026-03-01_PA_FCT2_A-entetes.docx", "en-tête « IADT » renommé « Type d'essai »",
     "Refusé dans les deux cas : aucun tableau de plan valide (Bloquant)",
     "Plan d'acceptation FCT2", [("FCT2", [row(e, t2) for e in plan_of(2)],
                                  ["Exigence", "Méthodes de test V&V", "Type d'essai", "MOC", "Délégation"])])
t3, e3 = titles_of(3), plan_of(3)
half = len(e3) // 2
case("2026-03-01_PA_FCT3_A-tableau-incomplet.docx",
     f"deux tableaux ; le second, sans colonne « MOC », porte {len(e3) - half} exigences",
     f"Chargeable, à vérifier : tableau 2 mis de côté (À vérifier) ; s'il écrase la base, les {len(e3) - half} "
     "exigences du second tableau restent sans ligne de plan (À vérifier)",
     "Plan d'acceptation FCT3", [("FCT3 — partie 1", [row(e, t3) for e in e3[:half]], HEADER),
                                 ("FCT3 — partie 2", [[r[0], r[1], r[2], r[4]] for r in (row(e, t3) for e in e3[half:])],
                                  ["Exigence", "Méthodes de test V&V", "IADT", "Délégation"])])
t4, e4 = titles_of(4), plan_of(4)
tested = [i for i, e in enumerate(e4) if e["iadt"] in ("D", "T")]
broken = {tested[0]: ("X", "MOC4"), tested[1]: ("T", "MOC7"), tested[2]: ("T", ""), tested[3]: ("", "MOC4")}
rows4 = []
for i, e in enumerate(e4):
    if i in broken:
        e = {**e, "iadt": broken[i][0], "moc": broken[i][1]}
    rows4.append(row(e, t4))
ids4 = [e4[i]["requirement"] for i in broken]
case("2026-03-01_PA_FCT4_A-regles.docx",
     f"{ids4[0]} en IADT X ; {ids4[1]} en T avec MOC7 ; {ids4[2]} sans MOC ; {ids4[3]} sans IADT",
     "Chargeable avec des choix à faire dans les deux cas : IADT inconnu, MOC hors règle, MOC vide, IADT vide (À "
     "régler), avec le MOC attendu en correction ; lignes de plan de la base modifiées (Information)",
     "Plan d'acceptation FCT4", [("FCT4", rows4, HEADER)])
t5, e5 = titles_of(5), plan_of(5)
tested5 = [i for i, e in enumerate(e5) if e["iadt"] != "N/A"]
e5[tested5[0]] = {**e5[tested5[0]], "delegation": "PLT99"}
e5[tested5[1]] = {**e5[tested5[1]], "delegation": "TOP, PLT1_N-1"}
case("2026-03-01_PA_FCT5_A-delegations.docx",
     f"{e5[tested5[0]]['requirement']} délégué à PLT99 (inconnue) ; {e5[tested5[1]]['requirement']} à deux plateformes",
     "Chargeable avec des choix à faire dans les deux cas : plateforme inconnue PLT99, plusieurs délégations (À "
     "régler) ; délégations modifiées (Information)",
     "Plan d'acceptation FCT5", [("FCT5", [row(e, t5) for e in e5], HEADER)])
t6 = titles_of(6)
case("2026-03-01_PA_FCT6_A-inconnues.docx", "3 lignes pour des exigences absentes de la base (SSS_FCT6_REQ_990 à 992)",
     "Chargeable avec des choix à faire dans les deux cas : 3 exigences inconnues de la base, lignes ignorées (À régler)",
     "Plan d'acceptation FCT6", [("FCT6", [row(e, t6) for e in plan_of(6) + new_entries(6, 990, 3)], HEADER)])
t7, t8, e8 = titles_of(7), titles_of(8), plan_of(8)
case("2026-03-01_PA_FCT7_A-deux-tomes.docx", f"plan fct7 avec 3 lignes du tome fct8 ({', '.join(e['requirement'] for e in e8[:3])})",
     "Chargeable, à vérifier : exigences de 2 tomes, fct7 et fct8 (À vérifier) ; s'il écrase la base, les autres "
     "exigences du tome fct8 restent sans ligne de plan (À vérifier)",
     "Plan d'acceptation FCT7", [("FCT7", [row(e, t7) for e in plan_of(7)] + [row(e, t8) for e in e8[:3]], HEADER)])
t9 = titles_of(9)
case("plan-sans-tome.docx", "lignes du plan fct9, nom de fichier sans date ni tome",
     "Chargeable dans les deux cas, aucun problème : le nom du fichier n'est pas utilisé, les exigences sont reconnues "
     "dans le tome fct9", "Plan d'acceptation", [("Exigences", [row(e, t9) for e in plan_of(9)], HEADER)])
t10, e10 = titles_of(10), plan_of(10)
cut = len(e10) // 2
write_plan(OUT / "2026-03-01_PA_FCT10_A-partie1.docx", "Plan d'acceptation FCT10 — partie 1",
           [("FCT10", [row(e, t10) for e in e10[:cut + 5]], HEADER)])
write_plan(OUT / "2026-03-01_PA_FCT10_A-partie2.docx", "Plan d'acceptation FCT10 — partie 2",
           [("FCT10", [row(e, t10) for e in e10[cut:]], HEADER)])
cases.append(("2026-03-01_PA_FCT10_A-partie1.docx et 2026-03-01_PA_FCT10_A-partie2.docx",
              "le plan fct10 en deux fichiers, 5 exigences dans les deux",
              "Vérifiés ensemble : 5 exigences dans plusieurs plans (À régler) ; s'ils écrasent la base, le tome fct10 "
              "concerné par 2 plans (À régler) et, pour chacun seul, les exigences de l'autre partie sans ligne de plan"))

lines = ["# Plans pour essayer la vérification avant import", "",
         "Prérequis : les exigences et les plans de test (`exigences/2026-01-01-fctk.xlsx`, `plans/2026-01-01_PA_FCTk_A-"
         "plan.docx`) sont chargés dans la base, ainsi que les plateformes PLT1_N-1 à PLT14_N-1. Vérifier ces fichiers "
         "avec la loupe de la boîte Plan d'acceptation ; rien n'est importé.", "",
         "La vérification ne se sert pas du nom du fichier : chaque exigence est cherchée dans toute la base, et sa ligne "
         "comparée à celle des plans de la base. Le rapport donne pour chaque fichier ce qui se passe s'il écrase la "
         "base (les plans des tomes concernés sont remplacés) et s'il la modifie (ses lignes remplacent celles des "
         "mêmes exigences) ; les résultats ci-dessous valent pour les deux cas, sauf mention.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_plans.py ../../base-test` depuis "
         "`nexus-v/programme`.", "", "| Fichier | Contenu | Résultat attendu |", "| --- | --- | --- |"]
lines += [f"| {f} | {what} | {expected} |" for f, what, expected in cases]
write_html(OUT / "LISEZMOI.html", lines)

# ------------------------------------------------------------------ complement and replacement
report = []
for k in range(1, 11):
    rng = random.Random(2500 + k)
    titles, entries = titles_of(k), plan_of(k)
    n = max(1, math.ceil(SHARE * len(entries)))
    modified = pick_modified(entries, n, rng)
    added = new_entries(k, 500, n)
    rows = [row(modify(entries[i], f), titles) for i, f in modified.items()] + [row(e, titles) for e in added]
    write_plan(OUT / "complement" / f"2026-03-15_PA_FCT{k}_A-complement.docx", f"Plan d'acceptation FCT{k} — complément",
               [(f"FCT{k}", rows, HEADER)])
    by = lambda changed, f: [entries[i]["requirement"] for i, g in changed.items() if g == f]
    report.append((f"2026-03-15_PA_FCT{k}_A-complement.docx", "complément",
                   {"nouvelles (exigences du complément)": [e["requirement"] for e in added],
                    "IADT modifié": by(modified, "IADT"), "MOC modifié": by(modified, "MOC"),
                    "délégation modifiée": by(modified, "délégation")}))

    removed = set(rng.sample([i for i in range(len(entries))], n))
    changed = pick_modified(entries, n, rng, avoid=removed)
    added = new_entries(k, 600, n)
    rows = [row(modify(e, changed[i]) if i in changed else e, titles) for i, e in enumerate(entries) if i not in removed]
    rows += [row(e, titles) for e in added]
    write_plan(OUT / "remplacement" / f"2026-03-15_PA_FCT{k}_A-remplacement.docx",
               f"Plan d'acceptation FCT{k} — remplacement", [(f"FCT{k}", rows, HEADER)])
    report.append((f"2026-03-15_PA_FCT{k}_A-remplacement.docx", "remplacement",
                   {"retirées": [entries[i]["requirement"] for i in sorted(removed)],
                    "nouvelles (exigences du remplacement)": [e["requirement"] for e in added],
                    "IADT modifié": by(changed, "IADT"), "MOC modifié": by(changed, "MOC"),
                    "délégation modifiée": by(changed, "délégation")}))

kinds = ("retirées", "nouvelles", "IADT modifié", "MOC modifié", "délégation modifiée")
count = lambda found, kind: next((len(v) for c, v in found.items() if c.startswith(kind)), None)
lines = ["# Plans de complément et de remplacement", "",
         "Prérequis : les exigences et les plans de test fct1 à fct10 sont chargés dans la base.", "",
         "- `complement/` : plans à importer en complément : 5 % des lignes du plan modifiées, et des lignes pour les "
         "exigences nouvelles des fichiers d'exigences de complément (`SSS_FCTk_REQ_500…`). Dans le rapport, regarder "
         "la partie « Si le fichier modifie la base ».",
         "- `remplacement/` : plans à importer en remplacement : tout le plan, moins 5 % de lignes retirées, 5 % de "
         "lignes modifiées, et des lignes pour les exigences nouvelles des fichiers d'exigences de remplacement "
         "(`SSS_FCTk_REQ_600…`). Dans le rapport, regarder la partie « Si le fichier écrase la base ».", "",
         "Tant que les fichiers d'exigences correspondants ne sont pas importés, les exigences nouvelles sont "
         "signalées « exigence inconnue de la base » (À régler) ; une fois importées, elles apparaissent comme "
         "« sans ligne de plan dans la base, couvertes par ce plan ».", "",
         "Une ligne modifiée change un champ, à tour de rôle, dans le respect des règles IADT / MOC : l'IADT (D et T "
         "échangés, MOC4 gardé), le MOC (MOC2 et MOC3 échangés, IADT A) ou la délégation (TOP devient PLT2_N-1, une "
         "plateforme devient TOP). Les lignes retirées laissent leurs exigences sans ligne de plan si le fichier "
         "écrase la base.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_plans.py ../../base-test` depuis "
         "`nexus-v/programme`.", "",
         "| Fichier | Mode | " + " | ".join(k.capitalize() for k in kinds) + " |",
         "| --- | --- | " + " | ".join("---" for _ in kinds) + " |"]
for f, mode, found in report:
    lines.append(f"| {f} | {mode} | " + " | ".join(str(count(found, k)) if count(found, k) is not None else "—"
                                                    for k in kinds) + " |")
lines += ["", "## Numéros par fichier", ""]
for f, mode, found in report:
    lines += [f"**{f}** ({mode})", ""] + [f"- {c} : {', '.join(v)}" for c, v in found.items() if v] + [""]
write_html(OUT / "LISEZMOI-complement-remplacement.html", lines)
print(f"{len(list(OUT.glob('*.docx')))} plans de cas, {len(report)} plans de complément et de remplacement écrits dans {OUT}")
