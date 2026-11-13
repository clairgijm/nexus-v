"""Requirement files of complement and of replacement, to try the check before import (and the import itself).

For each tome fct1 to fct10, from its requirements of release A (exigences/2026-01-01-fctk.xlsx, loaded in the base):
- complement/2026-03-15-fctk-complement.xlsx: only new requirements (5 % of the tome) and modified ones (5 %);
- remplacement/2026-03-15-fctk-remplacement.xlsx: the whole tome, 5 % removed, 5 % new and 5 % modified.
A modified requirement changes one field, in turn: its title, its STDHL, or its release (A becomes B: the requirement
leaves the release of the base, so it is reported outside the release). LISEZMOI-complement-remplacement.html lists, file by file, the numbers
of each case: what the check (and the import) must find.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_verification_modifications.py ../../base-test
"""
import math
import random
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, next to this script
from lisezmoi import write_html

BASE = Path(sys.argv[1])
OUT = BASE / "verification-import" / "exigences"
SHARE = 0.05
FIELDS = ("titre", "STDHL", "release")
in_a = lambda r: "A" in str(r[2] or "").replace(" ", "").split(",")


def rows_of(k):
    """Rows [ID, Titre, Specification releases, STDHL, IADT, MOC, Délégation, Validation releases] of release A of the
    test file of fctk, as the base holds them (specification releases reduced to A)."""
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    return [validated(["A" if c == 2 else v for c, v in enumerate(r[:8])])
            for r in ws.iter_rows(min_row=2, values_only=True) if r[0] and in_a(r)]


def validated(row):
    """The row with its validation releases within its specification releases (A); a qualified requirement keeps
    them as they are."""
    if "QUALIFIE" not in str(row[7] or ""):
        row[7] = "A"
    return row


def save(path, rows):
    """A requirements workbook, printed in A4 portrait."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exigences"
    ws.append(["ID", "Titre", "Specification releases", "STDHL", "IADT", "MOC", "Délégation", "Validation releases"])
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
    for col, w in zip("ABCDEFGH", (24, 60, 16, 12, 8, 9, 14, 16)):
        ws.column_dimensions[col].width = w
    ws.page_setup.paperSize, ws.page_setup.orientation = ws.PAPERSIZE_A4, "portrait"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def modify(row, field):
    """A copy of the row with one field changed."""
    row = list(row)
    if field == "titre":
        row[1] = f"{row[1]} — version modifiée"
    elif field == "STDHL":
        row[3] = f"{row[3]}-bis"
    else:
        row[2] = "B"   # validated in B too, unless qualified
        if "QUALIFIE" not in str(row[7] or ""):
            row[7] = "B"
    return row


def new_rows(k, first, count):
    """New requirements of fctk, numbered from `first`."""
    return [[f"SSS_FCT{k}_REQ_{n}", f"nouvelle exigence {n} de FCT{k}", "A", f"fct{k}", "T", "MOC4", "TOP", "A"]
            for n in range(first, first + count)]


report = []   # (file, mode, {case: [numbers]})
for k in range(1, 11):
    rng = random.Random(1500 + k)
    rows = rows_of(k)
    n = max(1, math.ceil(SHARE * len(rows)))
    ids = [r[0] for r in rows]

    # Complement: new and modified requirements only; nothing is removed by a complement.
    chosen = rng.sample(range(len(rows)), n)
    modified = {i: FIELDS[j % 3] for j, i in enumerate(sorted(chosen))}
    added = new_rows(k, 500, n)
    save(OUT / "complement" / f"2026-03-15-fct{k}-complement.xlsx",
         [modify(rows[i], f) for i, f in modified.items()] + added)
    report.append((f"2026-03-15-fct{k}-complement.xlsx", "complément",
                   {"nouvelles": [r[0] for r in added],
                    **{f"{f} modifié" if f != "release" else "release passée en B": [ids[i] for i, g in modified.items() if g == f]
                       for f in FIELDS}}))

    # Replacement: the whole tome, some removed, some modified (others than the removed ones), some new.
    picked = rng.sample(range(len(rows)), 2 * n)
    removed, changed = set(picked[:n]), {i: FIELDS[j % 3] for j, i in enumerate(sorted(picked[n:]))}
    added = new_rows(k, 600, n)
    out = []
    for i, r in enumerate(rows):
        if i in removed:
            continue
        out.append(modify(r, changed[i]) if i in changed else r)
    save(OUT / "remplacement" / f"2026-03-15-fct{k}-remplacement.xlsx", out + added)
    report.append((f"2026-03-15-fct{k}-remplacement.xlsx", "remplacement",
                   {"retirées": [ids[i] for i in sorted(removed)], "nouvelles": [r[0] for r in added],
                    **{f"{f} modifié" if f != "release" else "release passée en B": [ids[i] for i, g in changed.items() if g == f]
                       for f in FIELDS}}))

cases = ("retirées", "nouvelles", "titre modifié", "STDHL modifié", "release passée en B")
lines = ["# Fichiers de complément et de remplacement", "",
         "Prérequis : les exigences de test `exigences/2026-01-01-fct1.xlsx` à `fct10.xlsx` sont chargées dans la base "
         "(tomes fct1 à fct10, release A). Chaque fichier part des exigences de la release A de son tome.", "",
         "- `complement/` : fichiers à importer en complément. Ils ne contiennent que les exigences nouvelles (5 % du "
         "tome) et modifiées (5 %). Dans le rapport, regarder la partie « Si le fichier modifie la base ».",
         "- `remplacement/` : fichiers à importer en remplacement. Ils contiennent tout le tome, moins 5 % d'exigences "
         "retirées, plus 5 % de nouvelles et 5 % de modifiées. Dans le rapport, regarder la partie « Si le fichier "
         "écrase la base ».", "",
         "Le rapport donne toujours les deux cas : un fichier de complément qui écraserait la base en retirerait les "
         "95 % d'exigences qu'il ne contient pas, ce qu'il doit signaler.", "",
         "Une exigence modifiée change un seul champ, à tour de rôle : le titre (« — version modifiée » ajouté), le "
         "STDHL (« -bis » ajouté) ou la release (A devient B). Une exigence passée en B sort de la release de la base : "
         "le rapport la signale dans un groupe hors release A (À régler) ; en remplacement, elle est aussi comptée parmi "
         "les exigences absentes du fichier. Les exigences nouvelles n'ont pas de ligne de plan : la couverture du plan "
         "baisse d'autant.", "",
         "Le rapport doit donner, pour chaque fichier, les numéros ci-dessous : « exigences nouvelles », « exigences de la "
         "base au titre modifié », « exigences de la base au STDHL modifié », le groupe « release B » hors release et, "
         "s'il écrase la base, « exigences du tome absentes du fichier » (les retirées et celles passées en B).", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_modifications.py ../../base-test` depuis "
         "`nexus-v/programme`.", "",
         "| Fichier | Mode | " + " | ".join(c.capitalize() for c in cases) + " |",
         "| --- | --- | " + " | ".join("---" for _ in cases) + " |"]
for f, mode, found in report:
    lines.append(f"| {f} | {mode} | " + " | ".join(str(len(found[c])) if c in found else "—" for c in cases) + " |")
lines += ["", "## Numéros par fichier", ""]
for f, mode, found in report:
    lines.append(f"**{f}** ({mode})")
    lines.append("")
    lines += [f"- {c} : {', '.join(found[c])}" for c in cases if found.get(c)]
    lines.append("")
write_html(OUT / "LISEZMOI-complement-remplacement.html", lines)
print(f"{len(report)} fichiers écrits dans {OUT}")
