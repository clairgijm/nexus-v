"""Requirement files to try the check before import (loupe of the Exigences box), one case per file.

They start from the test requirements exigences/2026-01-01-fctk.xlsx, which must be loaded in the base (tomes fct1 to
fct10, release A). Each file triggers one or a few checks; LISEZMOI.html, written next to them, gives the expected
result of each one.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_verification_exigences.py ../../base-test
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, next to this script
from lisezmoi import write_html

BASE = Path(sys.argv[1])
OUT = BASE / "verification-import" / "exigences"
OUT.mkdir(parents=True, exist_ok=True)
HEADER = ["ID", "Titre", "Specification releases", "STDHL", "IADT", "MOC", "Délégation", "Validation releases"]
REF = ["T", "MOC4", "TOP", "A"]   # reference and validation release of a new requirement
in_a = lambda r: "A" in str(r[2] or "").replace(" ", "").split(",")


def rows_of(k):
    """Rows (ID, Titre, Specification releases, STDHL, IADT, MOC, Délégation, Validation releases) of release A of the test file of fctk: the tome as the base holds it, so that
    each file only shows the case it is made for (the rows of releases B or C alone would add groups outside the
    release to every file)."""
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    return [list(r[:8]) for r in ws.iter_rows(min_row=2, values_only=True) if r[0] and in_a(r)]


def save(name, rows, header=HEADER, extra_sheet=None):
    """A workbook: header in row 1, then the rows; optionally a second sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exigences"
    ws.append(header)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
    for col, w in zip("ABCDEFGH", (24, 60, 16, 12, 8, 9, 14, 16)):
        ws.column_dimensions[col].width = w
    if extra_sheet:
        other = wb.create_sheet(extra_sheet[0])
        for r in extra_sheet[1]:
            other.append(r)
    for sheet in wb.worksheets:   # printed in A4 portrait
        sheet.page_setup.paperSize, sheet.page_setup.orientation = sheet.PAPERSIZE_A4, "portrait"
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(OUT / name)


cases = []   # (file, what it holds, expected result)

# 1. Same content as the tome in the base: nothing to report.
save("2026-03-01-fct1-identique.xlsx", rows_of(1))
cases.append(("2026-03-01-fct1-identique.xlsx", "copie exacte du tome fct1 de la base",
              "Chargeable dans les deux cas, aucun problème : 0 nouvelle, 0 modifiée"))

# 2. Header « Titre » renamed: refused.
save("2026-03-01-fct2-entetes.xlsx", rows_of(2), header=["ID", "Libellé", "Specification releases", "STDHL", "IADT", "MOC", "Délégation", "Validation releases"])
cases.append(("2026-03-01-fct2-entetes.xlsx", "en-tête « Titre » renommé « Libellé »",
              "Refusé dans les deux cas : en-tête attendu introuvable « Titre » (Bloquant)"))

# 3. No requirement in release A: refused.
save("2026-03-01-fct3-sans-release-a.xlsx", [[r[0], r[1], "B, C", *r[3:]] for r in rows_of(3)])
cases.append(("2026-03-01-fct3-sans-release-a.xlsx", "toutes les lignes en releases B, C",
              "Refusé dans les deux cas : aucune exigence en release A (Bloquant), un groupe « releases B, C » hors "
              "release A, rejetées à l'import, dont 120 déjà en base (Rejeté) ; s'il écrase la base : ces 120 exigences "
              "en sont retirées et le tome fct3 les perd (À vérifier) ; s'il la modifie : la base les garde en release A"))

# 4. Repeated number, several STDHL values, numbers of another tome (different, then identical content).
rows = rows_of(4)
fct5 = [r for r in rows_of(5) if in_a(r)]
first_a = next(r for r in rows if in_a(r))
rows.append([first_a[0], "doublon du numéro avec un autre titre", "A", *first_a[3:]])
second_a = [r for r in rows if in_a(r)][1]
second_a[3] = "fct4, FCT-AUTRE"
rows.append([fct5[0][0], "exigence de fct5 avec un autre titre", "A", *fct5[0][3:]])
rows.append(list(fct5[1]))
save("2026-03-01-fct4-conflits.xlsx", rows)
cases.append(("2026-03-01-fct4-conflits.xlsx",
              f"{first_a[0]} sur deux lignes ; {second_a[0]} avec deux valeurs STDHL ; {fct5[0][0]} (tome fct5) avec "
              f"un autre titre ; {fct5[1][0]} (tome fct5) à l'identique",
              "Chargeable avec des choix à faire dans les deux cas : numéro sur plusieurs lignes et plusieurs valeurs "
              "STDHL (À régler), exigences de 2 tomes de la base, fct4 et fct5 : les 2 de fct5 à garder ou à prendre du fichier "
              "(À régler), titre et STDHL modifiés "
              "(Information) ; s'il écrase la base : les autres exigences du tome fct5 absentes du fichier (À vérifier)"))

# 5. Requirements removed and added, a number out of format, a requirement without title.
rows = rows_of(5)
removed = [r[0] for r in rows if in_a(r)][:3]
rows = [r for r in rows if r[0] not in removed]
rows += [["SSS_FCT5_REQ_900", "nouvelle exigence", "A", "fct5", *REF], ["SSS_FCT5_REQ_901", "", "A", "fct5", *REF],
         ["FCT5-REQ-902", "numéro hors format", "A", "fct5", *REF]]
save("2026-03-01-fct5-modifications.xlsx", rows)
cases.append(("2026-03-01-fct5-modifications.xlsx",
              f"{', '.join(removed)} retirées ; SSS_FCT5_REQ_900 et 901 ajoutées (901 sans titre) ; FCT5-REQ-902 hors format",
              "Chargeable, avec des lignes rejetées, dans les deux cas : 1 numéro hors format, rejeté à l'import (Rejeté), "
              "1 exigence sans titre (À vérifier), 2 exigences nouvelles (Information) ; s'il écrase la base : 3 exigences du tome fct5 absentes du fichier "
              "(À vérifier) ; la couverture du plan baisse (exigences nouvelles sans ligne de plan)"))

# 6. Lines outside release A, in several groups, one without release.
rows = rows_of(6)
a_rows = [r for r in rows if in_a(r)]
a_rows[0][2], a_rows[1][2], a_rows[2][2] = "B", "C", None
save("2026-03-01-fct6-hors-release.xlsx", rows)
cases.append(("2026-03-01-fct6-hors-release.xlsx",
              f"{a_rows[0][0]} passée en B, {a_rows[1][0]} en C, {a_rows[2][0]} sans release",
              "Chargeable, avec des lignes rejetées, dans les deux cas : 3 groupes hors release A, « release B », "
              "« release C » et « sans release », rejetés à l'import, chacun avec 1 exigence déjà en base (Rejeté) : s'il "
              "écrase la base, elles en sont retirées (et le tome fct6 les perd, À vérifier) ; s'il la modifie, la base "
              "les garde en release A"))

# 7. Workbook of two sheets.
save("2026-03-01-fct7-deux-feuilles.xlsx", rows_of(7), extra_sheet=("Historique", [["Version", "Date"], ["1", "2026-03-01"]]))
cases.append(("2026-03-01-fct7-deux-feuilles.xlsx", "classeur avec une deuxième feuille « Historique »",
              "Chargeable dans les deux cas : classeur de 2 feuilles, la première vérifiée (Information)"))

# 8. No tome in the file name.
save("exigences-sans-tome.xlsx", rows_of(8))
cases.append(("exigences-sans-tome.xlsx", "exigences du tome fct8, nom de fichier sans date ni tome",
              "Chargeable dans les deux cas, aucun problème : le nom du fichier n'est pas utilisé, les exigences sont "
              "reconnues dans le tome fct8 de la base"))

# 9 and 10. A new tome, in two files sharing numbers: checked together.
new = [[f"SSS_FCT11_REQ_{n:03d}", f"exigence {n} du tome fct11", "A", "fct11", *REF] for n in range(1, 21)]
save("2026-03-01-fct11.xlsx", new)
save("2026-03-01-fct11-suite.xlsx", new[15:] + [[f"SSS_FCT11_REQ_{n:03d}", f"exigence {n} du tome fct11", "A", "fct11", *REF]
                                                for n in range(21, 31)])
cases.append(("2026-03-01-fct11.xlsx et 2026-03-01-fct11-suite.xlsx",
              "nouveau tome fct11 en deux fichiers, 5 numéros (016 à 020) dans les deux",
              "Chacun chargeable : toutes ses exigences sont nouvelles, leur tome sera à choisir à l'import "
              "(Information) ; vérifiés ensemble : 5 exigences dans plusieurs fichiers (À régler, entre les fichiers)"))

# 11. Wrong reference: an empty delegation, an IADT unknown of the configuration, two MOC codes.
rows = rows_of(9)
a_rows = [r for r in rows if in_a(r)]
a_rows[0][6], a_rows[1][4], a_rows[2][5] = None, "X", "MOC2, MOC3"
save("2026-03-01-fct9-reference.xlsx", rows)
cases.append(("2026-03-01-fct9-reference.xlsx",
              f"{a_rows[0][0]} sans délégation, {a_rows[1][0]} en IADT « X », {a_rows[2][0]} avec deux MOC « MOC2, MOC3 »",
              "Refusé dans les deux cas : 3 erreurs dans la référence des exigences (IADT, MOC, délégation) (Bloquant)"))

# 12. Accepted: the validation changed (QUALIFIE alone, an old release with QUALIFIE, an empty cell, B only), and a
# platform absent from the configuration.
rows = rows_of(10)
a_rows = [r for r in rows if in_a(r) and r[7] == "A"]
a_rows[0][7], a_rows[1][7], a_rows[2][7], a_rows[3][7], a_rows[4][6] = "QUALIFIE", "0, QUALIFIE", None, "B", "PLT99_NOUVELLE"
save("2026-03-01-fct10-validation.xlsx", rows)
cases.append(("2026-03-01-fct10-validation.xlsx",
              f"{a_rows[0][0]} QUALIFIE et {a_rows[1][0]} « 0, QUALIFIE » (qualifiées), {a_rows[2][0]} sans Validation "
              f"releases et {a_rows[3][0]} validée en B seulement (hors scope), {a_rows[4][0]} déléguée à une plateforme "
              "absente de la configuration, PLT99_NOUVELLE",
              "Chargeable, à vérifier, dans les deux cas : nouvelle plateforme à ajouter PLT99_NOUVELLE (À vérifier) ; "
              "4 exigences à la validation modifiée et 1 à la délégation modifiée (Information)"))

lines = ["# Fichiers d'exigences pour essayer la vérification avant import", "",
         "Prérequis : les exigences de test `exigences/2026-01-01-fct1.xlsx` à `fct10.xlsx` sont chargées dans la base "
         "(tomes fct1 à fct10, release A). Vérifier ces fichiers avec la loupe de la boîte Exigences, un par un ou "
         "ensemble ; rien n'est importé.", "",
         "Chaque fichier a les colonnes ID, Titre, Specification releases, STDHL, IADT, MOC, Délégation et Validation "
         "releases : IADT, MOC et Délégation sont la référence de l'exigence (une valeur de chaque, la délégation "
         "obligatoire) ; Validation releases, qui peut être vide, donne les releases où elle est testée et/ou QUALIFIE : "
         "pour la release A de la base, l'exigence est qualifiée (QUALIFIE), à qualifier (A dans la liste) ou hors "
         "scope (sinon). "
         "Chaque fichier reprend les exigences de la release A de son tome, telles que la base les contient, puis y "
         "introduit le cas décrit ; un fichier ne signale donc que son cas.", "",
         "La vérification ne se sert pas du nom du fichier : chaque exigence est cherchée dans toute la base. Le rapport "
         "donne pour chaque fichier ce qui se passe s'il écrase la base (import en remplacement) et s'il la modifie "
         "(import en complément) ; les résultats ci-dessous valent pour les deux cas, sauf mention.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_exigences.py ../../base-test` depuis "
         "`nexus-v/programme`.", "",
         "| Fichier | Contenu | Résultat attendu |", "| --- | --- | --- |"]
lines += [f"| {f} | {what} | {expected} |" for f, what, expected in cases]
write_html(OUT / "LISEZMOI.html", lines)
print(f"{len(list(OUT.glob('*.xlsx')))} fichiers écrits dans {OUT}")
