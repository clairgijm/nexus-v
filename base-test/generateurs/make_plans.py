"""Plans d'acceptation de test : plans/2026-01-01_PA_FCTk_A-plan.docx, pour k de 1 à 10.

Chaque plan est construit à partir du fichier d'exigences de son tome (exigences/2026-01-01-fctk.xlsx), qui est la
référence : une ligne par exigence de la release de la base (colonne « Specification releases »), avec son IADT, son
MOC, sa délégation et sa validation. La validation est écrite avec les textes de la configuration
(plan_validation_to_qualify, plan_validation_qualified, plan_validation_out_of_scope), d'après la colonne « Validation
releases » (QUALIFIE : déjà qualifiée ; la release de la base listée : à qualifier ; sinon hors scope). Chaque plan est
donc identique aux exigences de son tome : il est accepté à l'import.

Les tableaux ont les colonnes Exigence, Méthodes de test V&V, IADT, MOC, Délégation, Validation. Documents Word en A4
portrait, sur les styles de stdhl/2026-11-11-STDHL-template2.docx.
Usage (depuis nexus-v/programme) : python ../../base-test/generateurs/make_plans.py ../../base-test
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, ".")
from configuration import Configuration, fold
from requirements_db import RequirementsDatabase

BASE = Path(sys.argv[1])
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # lecture seule : textes et séparateurs
RELEASE = "A"
HEADER = ["Exigence", "Méthodes de test V&V", "IADT", "MOC", "Délégation", "Validation"]
VALIDATION_TEXTS = {"to_qualify": cfg.get("plan_validation_to_qualify"), "qualified": cfg.get("plan_validation_qualified"),
                    "out_of_scope": cfg.get("plan_validation_out_of_scope")}


def p(text, style=""):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="1600"/>' for _ in rows[0])
    body = "".join("<w:tr>" + "".join(f"<w:tc>{p(c)}</w:tc>" for c in r) + "</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>")


def write_plan(path, title, heading, rows):
    """Un plan : un titre, un intertitre et le tableau du plan ; A4 portrait."""
    body = [p(title, "Titre"), p(heading, "Titre2"), table([HEADER] + rows)]
    with zipfile.ZipFile(SKELETON) as z, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "word/document.xml":
                x = data.decode("utf-8")
                a, b = x.index("<w:body>") + len("<w:body>"), x.rindex("<w:sectPr")
                x = x[:a] + "".join(body) + x[b:]
                data = re.sub(r"<w:pgSz [^>]*/>", '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>', x).encode("utf-8")
            w.writestr(info, data)


def requirements_of(k):
    """Exigences du fichier de fctk qui sont importées (release de la base, numéro au bon format), avec leur
    référence, dans l'ordre du fichier."""
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    rows = ws.iter_rows(values_only=True)
    head = [fold(str(h or "")).strip() for h in next(rows)]
    col = lambda name: next(i for i, h in enumerate(head) if fold(name) in h)   # en-tête qui contient le texte
    c = {f: col(n) for f, n in (("id", "ID"), ("title", "Titre"), ("releases", "Specification releases"),
                                ("validation", "Validation releases"), ("iadt", "IADT"), ("moc", "MOC"),
                                ("delegation", "Délégation"))}
    number = re.compile(cfg.get("requirement_regex"))
    separator = cfg.get("release_separator_regex")
    found, seen = [], set()
    for r in rows:
        rid = str(r[c["id"]] or "").strip()
        if not rid or not number.fullmatch(rid) or rid in seen:
            continue
        releases = [x.strip().upper() for x in re.split(separator, str(r[c["releases"]] or "")) if x.strip()]
        if RELEASE not in releases:
            continue
        seen.add(rid)
        values = RequirementsDatabase.validation_values(str(r[c["validation"]] or ""), separator)
        state = RequirementsDatabase.validation_state(values, RELEASE)
        found.append({"id": rid, "title": str(r[c["title"]] or ""), "iadt": str(r[c["iadt"]] or "").strip(),
                      "moc": str(r[c["moc"]] or "").strip(), "delegation": str(r[c["delegation"]] or "").strip(),
                      "validation": VALIDATION_TEXTS[state]})
    return found


for k in range(1, 11):
    reqs = requirements_of(k)
    rows = [[f"{r['id']} {r['title']}", f"V&V [{r['delegation']}] : méthode de test et moyens", r["iadt"], r["moc"],
             r["delegation"], r["validation"]] for r in reqs]
    path = BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx"
    try:
        write_plan(path, f"Plan d'acceptation FCT{k}", f"FCT{k} — release {RELEASE}", rows)
    except PermissionError:   # document ouvert dans Word
        print(f"{path.name} : NON ÉCRIT, le fichier est ouvert (le fermer puis relancer)")
        continue
    print(f"{path.name} : {len(rows)} exigences")
