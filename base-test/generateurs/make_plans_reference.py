"""Test acceptance plans plans/2026-01-01_PA_FCTk_A-plan.docx made consistent with the reference of the requirements.

The requirements file is the reference (columns IADT, MOC and Délégation of exigences/2026-01-01-fctk.xlsx, see
make_exigences_reference.py); the plan is only checked against it. Each plan document is rewritten in place, its
titles, themes and layout kept:
- the cells IADT, MOC and Délégation of a requirement row take the values of its reference (the delegation always
  written, TOP included);
- a requirement on several rows keeps its first row, the others are removed (one reference, one row).
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_plans_reference.py ../../base-test
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl

sys.path.insert(0, ".")
from configuration import Configuration

BASE = Path(sys.argv[1])
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only
number_regex = re.compile(cfg.get("requirement_regex"))


def text_of(xml):
    # Paragraphs joined by a space: « SSS_FCT1_REQ_001 » then its title must stay two words.
    return " ".join("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", para))
                    for para in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S))


def set_text(cell, value):
    """The cell with `value` as its only text (first run kept, the others emptied)."""
    found = [False]

    def put(m):
        if found[0]:
            return m.group(1) + "</w:t>"
        found[0] = True
        return m.group(1) + escape(value) + "</w:t>"
    cell = re.sub(r"(<w:t(?: [^>]*)?>)[^<]*</w:t>", put, cell)
    if not found[0]:   # an empty cell: a run in its first paragraph
        run = f'<w:r><w:t xml:space="preserve">{escape(value)}</w:t></w:r>'
        cell = re.sub(r"<w:p(\s[^>]*)?/>", lambda m: f"<w:p{m.group(1) or ''}>{run}</w:p>", cell, count=1) \
            if "</w:p>" not in cell else cell.replace("</w:p>", run + "</w:p>", 1)
    return cell


for k in range(1, 11):
    ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
    header = [c.value for c in ws[1]]
    at = header.index("IADT")
    reference = {r[0]: r[at:at + 3] for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}
    path = BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx"
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    seen, changed, removed = set(), 0, 0
    for n, (info, data) in enumerate(items):
        if info.filename != "word/document.xml":
            continue
        xml = data.decode("utf-8")

        def fix(m):
            global changed, removed
            row = m.group(0)
            cells = re.findall(r"<w:tc>.*?</w:tc>|<w:tc .*?</w:tc>", row, flags=re.S)
            if len(cells) < 5:
                return row
            numbers = number_regex.findall(text_of(cells[0]))
            if not numbers or numbers[0] not in reference:
                return row
            if numbers[0] in seen:
                removed += 1
                return ""
            seen.add(numbers[0])
            values = reference[numbers[0]]
            new = list(cells)
            for i, v in zip((2, 3, 4), values):
                if text_of(cells[i]).strip() != str(v):
                    new[i] = set_text(cells[i], str(v))
            if new != cells:
                changed += 1
                for old, cell in zip(cells, new):
                    row = row.replace(old, cell, 1)
            return row
        xml = re.sub(r"<w:tr[ >].*?</w:tr>", fix, xml, flags=re.S)
        items[n] = (info, xml.encode("utf-8"))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for info, data in items:
            w.writestr(info, data)
    print(f"FCT{k} : {len(seen)} exigences, {changed} lignes mises à la référence, {removed} lignes en double retirées")
