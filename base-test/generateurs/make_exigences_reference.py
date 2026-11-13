"""Reference of the test requirements: the columns IADT, MOC and Délégation of exigences/2026-01-01-fctk.xlsx.

The requirements file is the reference of each requirement: one IADT code, one MOC code and one delegation (always
written: an empty delegation refuses the file). The values are taken from the test acceptance plans
plans/2026-01-01_PA_FCTk_A-plan.docx, so that the plans stay consistent with the requirements:
- a requirement with one plan row takes its IADT, MOC and delegation (an empty delegation of the plan becomes the
  default platform, TOP);
- a requirement with several plan rows (its codes joined, « MOC2, MOC3 ») takes the first value of each (the plan
  check then reports the difference);
- a requirement absent from the plans (rows of releases B or C only) gets T, MOC4, TOP.
The columns are added after STDHL, or replaced when they are there already.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_exigences_reference.py ../../base-test
"""
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

sys.path.insert(0, ".")
from configuration import Configuration
from plans_db import PlansDatabase

BASE = Path(sys.argv[1])
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only
DEFAULT = "TOP"
COLUMNS = ["IADT", "MOC", "Délégation"]
summary = []
for k in range(1, 11):
    plan = PlansDatabase.read_docx((BASE / "plans" / f"2026-01-01_PA_FCT{k}_A-plan.docx").read_bytes(), cfg)
    reference, several = {}, 0
    for e in PlansDatabase.entries(plan["rows"]):
        if len(re.split(r"[,;+\s]+", (e["moc"] or "").strip())) > 1 or len(re.split(r"[,;+\s]+", (e["iadt"] or "").strip())) > 1:
            several += 1
        # One value of each: the first one written (a requirement on several plan rows has its codes joined).
        first = lambda v: re.split(r"[,;+\s]+", (v or "").strip())[0]
        reference[e["requirement"]] = (first(e["iadt"]), first(e["moc"]), first(e["delegation"]) or DEFAULT)
    path = BASE / "exigences" / f"2026-01-01-fct{k}.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    at = header.index("IADT") + 1 if "IADT" in header else len([h for h in header if h]) + 1
    for i, name in enumerate(COLUMNS):
        cell = ws.cell(row=1, column=at + i, value=name)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
    absent = 0
    for row in ws.iter_rows(min_row=2):
        number = str(row[0].value or "").strip()
        if not number:
            continue
        values = reference.get(number)
        if values is None:
            values, absent = ("T", "MOC4", DEFAULT), absent + 1
        for i, v in enumerate(values):
            ws.cell(row=row[0].row, column=at + i, value=v)
    for i, width in enumerate((8, 9, 14)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(at + i)].width = width
    wb.save(path)
    summary.append(f"fct{k} : {len(reference)} exigences du plan, {several} à plusieurs codes dans le plan (premier gardé), "
                   f"{absent} absentes du plan (T, MOC4, TOP)")
print("\n".join(summary))
