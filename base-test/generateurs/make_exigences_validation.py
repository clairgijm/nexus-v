"""« Specification releases » and « Validation releases » of the test requirements exigences/2026-01-01-fctk.xlsx.

The column of the releases is named « Specification releases » (the releases where the requirement applies); the
column « Validation releases » (the releases where it is tested, or QUALIFIE) is added after Délégation, or rewritten
when it is there. Its values, drawn for each requirement (the same at each run):
- 6 % QUALIFIE alone (qualified, no release given);
- 6 % its first specification release then QUALIFIE (« A, QUALIFIE »: qualified, last tested in A);
- 8 % of the requirements of several specification releases: those releases but the first (« A, B » gives « B »: not
  tested in A, out of scope);
- 4 % empty (out of scope of the validation);
- the others: their specification releases.
For a base in release A, the validation of a requirement is then « qualifiée » with QUALIFIE, « à qualifier » when A
is in the list, « hors scope » otherwise.
Not to be run again: the requirements files are now the reference (their N/A were converted by
make_exigences_sans_na.py; this script, made from the old plans, would bring them back).
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_exigences_validation.py ../../base-test
"""
import random
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

BASE = Path(sys.argv[1])
SPEC, VALIDATION = "Specification releases", "Validation releases"
for k in range(1, 11):
    path = BASE / "exigences" / f"2026-01-01-fct{k}.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    spec = next(i for i, h in enumerate(header) if h and str(h).strip().lower().startswith(("release", "specification")))
    ws.cell(row=1, column=spec + 1, value=SPEC)
    at = header.index(VALIDATION) + 1 if VALIDATION in header else len([h for h in header if h]) + 1
    cell = ws.cell(row=1, column=at, value=VALIDATION)
    cell.font, cell.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
    ws.column_dimensions[openpyxl.utils.get_column_letter(at)].width = 18
    rng = random.Random(f"validation-{k}")
    counts = {"spécification": 0, "QUALIFIE": 0, "release, QUALIFIE": 0, "sans la première": 0, "vide": 0}
    for row in ws.iter_rows(min_row=2):
        if not row[0].value:
            continue
        releases = [r.strip().upper() for r in str(row[spec].value or "").split(",") if r.strip()]
        draw = rng.random()
        if draw < 0.06:
            value, kind = "QUALIFIE", "QUALIFIE"
        elif draw < 0.12 and releases:
            value, kind = f"{releases[0]}, QUALIFIE", "release, QUALIFIE"
        elif draw < 0.20 and len(releases) > 1:
            value, kind = ", ".join(releases[1:]), "sans la première"
        elif draw >= 0.96:
            value, kind = None, "vide"
        else:
            value, kind = ", ".join(releases) or "QUALIFIE", "spécification"
        counts[kind] += 1
        ws.cell(row=row[0].row, column=at, value=value)
    wb.save(path)
    print(f"fct{k} :", ", ".join(f"{n} {kind}" for kind, n in counts.items()))
