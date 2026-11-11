"""Fiches d'Essai XML (platform TOP) from the test STDHL FCT1 .. FCT10: same tests and test cases, 95 % of the
requirements of each STDHL (the others left uncovered for the tests)."""
import random, sys
sys.path.insert(0, ".")
from pathlib import Path
from xml.sax.saxutils import escape
import re
from configuration import DEFAULT_CONFIG
from file_database import WordDocument
from fiches_db import FichesDatabase
from stdhl_db import StdhlDatabase


def stdhl_rows(content):
    """Rows of a test STDHL with their test case: [{fiche, test_case, requirements}] (the STDHL reading of NEXUS-V
    keeps the fiche only)."""
    headers, regex = StdhlDatabase._settings(None)
    fiche_regex = re.compile(DEFAULT_CONFIG["fiche_regex"])
    rows, fiche, case = [], "", ""
    for kind, number, value in WordDocument(content).blocks():
        if kind == "p":
            if m := fiche_regex.search(value):
                fiche, case = (m[1] if m.re.groups else m[0]).strip(), ""
            elif value.strip().upper().startswith("TEST CASE"):
                case = value.strip()
            continue
        header = StdhlDatabase._find_header(value, headers)
        if header is None or header[1]:
            continue
        col = header[0]
        for cells in value[1:]:
            ids = list(dict.fromkeys(regex.findall(cells[col]))) if col < len(cells) else []
            if ids:
                rows.append({"fiche": fiche, "test_case": case, "requirements": ids})
    return rows

base = Path(sys.argv[1])
for fct in [f"FCT{k}" for k in range(1, 11)]:
    src = base / "stdhl" / "perTome" / f"2026-01-01_TEST_STDHL_{fct}_rev1.docx"
    out = base / "fiches" / f"2026-01-01_TEST_FICHES_{fct}_rev1.xml"
    rows = stdhl_rows(src.read_bytes())
    ids = sorted({i for r in rows for i in r["requirements"]})
    rng = random.Random(fct)
    kept = set(rng.sample(ids, round(len(ids) * .95)))
    tests = {}
    for r in rows:
        tests.setdefault(r["fiche"], {}).setdefault(r["test_case"], []).extend(i for i in r["requirements"] if i in kept)
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<fiches>"]
    for test, cases in tests.items():
        cases = {c: list(dict.fromkeys(v)) for c, v in cases.items() if v}
        if not cases:
            continue
        lines.append(f"  <test>{escape(test)}")
        for case, values in cases.items():
            lines.append(f"    <test_case>{escape(case)}")
            lines.append(f"      <requirements>{escape(', '.join(values))}</requirements>")
            lines.append("    </test_case>")
        lines.append("  </test>")
    lines.append("</fiches>")
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    back = FichesDatabase.read_xml(out.read_bytes())
    found = set(back["ids"])
    assert found == kept and found <= set(ids)
    print(f"{out.name} : {len(found)} / {len(ids)} = {100 * len(found) / len(ids):.1f} % | {len({r['test'] for r in back['rows']})} tests, {len(back['rows'])} test cases | absentes : {', '.join(sorted(set(ids) - found))}")
