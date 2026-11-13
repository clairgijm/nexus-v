"""Fiches d'essai XML of the format « campagne » (export TestSuite / TestCase / TestStep of the campaign tool), from the
reference STDHL « unique » stdhl/unique/2026-01-01_TEST_STDHL_UNIQUE_rev1.docx: one file per tome (FCT1 .. FCT10),
the same tests (TestSuite name) and the same test cases (TestCase Title = the full name of the test case of the
STDHL), one TestStep per requirement of a test case, citing it in VerifiedRequirements.

As for the fiches of the format « balises », 95 % of the requirements of each tome are covered (the others left
uncovered for the tests). Results: 10 % of the test cases not run (no RunResult in their steps); the steps of the
others 95 % PASSED, 5 % FAILED. A test case is FAILED if one of its steps failed, PASSED if its steps run all passed,
NOT_RUN without any step run.

Written into fiches/campagne/2026-01-01_TEST_FE_CAMPAGNE_FCTk_rev1.xml, with fiches/campagne/LISEZMOI.html.
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_fiches_campagne.py ../../base-test
"""
import random
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, ".")
sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, next to this script
from lisezmoi import write_html
from stdhl_db import StdhlDatabase

BASE = Path(sys.argv[1])
SOURCE = BASE / "stdhl" / "unique" / "2026-01-01_TEST_STDHL_UNIQUE_rev1.docx"
OUT = BASE / "fiches" / "campagne"
COVERED, NOT_RUN, FAILED = 0.95, 0.10, 0.05   # share of the test cases not run; of the steps failed
DATE = "2026-09-24"


def prop(name, value=""):
    return f'<Property name="{name}">{escape(value)}</Property>' if value else f'<Property name="{name}"/>'


def properties(indent, items):
    pad = " " * indent
    return [f"{pad}<Properties>"] + [f"{pad}    {prop(n, v)}" for n, v in items] + [f"{pad}</Properties>"]


tests = StdhlDatabase.read_text_format(SOURCE.read_bytes(), SOURCE.name)["tests"]
OUT.mkdir(parents=True, exist_ok=True)
lines_readme = ["# Fiches d'essai de test au format campagne", "",
                f"Générées par `generateurs/make_fiches_campagne.py` depuis `stdhl/unique/{SOURCE.name}` : mêmes tests "
                "(TestSuite) et mêmes test cases (titre complet du test case du STDHL), un TestStep par exigence. "
                f"{round(COVERED * 100)} % des exigences de chaque tome sont couvertes ; {round(NOT_RUN * 100)} % des test cases "
                f"ne sont pas exécutés (aucun RunResult), les steps des autres sont {100 - round(FAILED * 100)} % PASSED et "
                f"{round(FAILED * 100)} % FAILED. Un test case est FAILED si l'un de ses steps a échoué, PASSED si ses "
                "steps exécutés sont tous passés, NOT_RUN si aucun n'est exécuté.", "",
                "À charger avec le STDHL unique de référence (ou les STDHL par tome) : chaque fichier est importable.", "",
                "| Fichier | Tests | Test cases | Exigences | PASSED | FAILED | NOT_RUN |", "| --- | --- | --- | --- | --- | --- | --- |"]
for k in range(1, 11):
    tome = f"FCT{k}"
    own = [t for t in tests if re.search(rf"_{tome}_", t["name"])]
    ids = sorted({i for t in own for c in t["cases"] for i in c["requirements"]})
    rng = random.Random(f"campagne-{tome}")
    kept = set(rng.sample(ids, round(len(ids) * COVERED)))
    uid = 0
    out = ['<?xml version="1.0" encoding="UTF-8"?>', '<Project name="NEXUS-V tests" path="base-test">', "    <Tcr>",
           '        <Campaign name="Roadmap A">', "            <TestSuites>"]
    counts = {"PASSED": 0, "FAILED": 0, "NOT_RUN": 0}
    n_cases, n_tests = 0, 0
    for t in own:
        cases = [(c["name"], [i for i in dict.fromkeys(c["requirements"]) if i in kept]) for c in t["cases"]]
        cases = [(name, req) for name, req in cases if req]
        if not cases:
            continue
        n_tests += 1
        out.append(f'                <TestSuite name="{escape(t["name"])}">')
        out += properties(20, [("Title", t["name"]), ("RunDate", DATE), ("Method", "T"), ("VerifiedRequirements", "")])
        out.append("                    <TestCases>")
        for name, req in cases:
            uid += 1
            n_cases += 1
            out.append(f'                        <TestCase uid="{uid}">')
            out += properties(28, [("Title", name), ("RunDate", DATE), ("RunningStatus", "RUN"),
                                   ("VerifiedRequirements", "")])
            out.append("                            <TestSteps>")
            runs = []
            run = rng.random() >= NOT_RUN
            for i in req:
                uid += 1
                result = "" if not run else "FAILED" if rng.random() < FAILED else "PASSED"
                runs.append(result)
                out.append(f'                                <TestStep uid="{uid}">')
                out += properties(36, [("Description", f"Vérifier {i}"), ("VerifiedRequirements", i),
                                       ("ExpectedResult", "Comportement conforme à l'exigence"),
                                       ("RunDate", DATE if result else ""), ("RunResult", result),
                                       ("StepType", "NORMAL")])
                out.append("                                </TestStep>")
            out.append("                            </TestSteps>")
            out.append("                        </TestCase>")
            counts["FAILED" if "FAILED" in runs else "PASSED" if any(runs) else "NOT_RUN"] += 1
        out.append("                    </TestCases>")
        out.append("                </TestSuite>")
    out += ["            </TestSuites>", "        </Campaign>", "    </Tcr>", "</Project>"]
    name = f"2026-01-01_TEST_FE_CAMPAGNE_{tome}_rev1.xml"
    (OUT / name).write_text("\n".join(out) + "\n", encoding="utf-8")
    lines_readme.append(f"| {name} | {n_tests} | {n_cases} | {len(kept)} | {counts['PASSED']} | {counts['FAILED']} "
                        f"| {counts['NOT_RUN']} |")
lines_readme += ["", "Les colonnes PASSED, FAILED et NOT_RUN comptent des test cases."]
write_html(OUT / "LISEZMOI.html", lines_readme)
print("\n".join(lines_readme[-14:]))
