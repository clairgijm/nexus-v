"""Jeu de test « couverture complète » : un STDHL « unique » et des fiches d'essai alignées, pour obtenir 100 % aux
tuiles STDHL et Fiches d'Essai de la vérification.

Les exigences sont lues dans exigences/2026-01-01-fctk.xlsx (k de 1 à 10), avec leur validation dans la release A
(comme à l'import) :
- une exigence qualifiée ou hors scope n'est attendue ni dans un STDHL ni dans une fiche : elle est laissée de côté ;
- STDHL : chaque exigence attendue déléguée à la plateforme de niveau système (TOP) est citée par un test. Les tests
  reprennent ceux des STDHL par tome (stdhl/perTome) ; les exigences qu'ils ne citent pas vont dans un test
  SYS_FCTk_COMPLEMENT_900 par tome. Chaque test case a de 1 à 4 phases, chacune citant une exigence : la liste
  « Exigences associées » d'un test est exactement celle de ses test cases ;
- fiches : un fichier au format « campagne » par plateforme de délégation, qui couvre toutes les exigences attendues
  qui lui sont déléguées. Celui de TOP reprend exactement les tests (TestSuite) et les test cases (titre du test
  case du STDHL) ; ceux des autres plateformes ont un test par tome. 5 % des exigences attendues (tirage
  reproductible) sont FAILED : chacune est seule dans son test case (dans le STDHL comme dans les fiches), pour que
  son échec ne touche pas les autres exigences ; les autres sont PASSED. Le résultat ne change pas la couverture.

Écrit dans stdhl/unique/2026-01-01_TEST_STDHL_COMPLET_rev1.docx et fiches/complet/ (un fichier par plateforme, avec
LISEZMOI.html). À charger dans une base vide : les exigences des 10 tomes, les plans (plans/), le STDHL complet (sans
tome), puis chaque fichier de fiches sans tome, avec sa plateforme (celle du nom du fichier).
Usage (depuis nexus-v/programme) : python ../../base-test/generateurs/make_couverture_complete.py ../../base-test
"""
import random
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, ".")
sys.path.insert(0, str(Path(__file__).resolve().parent))   # lisezmoi.py, à côté de ce script
from lisezmoi import write_html
from configuration import Configuration
from requirements_db import RequirementsDatabase
from stdhl_db import StdhlDatabase

BASE = Path(sys.argv[1])
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
STDHL_OUT = BASE / "stdhl" / "unique" / "2026-01-01_TEST_STDHL_COMPLET_rev1.docx"
FICHES_OUT = BASE / "fiches" / "complet"
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # lecture seule
RELEASE = "A"
SYSTEM = "TOP"            # plateforme de niveau système : celle dont les exigences sont spécifiées dans le STDHL
PHASES_PER_CASE = 4       # phases (une exigence chacune) au plus par test case
DATE = "2026-09-24"
FAILED = 0.05             # part des exigences attendues FAILED (tirage reproductible) ; les autres sont PASSED
# Le style Titre3 du modèle met son texte hors de la page (voir make_stdhl_format2.py) : chaque titre l'annule.
TITLE3_FIX = ('<w:tabs><w:tab w:val="clear" w:pos="12137"/><w:tab w:val="num" w:pos="1080"/></w:tabs>'
              '<w:ind w:left="0" w:firstLine="0"/>')


# ------------------------------------------------------------------ exigences attendues
def requirements():
    """Exigences des 10 tomes : [{id, title, tome, delegation, validation}] (validation dans la release A)."""
    found = []
    for k in range(1, 11):
        name = f"2026-01-01-fct{k}.xlsx"
        data = RequirementsDatabase.read_excel((BASE / "exigences" / name).read_bytes(), None, cfg)
        for r in data["requirements"]:
            state = RequirementsDatabase.validation_state(r.get("validation_releases", []), RELEASE)
            found.append({"id": r["id"], "title": r["title"], "tome": k, "delegation": r["delegation"] or "",
                          "validation": state})
    return found


REQS = requirements()
EXPECTED = [r for r in REQS if r["validation"] not in ("qualified", "out_of_scope")]
TOP = {r["id"] for r in EXPECTED if r["delegation"].casefold() == SYSTEM.casefold()}
TITLES = {r["id"]: r["title"] for r in REQS}


def stdhl_tests(k):
    """Tests du STDHL par tome de fctk réduits aux exigences de TOP attendues, plus un test complément pour celles
    qu'aucun test ne cite : [(test, [exigences])]."""
    name = f"2026-01-01_TEST_STDHL_FCT{k}_rev1.docx"
    data = StdhlDatabase.read_docx((BASE / "stdhl" / "perTome" / name).read_bytes(), name, cfg)
    tests = {}
    for row in data["rows"]:
        for i in row["requirements"]:
            if i in TOP and i not in tests.setdefault(row["fiche"], []):
                tests[row["fiche"]].append(i)
    tests = {t: reqs for t, reqs in tests.items() if reqs}
    cited = {i for reqs in tests.values() for i in reqs}
    missing = [r["id"] for r in EXPECTED if r["tome"] == k and r["id"] in TOP and r["id"] not in cited]
    if missing:
        tests[f"SYS_FCT{k}_COMPLEMENT_900"] = missing
    return list(tests.items())


def layout(reqs):
    """Test cases d'un test : chaque exigence FAILED seule dans le sien, les autres par PHASES_PER_CASE."""
    return [[i] for i in reqs if i in FAILED_IDS] + chunks([i for i in reqs if i not in FAILED_IDS], PHASES_PER_CASE)


def chunks(values, size):
    return [values[i:i + size] for i in range(0, len(values), size)]


def case_name(c):
    """Nom complet du test case n° c, tel que le STDHL l'écrit et que les fiches le reprennent."""
    return f"TEST CASE#{c} : cas {c}"


# ------------------------------------------------------------------ STDHL unique
def p(text, style=""):
    fix = TITLE3_FIX if style == "Titre3" else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>{fix}</w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def test_block(name, reqs):
    """Paragraphes d'un test : titre, État, Objectif, Applicabilité, exigences associées, test cases et phases."""
    out = [p(f"{name} : vérification complète", "Titre3"), p("État : Revue"),
           p("Objectif : vérifier les exigences du test"), p(f"Applicabilité : RELEASE {RELEASE}"),
           p("Exigences associées :")]
    out += [p(f"{i} : {TITLES.get(i, '')}", "Paragraphedeliste") for i in reqs]
    out += [p("Conditions et configuration initiales :"), p("Banc de test en configuration nominale", "Paragraphedeliste"),
            p("DESCRIPTION DE LA FICHE DE TEST")]
    for c, phases in enumerate(layout(reqs), 1):
        out.append(p(case_name(c)))
        out += [p(f"Phase {j} : vérifier l'exigence : {i}", "Paragraphedeliste") for j, i in enumerate(phases, 1)]
    return out


def write_stdhl(path, tests):
    body = [p("STDHL de tous les tomes — couverture complète", "Titre"), p("Contents", "En-ttedetabledesmatires")]
    body += [p(f"2.{k} {t} : test {k}", "TM3") for k, (t, _) in enumerate(tests, 1)]
    body += [p("Description des tests", "Titre1")]
    for t, reqs in tests:
        body += test_block(t, reqs)
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


# ------------------------------------------------------------------ fiches au format campagne
def prop(name, value=""):
    return f'<Property name="{name}">{escape(value)}</Property>' if value else f'<Property name="{name}"/>'


def properties(indent, items):
    pad = " " * indent
    return [f"{pad}<Properties>"] + [f"{pad}    {prop(n, v)}" for n, v in items] + [f"{pad}</Properties>"]


def write_fiches(path, tests):
    """Fichier campagne : tests [(test, [test cases [exigences]])], un TestStep par exigence, FAILED pour une
    exigence de FAILED_IDS, PASSED sinon. Renvoie le nombre de steps FAILED."""
    uid, failed = 0, 0
    out = ['<?xml version="1.0" encoding="UTF-8"?>', '<Project name="NEXUS-V tests" path="base-test">', "    <Tcr>",
           '        <Campaign name="Couverture complète">', "            <TestSuites>"]
    for t, cases in tests:
        out.append(f'                <TestSuite name="{escape(t)}">')
        out += properties(20, [("Title", t), ("RunDate", DATE), ("Method", "T"), ("VerifiedRequirements", "")])
        out.append("                    <TestCases>")
        for c, reqs in enumerate(cases, 1):
            uid += 1
            out.append(f'                        <TestCase uid="{uid}">')
            out += properties(28, [("Title", case_name(c)), ("RunDate", DATE), ("RunningStatus", "RUN"),
                                   ("VerifiedRequirements", "")])
            out.append("                            <TestSteps>")
            for i in reqs:
                uid += 1
                result = "FAILED" if i in FAILED_IDS else "PASSED"
                failed += result == "FAILED"
                out.append(f'                                <TestStep uid="{uid}">')
                out += properties(36, [("Description", f"Vérifier {i}"), ("VerifiedRequirements", i),
                                       ("ExpectedResult", "Comportement conforme à l'exigence"), ("RunDate", DATE),
                                       ("RunResult", result), ("StepType", "NORMAL")])
                out.append("                                </TestStep>")
            out += ["                            </TestSteps>", "                        </TestCase>"]
        out += ["                    </TestCases>", "                </TestSuite>"]
    out += ["            </TestSuites>", "        </Campaign>", "    </Tcr>", "</Project>"]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(out) + "\n", encoding="utf-8")
    return failed


# ------------------------------------------------------------------ écriture
# Exigences FAILED : FAILED des exigences attendues, tirées une fois pour le STDHL et les fiches.
FAILED_IDS = set(random.Random("complet").sample([r["id"] for r in EXPECTED], round(len(EXPECTED) * FAILED)))
STDHL_TESTS = [x for k in range(1, 11) for x in stdhl_tests(k)]
write_stdhl(STDHL_OUT, STDHL_TESTS)
assert {i for _, reqs in STDHL_TESTS for i in reqs} == TOP

FICHES_OUT.mkdir(parents=True, exist_ok=True)
for old in FICHES_OUT.glob("*.xml"):
    old.unlink()
platforms = sorted({r["delegation"] for r in EXPECTED if r["delegation"]}, key=lambda x: (x != SYSTEM, x))
summary = []
for platform in platforms:
    if platform == SYSTEM:   # les tests et test cases du STDHL, à l'identique
        tests = [(t, layout(reqs)) for t, reqs in STDHL_TESTS]
    else:                    # un test par tome
        tests = []
        for k in range(1, 11):
            own = [r["id"] for r in EXPECTED if r["tome"] == k and r["delegation"] == platform]
            if own:
                tests.append((f"ESSAI_{platform}_FCT{k}_001", layout(own)))
    name = f"2026-01-01_TEST_FE_COMPLET_{platform}_rev1.xml"
    failed = write_fiches(FICHES_OUT / name, tests)
    covered = {i for _, cases in tests for c in cases for i in c}
    assert covered == {r["id"] for r in EXPECTED if r["delegation"] == platform}
    steps = sum(len(c) for _, cases in tests for c in cases)
    summary.append((name, platform, len(tests), sum(len(c) for _, c in tests), len(covered), steps, failed))

no_delegation = [r["id"] for r in EXPECTED if not r["delegation"]]
lines = ["# Couverture complète : STDHL unique et fiches alignées", "",
         "Générés par `generateurs/make_couverture_complete.py` depuis les exigences `exigences/2026-01-01-fctk.xlsx`, "
         f"validation dans la release {RELEASE}. Les exigences qualifiées ou hors scope ne sont attendues ni dans le "
         "STDHL ni dans les fiches : elles sont laissées de côté.", "",
         f"- `stdhl/unique/{STDHL_OUT.name}` : {len(STDHL_TESTS)} tests, les {len(TOP)} exigences attendues déléguées à "
         f"{SYSTEM} (plateforme de niveau système) ; les tests des STDHL par tome, plus un test SYS_FCTk_COMPLEMENT_900 "
         "par tome pour les exigences qu'ils ne citent pas.",
         f"- `fiches/complet/` : un fichier au format campagne par plateforme, qui couvre toutes les exigences attendues "
         f"qui lui sont déléguées ({len(EXPECTED) - len(no_delegation)} en tout). Le fichier de {SYSTEM} reprend "
         f"exactement les tests et test cases du STDHL. {len(FAILED_IDS)} exigences ({round(FAILED * 100)} %) sont "
         "FAILED, chacune seule dans son test case ; les autres sont PASSED. Le résultat ne change pas la couverture.",
         "",
         "## Chargement", "",
         "Dans une base vide (configuration : process_plan et process_stdhl cochés, plateforme de niveau système "
         f"{SYSTEM}) :", "",
         "- les 10 fichiers d'exigences `exigences/2026-01-01-fctk.xlsx`, chacun dans son tome fctk ;",
         "- les 10 plans `plans/2026-01-01_PA_FCTk_A-plan.docx` ;",
         f"- le STDHL `stdhl/unique/{STDHL_OUT.name}`, sans tome ;",
         "- chaque fichier de `fiches/complet/`, sans tome, avec la plateforme de son nom.", "",
         "Résultat attendu : 100 % aux tuiles Plan d'acceptation, STDHL et Fiches d'Essai.", "",
         "| Fichier | Plateforme | Tests | Test cases | Exigences | Steps | Steps FAILED |",
         "| --- | --- | --- | --- | --- | --- | --- |"]
lines += [f"| {n} | {pl} | {t} | {c} | {e} | {st} | {fa} |" for n, pl, t, c, e, st, fa in summary]
if no_delegation:
    lines += ["", f"{len(no_delegation)} exigences attendues sans délégation : aucune fiche ne peut les couvrir."]
write_html(FICHES_OUT / "LISEZMOI.html", lines)
print(f"STDHL : {len(STDHL_TESTS)} tests, {len(TOP)} exigences de {SYSTEM}")
print("\n".join(f"{n} : {e} exigences, {fa} steps FAILED sur {st}" for n, _, _, _, e, st, fa in summary))
print("sans délégation :", len(no_delegation))
