"""STDHL « unique » (format 2: one document, no STDHL table) to try the check before import and the import itself.

They take the content of the test STDHL stdhl/perTome/2026-01-01_TEST_STDHL_FCTk_rev1.docx (same tests SYS_, same
requirements), which must be loaded in the base with the requirements and the plans of the tests, and write it as in
stdhl/2026-11-11-STDHL-template2.docx: a title « <test> : <titre> » per test, État, Objectif, Applicabilité, the list
« Exigences associées » (the end of the list ends it), then the test cases and their phases, each phase citing one
requirement. Written into stdhl/unique/2026-01-01_TEST_STDHL_UNIQUE_rev1.docx (the 10 test STDHL in one document, the
reference STDHL « unique » of the tests) and verification-import/stdhl/unique/:
- one document per case, with LISEZMOI.html giving the expected result of each one;
- complement/2026-03-15_STDHL_UNIQUE-complement_rev1.docx: for the 10 tomes, 5 % of the requirements asked for another
  test, plus the new requirements of the requirement complement files (SSS_FCTk_REQ_500...);
- remplacement/2026-03-15_STDHL_UNIQUE_rev2.docx: the 10 tomes, 5 % of the requirements removed, 5 % asked for another
  test, plus the new requirements of the requirement replacement files (SSS_FCTk_REQ_600...).
Word documents in A4 portrait, the titles of the tests visible (see TITLE3_FIX).
Usage (from nexus-v/programme): python ../../base-test/generateurs/make_verification_stdhl_unique.py ../../base-test
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
from stdhl_db import StdhlDatabase

BASE = Path(sys.argv[1])
OUT = BASE / "verification-import" / "stdhl" / "unique"
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"
cfg = Configuration(Path("../fichiers-references/base/config.json"))   # read only
SHARE = 0.05
# The style Titre3 of the template puts its text off the page (see make_stdhl_format2.py): each title cancels it.
TITLE3_FIX = ('<w:tabs><w:tab w:val="clear" w:pos="12137"/><w:tab w:val="num" w:pos="1080"/></w:tabs>'
              '<w:ind w:left="0" w:firstLine="0"/>')


def p(text, style=""):
    fix = TITLE3_FIX if style == "Titre3" else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>{fix}</w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="3000"/>' for _ in rows[0])
    body = "".join("<w:tr>" + "".join(f"<w:tc>{p(c)}</w:tc>" for c in r) + "</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>")


def titles():
    found = {}
    for k in range(1, 11):
        ws = openpyxl.load_workbook(BASE / "exigences" / f"2026-01-01-fct{k}.xlsx").active
        found.update({r[0]: r[1] for r in ws.iter_rows(min_row=2, values_only=True) if r[0]})
    return found


TITLES = titles()


ACTIONS = ["Lancer le scénario nominal", "Provoquer le cas dégradé", "Rejouer la séquence opérateur",
           "Vérifier l'affichage de l'IHM", "Contrôler les journaux", "Couper puis rétablir la liaison"]


def test_block(name, listed, cases, release="RELEASE A", means=False):
    """The paragraphs of a test: title, État, Objectif, Applicabilité (None: absent), the list of its associated
    requirements, its conditions, then its test cases [[[requirements of a phase]]]."""
    out = [p(f"{name} : vérification de {name.split('_')[1].lower() if '_' in name else name}", "Titre3"),
           p("État : Revue"), p("Objectif : vérifier les exigences du test")]
    if release is not None:
        out.append(p(f"Applicabilité : {release}"))
    out.append(p("Exigences associées :"))
    out += [p(f"{i} : {TITLES.get(i, 'nouvelle exigence')}", "Paragraphedeliste") for i in listed]
    out += [p("Conditions et configuration initiales :"), p("Banc de test en configuration nominale", "Paragraphedeliste")]
    if means:   # a table that is not a STDHL table of the format par tome
        out += [p("Moyens d'essai :"), table([["Moyen", "Configuration"], ["Banc de test V&V", "Nominale"]])]
    out.append(p("DESCRIPTION DE LA FICHE DE TEST"))
    for c, phases in enumerate(cases, 1):
        out.append(p(f"TEST CASE#{c} : cas {c}"))
        out += [p(f"Phase {j} : {ACTIONS[(c + j) % len(ACTIONS)]}" + (" : " + ", ".join(ph) if ph else ""),
                  "Paragraphedeliste") for j, ph in enumerate(phases, 1)]
    return out


def write_unique(path, title, blocks, history=False):
    """A STDHL « unique »: a title, the table of contents (left out by the reading), then the tests."""
    names = [b[0] for b in blocks]
    body = [p(title, "Titre"), p("Contents", "En-ttedetabledesmatires")]
    body += [p(f"2.{k} {n} : test {k}", "TM3") for k, n in enumerate(names, 1)]
    if history:
        body += [p("Historique des versions", "Titre1"),
                 table([["Version", "Date", "Auteur"], ["1.0", "2026-03-01", "V&V"]])]
    body += [p("Description des tests", "Titre1")]
    for name, listed, cases, kw in blocks:
        body += test_block(name, listed, cases, **kw)
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


def stdhl_of(k):
    """Tests of the test STDHL of fctk: [(test, [requirements])]."""
    name = f"2026-01-01_TEST_STDHL_FCT{k}_rev1.docx"
    data = StdhlDatabase.read_docx((BASE / "stdhl" / "perTome" / name).read_bytes(), name, cfg)
    tests = {}
    for r in data["rows"]:
        for i in r["requirements"]:
            if i not in tests.setdefault(r["fiche"], []):
                tests[r["fiche"]].append(i)
    return list(tests.items())


SIZES = ((0, 0.50), (1, 0.25), (2, 0.10), (3, 0.10), (None, 0.05))   # requirements per phase (None: 4 to 6), share
STATS = {}   # document -> number of phases by number of requirements, for the LISEZMOI


def phases_of(tests, rng, n=None):
    """{test: [test cases [phases [requirements]]]} for consistent tests [(test, [requirements])]: the phases of the
    document in the exact shares of SIZES, every requirement of a test cited in one of its phases at least, the places
    left citing again a requirement of the same test (another one than in the same phase), 2 to 6 phases per test case."""
    total = sum(len(r) for _, r in tests)
    n = n or max(total, len(tests) * 2)
    counts = [round(n * share) for _, share in SIZES]
    counts[-1] = n - sum(counts[:-1])
    sizes = [size for (size, _), c in zip(SIZES, counts) if size is not None for _ in range(c)]
    sizes += [rng.randint(4, 6) for _ in range(counts[-1])]
    while sum(sizes) < total:   # enough places for every requirement
        k = next((i for i, x in enumerate(sizes) if 4 <= x < 6), None)
        if k is None:
            break
        sizes[k] += 1
    largest = max(len(r) for _, r in tests)
    full = sorted((min(x, largest) for x in sizes if x), reverse=True)
    empty = sum(1 for x in sizes if not x)
    own = {t: [] for t, _ in tests}
    pool = list(full)
    # Each test takes phases, the largest it can fill first, until it has a place for each of its requirements. When
    # no phase is left small enough for a test, the document gets more phases, still in the shares of SIZES.
    for t, reqs in sorted(tests, key=lambda x: -len(x[1])):
        while sum(own[t]) < len(reqs):
            k = next((i for i, x in enumerate(pool) if x <= len(reqs)), None)
            if k is None:
                if n < 4 * total:
                    return phases_of(tests, rng, n + max(4, n // 20))
                pool.append(1)
                k = len(pool) - 1
            own[t].append(pool.pop(k))
    for x in pool:   # phases left: to the tests that can fill them
        own[rng.choice([t for t, reqs in tests if len(reqs) >= x])].append(x)
    weights = [len(r) for _, r in tests]
    for _ in range(empty):   # empty phases in proportion to the requirements of each test
        own[rng.choices([t for t, _ in tests], weights)[0]].append(0)
    result = {}
    for t, reqs in tests:
        phases, queue = [], list(reqs)
        for x in sorted(own[t], reverse=True):
            ph = []
            for _ in range(x):
                ph.append(queue.pop(0) if queue else rng.choice([i for i in reqs if i not in ph]))
            phases.append(ph)
        rng.shuffle(phases)
        cases, at = [], 0
        while at < len(phases):
            step = rng.randint(2, 6)
            cases.append(phases[at:at + step])
            at += step
        result[t] = cases
    return result


def blocks_of(tests, seed, **kw):
    """Blocks of consistent tests (the list and the test cases hold the same requirements), phases in the shares of
    SIZES."""
    found = phases_of(tests, random.Random(seed))
    return [(t, reqs, found[t], dict(kw)) for t, reqs in tests]


def distribution(blocks):
    """Phases of a document by number of requirements: {"0": n, "1": n, "2": n, "3": n, "4-6": n}."""
    out = {"0": 0, "1": 0, "2": 0, "3": 0, "4-6": 0}
    for _, _, cases, _ in blocks:
        for c in cases:
            for ph in c:
                out[str(len(ph)) if len(ph) < 4 else "4-6"] += 1
    return out


ALL = [x for k in range(1, 11) for x in stdhl_of(k)]
# The reference STDHL « unique » of the tests: the 10 test STDHL in one document, consistent, to load instead of them.
reference = blocks_of(ALL, 5000)
write_unique(BASE / "stdhl" / "unique" / "2026-01-01_TEST_STDHL_UNIQUE_rev1.docx", "STDHL de tous les tomes",
             reference, history=True)
STATS["stdhl/unique/2026-01-01_TEST_STDHL_UNIQUE_rev1.docx (référence)"] = distribution(reference)
# The same content with INCONSISTENCY of its requirements in one part only of their test (half in the list only, half
# in the test cases only): a case of its own, refused.
INCONSISTENCY = 0.05
reference = [(t, list(listed), [[list(ph) for ph in c] for c in cases_], dict(kw)) for t, listed, cases_, kw in reference]
pick = random.Random(5001)
pool = [(k, i) for k, (_, listed, _, _) in enumerate(reference) for i in listed]
picked = pick.sample(pool, round(INCONSISTENCY * len({i for _, i in pool})))
inconsistent = []
for j, (k, i) in enumerate(picked):
    t, listed, cases_, kw = reference[k]
    if j % 2 == 0:   # in the list, cited in no test case: in its phases, another requirement of the test instead
        others = [x for x in listed if x != i]
        cases_ = [[[x if x != i else pick.choice([y for y in others if y not in ph] or others) for x in ph] if i in ph
                   else ph for ph in c] for c in cases_]
        cases_ = [[list(dict.fromkeys(ph)) for ph in c] for c in cases_]
        inconsistent.append((t, i, "dans les exigences associées, citée dans aucun test case"))
    else:            # cited in the test cases, absent from the list
        listed = [x for x in listed if x != i]
        inconsistent.append((t, i, "citée dans les test cases, absente des exigences associées"))
    reference[k] = (t, listed, cases_, kw)
cases = []


def case(name, what, expected, title, blocks, **kw):
    write_unique(OUT / name, title, blocks, **kw)
    STATS[name] = distribution(blocks)
    cases.append((name, what, expected))


case("2026-03-01_STDHL_UNIQUE-incoherences_rev1.docx",
     f"les 10 STDHL de test en un document, {len(inconsistent)} exigences ({round(INCONSISTENCY * 100)} %) dans une seule "
     "des deux parties de leur test (liste ci-dessous)",
     "Refusé dans les deux cas : chaque test touché a une exigence dans une seule des deux parties (Bloquant)",
     "STDHL de tous les tomes", reference, history=True)
case("2026-03-01_STDHL_UNIQUE-identique_rev1.docx", "les 10 STDHL de test en un seul document, mêmes tests et exigences",
     "Chargeable dans les deux cas : format unique, aucune exigence nouvelle ni modifiée ; les exigences de TOP déjà "
     "sans STDHL dans la base sont rappelées (À vérifier)", "STDHL de tous les tomes", blocks_of(ALL, 5100))
t1 = stdhl_of(1)
b = blocks_of(t1, 5101)
b[0] = (b[0][0], b[0][1], b[0][2], {"release": "RELEASE B"})
case("2026-03-01_STDHL_UNIQUE-release-b_rev1.docx", f"STDHL fct1, le test {t1[0][0]} en RELEASE B",
     f"Refusé dans les deux cas : applicabilité B du test {t1[0][0]}, la base est en release A (Bloquant)",
     "STDHL FCT1", b)
b = blocks_of(stdhl_of(2), 5102)
b[0] = (b[0][0], b[0][1], b[0][2], {"release": None})
case("2026-03-01_STDHL_UNIQUE-sans-release_rev1.docx", f"STDHL fct2, le test {b[0][0]} sans applicabilité",
     f"Chargeable, à vérifier, dans les deux cas : test {b[0][0]} : applicabilité absente (À vérifier)", "STDHL FCT2", b)
t3 = stdhl_of(3)
b = blocks_of(t3, 5103)
extra = next(i for _, reqs in stdhl_of(4) for i in reqs)
first, second = b[0], b[1] if len(b) > 1 else b[0]
# extra: cited in a phase of the first test, not in its list; alone: in the list of the second test, in no phase.
next(ph for c in first[2] for ph in c if ph).append(extra)
alone = second[1][0]
b[1 if len(b) > 1 else 0] = (second[0], second[1], [[[i for i in ph if i != alone] for ph in c] for c in second[2]], {})
case("2026-03-01_STDHL_UNIQUE-ecarts_rev1.docx",
     f"STDHL fct3 : {extra} citée dans un test case de {first[0]} sans être dans sa liste ; {alone} dans la liste de "
     f"{second[0]} sans être citée dans ses test cases",
     f"Refusé dans les deux cas : 1 exigence citée dans les test cases mais absente des exigences associées "
     f"({first[0]}), 1 exigence associée citée dans aucun test case ({second[0]}) (Bloquant)",
     "STDHL FCT3", b)
t5 = stdhl_of(5)
case("2026-03-01_STDHL_UNIQUE-nom-double_rev1.docx", f"STDHL fct5, le test {t5[0][0]} écrit deux fois",
     f"Chargeable, à vérifier, dans les deux cas : test {t5[0][0]} présent 2 fois (À vérifier)", "STDHL FCT5",
     blocks_of(t5, 5105) + blocks_of(t5[:1], 5106))
t6 = stdhl_of(6)
unknown = [f"SSS_FCT6_REQ_{n}" for n in range(990, 993)]
case("2026-03-01_STDHL_UNIQUE-inconnues_rev1.docx", "STDHL fct6, 3 exigences absentes de la base (SSS_FCT6_REQ_990 à 992)",
     "Chargeable avec des choix à faire dans les deux cas : 3 exigences inconnues de la base, non enregistrées (À régler)",
     "STDHL FCT6", blocks_of(t6[:-1] + [(t6[-1][0], t6[-1][1] + unknown)], 5107))
case("2026-03-01_STDHL_UNIQUE-tableaux_rev1.docx",
     "STDHL fct7 avec un historique des versions et un tableau des moyens d'essai par test (pas des tableaux STDHL)",
     "Chargeable dans les deux cas : lu au format unique, les tableaux sont ignorés", "STDHL FCT7",
     blocks_of(stdhl_of(7), 5108, means=True), history=True)
write_unique(OUT / "2026-03-01_STDHL_UNIQUE-sans-test_rev1.docx", "Document sans test", [])
cases.append(("2026-03-01_STDHL_UNIQUE-sans-test_rev1.docx", "document sans tableau STDHL ni titre de test",
              "Refusé dans les deux cas : aucun tableau STDHL ni test du format unique (Bloquant)"))




def moved(tests, chosen, rng):
    """The tests with the requirements `chosen` asked for another test of the same tome; (tests, {requirement: test})."""
    names = [t for t, _ in tests]
    changes = {}
    for i in chosen:
        own = next(t for t, rs in tests if i in rs)
        tome = re.search(r"FCT\d+", i).group(0)
        others = [t for t in names if t != own and tome in t] or [t for t in names if t != own]
        changes[i] = rng.choice(others)
    out = {t: [i for i in rs if i not in changes] for t, rs in tests}
    for i, t in changes.items():
        out[t].append(i)
    return [(t, rs) for t, rs in out.items() if rs], changes


rng = random.Random(4500)
reqs = list(dict.fromkeys(i for _, rs in ALL for i in rs))
n = max(1, math.ceil(SHARE * len(reqs)))
report = []
# Complement: the requirements whose test changes, and the new ones (one test per tome); nothing else.
chosen = rng.sample(reqs, n)
full, changes = moved(ALL, chosen, rng)
only = [(t, [i for i in rs if i in changes]) for t, rs in full]
added = {k: [f"SSS_FCT{k}_REQ_{m}" for m in range(500, 502)] for k in range(1, 11)}
blocks = blocks_of([(t, rs) for t, rs in only if rs] + [(f"SYS_NOUVEAU_FCT{k}_001", ids) for k, ids in added.items()], 5200)
STATS["complement/2026-03-15_STDHL_UNIQUE-complement_rev1.docx"] = distribution(blocks)
write_unique(OUT / "complement" / "2026-03-15_STDHL_UNIQUE-complement_rev1.docx", "STDHL unique — complément", blocks)
report.append(("2026-03-15_STDHL_UNIQUE-complement_rev1.docx", "complément",
               {"nouvelles": [i for ids in added.values() for i in ids], "test modifié": sorted(changes)}))
# Replacement: the whole content, some requirements removed, some with another test, some new.
picked = rng.sample(reqs, 2 * n)
removed, chosen = set(picked[:n]), picked[n:]
kept = [(t, [i for i in rs if i not in removed]) for t, rs in ALL]
full, changes = moved([(t, rs) for t, rs in kept if rs], chosen, rng)
added = {k: [f"SSS_FCT{k}_REQ_{m}" for m in range(600, 602)] for k in range(1, 11)}
blocks = blocks_of(full + [(f"SYS_NOUVEAU_FCT{k}_001", ids) for k, ids in added.items()], 5201)
STATS["remplacement/2026-03-15_STDHL_UNIQUE_rev2.docx"] = distribution(blocks)
write_unique(OUT / "remplacement" / "2026-03-15_STDHL_UNIQUE_rev2.docx", "STDHL unique — remplacement", blocks)
report.append(("2026-03-15_STDHL_UNIQUE_rev2.docx", "remplacement",
               {"retirées": sorted(removed), "nouvelles": [i for ids in added.values() for i in ids],
                "test modifié": sorted(changes)}))

lines = ["# STDHL « unique » pour essayer la vérification avant import", "",
         "Prérequis : les exigences, les plans et les STDHL de test fct1 à fct10 sont chargés dans la base. Vérifier ces "
         "documents avec la loupe de la boîte STDHL ; rien n'est importé. Ils sont au format unique : aucun tableau "
         "STDHL, un titre par test, sa liste « Exigences associées » et ses test cases.", "",
         "La couverture d'un test est l'ensemble des exigences de ses test cases ; ses exigences associées doivent être "
         "exactement celles-là : une exigence dans une seule des deux parties refuse le fichier (Bloquant), comme une "
         "applicabilité d'une autre release que la base. Le rapport donne pour chaque fichier ce qui se passe s'il écrase "
         "la base et s'il la modifie ; les résultats ci-dessous valent pour les deux cas, sauf mention. Chaque rapport "
         "rappelle aussi les exigences de TOP déjà sans STDHL dans la base (À vérifier), non répétées ci-dessous.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_stdhl_unique.py ../../base-test` depuis "
         "`nexus-v/programme`.", "", "| Fichier | Contenu | Résultat attendu |", "| --- | --- | --- |"]
lines += [f"| {f} | {what} | {expected} |" for f, what, expected in cases]
lines += ["", "## Phases par nombre d'exigences", "",
          "Chaque document a ses phases dans les proportions 50 % sans exigence, 25 % avec une, 10 % avec deux, 10 % avec "
          "trois, 5 % avec quatre à six ; les places en trop citent à nouveau une exigence du même test.", "",
          "| Document | 0 | 1 | 2 | 3 | 4 à 6 |", "| --- | --- | --- | --- | --- | --- |"]
lines += [f"| {d} | " + " | ".join(str(v) for v in st.values()) + " |" for d, st in STATS.items()]
lines += ["", "## Incohérences de 2026-03-01_STDHL_UNIQUE-incoherences_rev1.docx", "",
          f"Le fichier contient {len(inconsistent)} incohérences ({round(INCONSISTENCY * 100)} % de ses exigences), "
          "chacune Bloquant à la vérification ; le STDHL unique de référence "
          "(`stdhl/unique/2026-01-01_TEST_STDHL_UNIQUE_rev1.docx`), lui, n'en a aucune et peut être importé :", ""]
lines += [f"- {t} : {i}, {what}" for t, i, what in inconsistent]
write_html(OUT / "LISEZMOI.html", lines)

kinds = ("retirées", "nouvelles", "test modifié")
lines = ["# STDHL « unique » de complément et de remplacement", "",
         "Prérequis : les exigences, les plans et les STDHL de test fct1 à fct10 sont chargés dans la base. Chaque "
         "document couvre les 10 tomes, au format unique.", "",
         "- `complement/` : 5 % des exigences demandées pour un autre test du même tome, et 2 exigences nouvelles par tome "
         "(`SSS_FCTk_REQ_500`, `501`, celles des fichiers d'exigences de complément) sous un test `SYS_NOUVEAU_FCTk_001`. "
         "Dans le rapport, regarder la partie « Si le fichier modifie la base ».",
         "- `remplacement/` : tout le contenu, moins 5 % d'exigences retirées, 5 % demandées pour un autre test, et 2 "
         "exigences nouvelles par tome (`SSS_FCTk_REQ_600`, `601`). Dans le rapport, regarder la partie « Si le fichier "
         "écrase la base ».", "",
         "Tant que les exigences nouvelles ne sont pas importées, elles sont « exigences inconnues de la base » (À "
         "régler). Les exigences retirées « perdent leur STDHL avec l'import » si le fichier écrase la base.", "",
         "Régénérer : `python ../../base-test/generateurs/make_verification_stdhl_unique.py ../../base-test` depuis "
         "`nexus-v/programme`.", "",
         "| Fichier | Mode | " + " | ".join(k.capitalize() for k in kinds) + " |",
         "| --- | --- | " + " | ".join("---" for _ in kinds) + " |"]
for f, mode, found in report:
    lines.append(f"| {f} | {mode} | " + " | ".join(str(len(found[k])) if k in found else "—" for k in kinds) + " |")
lines += ["", "## Numéros par fichier", ""]
for f, mode, found in report:
    lines += [f"**{f}** ({mode})", ""] + [f"- {c} : {', '.join(v)}" for c, v in found.items() if v] + [""]
write_html(OUT / "LISEZMOI-complement-remplacement.html", lines)
print(f"{len(list(OUT.glob('*.docx')))} STDHL unique de cas, {len(report)} de complément et de remplacement dans {OUT}")
