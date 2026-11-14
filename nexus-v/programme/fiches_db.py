"""Fiches d'essai of NEXUS-V, in the SQL database (tables files, fe_tests and fiches_essai).

The XML fiches come from the default platform; the Excel files of the other platforms (see traceability_db.py) go
into the same table fiches_essai. Each file keeps its platform (files.platform_id).

Like a STDHL, a fiches file (XML or Excel) belongs to a tome, or to none: it then applies to every tome. Only the
requirements of its tome (of the whole base without tome) are stored; another number is shown at import, where it
can only be ignored. Removing the file removes its content (cascade).

The XML is read through three configurable tags (test_tag, test_case_tag, requirement_tag): a test contains test
cases, which contain the requirements tested, separated by commas in one tag. Each value is the text between the
opening and closing tags (without the text of the tags it contains); attributes are not read. The table
fiches_essai has one row per requirement of each test case: the test (referential fe_tests), the test case and the
requirement. A file without any number found that way is refused.

A second XML format, « campagne » (export of the campaign tool), is recognised by its TestSuite elements: a TestSuite
is a test (attribute name, else its property Title), its TestCases the test cases (property Title);
the requirements of a test case are the property VerifiedRequirements of its TestSteps, of itself and of its
TestSuite; each one takes the result of the test case, from the RunResult of its steps (FAILED if one failed, PASSED
if all those run passed, NOT_RUN otherwise). The campaign name (release) and the configuration of the
tool are not read. The format recognised is kept in files.format (« balises » or « campagne »).

For the page (to_dict):

    {"fiches": {"12": {"file_id": 12, "name": "...xml", "releases": ["A"], "tome": "fct1" (None: every tome),
                       "ids": [ID], "rows": [{"test_id", "test", "test_case", "requirements": [ID]}, ...]}}}
"""
import re
import xml.etree.ElementTree as ET

from configuration import DEFAULT_CONFIG, plural
from file_database import FileDatabase, WordDocument


def local_name(el):
    """Name of an element without its namespace, in lower case."""
    # ElementTree writes a namespaced tag as {uri}name: only the name counts, whatever its case.
    return el.tag.rsplit("}", 1)[-1].lower()


# Formats of fiches XML recognised at import, kept in files.format.
FORMATS = {"tags": "balises", "campaign": "campagne"}


def properties(el):
    """{name: text} of the properties <Properties><Property name="...">text</Property></Properties> of an element
    (its own ones, not those of the elements it contains)."""
    found = {}
    for block in el:
        if local_name(block) == "properties":
            for prop in block:
                if local_name(prop) == "property" and prop.get("name"):
                    found[prop.get("name")] = (prop.text or "").strip()
    return found


def own_text(el):
    """Text between the opening and closing tags of an element, without the text of the elements it contains."""
    # The text of an element is its leading text plus the text after each child (tail); the children's own text is left
    # out. Runs of spaces and line breaks become one space.
    return " ".join(" ".join([el.text or "", *(child.tail or "" for child in el)]).split())


class FichesDatabase:
    """Content of the Fiches d'Essai files (tests, test cases, requirements), in the SQL database."""

    def __init__(self, sql):
        """sql: the shared NexusDatabase."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def read_xml(content, config=None):
        """Read a Fiches d'Essai file (XML bytes). Returns {"rows": [{test, test_case, requirements: [ID]}], "ids":
        requirement numbers tested}, through the three configured tags (a test contains test cases, which contain
        the requirements)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        regex = re.compile(get("requirement_regex"))
        try:
            root = ET.fromstring(content)
        except ET.ParseError as e:
            raise ValueError(f"XML mal formé : {e}") from None

        def children(parent, tag):
            """Elements under `parent` (itself included) with this tag, whatever their namespace and case."""
            tag = tag.lower()
            return [el for el in parent.iter() if local_name(el) == tag]

        if children(root, "testsuite"):
            return FichesDatabase.read_campaign(root, regex, children)
        rows = []
        # One row per test case of each test, with the requirement numbers of its requirement tags (separated by commas
        # in one tag, or in several tags). WordDocument.normalize turns the special dashes and spaces of a copy-paste
        # into plain ones, so that the numbers match.
        for test in children(root, get("test_tag")):
            for case in children(test, get("test_case_tag")):
                text = [own_text(el) for el in children(case, get("requirement_tag"))]
                ids = list(dict.fromkeys(regex.findall(WordDocument.normalize("\n".join(text)))))
                if ids:
                    rows.append({"test": own_text(test), "test_case": own_text(case), "requirements": ids})
        # ids: every requirement number of the file, once, in reading order.
        return {"rows": rows, "ids": list(dict.fromkeys(i for row in rows for i in row["requirements"])),
                "format": FORMATS["tags"]}

    @staticmethod
    def read_campaign(root, regex, children):
        """A fiches file of the format « campagne »: one row per TestCase with requirements, {test, test_case,
        requirements, result PASSED | FAILED | NOT_RUN of the test case} (see the module docstring)."""
        numbers = lambda text: regex.findall(WordDocument.normalize(text or ""))
        rows = []
        for suite in children(root, "testsuite"):
            suite_props = properties(suite)
            test = (suite.get("name") or suite_props.get("Title") or "").strip()
            for case in children(suite, "testcase"):
                case_props = properties(case)
                title = case_props.get("Title") or ""   # without title: ignored at import
                found, runs = [], []   # requirements of the test case; results of its steps
                for step in children(case, "teststep"):
                    step_props = properties(step)
                    found += numbers(step_props.get("VerifiedRequirements"))
                    if (step_props.get("RunResult") or "").upper() in ("PASSED", "FAILED"):
                        runs.append(step_props["RunResult"].upper())
                found += numbers(case_props.get("VerifiedRequirements")) + numbers(suite_props.get("VerifiedRequirements"))
                # The result of the test case: FAILED if a step failed, PASSED if every step run passed, else NOT_RUN.
                result = "FAILED" if "FAILED" in runs else "PASSED" if runs else "NOT_RUN"
                if found:
                    rows.append({"test": test, "test_case": title, "requirements": list(dict.fromkeys(found)),
                                 "result": result})
        return {"rows": rows, "ids": list(dict.fromkeys(i for row in rows for i in row["requirements"])),
                "format": FORMATS["campaign"]}

    @classmethod
    def _readable(cls, content, config=None):
        """Content of a file that can be loaded: ValueError when no requirement number was found."""
        data = cls.read_xml(content, config)
        if not data["ids"]:
            raise ValueError("aucun numéro d'exigence trouvé dans les balises attendues (voir Configuration > "
                             "En-têtes des tableaux) : fichier non chargé")
        return data

    def analyse(self, content, config=None, releases=(), tome=None):
        """What loading a fiches file would do: number of requirement numbers and the numbers to ignore (absent from
        the requirements of the `tome`, of the whole base without tome). Refused (ValueError) when malformed or without
        any number."""
        data = self._readable(content, config)
        found = self.screen(data, tome or None, config)
        return {"ids": len(data["ids"]), "format": data["format"], "to_confirm": [],
                **{k: v for k, v in found.items() if k != "rows"}}

    def screen(self, data, tome=None, config=None, platform=""):
        """The rules of the import on the rows read in a fiches file XML: {"rows": the rows to store [(test, test case,
        [numbers], result)], "unknown": numbers absent from the base (of the `tome`), ignored; "untitled": test cases
        without title, ignored; "ignored_tests": tests that no STDHL asks for, their rows ignored; "not_specified":
        requirements in no STDHL of their tome (it refuses the file);
        "delegated": requirements not delegated to the `platform` of the file (required), loaded
        without their result. The rules of the STDHL (tests and test cases asked, requirement in no STDHL) apply to a
        fiche of the platform of system level (system_platform), on the requirements delegated to it;
        "ignored_cases": test cases that the STDHL « unique » of their test does not ask for, ignored}."""
        # The STDHL specify the tests of their platform (configuration system_platform): its fiches are checked by them.
        system_platform = (config.get("system_platform") if config is not None else "") or ""
        with_stdhl = bool(platform) and platform.casefold() == system_platform.casefold()
        tome_of = {r["req_id"]: r["tome"] for r in self.sql.query(
            "SELECT q.req_id, t.name AS tome FROM requirements q JOIN list_tomes t ON t.id = q.tome_id"
            + (" WHERE t.name = ?" if tome else ""), (tome,) if tome else ())}
        asked = {r["name"] for r in self.sql.query(
            "SELECT DISTINCT t.name FROM stdhl_exi x JOIN fe_tests t ON t.id = x.test_id WHERE x.deleted = 0")}
        # Test cases asked by the STDHL « unique » (a STDHL « par tome » has none: its tests are checked alone).
        cases = {(r["test"], r["test_case"]) for r in self.sql.query(
            "SELECT t.name AS test, c.test_case FROM stdhl_test_case s JOIN fe_test_case c ON c.id = s.test_case_id "
            "JOIN fe_tests t ON t.id = c.test_id")}
        with_cases = {t for t, _ in cases}
        specified = {(r["req_id"], r["tome"]) for r in self.sql.query(
            "SELECT q.req_id, t.name AS tome FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
            "JOIN files f ON f.id = x.file_id LEFT JOIN list_tomes t ON t.id = f.tome_id WHERE x.deleted = 0")}
        # Delegated elsewhere: none of the plan rows of the requirement is delegated to the platform of the file.
        own = {r["req_id"] for r in self.sql.query(
            "SELECT DISTINCT q.req_id FROM plan_acceptation p JOIN requirements q ON q.id = p.requirement_id "
            "JOIN list_plateformes l ON l.id = p.delegation_id WHERE p.deleted = 0 AND LOWER(l.name) = LOWER(?)",
            (platform,))}
        planned = {r["req_id"] for r in self.sql.query(
            "SELECT DISTINCT q.req_id FROM plan_acceptation p JOIN requirements q ON q.id = p.requirement_id "
            "WHERE p.deleted = 0")}
        rows, unknown, untitled, ignored_tests, ignored_cases = [], [], [], [], []
        for r in data["rows"]:
            if not (r["test_case"] or "").strip():
                untitled.append(r["test"])
                continue
            if with_stdhl and r["test"] not in asked:
                if r["test"] not in ignored_tests:
                    ignored_tests.append(r["test"])
                continue
            if with_stdhl and r["test"] in with_cases and (r["test"], r["test_case"].strip()) not in cases:
                ignored_cases.append(f"{r['test']} / {r['test_case']}")
                continue
            keep = [i for i in r["requirements"] if i in tome_of]
            unknown += [i for i in r["requirements"] if i not in tome_of and i not in unknown]
            if keep:
                rows.append((r["test"], r["test_case"], keep, r.get("result")))
        cited = list(dict.fromkeys(i for _, _, ids, _ in rows for i in ids))
        delegated = [i for i in cited if i in planned and i not in own]
        # In no STDHL of its tome: only a requirement delegated to the default platform (the others have no STDHL).
        not_specified = [i for i in cited if with_stdhl and i not in delegated
                         and (i, tome_of[i]) not in specified and (i, None) not in specified]
        # A requirement delegated elsewhere is loaded without its result: a row of its own, NOT_RUN.
        split = []
        for test, case, ids, result in rows:
            split.append((test, case, [i for i in ids if i not in delegated], result))
            if any(i in delegated for i in ids):
                split.append((test, case, [i for i in ids if i in delegated], "NOT_RUN"))
        return {"rows": [x for x in split if x[2]], "unknown": unknown, "untitled": untitled,
                "ignored_tests": ignored_tests, "ignored_cases": ignored_cases,
                "not_specified": not_specified, "delegated": delegated}

    @staticmethod
    def refusal(found):
        """The message refusing a fiches file whose requirements are in no STDHL, or None."""
        parts = []
        if found["not_specified"]:
            parts.append(f"{plural(len(found['not_specified']), 'exigence')} dans aucun STDHL de son tome : "
                         + ", ".join(found["not_specified"]) + " (les spécifier d'abord dans le STDHL)")
        return " ; ".join(parts) + " : fichier non chargé" if parts else None

    @staticmethod
    def _write(db, file_id, data, rows):
        """Replace the stored content of a file (inside a transaction): its `rows` kept by screen, only the
        requirements of its tome (of the whole base for a file without tome)."""
        # The tome of the file (None: every tome) limits the requirements stored.
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        FichesDatabase.store_tests(db, file_id, rows, tome_id)
        db.execute("UPDATE files SET format = ? WHERE id = ?", (data.get("format", ""), file_id))

    @staticmethod
    def store_tests(db, file_id, tests, tome_id=None, by_case=True):
        """Replace the coverage of a fiches file (inside a transaction): one row of fiches_essai per requirement of a
        test case [(test name, test case, [ID], optionally the result of the test case)], the result PASSED or FAILED
        when the file gives one (RunResult), NOT_RUN otherwise; the test taken in the referential fe_tests and the test
        case in fe_test_case (added when new); only the requirements of the base (of the tome `tome_id` if given) are
        stored. A row always points to its test; with `by_case` (fiches XML) also to its test case."""
        from nexus_db import test_id as test_of
        db.execute("DELETE FROM fiches_essai WHERE file_id = ?", (file_id,))   # tests no longer cited: removed
        # Requirement number -> id, of the tome if the file has one, of the whole base otherwise.
        known = {r[0]: r[1] for r in db.execute(
            "SELECT req_id, id FROM requirements" + (" WHERE tome_id = ?" if tome_id is not None else ""),
            (tome_id,) if tome_id is not None else ())}
        for name, test_case, ids, *rest in tests:
            result = rest[0] if rest and rest[0] in ("PASSED", "FAILED") else "NOT_RUN"
            # The test in the referential fe_tests (shared with the STDHL); the same requirement in the same test case
            # is stored once (INSERT OR IGNORE on the UNIQUE key).
            test_id = test_of(db, name)
            if by_case:
                db.execute("INSERT OR IGNORE INTO fe_test_case (test_id, test_case) VALUES (?, ?)", (test_id, test_case))
                case_id = db.execute("SELECT id FROM fe_test_case WHERE test_id = ? AND test_case = ?",
                                     (test_id, test_case)).fetchone()[0]
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_id, test_case_id, result) "
                               "VALUES (?, ?, ?, ?, ?)", [(file_id, known[i], test_id, case_id, result) for i in ids if i in known])
            else:
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_id, result) "
                               "VALUES (?, ?, ?, ?)", [(file_id, known[i], test_id, result) for i in ids if i in known])

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None, platform=""):
        """Load a fiches file (XML) of the `platform` chosen: it belongs to the `tome` given, or to none ("": every tome).
        The numbers absent from its tome must be ignored by the user (`decisions`). The file is stored, then its
        content (the file is removed again on failure); a file of the same name is replaced, whatever its tome."""
        tome = files.check_tome("de fiches d'essai", requirements, tome)
        data = self._readable(content, config)
        # The rules of the import (see screen): refused for a requirement in no STDHL.
        # The platform the file comes from: required, a platform of the configuration.
        if not platform:
            raise ValueError("indiquer la plateforme d'où vient ce fichier de fiches d'essai")
        if config is not None and not config.is_platform(platform):
            raise ValueError(f"plateforme {platform} absente de la configuration")
        found = self.screen(data, tome, config, platform)
        if (message := self.refusal(found)) is not None:
            raise ValueError(message)
        # Copy stored first, content next, in one transaction: on failure the copy is removed again.
        record = files.add("xml", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data, found["rows"])
                # The platform the fiche comes from.
                db.execute("UPDATE files SET platform_id = (SELECT id FROM list_plateformes WHERE LOWER(name) = LOWER(?)) "
                           "WHERE id = ?", (platform, record["id"]))
        except Exception:
            files.remove(record["id"])
            raise
        stored = len({i for _, _, ids, _ in found["rows"] for i in ids})
        return {"file_id": record["id"], "ids": stored, "ignored": len(found["unknown"]), "format": data["format"],
                "unknown": found["unknown"], "untitled": found["untitled"], "ignored_tests": found["ignored_tests"],
                "ignored_cases": found["ignored_cases"],
                "delegated": found["delegated"], "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Read and store the content of the fiches files without any row yet, and of all of them with `again`
        (reading settings or requirements changed). Returns the names read; a file that cannot be read keeps its stored
        content."""
        known = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM fiches_essai")}   # rows stored
        done = []
        for record in files.list(slot="xml"):
            if record["id"] not in known or again:
                try:
                    data = self._readable(files.content(record["id"]) or b"", config)
                except ValueError:
                    continue
                # Read again: the same rules, a requirement now in no STDHL left out instead of refusing.
                found = self.screen(data, record["tome"], config, self.sql.scalar(
                    "SELECT l.name FROM files f JOIN list_plateformes l ON l.id = f.platform_id WHERE f.id = ?",
                    (record["id"],)) or "")
                blocked = set(found["not_specified"])
                rows = [(t, c, [i for i in ids if i not in blocked], r) for t, c, ids, r in found["rows"]]
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data, [x for x in rows if x[2]])
                done.append(record["name"])
        return done

    def documents(self, source="xml"):
        """The fiches files of a type (xml: Fiches d'essai of the default platform; traceability: Excel files of the
        other platforms) with their platform, tome (None: every tome), releases, requirement numbers and rows (test,
        test case, requirements)."""
        heads = self.sql.query("SELECT f.id AS file_id, f.name, p.name AS platform, t.name AS tome "
                               "FROM files f LEFT JOIN list_plateformes p ON p.id = f.platform_id "
                               "LEFT JOIN list_tomes t ON t.id = f.tome_id "
                               "WHERE f.source = ? ORDER BY f.id", (source,))
        releases = self.sql.file_releases()
        wanted = {h["file_id"] for h in heads}
        rows = {}
        # A row points to its test, and for the fiches XML to its test case too (none for the other platforms).
        for r in self.sql.query("SELECT x.file_id, x.test_id, t.name AS test, COALESCE(c.test_case, '') AS test_case, "
                                "q.req_id, x.result FROM fiches_essai x LEFT JOIN fe_test_case c ON c.id = x.test_case_id "
                                "JOIN fe_tests t ON t.id = x.test_id "
                                "JOIN requirements q ON q.id = x.requirement_id ORDER BY x.file_id, x.id"):
            if r["file_id"] not in wanted:
                continue
            cases = rows.setdefault(r["file_id"], [])
            if not cases or (cases[-1]["test_id"], cases[-1]["test_case"], cases[-1]["result"]) != \
                    (r["test_id"], r["test_case"], r["result"]):
                cases.append({"test_id": r["test_id"], "test": r["test"], "test_case": r["test_case"],
                              "result": r["result"], "requirements": []})
            cases[-1]["requirements"].append(r["req_id"])
        return [{"file_id": h["file_id"], "name": h["name"], "source": source, "platform": h["platform"] or "",
                 "tome": h["tome"], "releases": releases.get(h["file_id"], []),
                 "ids": list(dict.fromkeys(i for c in rows.get(h["file_id"], []) for i in c["requirements"])),
                 "rows": rows.get(h["file_id"], [])} for h in heads]

    def to_dict(self):
        """The fiches files by file number, for the page."""
        return {"fiches": {str(d["file_id"]): d for d in self.documents()}}
