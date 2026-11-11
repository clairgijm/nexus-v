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

For the page (to_dict):

    {"fiches": {"12": {"file_id": 12, "name": "...xml", "releases": ["A"], "tome": "fct1" (None: every tome),
                       "ids": [ID], "rows": [{"test_id", "test", "test_case", "requirements": [ID]}, ...]}}}
"""
import re
import xml.etree.ElementTree as ET

from configuration import DEFAULT_CONFIG
from file_database import FileDatabase, WordDocument


def local_name(el):
    """Name of an element without its namespace, in lower case."""
    # ElementTree writes a namespaced tag as {uri}name: only the name counts, whatever its case.
    return el.tag.rsplit("}", 1)[-1].lower()


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
        return {"rows": rows, "ids": list(dict.fromkeys(i for row in rows for i in row["requirements"]))}

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
        ids = self._readable(content, config)["ids"]
        return {"ids": len(ids), "to_confirm": FileDatabase.to_ignore(self.sql.db, ids, tome or None)}

    @staticmethod
    def _write(db, file_id, data):
        """Replace the stored content of a file (inside a transaction): only the requirements of its tome (of the
        whole base for a file without tome) are stored."""
        # The tome of the file (None: every tome) limits the requirements stored.
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        FichesDatabase.store_tests(db, file_id, [(r["test"], r["test_case"], r["requirements"]) for r in data["rows"]],
                                   tome_id)

    @staticmethod
    def store_tests(db, file_id, tests, tome_id=None, by_case=True):
        """Replace the coverage of a fiches file (inside a transaction): one row of fiches_essai per requirement of a
        test case [(test name, test case, [ID], optionally {ID: result})], the result PASSED or FAILED when the file
        gives one (RunResult), NOT_RUN otherwise; the test taken in the referential fe_tests and the test case in
        fe_test_case (added when new); only the requirements of the base (of the tome `tome_id` if given) are stored.
        `by_case`: the rows point to the test case (fiches XML); otherwise to the test (fiches of the other platforms,
        without test case)."""
        from nexus_db import test_id as test_of
        db.execute("DELETE FROM fiches_essai WHERE file_id = ?", (file_id,))   # tests no longer cited: removed
        # Requirement number -> id, of the tome if the file has one, of the whole base otherwise.
        known = {r[0]: r[1] for r in db.execute(
            "SELECT req_id, id FROM requirements" + (" WHERE tome_id = ?" if tome_id is not None else ""),
            (tome_id,) if tome_id is not None else ())}
        for name, test_case, ids, *rest in tests:
            results = rest[0] if rest else {}
            result = lambda i: results.get(i) if results.get(i) in ("PASSED", "FAILED") else "NOT_RUN"
            # The test in the referential fe_tests (shared with the STDHL); the same requirement in the same test case
            # is stored once (INSERT OR IGNORE on the UNIQUE key).
            test_id = test_of(db, name)
            if by_case:
                db.execute("INSERT OR IGNORE INTO fe_test_case (test_id, test_case) VALUES (?, ?)", (test_id, test_case))
                case_id = db.execute("SELECT id FROM fe_test_case WHERE test_id = ? AND test_case = ?",
                                     (test_id, test_case)).fetchone()[0]
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_case_id, result) "
                               "VALUES (?, ?, ?, ?)", [(file_id, known[i], case_id, result(i)) for i in ids if i in known])
            else:
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_id, result) "
                               "VALUES (?, ?, ?, ?)", [(file_id, known[i], test_id, result(i)) for i in ids if i in known])

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None):
        """Load a fiches file (XML, default platform): it belongs to the `tome` given, or to none ("": every tome).
        The numbers absent from its tome must be ignored by the user (`decisions`). The file is stored, then its
        content (the file is removed again on failure); a file of the same name is replaced, whatever its tome."""
        tome = files.check_tome("de fiches d'essai", requirements, tome)
        data = self._readable(content, config)
        # Numbers absent from the requirements of the tome: ignored by the user, not stored.
        ignored = FileDatabase.to_ignore(self.sql.db, data["ids"], tome)
        # Refused unless the user ignored every one of them (the page asks before loading).
        FileDatabase.check_ignored(ignored, decisions)
        # Copy stored first, content next, in one transaction: on failure the copy is removed again.
        record = files.add("xml", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data)
                # An XML fiche comes from the default platform.
                db.execute("UPDATE files SET platform_id = (SELECT id FROM list_plateformes WHERE is_default = 1) "
                           "WHERE id = ?", (record["id"],))
        except Exception:
            files.remove(record["id"])
            raise
        return {"file_id": record["id"], "ids": len(data["ids"]) - len(ignored), "ignored": len(ignored),
                "replaced": files.remove_older(record)}

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
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data)
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
        # A row points to a test case (fiches XML: its test through it) or to a test (other platforms: no test case).
        for r in self.sql.query("SELECT x.file_id, COALESCE(x.test_id, c.test_id) AS test_id, t.name AS test, "
                                "COALESCE(c.test_case, '') AS test_case, q.req_id FROM fiches_essai x "
                                "LEFT JOIN fe_test_case c ON c.id = x.test_case_id "
                                "JOIN fe_tests t ON t.id = COALESCE(x.test_id, c.test_id) "
                                "JOIN requirements q ON q.id = x.requirement_id ORDER BY x.file_id, x.id"):
            if r["file_id"] not in wanted:
                continue
            cases = rows.setdefault(r["file_id"], [])
            if not cases or (cases[-1]["test_id"], cases[-1]["test_case"]) != (r["test_id"], r["test_case"]):
                cases.append({"test_id": r["test_id"], "test": r["test"], "test_case": r["test_case"], "requirements": []})
            cases[-1]["requirements"].append(r["req_id"])
        return [{"file_id": h["file_id"], "name": h["name"], "source": source, "platform": h["platform"] or "",
                 "tome": h["tome"], "releases": releases.get(h["file_id"], []),
                 "ids": list(dict.fromkeys(i for c in rows.get(h["file_id"], []) for i in c["requirements"])),
                 "rows": rows.get(h["file_id"], [])} for h in heads]

    def to_dict(self):
        """The fiches files by file number, for the page."""
        return {"fiches": {str(d["file_id"]): d for d in self.documents()}}
