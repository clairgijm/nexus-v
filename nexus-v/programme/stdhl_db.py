"""STDHL of NEXUS-V, in the SQL database (tables files, stdhl_exi and fe_tests).

When a STDHL is loaded, its useful content (read in the Word document) is stored here. It is identified by the number
of its file (table files); removing the file removes its content (cascade). At load, the user attaches the STDHL to a
tome, or to none: it then applies to every tome (a single STDHL for the whole specification).

The name of the STDHL comes from the name of its file: [date]_[text]_STDHL_[STDHL name]_rev[text].docx. The fiches
(tests) are the paragraphs SYS_[type]_[features]_[N°]; each is followed by tables whose header matches the setting
headers_stdhl. Only the requirement numbers of the column « Exigences associées » are stored, each with its fiche
(the test asked); the check then makes sure that each test asked covers the requirement in the fiches.

A table without the expected columns is left out without stopping the reading; the file is refused only when no table
has a correct header. The anomalies of the document (file name, table without fiche, table left out, no requirement
number recognised) are stored with it and shown on the page STDHL.

For the page (to_dict):

    {"stdhl": {"12": {"file_id": 12, "name": "...docx", "releases": ["A"], "stdhl_name": "FCT1",
                      "tome": "fct1" (None: every tome),
                      "rows": [{"requirements": [ID], "test": fiche SYS_}, ...] (one row per requirement and test),
                      "issues": [text]}}}
"""
import re

from configuration import DEFAULT_CONFIG, fold, header_column, header_text, match_header, plural, table_left_out
from file_database import WordDocument


def found(match):
    """Text recognised by a configured expression: its group 1 if it has one, otherwise the whole match."""
    return (match[1] if match.re.groups else match[0]).strip()


# Formats of STDHL recognised at import, kept in files.format: 1 = tables (one STDHL per tome), 2 = one document in text.
FORMATS = {1: "par tome", 2: "unique"}


def unique(values):
    """Non-empty values without duplicates, in their order."""
    return list(dict.fromkeys(v for v in values if v))


class StdhlDatabase:
    """Content of the STDHL documents, in the SQL database."""

    def __init__(self, sql):
        """sql: the shared NexusDatabase."""
        self.sql = sql
        self.path = sql.path

    # ------------------------------------------------------------------ #
    # Reading a STDHL document
    # ------------------------------------------------------------------ #

    @staticmethod
    def _settings(config):
        """Expected headers and requirement number format, from a configuration (defaults without one)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        return get("headers_stdhl"), re.compile(get("requirement_regex"))

    @staticmethod
    def _find_header(grid, headers):
        """(requirements column, [expected texts missing]) of the header of a table, always its first row, or None.
        A STDHL table is recognised by the header of its requirements' column."""
        text = header_text(headers, "requirements")
        cells = grid[0] if grid else []
        # Only the requirements column identifies a STDHL table; the other expected columns are then checked, and a
        # table missing one of them is reported and left out.
        if header_column(cells, text) is not None:
            cols, missing = match_header(cells, headers)
            return cols["requirements"], missing
        return None

    @classmethod
    def read_docx(cls, content, file_name, config=None):
        """Read a STDHL (.docx bytes). Returns {"stdhl_name", "rows": [{fiche, requirements: [ID]}], "skipped":
        tables left out, "issues": anomalies of the document}. A row is the cell of a table under the requirements
        header; its fiche (the test) is the last SYS_ test sheet before the table. Tables are numbered at top level."""
        headers, regex = cls._settings(config)
        get = config.get if config is not None else DEFAULT_CONFIG.get
        fiche_regex = re.compile(get("fiche_regex"))
        match = re.search(get("stdhl_name_regex"), file_name)
        rows, skipped, orphans = [], [], []
        tables = 0   # tables with a correct header, whatever their cells hold
        fiche = ""
        # The document is walked in reading order: a paragraph matching fiche_regex opens a fiche (the test), and every
        # table after it belongs to that fiche until the next one.
        for kind, number, value in WordDocument(content).blocks():
            if kind == "p":
                if m := fiche_regex.search(value):
                    fiche = found(m)
                continue
            header = cls._find_header(value, headers)
            if header is None:
                continue
            col, missing = header
            if missing:
                skipped.append(table_left_out(number, missing))
                continue
            tables += 1
            # A valid table before any fiche: its requirements are kept, without test, and the document is reported.
            if not fiche:
                orphans.append(number)
            # Each row of the table under the header: the requirement numbers found in its requirements cell.
            for cells in value[1:]:
                ids = unique(regex.findall(cells[col])) if col < len(cells) else []
                if ids:
                    rows.append({"fiche": fiche, "requirements": ids})
        label = header_text(headers, "requirements")
        # Anomalies of the document, kept with it and shown on the page STDHL: file name, tables left out, tables
        # without fiche, no requirement number at all.
        issues = ([] if match else ["nom de fichier hors format [date]_[texte]_STDHL_[Nom du STDHL]_rev[texte].docx"]) \
            + skipped + [f"tableau {t} « {label} » sans fiche SYS_ avant lui" for t in orphans] \
            + ([f"aucun numéro d'exigence reconnu dans {'les ' + str(tables) + ' tableaux' if tables > 1 else 'le tableau'}"
                f" : vérifier la colonne « {label} » et le format des numéros (Configuration)"] if tables and not rows else [])
        if not tables:
            # No STDHL table: the document may be a STDHL of format 2 (every STDHL in one document, in text).
            single = cls.read_single(content, file_name, config)
            if single["tables"]:
                return single
        return {"stdhl_name": found(match) if match else file_name.rsplit(".", 1)[0], "rows": rows, "tables": tables,
                "skipped": skipped, "issues": issues, "format": FORMATS[1], "gaps": [], "refusals": []}

    @classmethod
    def _readable(cls, content, file_name, config=None):
        """Content of a STDHL that can be loaded: refused (ValueError) when no table has a correct header and no test
        of format 2 is found, or when a test of format 2 is for another release than the base. Tables without any
        recognised requirement number are an anomaly of the document, not a refusal."""
        data = cls.read_docx(content, file_name, config)
        if not data["tables"]:
            raise ValueError("aucun tableau STDHL valide (format par tome) ni test du format unique trouvé (voir "
                             "Configuration > STDHL) : fichier non chargé")
        refusals = data["refusals"] + [
            f"test {g['test']} : {g['requirement']} " + ("citée dans les test cases mais absente des exigences associées"
                                                         if g["side"] == "test case" else
                                                         "dans les exigences associées mais citée dans aucun test case")
            for g in data["gaps"]]
        if refusals:
            raise ValueError(" ; ".join(refusals) + " : fichier non chargé")
        return data

    @classmethod
    def read_single(cls, content, file_name, config=None):
        """A STDHL of format 2 (« unique ») in the form of read_docx: one row per test and requirement, "tables"
        being the number of tests. The coverage of a test is the requirements of its test cases; its list « Exigences
        associées » must hold exactly them: a requirement in one part only refuses the file ("gaps": [{test,
        requirement, side « liste » or « test case »}]), as a release of a test that is not the one of the base
        ("refusals"). A test without release, a test name repeated, a test without requirement are anomalies of the
        document ("issues")."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        data = cls.read_text_format(content, file_name, config)
        release = (get("release") or "").upper()
        tests, rows, issues, refusals, gaps = data["tests"], [], [], [], []
        names = [t["name"] for t in tests]
        issues += [f"test {n} présent {names.count(n)} fois" for n in sorted({n for n in names if names.count(n) > 1})]
        for t in tests:
            if t["releases"] is None:
                issues.append(f"test {t['name']} : applicabilité absente")
            elif release and release not in t["releases"]:
                refusals.append(f"test {t['name']} : applicabilité {', '.join(t['releases']) or 'illisible'}, la base "
                                f"est en release {release}")
            covered = unique(i for c in t["cases"] for i in c["requirements"])
            listed = t["listed"]
            if not covered and not listed:
                issues.append(f"test {t['name']} : aucune exigence")
            gaps += [{"test": t["name"], "requirement": i, "side": "test case"} for i in covered if i not in listed]
            gaps += [{"test": t["name"], "requirement": i, "side": "liste"} for i in listed if i not in covered]
            rows += [{"fiche": t["name"], "requirements": [i]} for i in covered]
        return {"stdhl_name": data["stdhl_name"], "rows": rows, "tables": len(tests), "skipped": [], "issues": issues,
                "format": FORMATS[2], "gaps": gaps, "refusals": refusals,
                "cases": sum(len(t["cases"]) for t in tests),
                # The test cases asked, by their full name: a fiche of another test case is ignored.
                "test_cases": list(dict.fromkeys((t["name"], c["name"]) for t in tests for c in t["cases"]))}

    @classmethod
    def read_text_format(cls, content, file_name, config=None):
        """Read a STDHL of format 2 (no STDHL table): its tests, as they are written (see read_single).
        A test is a title « <name> : <title> » (stdhl2_test_regex, a line starting with a requirement number being
        never a test; the table of contents, styles TM* / TOC*, is left out). After it come labelled sections: its
        release (stdhl2_release_label, read by stdhl2_release_regex), the list of its associated requirements
        (stdhl2_requirements_label, then the list items that follow: the end of the list ends it), then its test cases
        (stdhl2_test_case_regex),
        whose phases cite requirements. Returns {"stdhl_name", "tests": [{name, releases (None when absent), listed:
        [numbers], cases: [{name, requirements}]}]}."""
        _, regex = cls._settings(config)
        get = config.get if config is not None else DEFAULT_CONFIG.get
        test_regex = re.compile(get("stdhl2_test_regex"))
        case_regex = re.compile(get("stdhl2_test_case_regex"))
        release_regex = re.compile(get("stdhl2_release_regex"))
        requirements_label, release_label = fold(get("stdhl2_requirements_label")), fold(get("stdhl2_release_label"))
        tests, part = [], None   # part: "list" in the associated requirements, "case" in a test case, None elsewhere
        for style, text, item in WordDocument(content).paragraphs():
            if not text.strip() or style.upper().startswith(("TM", "TOC")):
                continue
            if part == "list" and not item:
                part = None   # the end of the list ends the associated requirements
            starts_with_number = (m := regex.search(text)) is not None and not text[:m.start()].strip()
            if not starts_with_number and (t := test_regex.search(text)):
                tests.append({"name": found(t), "releases": None, "listed": [], "cases": []})
                part = None
                continue
            if not tests:
                continue   # before the first test: introduction of the document
            test = tests[-1]
            label, _, value = text.partition(":")
            label = fold(label)
            if label == release_label:
                test["releases"] = [r.upper() for r in release_regex.findall(value)]
                part = None
            elif label == requirements_label:
                part = "list"
                test["listed"] += regex.findall(value)
            elif case_regex.search(text):
                test["cases"].append({"name": text.strip(), "requirements": regex.findall(text)})
                part = "case"
            elif part == "list":
                test["listed"] += regex.findall(text)
            elif part == "case":
                test["cases"][-1]["requirements"] += regex.findall(text)
        for t in tests:
            t["listed"] = unique(t["listed"])
            for c in t["cases"]:
                c["requirements"] = unique(c["requirements"])
        match = re.search(get("stdhl_name_regex"), file_name)
        return {"stdhl_name": found(match) if match else file_name.rsplit(".", 1)[0], "tests": tests}

    @classmethod
    def analyse(cls, content, file_name, config=None):
        """What loading a STDHL would do. Refused (ValueError) when unreadable or without a valid table."""
        data = cls._readable(content, file_name, config)
        return {"tables": data["tables"], "rows": len(data["rows"]), "stdhl_name": data["stdhl_name"],
                "skipped": data["skipped"], "issues": data["issues"], "format": data["format"],
                "gaps": data["gaps"]}

    # ------------------------------------------------------------------ #
    # Storing
    # ------------------------------------------------------------------ #

    @staticmethod
    def _write(db, file_id, data):
        """Replace the stored content of a STDHL (inside a transaction): one row of stdhl_exi per requirement it
        cites and test (fiche SYS_) it asks for it, the test taken in the referential fe_tests. A number absent from
        the requirements of the base is not stored: it is an anomaly of the document. A requirement N/A in the plans
        (IADT and MOC N/A on all its rows: not tested) is not stored either; returns their numbers, reported at the
        load of the file. A row deleted by the user (page STDHL) stays deleted when the file is read again."""
        from nexus_db import test_id as test_of
        deleted = {f"{r[0] or ''}|{r[1]}" for r in db.execute(
            "SELECT t.name, q.req_id FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
            "LEFT JOIN fe_tests t ON t.id = x.test_id WHERE x.file_id = ? AND x.deleted = 1", (file_id,))}
        db.execute("DELETE FROM stdhl_exi WHERE file_id = ?", (file_id,))
        # Requirement number -> id of the base, to store the coverage by id.
        known = {r[0]: r[1] for r in db.execute("SELECT req_id, id FROM requirements")}
        # Requirements N/A in the plans: every one of their plan rows not deleted has IADT N/A and MOC N/A (MIN of the
        # condition over the rows is 1).
        ended = {r[0] for r in db.execute(
            "SELECT q.req_id FROM plan_acceptation p JOIN requirements q ON q.id = p.requirement_id "
            "LEFT JOIN list_iadt i ON i.id = p.iadt_id LEFT JOIN list_moc m ON m.id = p.moc_id WHERE p.deleted = 0 "
            "GROUP BY p.requirement_id HAVING MIN(COALESCE(i.code, '') = 'N/A' AND COALESCE(m.code, '') = 'N/A') = 1")}
        unknown, left_out = [], []
        for r in data["rows"]:
            for number in r["requirements"]:
                if number in ended:
                    if number not in left_out:
                        left_out.append(number)
                elif number in known:
                    # The test of the row in the referential fe_tests (added when new); a row without fiche has no test.
                    # The same requirement and test may be cited on several lines: stored once (UNIQUE does not cover a
                    # NULL test, hence the check).
                    test = test_of(db, r["fiche"]) if r["fiche"] else None
                    if not db.execute("SELECT 1 FROM stdhl_exi WHERE file_id = ? AND requirement_id = ? "
                                      "AND test_id IS ?", (file_id, known[number], test)).fetchone():
                        key = f"{r['fiche']}|{number}"   # deleted by the user: stays deleted
                        db.execute("INSERT INTO stdhl_exi (file_id, requirement_id, test_id, deleted) VALUES (?, ?, ?, ?)",
                                   (file_id, known[number], test, int(key in deleted)))
                elif number not in unknown:
                    unknown.append(number)
        issues = data["issues"] + ([plural(len(unknown), "numéro d'exigence absent de la base, non enregistré",
                                           "numéros d'exigence absents de la base, non enregistrés")
                                    + " : " + ", ".join(unknown)] if unknown else [])
        # The test cases asked by a STDHL « unique », in the referential fe_test_case (added when new).
        db.execute("DELETE FROM stdhl_test_case WHERE file_id = ?", (file_id,))
        for test, case in data.get("test_cases", []):
            test = test_of(db, test)
            db.execute("INSERT OR IGNORE INTO fe_test_case (test_id, test_case) VALUES (?, ?)", (test, case))
            db.execute("INSERT OR IGNORE INTO stdhl_test_case (file_id, test_case_id) SELECT ?, id FROM fe_test_case "
                       "WHERE test_id = ? AND test_case = ?", (file_id, test, case))
        # The STDHL name and the anomalies are kept on the file row: the page STDHL shows them.
        db.execute("UPDATE files SET stdhl_name = ?, issues = ?, format = ? WHERE id = ?",
                   (data["stdhl_name"], "\n".join(issues), data.get("format", ""), file_id))
        return left_out

    def load_file(self, files, requirements, content, name, tome, releases, config=None):
        """Load a STDHL: it belongs to the `tome` given, or to none ("": every tome). The file is stored, then its
        content (the file is removed again on failure); a STDHL of the same name is replaced, whatever its tome.
        "na_ignored": the requirements N/A in the plans that it cites, left out. A STDHL « unique » whose associated
        requirements and test cases do not match is refused (see _readable)."""
        tome = files.check_tome("STDHL", requirements, tome)
        data = self._readable(content, name, config)
        # The copy is stored first: if writing the content fails, the copy is removed again, nothing half-loaded stays.
        record = files.add("stdhl", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                left_out = self._write(db, record["id"], data)
        except Exception:
            files.remove(record["id"])
            raise
        return {"file_id": record["id"], "rows": len(data["rows"]), "skipped": data["skipped"], "na_ignored": left_out,
                "format": data["format"], "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Read and store the content of the STDHL files that have none yet, and of all of them with `again` (reading
        settings changed, plans or requirements changed). Returns the names read; unreadable files are left as they
        are."""
        # A STDHL without any row stored has never been read (or its rows were emptied): it is read again.
        stored = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM stdhl_exi")}   # rows stored
        done = []
        for record in files.list(slot="stdhl"):
            if record["id"] not in stored or again:
                try:
                    data = self.read_docx(files.content(record["id"]) or b"", record["name"], config)
                except ValueError:
                    continue
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data)
                done.append(record["name"])
        return done

    # ------------------------------------------------------------------ #
    # Access
    # ------------------------------------------------------------------ #

    def documents(self, with_deleted=False):
        """The STDHL documents with their tome (None: every tome), rows and anomalies. The rows deleted by the user
        are left out, unless `with_deleted` (page STDHL: shown struck through, "deleted" true)."""
        heads = self.sql.query("SELECT f.id AS file_id, f.name, f.stdhl_name, f.issues, t.name AS tome FROM files f "
                               "LEFT JOIN list_tomes t ON t.id = f.tome_id WHERE f.source = 'stdhl' ORDER BY f.id")
        rows, releases = {}, self.sql.file_releases()
        # One row per requirement and test asked, in the order of the document.
        for r in self.sql.query("SELECT x.file_id, q.req_id, t.name AS test, x.deleted FROM stdhl_exi x "
                                "JOIN requirements q ON q.id = x.requirement_id LEFT JOIN fe_tests t ON t.id = x.test_id "
                                f"{'' if with_deleted else 'WHERE x.deleted = 0 '}ORDER BY x.file_id, x.id"):
            # Rows of a file deleted in between are ignored by the heads below.
            rows.setdefault(r["file_id"], []).append({"requirements": [r["req_id"]], "test": r["test"] or "",
                                                     "deleted": bool(r["deleted"])})
        return [{"file_id": h["file_id"], "name": h["name"], "releases": releases.get(h["file_id"], []),
                 "stdhl_name": h["stdhl_name"], "tome": h["tome"], "rows": rows.get(h["file_id"], []),
                 "issues": [i for i in h["issues"].split("\n") if i]} for h in heads]

    def to_dict(self):
        """The STDHL by file number, for the page."""
        return {"stdhl": {str(d["file_id"]): d for d in self.documents(with_deleted=True)}}

    def edit(self, file_id, action, requirement, test=""):
        """Row of a requirement and a test in a STDHL deleted by the user (« delete »: kept, struck through, left out
        of the check) or restored (« restore »). ValueError when the row is unknown."""
        if action not in ("delete", "restore"):
            raise ValueError(f"action inconnue : {action}")
        with self.sql.transaction() as db:
            row = db.execute(
                "SELECT x.id FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
                "LEFT JOIN fe_tests t ON t.id = x.test_id WHERE x.file_id = ? AND q.req_id = ? AND COALESCE(t.name, '') = ?",
                (file_id, str(requirement or ""), str(test or ""))).fetchone()
            if row is None:
                raise ValueError(f"ligne inconnue dans ce STDHL : {requirement} / {test or 'sans test'}")
            db.execute("UPDATE stdhl_exi SET deleted = ? WHERE id = ?", (int(action == "delete"), row[0]))
        return {"deleted": action == "delete"}
