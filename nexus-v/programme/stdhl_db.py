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

from configuration import DEFAULT_CONFIG, header_column, header_text, match_header, plural, table_left_out
from file_database import WordDocument


def found(match):
    """Text recognised by a configured expression: its group 1 if it has one, otherwise the whole match."""
    return (match[1] if match.re.groups else match[0]).strip()


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
        return {"stdhl_name": found(match) if match else file_name.rsplit(".", 1)[0], "rows": rows, "tables": tables,
                "skipped": skipped, "issues": issues}

    @classmethod
    def _readable(cls, content, file_name, config=None):
        """Content of a STDHL that can be loaded: refused (ValueError) only when no table has a correct header.
        Tables without any recognised requirement number are an anomaly of the document, not a refusal."""
        data = cls.read_docx(content, file_name, config)
        if not data["tables"]:
            raise ValueError("aucun tableau STDHL valide trouvé (voir Configuration > En-têtes des tableaux) : "
                             "fichier non chargé")
        return data

    @classmethod
    def analyse(cls, content, file_name, config=None):
        """What loading a STDHL would do. Refused (ValueError) when unreadable or without a valid table."""
        data = cls._readable(content, file_name, config)
        return {"tables": data["tables"], "rows": len(data["rows"]), "stdhl_name": data["stdhl_name"],
                "skipped": data["skipped"], "issues": data["issues"]}

    # ------------------------------------------------------------------ #
    # Storing
    # ------------------------------------------------------------------ #

    @staticmethod
    def _write(db, file_id, data):
        """Replace the stored content of a STDHL (inside a transaction): one row of stdhl_exi per requirement it
        cites and test (fiche SYS_) it asks for it, the test taken in the referential fe_tests. A number absent from
        the requirements of the base is not stored: it is an anomaly of the document. A requirement N/A in the plans
        (IADT and MOC N/A on all its rows: not tested) is not stored either; returns their numbers, reported at the
        load of the file."""
        from nexus_db import test_id as test_of
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
                        db.execute("INSERT INTO stdhl_exi (file_id, requirement_id, test_id) VALUES (?, ?, ?)",
                                   (file_id, known[number], test))
                elif number not in unknown:
                    unknown.append(number)
        issues = data["issues"] + ([plural(len(unknown), "numéro d'exigence absent de la base, non enregistré",
                                           "numéros d'exigence absents de la base, non enregistrés")
                                    + " : " + ", ".join(unknown)] if unknown else [])
        # The STDHL name and the anomalies are kept on the file row: the page STDHL shows them.
        db.execute("UPDATE files SET stdhl_name = ?, issues = ? WHERE id = ?",
                   (data["stdhl_name"], "\n".join(issues), file_id))
        return left_out

    def load_file(self, files, requirements, content, name, tome, releases, config=None):
        """Load a STDHL: it belongs to the `tome` given, or to none ("": every tome). The file is stored, then its
        content (the file is removed again on failure); a STDHL of the same name is replaced, whatever its tome.
        "na_ignored": the requirements N/A in the plans that it cites, left out."""
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
                "replaced": files.remove_older(record)}

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

    def documents(self):
        """The STDHL documents with their tome (None: every tome), rows and anomalies."""
        heads = self.sql.query("SELECT f.id AS file_id, f.name, f.stdhl_name, f.issues, t.name AS tome FROM files f "
                               "LEFT JOIN list_tomes t ON t.id = f.tome_id WHERE f.source = 'stdhl' ORDER BY f.id")
        rows, releases = {}, self.sql.file_releases()
        # One row per requirement and test asked, in the order of the document.
        for r in self.sql.query("SELECT x.file_id, q.req_id, t.name AS test FROM stdhl_exi x "
                                "JOIN requirements q ON q.id = x.requirement_id LEFT JOIN fe_tests t ON t.id = x.test_id "
                                "ORDER BY x.file_id, x.id"):
            # Rows of a file deleted in between are ignored by the heads below.
            rows.setdefault(r["file_id"], []).append({"requirements": [r["req_id"]], "test": r["test"] or ""})
        return [{"file_id": h["file_id"], "name": h["name"], "releases": releases.get(h["file_id"], []),
                 "stdhl_name": h["stdhl_name"], "tome": h["tome"], "rows": rows.get(h["file_id"], []),
                 "issues": [i for i in h["issues"].split("\n") if i]} for h in heads]

    def to_dict(self):
        """The STDHL by file number, for the page."""
        return {"stdhl": {str(d["file_id"]): d for d in self.documents()}}
