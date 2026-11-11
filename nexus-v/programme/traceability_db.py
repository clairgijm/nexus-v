"""Fiches d'essai of the other platforms: Excel files, in the SQL database (tables files and fiches_essai).

These Excel files run beside the XML fiches: they trace the requirements delegated to other platforms. Like an XML
fiches file, a file belongs to a tome or to none (it then applies to every tome); removing it removes its content
(cascade).

Their format is not described yet: only the requirement numbers of its tome found in its cells are stored, in the
table fiches_essai (one test without name, no test case). The file keeps the platform it comes from, chosen at load.
A number unknown in its tome is shown at import, where it can only be ignored. A file without any number is refused.

For the page (to_dict):

    {"traceability": {"12": {"file_id": 12, "name": "...xlsx", "tome": "fct1" (None: every tome), "releases": ["A"],
                             "platform": "PLT1_N-1", "ids": [ID]}}}
"""
import io
import re

import openpyxl

from configuration import DEFAULT_CONFIG
from file_database import FileDatabase


class TraceabilityDatabase:
    """Requirement numbers cited by the traceability files of the other platforms, in the SQL database."""

    def __init__(self, sql):
        """sql: the shared NexusDatabase."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def read_excel(content, config=None):
        """Read a traceability workbook (bytes). Returns {"ids": requirement numbers cited in any cell}."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        regex = re.compile(get("requirement_regex"))
        try:
            # Read only and values only (not formulas): the format of these files is not defined yet, so every cell of
            # every sheet is searched for requirement numbers.
            workbook = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except Exception as e:   # not an .xlsx / .xlsm workbook
            raise ValueError(f"classeur Excel illisible (formats .xlsx et .xlsm) : {e}") from None
        try:
            ids = {}
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for value in row:
                        if value is not None:
                            # A dict keeps the numbers once, in the order they are met.
                            ids.update(dict.fromkeys(regex.findall(str(value))))
            return {"ids": list(ids)}
        finally:
            workbook.close()

    @classmethod
    def _readable(cls, content, config=None):
        """Content of a file that can be loaded: ValueError when no requirement number was found."""
        data = cls.read_excel(content, config)
        if not data["ids"]:
            raise ValueError("aucun numéro d'exigence trouvé : fichier non chargé")
        return data

    def _known(self, tome):
        """Numbers of the requirements of a tome (None without tome)."""
        # Without tome the numbers are checked against the whole base (FileDatabase.to_ignore).
        if not tome:
            return None
        tome_id = self.sql.tome_id(tome)
        return set(FileDatabase.tome_requirements(self.sql.db, tome_id)) if tome_id is not None else set()

    def analyse(self, content, config=None, tome=None, releases=()):
        """What loading a traceability file would do: number of requirement numbers and the numbers unknown in the
        `tome` (in the whole base without tome), to ignore (same form as the plan entries to settle). Refused
        (ValueError) when unreadable or without any number."""
        ids = self._readable(content, config)["ids"]
        return {"ids": len(ids), "to_confirm": FileDatabase.to_ignore(self.sql.db, ids, tome or None)}

    @staticmethod
    def _write(db, file_id, ids):
        """Replace the requirements of a file in fiches_essai (inside a transaction): the requirements of its tome (of
        the whole base without tome), under one test without name (fe_tests) nor test case."""
        from fiches_db import FichesDatabase
        # Stored like a fiche: one test without name and without test case, the requirements of its tome (all of them
        # without tome).
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        FichesDatabase.store_tests(db, file_id, [("", "", list(dict.fromkeys(ids)))], tome_id, by_case=False)

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None, platform=""):
        """Load a fiches file of another platform (Excel): attached to a tome that has requirements or to none ("":
        every tome), and to the `platform` it comes from. The numbers unknown in its tome must be ignored by the user
        (`decisions` {"rows": {number: {"ignore": true}}}, or {"accept_all": true} on the command line). The file is
        stored, then its content (the file is removed again on failure); a file of the same name is replaced,
        whatever its tome."""
        tome = files.check_tome("de fiches d'une autre plateforme", requirements, tome)
        # The platform the file comes from is required: coverage is judged against the delegation of the plans.
        if config is not None and not config.is_platform(platform or ""):
            raise ValueError("indiquer la plateforme d'où vient ce fichier (une plateforme de la configuration)")
        data = self._readable(content, config)
        # Numbers absent from the tome (from the base without tome): ignored by the user, not stored.
        unknown = FileDatabase.to_ignore(self.sql.db, data["ids"], tome)
        FileDatabase.check_ignored(unknown, decisions)
        record = files.add("traceability", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data["ids"])
                if config is not None:
                    # The platform as written in the configuration (the page may send it in another case).
                    name_ok = next(p for p in config.get("platforms") if p.casefold() == platform.casefold())
                    db.execute("UPDATE files SET platform_id = (SELECT id FROM list_plateformes WHERE name = ?) "
                               "WHERE id = ?", (name_ok, record["id"]))
        except Exception:
            files.remove(record["id"])
            raise
        return {"file_id": record["id"], "tome": tome, "ids": len(data["ids"]) - len(unknown), "ignored": len(unknown),
                "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Read and store the content of the traceability files that have none yet, and of all of them with `again`
        (the requirements changed). Returns the names read."""
        # Files without any row in fiches_essai have not been read yet (new base, tables emptied).
        known = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM fiches_essai")}   # rows stored
        done = []
        for record in files.list(slot="traceability"):
            if again or record["id"] not in known:
                try:
                    data = self.read_excel(files.content(record["id"]) or b"", config)
                except ValueError:
                    continue
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data["ids"])
                done.append(record["name"])
        return done

    def documents(self):
        """The fiches files of the other platforms, in the same form as the Fiches d'essai (FichesDatabase)."""
        from fiches_db import FichesDatabase
        return FichesDatabase(self.sql).documents("traceability")

    def to_dict(self):
        """The traceability files by file number, for the page."""
        return {"traceability": {str(d["file_id"]): d for d in self.documents()}}
