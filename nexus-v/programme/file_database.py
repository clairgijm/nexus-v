"""Files loaded in the page, in the SQL database (table files). Every file belongs to the release of the base.

Each loaded file is copied into the storage folder as « <n°>_<name of the file> », the number growing at each load
(never reused, even after a removal). A record (dict) has the fields:

    id, slot, name, stored_name, tome, releases

The tome is the one of a requirements file or a plan, and for a STDHL or fiches file the one chosen at load (None:
every tome). The module also reads Word documents (WordDocument), with the standard library only.
"""
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from configuration import DEFAULT_CONFIG, plural

SLOTS = ("requirements", "plan", "stdhl", "xml", "traceability")
# Setting giving the sub-folder of the storage folder for each type of file.
FOLDER_SETTINGS = {"requirements": "requirements_folder", "plan": "plan_folder", "stdhl": "stdhl_folder",
                   "xml": "fiches_folder", "traceability": "traceability_folder"}

COLUMNS = "id, source, name, stored_name"


# Reading of a Word document (.docx), with the standard library only.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Characters Word inserts in place of ordinary spaces / hyphens.
# Non-breaking and narrow spaces become spaces, the various Word hyphens become "-", the soft hyphen disappears: a
# requirement number typed in Word then matches requirement_regex.
NORMALIZATION = str.maketrans({" ": " ", " ": " ", "‑": "-", "‐": "-", "–": "-", "­": ""})


class WordDocument:
    """A .docx document read from its bytes."""

    # Parts of the document whose text counts: body, headers, footers, notes, comments.
    TEXT_PARTS = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml")

    def __init__(self, content):
        """Open a .docx from its bytes. ValueError when it is not a readable Word document."""
        try:
            self._zip = zipfile.ZipFile(io.BytesIO(content))
            self.body = ET.fromstring(self._zip.read("word/document.xml"))
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
            raise ValueError(f"document Word illisible (format .docx attendu) : {e}") from None

    @staticmethod
    def normalize(text):
        """Text with the special spaces and hyphens of Word replaced by ordinary ones."""
        return (text or "").translate(NORMALIZATION)

    @classmethod
    def paragraph_text(cls, p):
        """Text of a paragraph, with its "runs" stitched back together
        (Word often splits a single word into several pieces, e.g. "REQ-" + "001")."""
        pieces = []
        # Text runs (w:t) in order; a non-breaking hyphen, a tab or a line break are separate elements in Word XML.
        for el in p.iter():
            if el.tag == f"{W}t":
                pieces.append(el.text or "")
            elif el.tag == f"{W}noBreakHyphen":
                pieces.append("-")
            elif el.tag in (f"{W}tab", f"{W}br"):
                pieces.append(" ")
        return cls.normalize("".join(pieces))

    @classmethod
    def table_grid(cls, tbl):
        """Text of the cells of a table, row by row (paragraphs of a cell joined by line breaks)."""
        # One list per row (w:tr), one string per cell (w:tc); a merged cell appears once.
        return [["\n".join(cls.paragraph_text(p) for p in tc.iter(f"{W}p")).strip()
                 for tc in tr if tc.tag == f"{W}tc"]
                for tr in tbl if tr.tag == f"{W}tr"]

    def tables(self):
        """Every table of the body (nested ones included), in document order: (number from 1, grid)."""
        return [(i, self.table_grid(tbl)) for i, tbl in enumerate(self.body.iter(f"{W}tbl"), 1)]

    def blocks(self):
        """Top-level paragraphs and tables of the body, in document order: ("p", None, text) or
        ("tbl", number from 1 among top-level tables, grid). What is inside a table belongs to the table."""
        number = 0

        # Recursive walk: paragraphs inside a text box or a section are found too, but a table is returned whole (its
        # own paragraphs are not repeated as paragraphs).
        def walk(parent):
            """Paragraphs and tables under `parent`, in order, without going inside a table."""
            nonlocal number
            for el in parent:
                if el.tag == f"{W}tbl":
                    number += 1
                    yield "tbl", number, self.table_grid(el)
                    continue
                if el.tag == f"{W}p":
                    yield "p", None, self.paragraph_text(el)
                yield from walk(el)

        yield from walk(self.body)

    def paragraphs(self):
        """Paragraphs of the body outside the tables, in document order: [(style id, text, list item)]. The style id
        is the one of the Word file (« Titre3 », « TM1 »...), "" for a paragraph without style; a list item is a
        paragraph numbered or bulleted by Word (numPr), or of a list style (« Paragraphedeliste », « ListParagraph »)."""
        found = []

        def walk(parent):
            """Paragraphs under `parent`, in order, without going inside a table."""
            for el in parent:
                if el.tag == f"{W}tbl":
                    continue
                if el.tag == f"{W}p":
                    style = el.find(f"{W}pPr/{W}pStyle")
                    style = style.get(f"{W}val") if style is not None else ""
                    item = el.find(f"{W}pPr/{W}numPr") is not None or "list" in style.lower() or "liste" in style.lower()
                    found.append((style, self.paragraph_text(el), item))
                walk(el)

        walk(self.body)
        return found

    def text(self):
        """Full text: body, tables, headers, footers, notes, comments."""
        paragraphs = []
        for part in self._zip.namelist():
            if self.TEXT_PARTS.fullmatch(part):
                root = ET.fromstring(self._zip.read(part))
                paragraphs.extend(self.paragraph_text(p) for p in root.iter(f"{W}p"))
        return "\n".join(paragraphs)


class FileDatabase:
    """Loaded files: records in the SQL database, copies in the storage folder.

    A record is a dict: id (int), slot (str), name (str), stored_name (str), tome (str, or None for a STDHL or
    fiches file that applies to every tome), releases (list: the release of the base).
    """

    EDITABLE = ("slot", "name")

    def __init__(self, directory, sql, config=None):
        """directory: storage folder (copies of the files); sql: NexusDatabase; config: gives the sub-folder of each
        type of file (defaults without one)."""
        self.directory = Path(directory)
        self.sql = sql
        self.path = sql.path
        self.config = config

    def folder(self, slot):
        """Sub-folder of the storage folder for a type of file."""
        get = self.config.get if self.config is not None else DEFAULT_CONFIG.get
        return self.directory / get(FOLDER_SETTINGS[slot])

    @staticmethod
    def _insert(db, r):
        """Add a file record and its releases (inside a transaction); r["tome"] is a tome name."""
        db.execute(f"INSERT INTO files ({COLUMNS}, tome_id) "
                   "VALUES (?, ?, ?, ?, (SELECT id FROM list_tomes WHERE name = ?))",
                   (r["id"], r["slot"], r["name"], r["stored_name"], r["tome"]))
        # No release stored: every file belongs to the release of the base.

    # ------------------------------------------------------------------ #
    # Access
    # ------------------------------------------------------------------ #

    def _records(self, where="", params=()):
        """Records matching an SQL condition, each with its tome name and its releases."""
        # The column source of the table is the slot of the record; the tome is given by its name.
        rows = self.sql.query(f"SELECT {COLUMNS}, format, source AS slot, (SELECT t.name FROM list_tomes t WHERE t.id = tome_id) AS tome "
                              f"FROM files {where} ORDER BY id", params)
        releases = self.sql.file_releases()
        return [{**r, "releases": releases.get(r["id"], [])} for r in rows]

    def list(self, slot=None):
        """Records, optionally those of one slot."""
        conditions, params = [], []
        if slot is not None:
            conditions.append("source = ?")
            params.append(slot)
        return self._records(("WHERE " + " AND ".join(conditions)) if conditions else "", params)

    def get(self, file_id):
        """Record with this ID, or None."""
        found = self._records("WHERE id = ?", (file_id,))
        return found[0] if found else None

    def set(self, file_id, **fields):
        """Update editable fields of a record. Returns the record, or None if unknown."""
        unknown = set(fields) - set(self.EDITABLE)
        if unknown:
            raise ValueError(f"champs non modifiables : {', '.join(sorted(unknown))}")
        if "slot" in fields and fields["slot"] not in SLOTS:
            raise ValueError(f"type de fichier inconnu : {fields['slot']}")
        with self.sql.transaction() as db:
            if not db.execute("SELECT 1 FROM files WHERE id = ?", (file_id,)).fetchone():
                return None
            if fields:
                # Only the editable fields reach this SQL (checked above), so the column names are safe.
                assignments = ', '.join([f"{k.replace('slot', 'source')} = ?" for k in fields])

                query = f"UPDATE files SET {assignments} WHERE id = ?"
                params = (*fields.values(), file_id)

                db.execute(query, params)
        return self.get(file_id)

    def path_of(self, file_id):
        """Path of the stored copy of a file, or None if unknown."""
        record = self.get(file_id)
        return self.folder(record["slot"]) / record["stored_name"] if record else None

    def content(self, file_id):
        """Bytes of the stored copy of a file, or None if unknown or missing on disk."""
        path = self.path_of(file_id)
        return path.read_bytes() if path and path.is_file() else None

    # ------------------------------------------------------------------ #
    # Adding / removing files
    # ------------------------------------------------------------------ #

    @staticmethod
    def safe_name(name):
        """File name without directory parts nor characters Windows forbids."""
        # Keep only the last part of a path (a browser may send C:\fakepath\name), then replace the characters Windows
        # forbids.
        name = Path(name.replace("\\", "/")).name
        return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .") or "fichier"

    def add(self, slot, name, content, tome=None, releases=None):
        """Store a new file as "<n°>_<name>" and record it. Returns the new record.

        tome: tome of a requirements file (created if new), of an acceptance plan, or of a STDHL or fiches file
        (must exist, checked by the caller; None: every tome); releases: kept for the callers, every file belongs to
        the release of the base.
        """
        if slot not in SLOTS:
            raise ValueError(f"type de fichier inconnu : {slot}")
        name = self.safe_name(name)
        releases = list(dict.fromkeys(r.strip().upper() for r in releases or [] if r.strip()))
        with self.sql.transaction() as db:
            # The number comes from meta.next_file_id and is never reused, even after a deletion: a stored copy always
            # has a new name.
            file_id = int(self.sql.meta("next_file_id", 1))
            stored_name = f"{file_id:04d}_{name}"
            if tome:
                self.sql.ensure("tomes", [tome], db)
            self._insert(db, {"id": file_id, "slot": slot, "name": name, "stored_name": stored_name,
                              "tome": tome or None, "releases": releases})
            self.sql.set_meta("next_file_id", file_id + 1, db)
            # The copy is written inside the transaction: if it fails, nothing is recorded.
            folder = self.folder(slot)
            folder.mkdir(parents=True, exist_ok=True)
            (folder / stored_name).write_bytes(content)
        return self.get(file_id)

    @staticmethod
    def to_ignore(db, ids, tome=None):
        """Numbers of a fiches file that can only be ignored at import, in the form of the rows to settle: absent from
        the requirements (of the `tome` if given)."""
        # The numbers known: of the tome if given, of the whole base otherwise.
        own = {r[0] for r in db.execute("SELECT q.req_id FROM requirements q LEFT JOIN list_tomes t ON t.id = q.tome_id"
                                        + (" WHERE t.name = ?" if tome else ""), (tome,) if tome else ())}
        return [{"key": i, "requirements": [i], "places": [], "iadt": "", "moc": "", "delegation": "", "choose": [],
                 "unknown": True, "expected_mocs": [], "duplicate": [],
                 "issues": [f"exigence absente de la base des exigences{' du tome' if tome else ''} — seule possibilité, l'ignorer"]}
                for i in dict.fromkeys(ids) if i not in own]

    @staticmethod
    def check_ignored(found, decisions):
        """Every number of `found` (see to_ignore) must be ignored by the user (`decisions` {"rows": {number: {"ignore":
        true}}}), or {"accept_all": true} on the command line. ValueError otherwise."""
        chosen = (decisions or {}).get("rows") or {}
        # Every unknown number must have been ignored by the user; on the command line accept_all ignores them all.
        left = [c["key"] for c in found if not (chosen.get(c["key"]) or {}).get("ignore")]
        if left and not (decisions or {}).get("accept_all"):
            raise ValueError(f"{plural(len(left), 'numéro', 'numéros')} à ignorer avant le chargement (absents de la base "
                             f"des exigences) : {', '.join(left)}")

    @staticmethod
    def tome_requirements(db, tome_id):
        """{requirement number: id} of the requirements of a tome."""
        return {r["req_id"]: r["id"] for r in db.execute("SELECT id, req_id FROM requirements WHERE tome_id = ?", (tome_id,))}

    @staticmethod
    def check_attachment(kind, requirements, tome):
        """A plan or traceability file belongs to a tome that has requirements: ValueError otherwise. `kind`: name of
        the file type in the messages."""
        if not requirements.get(tome):
            raise ValueError(f"un fichier {kind} doit être rattaché à un tome contenant des exigences")


    @staticmethod
    def check_tome(kind, requirements, tome):
        """A STDHL or fiches file belongs to one tome that has requirements, or to none: it then applies to every tome
        (a single STDHL for all of them, for instance). ValueError otherwise. Returns the tome, or None."""
        if not requirements.tomes():
            raise ValueError("importer d'abord les Exigences")
        tome = (tome or "").strip()
        if tome and not requirements.get(tome):
            raise ValueError(f"un fichier {kind} est rattaché à un tome contenant des exigences, ou à aucun")
        return tome or None

    def remove_older(self, record):
        """Remove the other files of the same type and name as `record` (and the same tome for a requirements file or a
        plan; a STDHL or fiches file loaded again replaces the old one whatever its tome). Returns their numbers."""
        # A requirements file or a plan of the same name in another tome is another document; a STDHL or fiches file
        # loaded again replaces the old one even if its tome changed.
        same_tome = record["slot"] in ("requirements", "plan")
        older = [r["id"] for r in self.list(slot=record["slot"]) if r["id"] != record["id"] and r["name"] == record["name"]
                 and (not same_tome or (r["tome"] or "").casefold() == (record["tome"] or "").casefold())]
        for file_id in older:
            self.remove(file_id)
        return older

    def remove(self, file_id, delete_file=True):
        """Remove a record (and by default its stored copy). Returns False if unknown.

        The requirements imported from the file and the content of a plan are removed with it (cascade).
        """
        record = self.get(file_id)
        if record is None:
            return False
        with self.sql.transaction() as db:
            db.execute("DELETE FROM files WHERE id = ?", (file_id,))
        if delete_file:
            (self.folder(record["slot"]) / record["stored_name"]).unlink(missing_ok=True)
        return True
