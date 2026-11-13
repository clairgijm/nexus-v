"""Requirements of NEXUS-V, in the SQL database (table requirements). Only the requirements of the release of the
base (configuration general.release) are imported.

A requirement belongs to a tome (link to the table list_tomes); it comes from an imported requirements file (file_id:
it is removed with the file, by cascade).

An import into a tome that already exists is either a replacement (the tome gets exactly the new content) or a
complement (new requirements are added, those already there are updated).

Reading a requirements file (read_excel): a header row, then the columns found by their header (ID, title,
releases separated by commas, STDHL, IADT, MOC, delegation). The file is the reference of each requirement: one IADT
code, one MOC code and one delegation (platform), values of the configuration; a delegation is required. A row with
an empty delegation, an IADT or MOC code unknown of the configuration or several values refuses the file; a
delegation to a platform absent from the configuration is a new platform, added to the referential at import.
The column « Specification releases » gives the releases where the requirement applies (only those of the release of
the base are imported); « Validation releases » the releases where it is tested, separated by commas, and/or
QUALIFIE. It gives the field validation of the requirement, for the release of the base: qualified when QUALIFIE is
written (only that counts: the releases written with it are those of tests done before), to_qualify when the release
of the base is in the list, out_of_scope otherwise (an empty cell included: the requirement is outside the validation
of this release). It never refuses the file. The import has two steps: analysis (analyse: sheets, requirements,
requirements outside the release, differences with the tome), then, once approved in the page, storing of the file
and its requirements (import_file). The rows without the release of the base (another release, or none) are shown
to the user grouped by releases: each group is ignored or imported into the release of the base (choice "force").
Before the import, the user settles the conflicts: the row kept for a number found on several rows of the file, the
version kept for a number already in the base, the value kept for several STDHL values.

A requirement can be changed, added or deleted by hand (edit): it is then marked (manual = "added" / "modified",
deleted = 1: shown struck through and left out of the check, it can be restored). The next import of the tome
overwrites those changes: the file stays the reference.

For the page, requirements are given by tome (to_dict):

    {"tomes": {"fct1": {"description": "...", "files": [3, 7],
                        "requirements": [{"id", "title", "releases", "stdhl", "iadt", "moc", "delegation", "file_id",
                                          "manual", "deleted"}, ...]}}}
"""
import io
import re

import openpyxl

from configuration import DEFAULT_CONFIG, fold, match_header, plural

MODES = ("replace", "complement")

# Fields compared between the requirements of a tome and those of a file.
DIFF_FIELDS = ("title", "releases", "validation", "stdhl", "iadt", "moc", "delegation")
QUALIFIED = "QUALIFIE"   # value of « Validation releases »: the requirement is qualified
# Validation of a requirement in the release of the base (field validation), and its label.
VALIDATION_STATES = {"to_qualify": "à qualifier", "qualified": "qualifiée", "out_of_scope": "hors scope"}
# Reference of a requirement: (field, table, column of the value).
REFERENCE = (("iadt", "list_iadt", "code"), ("moc", "list_moc", "code"), ("delegation", "list_plateformes", "name"))

# Characters Word / Excel put in place of ordinary spaces and hyphens.
# A requirement number pasted from Word or a web page often carries such characters: they would not match.
NORMALIZATION = str.maketrans({"\u00a0": " ", "\u202f": " ", "\u2011": "-", "\u2010": "-", "\u2013": "-", "\u00ad": ""})


class RequirementsDatabase:
    """Requirements, by tome, in the SQL database.

    A requirement is a dict: id (str), title (str), releases (list of str), stdhl (str),
    file_id (int or None: number of the imported file it comes from).
    """

    def __init__(self, sql):
        """sql: the shared NexusDatabase."""
        self.sql = sql
        self.path = sql.path

    # ------------------------------------------------------------------ #
    # Reading / writing
    # ------------------------------------------------------------------ #

    @staticmethod
    def _clean(req, file_id=None):
        """Validated copy of a requirement."""
        req_id = str(req.get("id") or "").strip()
        if not req_id:
            raise ValueError("exigence sans ID")
        releases = req.get("releases") or []
        # The page may send the releases as one text "A, B" or as a list.
        if isinstance(releases, str):
            releases = releases.split(",")
        return {
            "id": req_id,
            "title": str(req.get("title") or "").strip(),
            "releases": list(dict.fromkeys(str(r).strip().upper() for r in releases if str(r).strip())),
            "stdhl": str(req.get("stdhl") or "").strip(),
            # Validation in the release of the base (see validation_state).
            "validation": req.get("validation") if req.get("validation") in VALIDATION_STATES else "out_of_scope",
            **{f: str(req.get(f) or "").strip() for f, _, _ in REFERENCE},
            "file_id": file_id if file_id is not None else (int(req["file_id"]) if req.get("file_id") is not None else None),
        }

    @staticmethod
    def validation_values(value, separator=","):
        """Values of « Validation releases » (text or list): the releases upper-cased, in order, then QUALIFIE when
        written (whatever its case and accent)."""
        if isinstance(value, str):
            value = re.split(separator, value)
        values = [str(v).strip().upper() for v in value or () if str(v).strip()]
        qualified = any(fold(v) == fold(QUALIFIED) for v in values)
        return list(dict.fromkeys(v for v in values if fold(v) != fold(QUALIFIED))) + ([QUALIFIED] if qualified else [])

    @staticmethod
    def validation_state(values, release):
        """Validation of a requirement in `release` from its « Validation releases »: qualified (QUALIFIE written,
        whatever the releases), to_qualify (the release is in the list) or out_of_scope."""
        if QUALIFIED in values:
            return "qualified"
        return "to_qualify" if release and release.upper() in values else "out_of_scope"

    @staticmethod
    def _stdhl_id(db, name):
        """Id of a value of the STDHL column, added to the table list_stdhl when new; None for an empty cell."""
        if not name:
            return None
        row = db.execute("SELECT id FROM list_stdhl WHERE name = ?", (name,)).fetchone()
        return row[0] if row else db.execute("INSERT INTO list_stdhl (name) VALUES (?)", (name,)).lastrowid

    @staticmethod
    def _write_tome(db, tome, reqs):
        """Replace the requirements of a tome (inside a transaction)."""
        tome_id = db.execute("SELECT id FROM list_tomes WHERE name = ?", (tome,)).fetchone()[0]
        # Deleted then written again: the lines of the plans, STDHL and fiches that pointed to them are removed by the
        # cascade and rebuilt by the server, which reads these documents again after an import.
        db.execute("DELETE FROM requirements WHERE tome_id = ?", (tome_id,))
        for r in reqs:
            db.execute(
                "INSERT INTO requirements (req_id, file_id, stdhl_id, tome_id, title, iadt_id, moc_id, delegation_id, "
                "validation) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (r["id"], r["file_id"], RequirementsDatabase._stdhl_id(db, r["stdhl"]), tome_id, r["title"],
                 *RequirementsDatabase._reference_ids(db, r), r["validation"]))

    @staticmethod
    def _reference_ids(db, req):
        """(iadt_id, moc_id, delegation_id) of the reference of a requirement (None for an empty or unknown value)."""
        ids = []
        for field, table, column in REFERENCE:
            row = db.execute(f"SELECT id FROM {table} WHERE {column} = ? COLLATE NOCASE", (req.get(field) or "",)).fetchone()
            ids.append(row[0] if row else None)
        return ids

    def _requirements(self, where="", params=()):
        """Requirements (with their tome), in tome then file order."""
        # Every requirement gets the release of the base: it is not stored, the whole base works on one release.
        rows = self.sql.query(
            f"""SELECT t.name AS tome, r.id, r.req_id, r.title, s.name AS stdhl, r.file_id, r.manual, r.deleted,
                       i.code AS iadt, m.code AS moc, p.name AS delegation, r.validation
                FROM requirements r JOIN list_tomes t ON t.id = r.tome_id LEFT JOIN list_stdhl s ON s.id = r.stdhl_id
                LEFT JOIN list_iadt i ON i.id = r.iadt_id LEFT JOIN list_moc m ON m.id = r.moc_id
                LEFT JOIN list_plateformes p ON p.id = r.delegation_id
                {where} ORDER BY t.id, r.id""", params)

        return [{"id": r["req_id"], "title": r["title"], "releases": [self.sql.release] if self.sql.release else [],
                 "stdhl": r["stdhl"] or "", "iadt": r["iadt"] or "", "moc": r["moc"] or "",
                 "validation": r["validation"],
                 "delegation": r["delegation"] or "",
                 "file_id": r["file_id"], "manual": r["manual"], "deleted": bool(r["deleted"]),
                 "tome": r["tome"]} for r in rows]

    # ------------------------------------------------------------------ #
    # Access
    # ------------------------------------------------------------------ #

    def tomes(self):
        """Names of the tomes that have requirements, in the order of the referential."""
        return [r["name"] for r in self.sql.query(
            "SELECT name FROM list_tomes t WHERE EXISTS (SELECT 1 FROM requirements r WHERE r.tome_id = t.id) ORDER BY t.id")]

    def get(self, tome):
        """A tome: {"description", "files", "requirements"}, or None if it has no requirements."""
        found = self.to_dict(tome)["tomes"]
        return next(iter(found.values()), None)

    def to_dict(self, only=None):
        """The requirements by tome (tomes without requirements are left out)."""
        result = {}
        where, params = ("WHERE t.name = ?", (only,)) if only is not None else ("", ())
        for r in self._requirements(where, params):
            tome = r.pop("tome")
            result.setdefault(tome, []).append(r)
        tomes = {}
        # For each tome: its description (Configuration > Tomes), the files imported into it and its requirements.
        for name, reqs in result.items():
            tomes[name] = {
                "description": self.sql.scalar("SELECT description FROM list_tomes WHERE name = ?", (name,)) or "",
                "files": [f["id"] for f in self.sql.query(
                    "SELECT f.id FROM files f JOIN list_tomes t ON t.id = f.tome_id "
                    "WHERE f.source = 'requirements' AND t.name = ? ORDER BY f.id", (name,))],
                "requirements": reqs,
            }
        # "stdhl_values": every value of the table list_stdhl, the ones a requirement added or changed by hand can take.
        return {"tomes": tomes, "stdhl_values": self.stdhl_names()}

    def distribution(self, colors=None):
        """Share of the requirements (not deleted) in each tome, for the cumulative bar of the page Exigences:
        [{tome, count, percent, color}] in the order of the tomes; `colors`: {tome: colour} (configuration)."""
        # Percentages with one decimal, so that small tomes still show a share.
        rows = self.sql.query("SELECT t.name AS tome, COUNT(q.id) AS count FROM list_tomes t "
                              "JOIN requirements q ON q.tome_id = t.id AND q.deleted = 0 GROUP BY t.id ORDER BY t.id")
        total = sum(r["count"] for r in rows)
        return [{**r, "percent": round(100 * r["count"] / total, 1) if total else 0,
                 "color": (colors or {}).get(r["tome"], "")} for r in rows]

    def stdhl_names(self):
        """Values of the table list_stdhl (column STDHL of the imported files), sorted."""
        return sorted((r["name"] for r in self.sql.query("SELECT name FROM list_stdhl")), key=str.casefold)

    def requirements(self, tomes=None):
        """Requirements of the given tomes (all of them by default), each with its "tome" key."""
        return [r for r in self._requirements() if tomes is None or r["tome"] in tomes]

    # ------------------------------------------------------------------ #
    # Reading a requirements file
    # ------------------------------------------------------------------ #

    @staticmethod
    def _cell_text(value):
        """Text of an Excel cell (a whole number is written without decimals)."""
        if value is None:
            return ""
        # Excel stores every number as a float: 12 would read "12.0".
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        return str(value).translate(NORMALIZATION)

    @staticmethod
    def _open(content):
        """Open a workbook from its bytes. ValueError when it is not a readable Excel file."""
        try:
            return openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except Exception as e:   # not an .xlsx / .xlsm workbook (an old .xls for instance)
            raise ValueError(f"classeur Excel illisible (formats .xlsx et .xlsm) : {e}") from None

    @classmethod
    def sheet_names(cls, content):
        """Names of the sheets of a workbook."""
        workbook = cls._open(content)
        try:
            return workbook.sheetnames
        finally:
            workbook.close()

    @classmethod
    def read_excel(cls, content, sheet=None, config=None, rows=None):
        """Read a requirements workbook (bytes). Returns {"sheets": names, "sheet": sheet read,
        "requirements": [{id, title, releases, stdhl, iadt, moc, delegation, row}], "releases": releases named in the file,
        "repeated": [{id, rows: [{row, title, releases, stdhl}], chosen}] numbers found on several rows}.

        Row 1 is the header: it must contain every expected text (configuration headers_requirements, text ->
        role); the columns of the roles id, title, releases, stdhl, iadt, moc, delegation are located by it. Otherwise,
        the file is refused (ValueError). The reference of the rows imported is checked after the choice of the
        releases (see check_reference, in analyse and import_file).
        Releases are upper-cased, their separator is in the configuration. A number found on several rows keeps the
        row chosen by the user, `rows` = {number: row number}; until it is chosen, its first row stands in and its
        "chosen" is None. The first sheet is read when `sheet` is not given.
        """
        get = config.get if config is not None else DEFAULT_CONFIG.get
        headers = get("headers_requirements")
        workbook = cls._open(content)
        try:
            names = workbook.sheetnames
            if sheet and sheet not in names:
                raise ValueError(f"feuille « {sheet} » introuvable : {', '.join(names)}")
            name = sheet or names[0]
            # Row 1 is the header: the columns are found by their text (headers_requirements), not by their position.
            first = next(workbook[name].iter_rows(max_row=1, values_only=True), ())
            cells = [cls._cell_text(v) for v in first]
            cols, missing = match_header(cells, headers)
            if missing:
                # One message: the header is not the expected one (the columns are found by their name, anywhere).
                raise ValueError(f"en-tête non conforme à l'attendu (ligne 1 de la feuille « {name} » : colonnes "
                                 + ", ".join(f"« {h} »" for h in headers if headers[h] not in ("check", "ignore"))
                                 + ")")
            found, releases = {}, {}
            # Every row after the header; `number` is the row number shown to the user (Excel numbering).
            for number, row in enumerate(workbook[name].iter_rows(min_row=2, values_only=True), start=2):
                value = lambda field: cls._cell_text(row[cols[field]] if cols[field] < len(row) else None)
                req_id, title, rels, validation, stdhl, iadt, moc, delegation = (
                    value(f) for f in ("id", "title", "releases", "validation", "stdhl", "iadt", "moc", "delegation"))
                rels = [r.strip().upper() for r in re.split(get("release_separator_regex"), rels) if r.strip()]
                releases.update(dict.fromkeys(rels))
                req_id = req_id.strip()
                # A STDHL cell may hold several values; one value written twice in different cases counts once.
                values = cls.stdhl_values(stdhl, config)
                stdhl = values[0] if len(values) == 1 else stdhl   # « F1, f1 » is one value
                if req_id:
                    found.setdefault(req_id, []).append(
                        {"id": req_id, "title": title.strip(), "releases": list(dict.fromkeys(rels)),
                         "stdhl": stdhl.strip(), "stdhl_values": values, "iadt": iadt.strip(), "moc": moc.strip(),
                         "delegation": delegation.strip(), "row": number,
                         # « Validation releases » as written; the field validation is set by check_reference.
                         "validation_releases": cls.validation_values(validation, get("release_separator_regex"))})
            # A number found on several rows: the row chosen by the user (`rows`), else the first row for now, with
            # "chosen" None so that the page asks.
            rows = {str(k): int(v) for k, v in (rows or {}).items()}
            requirements, repeated = [], []
            for req_id, versions in found.items():
                chosen = next((v for v in versions if v["row"] == rows.get(req_id)), None)
                requirements.append(chosen or versions[0])
                if len(versions) > 1:
                    repeated.append({"id": req_id, "chosen": chosen and chosen["row"],
                                     "rows": [{f: v[f] for f in ("row", "title", "releases", "stdhl")} for v in versions]})
            return {"sheets": names, "sheet": name, "requirements": requirements, "releases": list(releases),
                    "repeated": repeated}
        finally:
            workbook.close()

    @staticmethod
    def check_reference(requirements, config=None):
        """The reference of each requirement read: one IADT code, one MOC code and one delegation, values of the
        configuration (written as there), the delegation required; its field validation set for the release of the
        base from its « Validation releases » (never an error). ValueError listing the rows otherwise. A
        delegation to a platform absent from the configuration is not an error: returns those new platforms,
        {name: [numbers]}."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        known = {"iadt": list(get("iadt") or {}), "moc": list(get("mocs") or {}), "delegation": list(get("platforms") or [])}
        names = {"iadt": "IADT", "moc": "MOC", "delegation": "délégation"}
        separator = get("code_separator_regex") or r"[,;+\s]+"
        problems, new = [], {}
        for r in requirements:
            # « Validation releases » never refuses the file: it only gives the validation in the release of the base.
            r["validation"] = RequirementsDatabase.validation_state(r.get("validation_releases", []), get("release"))
            for field in ("iadt", "moc", "delegation"):
                value = r[field]
                if not value:
                    if field == "delegation":
                        problems.append(f"ligne {r['row']} ({r['id']}) : délégation vide")
                    continue
                match = next((k for k in known[field] if fold(k) == fold(value)), None)
                if match:
                    r[field] = match
                elif field == "delegation" and not re.search(r"[,;/+\n]", value):
                    # A new platform: added to the referential at import (one name whatever its case).
                    key = next((k for k in new if fold(k) == fold(value)), value)
                    new.setdefault(key, []).append(r["id"])
                    r[field] = key
                elif len([v for v in re.split(separator if field != "delegation" else r"[,;/+\n]+", value) if v]) > 1:
                    problems.append(f"ligne {r['row']} ({r['id']}) : plusieurs valeurs {names[field]} « {value} »")
                else:
                    problems.append(f"ligne {r['row']} ({r['id']}) : {names[field]} « {value} » absent de la configuration")
        if problems:
            more = len(problems) - 10
            raise ValueError(f"{plural(len(problems), 'erreur')} dans la référence des exigences (IADT, MOC, délégation) : "
                             + " ; ".join(problems[:10]) + (f" … et {more} autre{'s' if more > 1 else ''}" if more > 0 else "")
                             + " : fichier non importé")
        return new

    @staticmethod
    def filter_releases(requirements, selected, forced=()):
        """(requirements kept, groups outside) for the selected releases (None: all). A release not selected is
        removed from a requirement. A requirement without any selected release is outside: grouped by its releases
        ("" for none), [{"key", "count", "requirements"}]; the groups whose key is in `forced` are kept, with the
        selected releases."""
        if selected is None:
            return requirements, []
        # A requirement with at least one selected release is kept with those releases only; the others are grouped by
        # their exact list of releases ("B, C", "" for none), and a group forced by the user joins the selected
        # releases.
        chosen, kept, outside = [s.strip().upper() for s in selected], [], {}
        forced = set(forced or ())
        for r in requirements:
            releases = [x for x in r["releases"] if x in chosen]
            if releases:
                kept.append({**r, "releases": releases})
                continue
            key = ", ".join(r["releases"])
            outside.setdefault(key, []).append(r)
            if key in forced:
                kept.append({**r, "releases": chosen})
        groups = [{"key": key, "count": len(rs), "requirements": rs}
                  for key, rs in sorted(outside.items(), key=lambda g: (g[0] == "", g[0]))]
        return kept, groups

    def diff(self, tome, requirements, mode):
        """Differences between the requirements of a tome (the ones deleted by hand left out) and those of a
        file: {"added": [requirement], "updated": [{"id", "changes": [{"field", "before", "after"}]}],
        "removed": [requirement] (replacement only)}."""
        current = self.get(tome)
        # Compared field by field (releases as a text); in a complement, nothing is removed.
        old = {r["id"]: r for r in (current["requirements"] if current else []) if not r["deleted"]}
        value = lambda r, f: ", ".join(r[f]) if isinstance(r[f], list) else r[f]
        added, updated, fresh = [], [], {}
        for r in requirements:
            fresh.setdefault(r["id"], r)
        for req_id, r in fresh.items():
            if req_id not in old:
                added.append(r)
                continue
            changes = [{"field": f, "before": value(old[req_id], f), "after": value(r, f)}
                       for f in DIFF_FIELDS if value(old[req_id], f) != value(r, f)]
            if changes:
                updated.append({"id": req_id, "changes": changes})
        removed = [r for i, r in old.items() if i not in fresh] if mode == "replace" else []
        return {"added": added, "updated": updated, "removed": removed}

    @staticmethod
    def stdhl_values(value, config=None):
        """Values of a STDHL cell, without duplicates: a requirement may only have one."""
        separator = (config.get if config is not None else DEFAULT_CONFIG.get)("stdhl_separator_regex")
        # Read backwards so that, for values equal but for the case, the first one written wins.
        values = {v.strip().casefold(): v.strip() for v in reversed(re.split(separator, value or "")) if v.strip()}
        return list(reversed(values.values()))

    def conflicts(self, requirements, tome, mode):
        """What the user must settle before importing: {"duplicates": requirements of the file already in the base,
        whatever the tome, except the ones a replacement of their tome overwrites; "multi_stdhl": requirements of the
        file with several STDHL values}."""
        # A number already in the base is a conflict, unless the import replaces the very tome that holds it.
        existing = {r["id"]: r for r in self._requirements()}
        names = {r["id"]: r["name"] for r in self.sql.query("SELECT id, name FROM files")}
        fields = DIFF_FIELDS
        duplicates = []
        for r in requirements:
            e = existing.get(r["id"])
            if e and not (mode == "replace" and e["tome"].casefold() == (tome or "").casefold()):
                duplicates.append({"id": r["id"], "same": all(e[f] == r[f] for f in fields),
                                   "existing": {**{f: e[f] for f in fields}, "tome": e["tome"], "file": names.get(e["file_id"], "")},
                                   "new": {f: r[f] for f in fields}})
        multi = [{"id": r["id"], "title": r["title"], "values": r["stdhl_values"]} for r in requirements
                 if len(r["stdhl_values"]) > 1]
        return {"duplicates": duplicates, "multi_stdhl": multi}

    @staticmethod
    def _resolve(requirements, found, choices, strict):
        """(requirements to import, numbers taken from another tome) with the choices of the user: {"row": {number:
        row of the file}, "keep": {number: "existing" or "new"}, "stdhl": {number: value}}; the rows are applied when
        reading (read_excel). With `strict`, ValueError when something is not settled."""
        # Count what is still unsettled: duplicates without a choice, STDHL cells without a value chosen, repeated
        # numbers without a row.
        keep, stdhl = (choices or {}).get("keep") or {}, (choices or {}).get("stdhl") or {}
        duplicates = [d for d in found["duplicates"] if keep.get(d["id"]) in ("existing", "new")]
        settled = [m for m in found["multi_stdhl"] if stdhl.get(m["id"]) in m["values"]]
        repeated = sum(r["chosen"] is None for r in found.get("repeated", ()))
        left = len(found["duplicates"]) - len(duplicates) + len(found["multi_stdhl"]) - len(settled) + repeated
        if strict and left:
            raise ValueError(f"{plural(left, 'conflit')} à régler avant l'import : choisir la ligne d'une exigence "
                             "présente plusieurs fois dans le fichier, l'exigence à conserver ou la valeur STDHL")
        # "existing": the requirement of the base stays, the row of the file is dropped; "new": the row of the file is
        # imported and the requirement leaves its other tome (`moved`).
        kept = {d["id"] for d in duplicates if keep[d["id"]] == "existing"}
        values = {m["id"]: stdhl[m["id"]] for m in settled}
        result = [{**r, "stdhl": values.get(r["id"], r["stdhl"])} for r in requirements if r["id"] not in kept]
        return result, [d["id"] for d in duplicates if keep[d["id"]] == "new"]

    def analyse(self, content, sheet=None, tome=None, mode="replace", config=None, releases=None, choices=None):
        """What importing a requirements file would do, for the approval windows: sheets, releases found (with
        their number of requirements), requirements kept for the selected `releases` (None: all), groups of
        requirements outside them ("outside", see filter_releases; imported only when forced), conflicts to settle
        and, into a tome that already has requirements, the differences once the `choices` are applied.
        Without a sheet, a workbook of several sheets only gives its sheet names (the sheet is chosen first)."""
        # First call, without a sheet: a workbook of several sheets only returns its sheet names, the page asks which
        # one to read.
        if sheet is None:
            names = self.sheet_names(content)
            if len(names) > 1:
                return {"sheets": names, "sheet": None, "count": None}
        data = self.read_excel(content, sheet, config, (choices or {}).get("row"))
        # Number of requirements per release named in the file, before filtering.
        counts = {x: sum(x in r["releases"] for r in data["requirements"]) for x in data["releases"]}
        reqs, outside = self.filter_releases(data["requirements"], releases, (choices or {}).get("force"))
        # The rows imported: a wrong reference refuses the file; a new platform is added at import.
        new_platforms = self.check_reference(reqs, config)
        data = {**data, "requirements": reqs, "release_counts": counts, "outside": outside,
                "new_platforms": new_platforms}
        found = self.conflicts(reqs, tome, mode) if tome else {"duplicates": [], "multi_stdhl": []}
        found["repeated"] = data["repeated"]
        resolved, _ = self._resolve(reqs, found, choices, False)
        current = self.get(tome) if tome else None
        return {**data, **found, "count": len(reqs),
                "tome_exists": current is not None,
                "tome_count": len([r for r in current["requirements"] if not r["deleted"]]) if current else 0,
                "diff": self.diff(tome, resolved, mode) if current else None}

    def tome_from_name(self, name, config):
        """Tome named in the name of a requirements file (configuration requirements_tome_regex): an existing tome
        (whatever the case), or the name found for a new tome; "" when the name does not match."""
        # The tome named in the file name, written as the existing tome if one matches whatever the case.
        match = re.search(config.get("requirements_tome_regex"), name or "")
        found = (match.group(match.lastindex or 0) if match else "").strip()
        if not found:
            return ""
        return next((r["name"] for r in self.sql.query("SELECT name FROM list_tomes") if fold(r["name"]) == fold(found)), found)

    def import_file(self, files, content, name, tome, mode="replace", sheet=None, config=None, releases=None,
                    choices=None):
        """Import a requirements file into a tome: the file is stored (FileDatabase `files`), then its
        requirements; if the import fails, the stored file is removed again. A replacement removes the files
        previously imported into the tome. Requirements outside the selected `releases` are imported only when their
        group is forced (choices "force"). Conflicts must be settled by `choices` (see _resolve), otherwise
        ValueError. Returns the summary of import_requirements, plus "removed_files", "ignored" (outside, not
        forced), "forced" (outside, imported into the release), "kept" (already in the base, kept as they were) and
        "releases" (named in the file)."""
        data = self.read_excel(content, sheet, config, (choices or {}).get("row"))
        forced = (choices or {}).get("force") or ()
        selected, outside = self.filter_releases(data["requirements"], releases, forced)
        new_platforms = self.check_reference(selected, config)
        found = {**self.conflicts(selected, tome, mode), "repeated": data["repeated"]}
        reqs, moved = self._resolve(selected, found, choices, True)
        if not reqs:
            raise ValueError("aucune exigence à importer" + (f" dans la feuille « {data['sheet']} »" if sheet else ""))
        # The tome before the import: its files become useless after a replacement.
        previous = self.get(tome)
        record = files.add("requirements", name, content, tome=tome)
        try:
            # New platforms of the delegations: added to the referential first (their colour drawn at reading).
            if new_platforms:
                with self.sql.transaction() as db:
                    self.sql.ensure("platforms", list(new_platforms), db)
            summary = self.import_requirements(tome, reqs, mode, record["id"], moved)
        except Exception:
            files.remove(record["id"])   # nothing imported: the file is not kept
            raise
        summary["removed_files"] = []
        if previous and summary["mode"] == "replace":
            # A replacement makes the files previously imported into the tome useless.
            for file_id in previous["files"]:
                old = files.get(file_id)
                if file_id != record["id"] and old and files.remove(file_id):
                    summary["removed_files"].append(old["name"])
        forced_count = sum(g["count"] for g in outside if g["key"] in forced)
        return {**summary, "file_id": record["id"], "kept": len(selected) - len(reqs), "forced": forced_count,
                "new_platforms": list(new_platforms),
                "ignored": sum(g["count"] for g in outside) - forced_count, "releases": data["releases"]}

    # ------------------------------------------------------------------ #
    # Storing requirements
    # ------------------------------------------------------------------ #

    def import_requirements(self, tome, requirements, mode="replace", file_id=None, moved=()):
        """Store the requirements read from a file in a tome. Returns a summary of the changes.

        mode "replace": the tome receives exactly these requirements;
        mode "complement": new requirements are added, the ones already in the tome are updated.
        A new tome is created (added to the referential) whatever the mode.
        moved: numbers taken from another tome, where they are removed.
        """
        tome = str(tome or "").strip()
        if not tome:
            raise ValueError("nom de tome manquant")
        if not requirements:
            raise ValueError("aucune exigence à importer")
        if mode not in MODES:
            raise ValueError(f"mode d'import inconnu : {mode} (attendu : {', '.join(MODES)})")
        new = {}
        # Each requirement validated; an ID listed twice keeps its first row.
        for r in requirements:
            r = self._clean(r, file_id)
            new.setdefault(r["id"], r)   # an ID listed twice keeps its first row, as when reading the file
        with self.sql.lock:
            current = self.get(tome)
            old = {r["id"]: {k: r[k] for k in ("id", *DIFF_FIELDS, "file_id")}
                   for r in current["requirements"] if not r["deleted"]} if current else {}
            # Replace: the tome becomes exactly the file. Complement: the requirements of the file replace those of the
            # same ID, the others are added at the end; requirements deleted by hand are dropped either way (the file is
            # the reference).
            if mode == "replace" or current is None:
                merged = list(new.values())
                removed = [i for i in old if i not in new]
            else:
                merged = [new.get(i, r) for i, r in old.items()] + [r for i, r in new.items() if i not in old]
                removed = []
            with self.sql.transaction() as db:
                self.sql.ensure("tomes", [tome], db)
                # Numbers the user took from another tome leave it before the tome is written (one number, one tome).
                db.executemany("DELETE FROM requirements WHERE req_id = ? AND tome_id != (SELECT id FROM list_tomes WHERE name = ?)",
                               [(number, tome) for number in moved])
                self._write_tome(db, tome, merged)
            added = [i for i in new if i not in old]
            # Updated: same ID, different content (the file it comes from does not count).
            updated = [i for i in new if i in old and old[i] != {**new[i], "file_id": old[i]["file_id"]}]
            return {"tome": tome, "mode": "replace" if current is None else mode, "created": current is None,
                    "added": len(added), "updated": len(updated), "removed": len(removed), "total": len(merged)}

    def edit(self, tome, action, old_id=None, requirement=None, config=None):
        """Manual change of one requirement of a tome (it is overwritten by the next import of the tome).

        action "add": new requirement (file_id None: « saisie manuelle »), at the end of the tome, marked "added";
        action "update": requirement old_id replaced by `requirement` (its ID may change, its file is kept),
                         marked "modified" (or stays "added");
        action "delete": requirement old_id marked deleted (struck through, not checked) — removed for good when
                         it was added by hand;
        action "restore": deleted mark removed.
        Returns the requirement stored (None for a deletion). Raises ValueError when the change is refused.
        """
        if action not in ("add", "update", "delete", "restore"):
            raise ValueError(f"action inconnue : {action}")
        with self.sql.lock:
            tome_id = self.sql.tome_id(tome)
            if tome_id is None:
                raise ValueError(f"tome inconnu : « {tome} »")
            existing = None
            if action != "add":
                existing = self.sql.query("SELECT id, file_id, manual, deleted FROM requirements "
                                          "WHERE tome_id = ? AND req_id = ?", (tome_id, old_id))
                if not existing:
                    raise ValueError(f"exigence introuvable dans le tome « {tome} » : {old_id}")
                existing = existing[0]
            if action == "delete":
                with self.sql.transaction() as db:
                    # A requirement added by hand disappears; one coming from a file is only marked deleted, and can be
                    # restored.
                    if existing["manual"] == "added":
                        db.execute("DELETE FROM requirements WHERE id = ?", (existing["id"],))
                    else:
                        db.execute("UPDATE requirements SET deleted = 1 WHERE id = ?", (existing["id"],))
                return None
            if action == "restore":
                with self.sql.transaction() as db:
                    db.execute("UPDATE requirements SET deleted = 0 WHERE id = ?", (existing["id"],))
                return None
            if existing and existing["deleted"]:
                raise ValueError(f"l'exigence {old_id} est supprimée : la rétablir avant de la modifier")
            req = self._clean(requirement or {}, existing["file_id"] if existing else None)
            req["releases"] = [self.sql.release] if self.sql.release else []   # the release of the base
            if existing is not None:
                # A field of the reference not sent by the page keeps its value.
                before = self._requirements("WHERE r.id = ?", (existing["id"],))[0]
                for field in [f for f, _, _ in REFERENCE] + ["validation"]:
                    if field not in (requirement or {}) or (field == "validation"
                                                            and requirement[field] not in VALIDATION_STATES):
                        req[field] = before[field]
            self.check_reference([{**req, "row": "saisie"}], config)
            if existing is None:
                req["file_id"] = None
            else:
                # Saved without any change (releases in any order): nothing written, the requirement keeps its mark.
                # Saved without any change: nothing is written, so that an imported requirement does not become
                # "modified".
                current = self._requirements("WHERE r.id = ?", (existing["id"],))[0]
                if (current["id"], current["title"], current["stdhl"], sorted(current["releases"]),
                        current["iadt"], current["moc"], current["delegation"], current["validation"]) == \
                        (req["id"], req["title"], req["stdhl"], sorted(req["releases"]),
                         req["iadt"], req["moc"], req["delegation"], req["validation"]):
                    return current
            if not req["releases"]:
                raise ValueError("une exigence doit s'appliquer à au moins une release")
            if len(self.stdhl_values(req["stdhl"], config)) > 1:
                raise ValueError("une exigence ne peut avoir qu'une seule valeur STDHL")
            # Releases and STDHL are chosen among the existing values, not typed.
            declared = set((config.get if config is not None else DEFAULT_CONFIG.get)("releases") or {})
            unknown = [r for r in req["releases"] if r not in declared]
            if unknown:
                raise ValueError(f"{'releases inconnues' if len(unknown) > 1 else 'release inconnue'} : {', '.join(unknown)} "
                                 "(choisir parmi les releases de la configuration)")
            if req["stdhl"] and req["stdhl"] not in self.stdhl_names():
                raise ValueError(f"valeur STDHL inconnue : « {req['stdhl']} » (choisir parmi les valeurs existantes)")
            # A new or renamed ID must not exist anywhere in the base (one number, one requirement).
            if action == "add" or req["id"] != old_id:
                other = self.sql.scalar("SELECT t.name FROM requirements r JOIN list_tomes t ON t.id = r.tome_id WHERE r.req_id = ?",
                                        (req["id"],))
                if other is not None:
                    raise ValueError(f"l'exigence {req['id']} existe déjà dans le tome « {other} »")
            with self.sql.transaction() as db:
                stdhl_id = self._stdhl_id(db, req["stdhl"])
                reference = self._reference_ids(db, req)
                if action == "add":
                    db.execute(
                        "INSERT INTO requirements (req_id, stdhl_id, tome_id, title, iadt_id, moc_id, delegation_id, "
                        "validation, manual) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'added')",
                        (req["id"], stdhl_id, tome_id, req["title"], *reference, req["validation"]))
                else:
                    db.execute("UPDATE requirements SET req_id = ?, stdhl_id = ?, title = ?, iadt_id = ?, moc_id = ?, "
                               "delegation_id = ?, validation = ?, manual = ? WHERE id = ?",
                               (req["id"], stdhl_id, req["title"], *reference, req["validation"],
                                "added" if existing["manual"] == "added" else "modified", existing["id"]))
            return req

    def linked_to_file(self, file_id):
        """Number of requirements coming from a file, per tome: {tome: count} (only tomes > 0)."""
        return {r["tome"]: r["n"] for r in self.sql.query(
            "SELECT t.name AS tome, COUNT(*) AS n FROM requirements r JOIN list_tomes t ON t.id = r.tome_id "
            "WHERE r.file_id = ? GROUP BY t.id", (file_id,))}

    def remove_file(self, file_id):
        """Remove the requirements coming from a file. Returns {tome: number of requirements removed}."""
        return self.prune_files(lambda f: f != file_id)

    def prune_files(self, keep):
        """Remove the requirements whose file_id does not satisfy keep(file_id) (requirements without
        file_id are kept). Returns {tome: number removed}."""
        with self.sql.lock:
            # Count per tome before deleting, for the messages.
            gone = [f["file_id"] for f in self.sql.query(
                "SELECT DISTINCT file_id FROM requirements WHERE file_id IS NOT NULL") if not keep(f["file_id"])]
            removed = {}
            for file_id in gone:
                for tome, n in self.linked_to_file(file_id).items():
                    removed[tome] = removed.get(tome, 0) + n
            if gone:
                with self.sql.transaction() as db:
                    db.executemany("DELETE FROM requirements WHERE file_id = ?", [(f,) for f in gone])
            return removed

    def remove_tome(self, tome):
        """Remove the requirements of a tome (the tome stays in the referential). Returns False if none."""
        with self.sql.transaction() as db:
            return db.execute("DELETE FROM requirements WHERE tome_id = ?", (self.sql.tome_id(tome),)).rowcount > 0
