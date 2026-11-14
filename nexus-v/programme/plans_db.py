"""Acceptance plans of NEXUS-V, in the SQL database (tables files and plan_acceptation).

When an acceptance plan is loaded, its useful content (read in the Word document) is stored here. The plan is
identified by the number of its file (table files): its name and tome are those of the file; it applies to the
release of the base; removing the file removes its content (cascade).

Everything refers to the requirement: a plan row is a requirement of the tome of the plan, with its IADT and MOC codes
and its delegation (several rows of the document for one requirement are merged). The title and the column Méthode
are not stored. At import, a number unknown in the tome, a value empty or unknown in the configuration, a MOC outside
the IADT / MOC rules or several delegations are settled by the user: expected value chosen, or row ignored (an
unknown number can only be ignored).

The class gathers everything about plans: reading the Word document (read_docx: tables whose header matches the
setting headers_plan), rules of a row (check_row: IADT, MOC, delegation), check of a tome (check_tome: every
requirement of the tome must be in a plan of that tome), loading a file (analyse then load_file) and storing.

For the page, plans are given by file number (to_dict):

    {"plans": {"12": {"file_id": 12, "name": "...docx", "tome": "fct1", "releases": ["A"],
                      "rows": [{"key": ID, "requirements": [ID], "iadt", "moc", "delegation", "manual",
                                "deleted"}, ...]}}}

A row corrected at import, added, changed or deleted by hand (edit) is marked (manual = added, modified, deleted;
deleted = 1: struck through, no effect). When the plan is read again from its document, only the unmarked rows are
rewritten: a requirement with a marked row keeps it.
"""
import re

from configuration import DEFAULT_CONFIG, fold, header_column, header_text, match_header, plural, table_left_out
from file_database import FileDatabase, WordDocument
from nexus_db import value_id

# Columns of a plan table (their expected header texts are in the configuration: headers_plan).
PLAN_COLUMNS = ("exigence", "iadt", "moc", "delegation")


def natural_key(text):
    """Sort key putting VS2 before VS10."""
    return [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", str(text))]


def unique(values):
    """Non-empty values without duplicates, in their order."""
    return list(dict.fromkeys(v for v in values if v))


class PlansDatabase:
    """Content of the acceptance plans, in the SQL database."""

    def __init__(self, sql):
        """sql: the shared NexusDatabase."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def _write(db, file_id, entries):
        """Replace the rows of a plan read in its document (inside a transaction), from its entries (see entries).
        Rows marked manual stay, and the requirements that have one are not written again; a number that is not a
        requirement of the tome of the plan has no row, nor a requirement active in another plan of the tome that
        (a requirement is in one plan only of its tome: the other one was kept)."""
        # Rows read from the document are rewritten; rows marked by a correction or a manual change stay.
        db.execute("DELETE FROM plan_acceptation WHERE file_id = ? AND manual IS NULL", (file_id,))
        # Requirements not to write: those that still have a marked row in this plan, and those active in another plan
        # of the same tome (a requirement is in one plan only; a plan of the same name is the one being replaced).
        kept = {r[0] for r in db.execute("SELECT DISTINCT requirement_id FROM plan_acceptation WHERE file_id = ?", (file_id,))}
        kept |= {r[0] for r in db.execute(
            """SELECT DISTINCT a.requirement_id FROM plan_acceptation a
               JOIN files f ON f.id = a.file_id AND f.tome_id = (SELECT tome_id FROM files WHERE id = ?)
               WHERE a.file_id != ? AND a.deleted = 0
                 AND f.name != (SELECT name FROM files WHERE id = ?)   -- a plan of the same name is being replaced""",
            (file_id, file_id, file_id))}
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        known = FileDatabase.tome_requirements(db, tome_id)
        for e in entries:
            if e["requirement"] in known and known[e["requirement"]] not in kept:
                PlansDatabase._insert(db, file_id, known[e["requirement"]], e["iadt"], e["moc"], e["delegation"])

    @staticmethod
    def _insert(db, file_id, requirement_id, iadt, moc, delegation, manual=None, deleted=0):
        """Rows of one requirement of a plan: one for each IADT and MOC code (inside a transaction)."""
        delegation = (delegation or "").strip()
        platforms = {fold(r[1]): r[0] for r in db.execute("SELECT id, name FROM list_plateformes")}
        # The platform as written in the configuration whatever the case; a name unknown of it is only possible from the
        # command line (the page settles them first).
        delegation_id = (platforms.get(fold(delegation)) or value_id(db, "list_plateformes", "name", delegation)) if delegation else None
        iadt_ids = [value_id(db, "list_iadt", "code", c) for c in PlansDatabase.split_codes(iadt)] or [None]
        moc_ids = [value_id(db, "list_moc", "code", c) for c in PlansDatabase.split_codes(moc)] or [None]
        # One row per pair of codes: a cell "A" with MOC "MOC2, MOC3" gives two rows of the same requirement.
        db.executemany("INSERT INTO plan_acceptation (requirement_id, iadt_id, moc_id, delegation_id, file_id, manual, "
                       "deleted) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       [(requirement_id, i, m, delegation_id, file_id, manual, deleted) for i in iadt_ids for m in moc_ids])

    @classmethod
    def entries(cls, rows):
        """The rows read in a plan document by requirement: one entry per requirement number, with its IADT and MOC
        codes and its delegations joined (several rows of a requirement are merged), and the places of its rows."""
        found = {}
        for r in rows:
            for number in r["requirements"]:
                # A requirement cited on several rows (or a row citing several requirements): its codes and delegations
                # are merged, and its places (table, row) kept for the messages.
                e = found.setdefault(number, {"requirement": number, "iadt": [], "moc": [], "delegation": [], "places": []})
                e["iadt"] += cls.split_codes(r["iadt"])
                e["moc"] += cls.split_codes(r["moc"])
                if r["delegation"].strip():
                    e["delegation"].append(r["delegation"].strip())
                e["places"].append([r["table"], r["row"]])
        return [{**e, "iadt": ", ".join(unique(e["iadt"])), "moc": ", ".join(unique(e["moc"])),
                 "delegation": ", ".join(unique(e["delegation"]))} for e in found.values()]

    # ------------------------------------------------------------------ #
    # Access
    # ------------------------------------------------------------------ #

    def plans(self, tome=None, file_id=None, config=None, with_deleted=False):
        """Plans with their rows, optionally those of one tome or one file: one row per requirement, its IADT and MOC
        codes joined by ', '. Rows deleted (by hand or ignored at import) are left out unless `with_deleted` (page).
        With a configuration, each row also gets the result of check_row (codes, delegation, errors)."""
        conditions, params = [], []
        if tome is not None:
            conditions.append("tm.name = ?")
            params.append(tome)
        if file_id is not None:
            conditions.append("f.id = ?")
            params.append(file_id)
        heads = self.sql.query(
            f"""SELECT f.id AS file_id, f.name, tm.name AS tome FROM files f
                LEFT JOIN list_tomes tm ON tm.id = f.tome_id
                WHERE {" AND ".join(["f.source = 'plan'", *conditions])} ORDER BY f.id""", params)
        file_ids = [h["file_id"] for h in heads]
        if not file_ids:
            return []
        marks = ",".join("?" * len(file_ids))
        rows, groups, releases = {}, {}, self.sql.file_releases()
        for r in self.sql.query(
                f"""SELECT a.file_id, a.manual, a.deleted, q.req_id, i.code AS iadt, m.code AS moc, p.name AS delegation
                    FROM plan_acceptation a JOIN requirements q ON q.id = a.requirement_id
                    LEFT JOIN list_iadt i ON i.id = a.iadt_id LEFT JOIN list_moc m ON m.id = a.moc_id
                    LEFT JOIN list_plateformes p ON p.id = a.delegation_id
                    WHERE a.file_id IN ({marks}) {"" if with_deleted else "AND a.deleted = 0"} ORDER BY a.id""", file_ids):
            # The stored rows of one requirement in one plan (one per pair of codes) are grouped back into one row,
            # codes joined by ", ".
            g = groups.get((r["file_id"], r["req_id"]))
            if g is None:
                g = groups[(r["file_id"], r["req_id"])] = {
                    "key": r["req_id"], "requirements": [r["req_id"]], "iadt": [], "moc": [],
                    "delegation": r["delegation"] or "", "manual": r["manual"] or "", "deleted": bool(r["deleted"])}
                rows.setdefault(r["file_id"], []).append(g)
            g["iadt"].append(r["iadt"])
            g["moc"].append(r["moc"])
        for g in groups.values():
            g["iadt"], g["moc"] = ", ".join(unique(g["iadt"])), ", ".join(unique(g["moc"]))
        # With a configuration each row also gets its actual delegation (the default platform for an empty cell) and
        # its errors.
        check = (lambda row: self.check_row(row, config)) if config is not None else (lambda row: row)
        return [{"file_id": h["file_id"], "name": h["name"], "tome": h["tome"], "releases": releases.get(h["file_id"], []),
                 "rows": [check(r) for r in rows.get(h["file_id"], [])]} for h in heads]

    def get(self, file_id):
        """The plan of a file with its rows, or None."""
        found = self.plans(file_id=file_id)
        return found[0] if found else None

    def to_dict(self, config=None):
        """The plans by file number, for the page; rows are checked when a configuration is given."""
        plans = self.plans(config=config, with_deleted=True)
        return {"plans": {str(p["file_id"]): p for p in plans}}

    # ------------------------------------------------------------------ #
    # Reading a plan document
    # ------------------------------------------------------------------ #

    @staticmethod
    def _settings(config):
        """Expected headers and requirement number format, from a configuration (defaults without one)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        return get("headers_plan"), re.compile(get("requirement_regex"))

    @staticmethod
    def _row(table, row, exigence, iadt, moc, delegation, regex, separator):
        """Stored form of a plan row: the requirement numbers of the Exigence cell (not the title), IADT and MOC codes
        (upper case, separated by ", " whatever the document uses) and delegation."""
        # Codes written "A/T", "A ; T" or "a, t" are all stored "A, T".
        codes = lambda cell: ", ".join(c for c in re.split(separator, cell.upper()) if c)
        return {"table": table, "row": row, "requirements": unique(regex.findall(exigence)),
                "iadt": codes(iadt), "moc": codes(moc), "delegation": delegation.strip()}

    @staticmethod
    def _find_header(grid, headers):
        """({role: column index}, [expected texts missing]) of the header of a table, always its first row, or None.
        A plan table is recognised by the header of its exigence column and at least one other stored column
        (configuration headers_plan); every expected header must then be there."""
        exigence = header_text(headers, "exigence")
        others = [header_text(headers, role) for role in PLAN_COLUMNS if role != "exigence"]
        cells = grid[0] if grid else []
        # A table is a plan table when its first row has the requirement column and at least one other stored column;
        # then every expected header must be there, or the table is reported and left out.
        if header_column(cells, exigence) is not None and any(header_column(cells, t) is not None for t in others):
            return match_header(cells, headers)
        return None

    @classmethod
    def read_docx(cls, content, config=None):
        """Read an acceptance plan (.docx bytes). Returns {"rows": [stored rows], "ids": requirement numbers
        cited anywhere in the document, "platforms": platforms named in the delegations,
        "skipped": plan tables left out because their header lacks an expected column}. Such a table does
        not stop the reading. Table and row numbers are 1-based, the header row included."""
        headers, regex = cls._settings(config)
        separator = (config.get if config is not None else DEFAULT_CONFIG.get)("code_separator_regex")
        document = WordDocument(content)
        rows, skipped = [], []
        for t_index, grid in document.tables():
            header = cls._find_header(grid, headers)
            if header is None:
                continue
            cols, missing = header
            if missing:
                skipped.append(table_left_out(t_index, missing))
                continue
            # Every row after the header, numbered as in Word (the header is row 1).
            for r_index, cells in enumerate(grid[1:], 2):
                def cell(field):
                    """Text of the cell of a stored column in this row, empty when the row is too short."""
                    k = cols[field]
                    return cells[k] if k is not None and k < len(cells) else ""
                # Skip empty rows and header rows repeated inside the table.
                if not any(cells) or (header_column([cell("exigence")], header_text(headers, "exigence")) is not None
                                      and header_column([cell("iadt")], header_text(headers, "iadt")) is not None):
                    continue
                rows.append(cls._row(t_index, r_index, *(cell(f) for f in PLAN_COLUMNS), regex, separator))
        platforms = unique([r["delegation"] for r in rows])
        # ids: every number cited anywhere in the document (body, headers, notes), not only in the plan tables.
        return {"rows": rows, "ids": unique(regex.findall(document.text())), "platforms": platforms,
                "skipped": skipped}

    @classmethod
    def _readable(cls, content, config=None):
        """Content of a plan that can be loaded: refused (ValueError) only when no valid plan table is found;
        tables with an incomplete header are skipped."""
        headers, _ = cls._settings(config)
        data = cls.read_docx(content, config)
        if not data["rows"]:
            raise ValueError(f"aucun tableau de plan trouvé (en-tête « {header_text(headers, 'exigence')} » et colonnes "
                             "stockées, voir Configuration > En-têtes des tableaux) : fichier non chargé")
        return data

    # ------------------------------------------------------------------ #
    # Rules
    # ------------------------------------------------------------------ #

    @staticmethod
    def split_codes(value):
        """Codes of a stored IADT or MOC value."""
        return [c for c in (value or "").split(", ") if c]

    @classmethod
    def check_row(cls, row, config):
        """A plan row with its IADT and MOC codes, actual delegation and errors (configuration rules)."""
        errors = []
        iadt = cls.split_codes(row["iadt"])
        if not iadt:
            errors.append("IADT vide")
        # Errors against the configuration.
        errors += [f"IADT inconnu : « {c} »" for c in iadt if c not in config.get("iadt")]
        moc = cls.split_codes(row["moc"])
        if not moc:
            errors.append("MOC vide")
        errors += [f"MOC inconnu : « {c} »" for c in moc if c not in config.get("mocs")]
        # An empty Délégation cell means the default platform.
        delegation = config.delegation(row["delegation"])
        if not config.is_platform(delegation):
            errors.append(f"plateforme de délégation inconnue : « {delegation} »")
        return {**row, "iadt_codes": iadt, "moc_codes": moc, "effective_delegation": delegation, "errors": errors}

    @classmethod
    def rule_issues(cls, row, config):
        """Import rules of a plan row (configuration iadt_moc_rules): for each IADT with a rule, at least one of its
        MOC; no MOC outside the rules of the IADT of the row. Messages of the rules broken."""
        rules = config.get("iadt_moc_rules")
        iadt, moc = cls.split_codes(row["iadt"]), cls.split_codes(row["moc"])
        # Each IADT with a rule needs at least one of its MOC; and when every IADT of the row has a rule, no MOC outside
        # those rules is allowed.
        issues = [f"IADT {i} : {' ou '.join(rules[i])} attendu" for i in iadt
                  if i in rules and not set(moc) & set(rules[i])]
        if iadt and all(i in rules for i in iadt):
            allowed = {m for i in iadt for m in rules[i]}
            extra = [m for m in moc if m not in allowed]
            if extra:
                issues.append(f"MOC en plus de l'attendu : {', '.join(extra)}")
        return issues

    @staticmethod
    def delegation_values(row, config):
        """Platforms named in the delegation cell of a row (separator delegation_separator_regex), without duplicates."""
        values = {}
        for v in re.split(config.get("delegation_separator_regex"), row["delegation"] or ""):
            if v.strip():
                values.setdefault(fold(v), v.strip())
        return list(values.values())

    @staticmethod
    def expected_mocs(iadt_cell, config):
        """MOC values a row with these IADT codes may take, from the rules (iadt_moc_rules): each expected MOC, and all
        of them together when there are several; every MOC of the configuration when an IADT has no rule."""
        rules = config.get("iadt_moc_rules")
        iadt = PlansDatabase.split_codes(iadt_cell)
        if not iadt or any(i not in rules for i in iadt):
            return list(config.get("mocs"))
        allowed = unique(m for i in iadt for m in rules[i])
        # The choices offered: each expected MOC alone, then all of them together ("MOC2, MOC3").
        return allowed + ([", ".join(allowed)] if len(allowed) > 1 else [])

    @classmethod
    def to_confirm(cls, entries, config, known=None, elsewhere=None, releases=None, applies=None):
        """Entries of a plan (see entries) the user must settle at import, each with its key (the requirement number),
        its issues and "choose": the cells to replace by a value of the configuration (iadt, moc, delegation). Nothing
        is imported as it is: an empty or unknown IADT or MOC, a MOC breaking the IADT/MOC rules (the expected MOC is
        proposed, see expected_mocs), several or unknown delegations are replaced, or the entry is ignored. With
        `known` (numbers of the requirements of the tome), a number that is not one of them can only be ignored. With
        `elsewhere` (see PlansDatabase.elsewhere), a requirement already in another plan of the release is settled too:
        the other one kept (this entry is not stored), or this one taken (the row of the other plan is removed). With
        the `releases` of the plan and `applies` ({number: releases of the requirement}), a requirement that applies to
        none of them is an import error: it can only be ignored."""
        wanted = {r.strip().upper() for r in releases or () if r.strip()}
        elsewhere = elsewhere or {}
        found = []
        for r in entries:
            base = {"key": r["requirement"], "requirements": [r["requirement"]], "places": r["places"], "iadt": r["iadt"],
                    "moc": r["moc"], "delegation": r["delegation"]}
            # A number that is not a requirement of the tome can only be ignored.
            if known is not None and r["requirement"] not in known:
                found.append({**base, "choose": [], "unknown": True, "expected_mocs": [], "duplicate": [],
                              "issues": ["exigence inconnue du tome : seule possibilité, ignorer la ligne"]})
                continue
            if wanted and applies is not None and not applies.get(r["requirement"], set()) & wanted:
                own = ", ".join(sorted(applies.get(r["requirement"], set()))) or "aucune"
                found.append({**base, "choose": [], "unknown": True, "expected_mocs": [], "duplicate": [],
                              "issues": [f"exigence non applicable aux releases du plan ({', '.join(sorted(wanted))}) : "
                                         f"ses releases sont {own} — seule possibilité, ignorer la ligne"]})
                continue
            iadt, moc = cls.split_codes(r["iadt"]), cls.split_codes(r["moc"])
            issues, choose = [], []
            # Then, cell by cell, what the user must choose: the IADT (and its MOC with it), the MOC, the delegation.
            if r["requirement"] in elsewhere:
                issues.append(f"déjà dans un autre plan de la release : {', '.join(elsewhere[r['requirement']])} — "
                              "garder l'une des deux lignes")
            if not iadt or any(c not in config.get("iadt") for c in iadt):
                issues.append("IADT vide" if not iadt else
                              "IADT inconnu : " + ", ".join(c for c in iadt if c not in config.get("iadt")))
                choose += ["iadt", "moc"]   # the MOC expected depends on the IADT chosen
            elif not moc or any(c not in config.get("mocs") for c in moc):
                issues.append("MOC vide" if not moc else
                              "MOC inconnu : " + ", ".join(c for c in moc if c not in config.get("mocs")))
                choose.append("moc")
            else:
                broken = cls.rule_issues(r, config)
                if broken:
                    issues += broken
                    choose.append("moc")
            delegations = cls.delegation_values(r, config)
            if len(delegations) > 1:
                issues.append("plusieurs délégations : en choisir une")
                choose.append("delegation")
            elif delegations and not config.is_platform(delegations[0]):
                issues.append(f"plateforme de délégation inconnue : {delegations[0]}")
                choose.append("delegation")
            if issues:
                found.append({**base, "choose": choose, "unknown": False, "issues": issues,
                              "duplicate": elsewhere.get(r["requirement"], []),
                              "expected_mocs": cls.expected_mocs(r["iadt"], config) if "iadt" not in choose else []})
        return found

    @classmethod
    def _decide(cls, found, decisions, config):
        """Choices of the user for the entries to settle, as {requirement number: (ignored, iadt, moc, delegation)}.
        `decisions`: {"rows": {number: {"ignore": true} | {"iadt": code, "moc": codes, "delegation": platform}}} — a
        value of the configuration for each cell to choose, the corrected entry following the IADT/MOC rules; an
        unknown number can only be ignored —, or {"accept_all": true} (command line: entries loaded as they are,
        unknown numbers ignored). ValueError when an entry is not settled."""
        decisions = decisions or {}
        chosen, result, left = decisions.get("rows") or {}, {}, 0
        for c in found:
            d = chosen.get(c["key"]) or {}
            # (1, None, None, None) = ignored: the row will be stored struck through (manual = deleted).
            if d.get("ignore") or (c["unknown"] and decisions.get("accept_all")):
                result[c["key"]] = (1, None, None, None)
                continue
            values = {f: str(d.get(f) or "").strip() for f in c["choose"]}
            row = {"iadt": values.get("iadt", c["iadt"]), "moc": values.get("moc", c["moc"])}
            iadt, moc = cls.split_codes(row["iadt"]), cls.split_codes(row["moc"])
            # Every cell to choose has a value; the corrected IADT and MOC are known and follow the rules.
            valid = not c["unknown"] and all(values.values())
            valid = valid and bool(iadt) and all(x in config.get("iadt") for x in iadt)
            valid = valid and bool(moc) and all(x in config.get("mocs") for x in moc)
            valid = valid and not cls.rule_issues(row, config)
            valid = valid and ("delegation" not in values or config.is_platform(values["delegation"]))
            if valid:
                # (0, iadt, moc, delegation) = corrected: None for a cell the user did not have to choose.
                if "delegation" in values:   # the name as written in the configuration
                    values["delegation"] = next(p for p in config.get("platforms") if fold(p) == fold(values["delegation"]))
                result[c["key"]] = (0, values.get("iadt"), values.get("moc"), values.get("delegation"))
            elif not decisions.get("accept_all"):
                left += 1
        if left:
            raise ValueError(f"{plural(left, 'ligne')} du plan à régler avant le chargement : choisir les valeurs "
                             "attendues ou ignorer la ligne")
        return result

    def edit(self, file_id, action, requirement=None, values=None, config=None):
        """Manual change of the row of a requirement in a plan. `values`: {"iadt": code, "moc": codes, "delegation":
        platform or "" (default)}, taken from the configuration and following the IADT/MOC rules.
        action "add": row of a requirement of the tome not yet in the plan (manual = added); "update": values replaced
        (manual = modified, or stays added; a deleted row is restored at the same time; nothing is written when they
        do not change); "delete": row struck through and ignored (manual = deleted unless already changed), removed
        for good when it was added by hand; "restore": deletion undone, when the row follows the rules.
        ValueError when the change is refused."""
        if action not in ("add", "update", "delete", "restore"):
            raise ValueError(f"action inconnue : {action}")
        with self.sql.lock:
            # The requirement must belong to the tome of the plan.
            plan = self.sql.query("SELECT tome_id FROM files WHERE id = ? AND source = 'plan'", (file_id,))
            if not plan:
                raise ValueError(f"plan inconnu : n° {file_id}")
            req = self.sql.query("SELECT id FROM requirements WHERE req_id = ? AND tome_id = ?",
                                 (str(requirement or ""), plan[0]["tome_id"]))
            if not req:
                raise ValueError(f"exigence inconnue dans le tome du plan : « {requirement or ''} »")
            requirement_id = req[0]["id"]
            rows = self.sql.query("SELECT id, manual, deleted FROM plan_acceptation WHERE file_id = ? AND requirement_id = ?",
                                  (file_id, requirement_id))
            if action == "add" and rows:
                raise ValueError(f"l'exigence {requirement} est déjà dans ce plan : modifier sa ligne")
            # Adding or bringing back a row: the requirement must not be active in another plan of the tome.
            if action == "add" or (rows and rows[0]["deleted"] and action in ("update", "restore")):
                head = next(p for p in self.plans(file_id=file_id, with_deleted=True))
                own = self._releases_of(head["tome"]).get(str(requirement), set())
                if not own & set(head["releases"]):
                    raise ValueError(f"l'exigence {requirement} ne s'applique pas aux releases du plan "
                                     f"({', '.join(head['releases'])}) : ses releases sont {', '.join(sorted(own)) or 'aucune'}")
                other = self.elsewhere(head["tome"], exclude_file=file_id).get(str(requirement))
                if other:
                    raise ValueError(f"l'exigence {requirement} est déjà dans un autre plan de la release : "
                                     f"{', '.join(other)} (une exigence n'est que dans un plan par release)")
            if action != "add" and not rows:
                raise ValueError(f"l'exigence {requirement} n'est pas dans ce plan")
            manual, deleted = (rows[0]["manual"], rows[0]["deleted"]) if rows else (None, 0)
            # The row as the page shows it (codes joined), to compare with the new values.
            current = next((g for p in self.plans(file_id=file_id, with_deleted=True) for g in p["rows"]
                            if g["key"] == requirement), None)
            if action == "delete":
                with self.sql.transaction() as db:
                    if manual == "added":
                        db.execute("DELETE FROM plan_acceptation WHERE file_id = ? AND requirement_id = ?", (file_id, requirement_id))
                    else:
                        db.execute("UPDATE plan_acceptation SET deleted = 1, manual = ? WHERE file_id = ? AND requirement_id = ?",
                                   (manual or "deleted", file_id, requirement_id))
                return None
            if action == "restore":
                broken = self.check_row(current, config)["errors"] + self.rule_issues(current, config)
                if broken:
                    raise ValueError("cette ligne ne respecte pas les règles (" + " ; ".join(broken)
                                     + ") : la modifier (✎) pour la rétablir")
                with self.sql.transaction() as db:
                    db.execute("UPDATE plan_acceptation SET deleted = 0, manual = ? WHERE file_id = ? AND requirement_id = ?",
                               (None if manual == "deleted" else manual, file_id, requirement_id))
                return None
            # Add or update: the values must come from the configuration and follow the IADT/MOC rules, as at import.
            values = {k: str(v or "").strip() for k, v in (values or {}).items()}
            iadt, delegation = values.get("iadt", ""), values.get("delegation", "")
            codes = self.split_codes(", ".join(c.strip() for c in values.get("moc", "").split(",")))
            if iadt not in config.get("iadt"):
                raise ValueError(f"IADT à choisir dans la configuration{f' (« {iadt} » inconnu)' if iadt else ''}")
            if not codes or not all(c in config.get("mocs") for c in codes):
                raise ValueError("MOC à choisir parmi ceux attendus pour l'IADT")
            broken = self.rule_issues({"iadt": iadt, "moc": ", ".join(codes)}, config)
            if broken:
                raise ValueError("règle IADT / MOC non respectée : " + " ; ".join(broken))
            if delegation and not config.is_platform(delegation):
                raise ValueError(f"plateforme de délégation inconnue : « {delegation} »")
            delegation = next((p for p in config.get("platforms") if fold(p) == fold(delegation)), "")
            if action == "update" and not deleted and \
                    (current["iadt"], current["moc"], current["delegation"]) == (iadt, ", ".join(codes), delegation):
                return current   # saved without any change: nothing written, the row keeps its mark
            with self.sql.transaction() as db:
                # The rows of the requirement are written again (one per pair of codes), marked added or modified.
                db.execute("DELETE FROM plan_acceptation WHERE file_id = ? AND requirement_id = ?", (file_id, requirement_id))
                self._insert(db, file_id, requirement_id, iadt, ", ".join(codes), delegation,
                             "added" if action == "add" or manual == "added" else "modified")
            return {"key": requirement}

    def check_tome(self, tome, requirements, config, selected=()):
        """Check of the plans of a tome against its requirements (RequirementsDatabase `requirements`), release
        by release: a requirement must be found in a plan of the tome applying to one of its releases.
        `selected`: releases chosen in the interface (none = all). Requirements deleted by hand are ignored."""
        selected = {s.upper() for s in selected}
        in_selection = lambda releases: not selected or bool(selected & set(releases))
        current = requirements.get(tome)
        reqs = [r for r in (current["requirements"] if current else [])
                if not r["deleted"] and in_selection(r["releases"])]
        plans = [p for p in self.plans(tome=tome, config=config) if in_selection(p["releases"])]
        by_id = {r["id"]: r for r in reqs}
        # Covered: in a plan of the tome; anomalies: rows with configuration errors (shown in the check of the tome).
        covered, misplaced, anomalies = set(), {}, []
        for plan in plans:
            for row in plan["rows"]:
                for req_id in row["requirements"]:
                    r = by_id.get(req_id)
                    if r is None:   # requirement deleted by hand, or out of the selected releases
                        continue
                    if set(r["releases"]) & set(plan["releases"]):
                        covered.add(req_id)
                    else:
                        entry = misplaced.setdefault(req_id, {"requirement": r, "plans": []})
                        entry["plans"] = unique(entry["plans"] + [plan["name"]])
                if row["errors"]:
                    anomalies.append({"plan": plan["name"], "ids": ", ".join(row["requirements"]),
                                      "errors": row["errors"]})
        with_release = [r for r in reqs if r["releases"]]
        by_release = []
        # Per release: the plans that apply to it, the requirements expected and how many are missing.
        for release in sorted({x for r in with_release for x in r["releases"] if in_selection([x])}, key=natural_key):
            release_plans = [p for p in plans if release in p["releases"]]
            found = {i for p in release_plans for row in p["rows"] for i in row["requirements"]}
            expected = [r for r in with_release if release in r["releases"]]
            by_release.append({"release": release, "plans": [p["name"] for p in release_plans],
                             "expected": len(expected), "missing": len([r for r in expected if r["id"] not in found])})
        return {"tome": tome, "requirements": len(reqs), "with_release": len(with_release), "plans": [p["name"] for p in plans],
                "covered": len(covered), "releases": by_release,
                "missing": [r for r in with_release if r["id"] not in covered],
                "no_release": [r for r in reqs if not r["releases"]],
                "misplaced": list(misplaced.values()), "anomalies": anomalies}

    # ------------------------------------------------------------------ #
    # Loading a plan file
    # ------------------------------------------------------------------ #

    def analyse(self, content, config=None, tome=None, releases=(), name=None):
        """What loading a plan file would do: number of entries (requirements) read, platforms named, tables skipped
        and, with a configuration, the platforms named alone in a Délégation cell and unknown of it (offered to the
        user for adding before the rows are settled) and the entries to settle (with the `tome`: numbers unknown in it
        as well).
        Refused (ValueError) when the document is unreadable or has no valid plan table."""
        data = self._readable(content, config)
        entries = self.entries(data["rows"])
        # The tome named in the file name (configuration plan_tome_regex), when it is one of the base: preselected.
        tome_name = ""
        if config is not None and name:
            match = re.search(config.get("plan_tome_regex"), name)
            found = fold(match.group(match.lastindex or 0)) if match else ""
            tome_name = next((r["name"] for r in self.sql.query("SELECT name FROM list_tomes") if fold(r["name"]) == found), "")
        # Platforms named alone in a Délégation cell and unknown of the configuration: the page offers to add them
        # before the rows are settled (a cell with several platforms is settled as a row).
        unknown = unique([d[0] for e in entries for d in [self.delegation_values(e, config)]
                          if len(d) == 1 and not config.is_platform(d[0])]) if config is not None else []
        return {"rows": len(entries), "platforms": data["platforms"], "skipped": data["skipped"],
                "unknown_platforms": unknown, "tome_name": tome_name,
                "to_confirm": self.to_confirm(entries, config, self._known(tome), self.elsewhere(tome, exclude_name=name),
                                              releases, self._releases_of(tome) if tome else None)
                if config is not None else []}

    def elsewhere(self, tome, exclude_file=None, exclude_name=None):
        """{requirement number: [plan names]}: the requirements present (not deleted) in the other plans of the tome
        (one release for the base) — a requirement is in one plan only. The plan `exclude_file`
        and the plans named `exclude_name` (replaced by a new loading) do not count."""
        if not tome:
            return {}
        found = {}
        # Every requirement of every other plan of the tome, with the names of the plans that hold it.
        for p in self.plans(tome=tome):   # one release for the base: every plan of the tome
            if p["file_id"] == exclude_file or p["name"] == exclude_name:
                continue
            for row in p["rows"]:
                found.setdefault(row["key"], []).append(p["name"])
        return found

    def _releases_of(self, tome):
        """{requirement number: set of releases} of the requirements of a tome: the release of the base for all."""
        release = {self.sql.release} if self.sql.release else set()
        return {r["req_id"]: set(release) for r in self.sql.query(
            "SELECT q.req_id FROM requirements q JOIN list_tomes t ON t.id = q.tome_id WHERE t.name = ?", (tome,))}

    def _known(self, tome):
        """Numbers of the requirements of a tome (None without tome)."""
        if not tome:
            return None
        tome_id = self.sql.tome_id(tome)
        return set(FileDatabase.tome_requirements(self.sql.db, tome_id)) if tome_id is not None else set()

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None):
        """Load an acceptance plan: attached to a tome that has requirements and to releases (at least one).
        The file is stored (FileDatabase `files`) then its content; if storing the content fails, the file is
        removed again. A plan of the same name already loaded into the tome is replaced. The entries to settle
        (to_confirm) need a decision of the user (see _decide): corrected (row marked modified), or ignored (row
        struck through; nothing for an unknown number)."""
        files.check_attachment("Plan d'acceptation", requirements, tome)
        data = self._readable(content, config)
        entries = self.entries(data["rows"])
        found = self.to_confirm(entries, config, self._known(tome), self.elsewhere(tome, exclude_name=name),
                                releases, self._releases_of(tome)) if config is not None else []
        # Unknown numbers and requirements of other releases: ignored, nothing stored.
        # The decisions are checked before anything is stored: an unsettled row refuses the load.
        only_ignore = {c["key"] for c in found if c["unknown"]}
        decided = self._decide(found, decisions, config) if config is not None else {}
        # Requirements taken from this file although in another plan of the release: the other row is struck through.
        taken = {c["key"]: c["duplicate"] for c in found if c.get("duplicate") and not decided.get(c["key"], (1,))[0]}
        duplicates = {c["key"] for c in found if c.get("duplicate")}
        by_number = {e["requirement"]: e for e in entries}
        record = files.add("plan", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                known = FileDatabase.tome_requirements(db, db.execute("SELECT tome_id FROM files WHERE id = ?",
                                                                      (record["id"],)).fetchone()[0])
                # Requirements taken from this file although in another plan of the release: the row of the other plan
                # is removed — one row only for a requirement and a release (before the rows of this plan are written).
                for number, names in taken.items():
                    if number in known:
                        db.execute("DELETE FROM plan_acceptation "
                                   "WHERE requirement_id = ? AND file_id IN (SELECT id FROM files WHERE source = 'plan' "
                                   f"AND name IN ({','.join('?' * len(names))}) AND id != ?)",
                                   [known[number], *names, record["id"]])
                # Settled entries: corrected (modified) or ignored (deleted) rows, kept when the plan is read again;
                # a requirement kept in the other plan is not stored here at all.
                for number, (ignored, iadt, moc, delegation) in decided.items():
                    if not ignored and not (iadt or moc or delegation):
                        continue   # taken as it is: written with the document rows
                    if ignored and (number in duplicates or number in only_ignore):
                        continue   # the other plan keeps it, or it is not a requirement of these releases
                    if number in known:
                        e = by_number[number]
                        self._insert(db, record["id"], known[number], iadt or e["iadt"], moc or e["moc"],
                                     delegation or e["delegation"], "deleted" if ignored else "modified", ignored)
                self._write(db, record["id"], entries)
        except Exception:
            files.remove(record["id"])   # content not stored: the file is not kept
            raise
        # Summary for the page: rows stored and ignored, platforms used, tables left out, STDHL warnings.
        ignored = sum(values[0] for values in decided.values())
        stdhl_warnings = self.stdhl_requests(record["id"], config) if config is not None else []
        # Platforms of the rows stored, with the delegations chosen.
        platforms = unique([(decided.get(e["requirement"]) or (0, None, None, None))[3] or e["delegation"]
                            for e in entries if not (decided.get(e["requirement"]) or (0,))[0]])
        return {"file_id": record["id"], "tome": tome, "rows": len(entries) - ignored, "ignored": ignored,
                "platforms": platforms, "skipped": data["skipped"], "replaced": files.remove_older(record),
                "stdhl_warnings": stdhl_warnings}

    def stdhl_requests(self, file_id, config):
        """Rows of a plan delegated to another platform than the default one while the column STDHL of the
        requirement asks for a STDHL (only the requirements delegated to the default platform are in a STDHL; the
        column is optional): messages for the page."""
        plan = self.get(file_id)
        if plan is None:
            return []
        # Value of the column STDHL of each requirement of the tome of the plan.
        default = config.get("stdhl_platform")   # the platform the STDHL apply to
        stdhl = {r["req_id"]: r["name"] for r in self.sql.query(
            "SELECT q.req_id, s.name FROM requirements q LEFT JOIN list_stdhl s ON s.id = q.stdhl_id "
            "WHERE q.tome_id = (SELECT tome_id FROM files WHERE id = ?)", (file_id,))}
        rows = [self.check_row(r, config) for r in plan["rows"]]
        extra = [r["key"] for r in rows if r["effective_delegation"] != default and stdhl.get(r["key"])]
        warnings = []
        if extra:
            warnings.append(plural(len(extra), "exigence avec un STDHL demandé", "exigences avec un STDHL demandé")
                            + f" mais déléguée{'s' if len(extra) > 1 else ''} à une autre plateforme : " + ", ".join(extra))
        return warnings

    def read_missing(self, files, config=None, again=False):
        """Read and store the content of the plan files that have none yet, and of all of them with `again`: the lines
        of a plan point to the requirements, so they are read again when the requirements change. Returns the names
        read; unreadable files are left as they are."""
        done = []
        for record in files.list(slot="plan"):
            # Rows marked (corrections, manual changes) stay: _write only rewrites the rows read from the document.
            if again or not (self.get(record["id"]) or {}).get("rows"):   # no row stored yet
                content = files.content(record["id"])
                try:
                    data = self.read_docx(content or b"", config)
                except ValueError:
                    continue
                with self.sql.transaction() as db:
                    self._write(db, record["id"], self.entries(data["rows"]))
                done.append(record["name"])
        return done

    # ------------------------------------------------------------------ #
    # Removal
    # ------------------------------------------------------------------ #

    def remove(self, file_id):
        """Remove the content of the plan of a file. Returns False if unknown."""
        with self.sql.transaction() as db:
            return db.execute("DELETE FROM plan_acceptation WHERE file_id = ?", (file_id,)).rowcount > 0

    def prune_files(self, keep):
        """Remove the plans whose file_id does not satisfy keep(file_id). Returns the names removed."""
        gone = [p for p in self.sql.query("SELECT id, name FROM files WHERE source = 'plan'")
                if not keep(p["id"])]
        for p in gone:
            self.remove(p["id"])
        return [p["name"] for p in gone]
