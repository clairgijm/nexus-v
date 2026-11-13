"""SQL database of NEXUS-V: a SQLite file, nexus-v.db, in the folder base of NEXUS-V.

Queries and writes go straight to that file; a committed transaction is saved in it. dump() gives a backup of the
whole base as SQL (a .sql file), from which it can be built again.

The tables (columns, keys, links, indexes) and the version of the base are defined in nexus-v-schema.sql, next to the
base and config.json.

Common referentials (managed from the page Configuration, see configuration.py): tomes, platforms, iadt, mocs, with
their colour; the release, a single one, is a setting of config.json. Their order is the one of their id: a new
value comes last. The other tables refer to them by id: renaming a value changes it everywhere without touching
those tables. A value still used cannot be removed from the configuration. The referential of the tests (fe_tests)
is filled by the STDHL and the fiches; a test nothing cites any more is removed (SQL triggers).

Links (foreign keys): removing a file removes what came from it (requirements, plan rows, coverage of the STDHL and
fiches), by cascade. The table meta holds the version of the schema and the next file number. The module sqlite3 is
part of Python.
"""
import colorsys
import contextlib
import random
import re
import sqlite3
import threading
from pathlib import Path

# Files of NEXUS-V itself, in a folder of fixed name, relative to the folder of NEXUS-V: settings, definition of the
# tables, database. It sits in the folder of the imported files, beside their sub-folders.
BASE_FOLDER = "fichiers-references/base"
CONFIG_NAME = "config.json"
SCHEMA_NAME = "nexus-v-schema.sql"
DATABASE_NAME = "nexus-v.db"


def read_schema(path):
    """(version of the database, table names in the order of the file, SQL creating them) from a schema file. The
    version is the comment « -- version : x » of the file; it only changes when a migration is decided. ValueError when
    the file is missing or invalid."""
    try:
        sql = Path(path).read_text(encoding="utf-8")
    except OSError:
        raise ValueError(f"définition des tables introuvable : {path}") from None
    version = re.search(r"^--\s*version\s*:\s*(\S+)", sql, re.M)
    try:
        # The schema is tried on an empty base in memory: a syntax error is reported before the real base is touched,
        # and the table names come out in the order of the file.
        check = sqlite3.connect(":memory:")
        check.executescript(sql)
        tables = tuple(r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY rowid"))
        check.close()
    except sqlite3.Error as e:
        raise ValueError(f"{path} : définition des tables invalide ({e})") from None
    if version is None or not tables:
        raise ValueError(f"{path} : version ou tables absentes")
    return version[1], tables, sql


# Hues (degrees) of the reserved colours, never given to a platform: red (not covered), green (covered); greys are
# avoided by a saturation of at least 55 %.
RESERVED_HUES = ((0, 25), (330, 360), (62, 175))


NA_COLOR = "#8a8f98"   # grey of the code N/A (IADT, MOC), as the N/A parts of the charts


def random_platform_color(taken=()):
    """A colour drawn at random for a platform: saturated, of medium lightness, outside the red and green hues, and as
    far as possible from the colours already taken."""
    # Hue (0-360°) of a colour #rrggbb, and the distance between two hues on the colour wheel.
    def hue_of(color):
        r, g, b = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
        return colorsys.rgb_to_hls(r, g, b)[0] * 360

    def gap(h1, h2):
        return min(abs(h1 - h2), 360 - abs(h1 - h2))

    used = [hue_of(c) for c in taken if re.fullmatch(r"#[0-9a-fA-F]{6}", c or "")]
    best, best_gap = None, -1
    for _ in range(40):   # the farthest of 40 draws from the hues already used
        hue = random.uniform(0, 360)
        if any(low <= hue < high for low, high in RESERVED_HUES):
            continue
        distance = min((gap(hue, h) for h in used), default=360)
        if distance > best_gap:
            best, best_gap = hue, distance
    # Medium lightness and strong saturation: never a grey, readable on the light and the dark page.
    rgb = colorsys.hls_to_rgb(best / 360, random.uniform(0.40, 0.55), random.uniform(0.55, 0.80))
    return "#" + "".join(f"{round(v * 255):02x}" for v in rgb)


# First row of tabs of the page Tables de la base: the files and the content of the documents, along the chain.
FIRST_TABLES = ("files", "requirements", "plan_acceptation", "stdhl_exi", "fe_tests", "fe_test_case", "stdhl_test_case",
                "fiches_essai")

# Referential name (as in the configuration) -> table, key column, label column (None: a plain list).
REFERENTIALS = {
    "tomes": ("list_tomes", "name", None),
    "platforms": ("list_plateformes", "name", None),
    "iadt": ("list_iadt", "code", "label"),
    "mocs": ("list_moc", "code", "description"),
}

# Rows that use a referential value: (label, SQL counting the uses of each value given as ?).
USES = {
    "tomes": [("exigence", "exigences", "SELECT COUNT(*) FROM requirements WHERE tome_id = (SELECT id FROM list_tomes WHERE name = ?)"),
              ("fichier", "fichiers", "SELECT COUNT(*) FROM files WHERE tome_id = (SELECT id FROM list_tomes WHERE name = ?)")],
    "platforms": [("ligne de plan", "lignes de plan", "SELECT COUNT(*) FROM plan_acceptation WHERE delegation_id = "
                   "(SELECT id FROM list_plateformes WHERE name = ?)"),
                  ("fichier de fiches", "fichiers de fiches", "SELECT COUNT(*) FROM files WHERE platform_id = "
                   "(SELECT id FROM list_plateformes WHERE name = ?)")],
    "iadt": [("ligne de plan", "lignes de plan", "SELECT COUNT(*) FROM plan_acceptation WHERE iadt_id = "
              "(SELECT id FROM list_iadt WHERE code = ?)")],
    "mocs": [("ligne de plan", "lignes de plan", "SELECT COUNT(*) FROM plan_acceptation WHERE moc_id = "
              "(SELECT id FROM list_moc WHERE code = ?)")],
}


def value_id(db, table, key, value):
    """Id of a value of a referential, inside a transaction. A value it does not have is added (command line only:
    the page has every value settled at import)."""
    row = db.execute(f"SELECT id FROM {table} WHERE {key} = ?", (value,)).fetchone()
    return row[0] if row else db.execute(f"INSERT INTO {table} ({key}) VALUES (?)", (value,)).lastrowid


def test_id(db, name):
    """Id of a test of the referential fe_tests (inside a transaction), added when new."""
    db.execute("INSERT OR IGNORE INTO fe_tests (name) VALUES (?)", (name,))
    return db.execute("SELECT id FROM fe_tests WHERE name = ?", (name,)).fetchone()[0]


def migrate_1_0(db, schema):
    """Base 1.0 -> 2.0: files.format, the format recognised at import (STDHL « par tome » or « unique »);
    stdhl_exi.deleted, a row deleted by the user in the page STDHL; fe_test_case, the test cases of the tests, and
    fiches_essai created again, empty, to point always to a test and, for the fiches XML, also to a test case, with
    the result of the test case (its content is not migrated: the fiches files are read again at the start)."""
    if "format" not in {r[1] for r in db.execute("PRAGMA table_info(files)")}:
        db.execute("ALTER TABLE files ADD COLUMN format TEXT NOT NULL DEFAULT ''")
    if "deleted" not in {r[1] for r in db.execute("PRAGMA table_info(stdhl_exi)")}:
        db.execute("ALTER TABLE stdhl_exi ADD COLUMN deleted INTEGER NOT NULL DEFAULT 0")
    db.execute("UPDATE files SET format = 'par tome' WHERE source = 'stdhl' AND format = ''")
    if "test_case" in {r[1] for r in db.execute("PRAGMA table_info(fiches_essai)")}:
        # The old table and its triggers are dropped, then created again by the schema, empty: the content of the
        # fiches is not migrated; the fiches files still stored are read again at the start (read_missing).
        db.execute("DROP TRIGGER IF EXISTS fiches_essai_test_unused")
        db.execute("DROP TRIGGER IF EXISTS stdhl_exi_test_unused")   # rewritten by the schema (fe_test_case too)
        db.execute("DROP TABLE fiches_essai")
        db.executescript(schema)
    db.execute("UPDATE meta SET value = '2.0' WHERE key = 'schema_version'")
    db.commit()
    return "2.0"


def migrate_2_0(db, schema):
    """Base 2.0 -> 2.1: stdhl_test_case, the test cases asked by a STDHL « unique »; the trigger removing an unused
    test case rewritten by the schema (it must also look at stdhl_test_case). The STDHL are read again at the start."""
    db.execute("DROP TRIGGER IF EXISTS fiches_essai_test_case_unused")
    db.executescript(schema)
    # The STDHL « unique » read again (read_missing reads the files without rows): their test cases with them.
    db.execute("DELETE FROM stdhl_exi WHERE file_id IN (SELECT id FROM files WHERE source = 'stdhl' AND format = 'unique')")
    db.execute("UPDATE meta SET value = '2.1' WHERE key = 'schema_version'")
    db.commit()
    return "2.1"


# Migrations of the bases of an older version: {version: function(connection) -> next version}.
MIGRATIONS = {"1.0": migrate_1_0, "2.0": migrate_2_0}


class NexusDatabase:
    """SQLite database of NEXUS-V. Every module of documents (file_database, requirements_db, plans_db, stdhl_db,
    fiches_db, traceability_db) and the configuration work on it."""

    def __init__(self, path):
        """path: the database file (nexus-v.db), created when absent; its tables are defined by nexus-v-schema.sql in
        the same folder. ValueError when that file is missing or the base has another version."""
        self.path = Path(path)
        self.release = ""   # the release of the base, set by Configuration (general.release)
        self.version, self.tables, schema = read_schema(self.path.parent / SCHEMA_NAME)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.lock = threading.RLock()
        self._depth = 0           # nesting of transaction(): only the outer one commits
        # One connection shared by the server threads, protected by the lock.
        self.db = sqlite3.connect(self.path, check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        with self.lock:
            self.db.execute("PRAGMA foreign_keys = ON")
            # Write-ahead log: writes and small queries are much faster. It needs a local disk; after a power cut the
            # last transactions may be lost, the file stays valid. The files -wal and -shm exist while the base is open.
            self.db.execute("PRAGMA journal_mode = WAL")
            self.db.execute("PRAGMA synchronous = NORMAL")
            # CREATE TABLE IF NOT EXISTS: a new base gets its tables, an existing one is left as it is (a change of a
            # table needs a migration).
            # An older base of a known version is migrated first, step by step, up to the version of the schema: its
            # tables do not have the columns the indexes and triggers of the schema need.
            stored = self.db.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'meta'").fetchone()
            version = self.db.execute("SELECT value FROM meta WHERE key = 'schema_version'").fetchone() if stored else None
            version = version[0] if version else None
            while version in MIGRATIONS and version != self.version:
                version = MIGRATIONS[version](self.db, schema)
            self.db.executescript(schema)
            self.created = version is None   # a new base: Configuration writes the default referentials into it
            if version not in (None, self.version):
                self.db.close()
                raise ValueError(f"{self.path} : base en version {version}, version attendue : {self.version} ; "
                                 "une migration est nécessaire")
            with self.db:
                # Cleaning at each opening and after each transaction (see transaction).
                self._drop_unused(self.db)
                self.set_meta("schema_version", self.version)

    @staticmethod
    def _drop_unused(db):
        """Remove the STDHL values that no requirement names anymore."""
        db.execute("DELETE FROM list_stdhl WHERE id NOT IN (SELECT stdhl_id FROM requirements WHERE stdhl_id IS NOT NULL)")

    def checkpoint(self):
        """Move everything into the main file, so that it can be copied alone."""
        with self.lock:
            self.db.execute("PRAGMA wal_checkpoint(TRUNCATE)")

    def dump(self):
        """Backup of the whole database, as SQL statements."""
        with self.lock:
            return "\n".join(self.db.iterdump()) + "\n"

    def close(self):
        """Close the database file."""
        with self.lock:
            self.db.close()

    @contextlib.contextmanager
    def transaction(self):
        """Lock + transaction: everything inside is written, or nothing (exception)."""
        with self.lock:
            # Transactions may be nested (a method calling another one): only the outer one commits, or rolls everything
            # back on an exception. The lock is re-entrant (RLock) for the same reason.
            outer = self._depth == 0
            self._depth += 1
            try:
                if outer:
                    with self.db:
                        yield self.db
                        self._drop_unused(self.db)
                else:
                    yield self.db
            # A broken constraint (unique value, link to a missing row) becomes a message for the page.
            except sqlite3.IntegrityError as e:
                raise ValueError(f"base SQL : contrainte non respectée ({e})") from None
            finally:
                self._depth -= 1

    def query(self, sql, params=()):
        """Rows of a SELECT, as dicts."""
        with self.lock:
            return [dict(r) for r in self.db.execute(sql, params)]

    def scalar(self, sql, params=()):
        """First value of the first row of a query, or None."""
        with self.lock:
            row = self.db.execute(sql, params).fetchone()
            return row[0] if row else None

    # ------------------------------------------------------------------ #
    # meta
    # ------------------------------------------------------------------ #

    def rename_test(self, old, new):
        """Rename a test of the referential fe_tests: every STDHL and fiche that cites it shows the new name. Refused
        (ValueError) for an empty name, an unknown test, or a name already taken by another test. The documents are not
        changed: read again, they give back the name they hold."""
        old, new = str(old or "").strip(), str(new or "").strip()
        if not new:
            raise ValueError("nom de test vide")
        with self.transaction() as db:
            row = db.execute("SELECT id FROM fe_tests WHERE name = ?", (old,)).fetchone()
            if row is None:
                raise ValueError(f"test inconnu : « {old} »")
            if new == old:
                return new
            if db.execute("SELECT 1 FROM fe_tests WHERE name = ?", (new,)).fetchone():
                raise ValueError(f"le test « {new} » existe déjà : un nom de test est unique")
            db.execute("UPDATE fe_tests SET name = ? WHERE id = ?", (new, row[0]))
        return new

    def meta(self, key, default=None):
        """Internal setting of the table meta, or `default`."""
        value = self.scalar("SELECT value FROM meta WHERE key = ?", (key,))
        return default if value is None else value

    def set_meta(self, key, value, db=None):
        """Write an internal setting, inside the transaction `db` when given."""
        (db or self.db).execute("INSERT INTO meta (key, value) VALUES (?, ?) "
                                "ON CONFLICT (key) DO UPDATE SET value = excluded.value", (key, str(value)))

    # ------------------------------------------------------------------ #
    # Referentials
    # ------------------------------------------------------------------ #

    @staticmethod
    def _table(kind):
        """(table, key column, label column) of a referential. KeyError when unknown."""
        if kind not in REFERENTIALS:
            raise KeyError(kind)
        return REFERENTIALS[kind]

    def get(self, kind):
        """Content of a referential in the configuration form: a list of names, or a dict code -> label."""
        table, key, label = self._table(kind)
        if label is None:
            return [r[key] for r in self.query(f"SELECT {key} FROM {table} ORDER BY id")]
        return {r[key]: r[label] for r in self.query(f"SELECT {key}, {label} FROM {table} ORDER BY id")}

    def set_many(self, values):
        """Replace the content of several referentials in one transaction: {kind: list or dict}.

        The values must already be validated (see Configuration._check): unique names, codes in upper case.
        A value still used (tome with requirements or files, platform or code of a plan row) cannot be removed.
        """
        with self.lock:
            # Every referential is checked before anything is written: a refusal leaves the base unchanged.
            for kind, value in values.items():
                self._check_removal(kind, value)
            with self.transaction() as db:
                for kind, value in values.items():
                    table, key, label = self._table(kind)
                    # Values removed from the list are deleted (those still used were refused above); the others are
                    # inserted, or their label updated: an existing value keeps its id, so the rows that use it are
                    # untouched.
                    items = [(name, "") for name in value] if label is None else list(value.items())
                    kept = [name for name, _ in items]
                    others = f"WHERE {key} NOT IN ({','.join('?' * len(kept))})"
                    db.execute(f"DELETE FROM {table} {others}", kept)   # values in use: refused by _check_removal
                    changes = [f"{label} = excluded.{label}"] if label else []
                    action = f"DO UPDATE SET {', '.join(changes)}" if changes else "DO NOTHING"
                    for name, text in items:
                        columns, values = ((key, label), (name, text)) if label else ((key,), (name,))
                        db.execute(f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({', '.join('?' * len(values))}) "
                                   f"ON CONFLICT ({key}) {action}", values)

    def _check_removal(self, kind, value):
        """ValueError naming the values that would be removed while still in use."""
        if kind not in USES:
            return
        kept = {str(v).casefold() for v in value}
        removed = [v for v in self.get(kind) if v.casefold() not in kept]
        problems = []
        # For each value removed: how many rows still use it, per kind of use (see USES).
        for name in removed:
            uses = [f"{n} {several if n > 1 else one}" for one, several, sql in USES[kind] if (n := self.scalar(sql, (name,)))]
            if uses:
                problems.append(f"« {name} » ({', '.join(uses)})")
        if problems:
            # Noun (singular, plural) and agreement of each referential.
            noun, nouns, feminine = {"tomes": ("tome", "tomes", False), "platforms": ("plateforme", "plateformes", True),
                                     "iadt": ("code IADT", "codes IADT", False), "mocs": ("code MOC", "codes MOC", False)}[kind]
            many = len(problems) > 1
            used = "utilisé" + ("e" if feminine else "") + ("s" if many else "")
            raise ValueError(f"{nouns if many else noun} encore {used}, suppression impossible : {', '.join(problems)}")

    def colors(self, kind):
        """{name or code: colour} of the values of a referential (platforms, tomes, iadt, mocs); a value
        without colour gets one drawn at random (kept in the base), the code N/A the grey of the N/A parts."""
        table, key, _ = self._table(kind)
        with self.lock:
            rows = self.query(f"SELECT id, {key} AS name, color FROM {table} ORDER BY id")
            taken = {r["color"].lower() for r in rows if r["color"]}
            # A value without colour (added by the command line or an import) gets one now, kept in the base so that it
            # is the same everywhere.
            missing = [r for r in rows if not r["color"]]
            if missing:
                with self.transaction() as db:
                    for r in missing:
                        r["color"] = NA_COLOR if r["name"].upper() == "N/A" else random_platform_color(taken)
                        taken.add(r["color"])
                        db.execute(f"UPDATE {table} SET color = ? WHERE id = ?", (r["color"], r["id"]))
            return {r["name"]: r["color"] for r in rows}

    def tome_descriptions(self):
        """{tome: description} of the tomes (what each one holds, given by the user)."""
        return {r["name"]: r["description"] for r in self.query("SELECT name, description FROM list_tomes ORDER BY id")}

    def set_tome_descriptions(self, descriptions):
        """Descriptions of tomes: {name: text}."""
        with self.transaction() as db:
            for name, text in descriptions.items():
                db.execute("UPDATE list_tomes SET description = ? WHERE name = ?", (str(text).strip(), name))

    def set_colors(self, colors):
        """Colours of values of referentials: {kind: {name or code: "#rrggbb"}}."""
        with self.transaction() as db:
            for kind, values in colors.items():
                table, key, _ = self._table(kind)
                for name, color in values.items():
                    db.execute(f"UPDATE {table} SET color = ? WHERE {key} = ?", (color, name))

    def set(self, kind, value):
        """Replace the content of one referential (see set_many)."""
        self.set_many({kind: value})

    # ------------------------------------------------------------------ #
    # Direct access to the tables (page Tables de la base)
    # ------------------------------------------------------------------ #

    def _columns(self, table):
        """Columns of a table: [{name, type, required, key}]. KeyError for an unknown table."""
        if table not in self.tables:
            raise KeyError(table)
        return [{"name": c["name"], "type": c["type"].upper(), "required": bool(c["notnull"]), "key": bool(c["pk"])}
                for c in self.query(f"PRAGMA table_info({table})")]

    def table_list(self):
        """Tables with their number of rows and their group, the rows of tabs of the page Tables de la base: 1 the
        files and the content of the documents (FIRST_TABLES, in that order), 2 the other ones: meta first, then the
        configuration lists (list_*) and any other table, in the order of the schema file."""
        # Two rows of tabs: the content along the chain first, then meta and the referentials.
        group = lambda t: 1 if t in FIRST_TABLES else 2
        order = [t for t in FIRST_TABLES if t in self.tables] + sorted(
            (t for t in self.tables if t not in FIRST_TABLES),
            key=lambda t: (t != "meta", not t.startswith("list_"), self.tables.index(t)))
        return [{"name": t, "rows": self.scalar(f"SELECT COUNT(*) FROM {t}"), "group": group(t)} for t in order]

    def table(self, table):
        """Content of a table: {name, columns, rows: [{rowid, values: {column: value}}]}."""
        columns = self._columns(table)
        # rowid identifies a row even in a table whose key is not a number (meta).
        rows = self.query(f"SELECT rowid AS rowid_, * FROM {table} ORDER BY rowid")
        return {"name": table, "columns": columns,
                "rows": [{"rowid": r["rowid_"], "values": {c["name"]: r[c["name"]] for c in columns}} for r in rows]}

    def _cell_values(self, table, values):
        """Values typed for a row, converted: a whole number for an INTEGER column, NULL for an empty optional value."""
        columns = {c["name"]: c for c in self._columns(table)}
        clean = {}
        for name, value in values.items():
            if name not in columns:
                raise ValueError(f"colonne inconnue : {name}")
            text = "" if value is None else str(value)
            # The page sends every value as text: an INTEGER column takes a whole number, an empty optional value
            # becomes NULL.
            if columns[name]["type"] == "INTEGER":
                try:
                    clean[name] = int(text) if text.strip() else None
                except ValueError:
                    raise ValueError(f"{name} : nombre entier attendu") from None
            else:
                clean[name] = None if not text and not columns[name]["required"] else text
        return clean

    def insert_row(self, table, values):
        """Add a row; columns left empty take their default value."""
        if table == "files":
            raise ValueError("un fichier se charge depuis la page Vérification : il faut aussi sa copie stockée")
        clean = {k: v for k, v in self._cell_values(table, values).items() if v not in (None, "")}
        if not clean:
            raise ValueError("ligne vide")
        with self.transaction() as db:
            db.execute(f"INSERT INTO {table} ({', '.join(clean)}) VALUES ({', '.join('?' * len(clean))})", list(clean.values()))

    def update_row(self, table, rowid, values):
        """Change values of a row. The links between tables stay checked by the database."""
        clean = self._cell_values(table, values)
        if clean:
            with self.transaction() as db:
                db.execute(f"UPDATE {table} SET {', '.join(f'{k} = ?' for k in clean)} WHERE rowid = ?",
                           [*clean.values(), rowid])

    def delete_row(self, table, rowid):
        """Remove a row (and what depends on it by cascade); refused when other rows still use it."""
        self._columns(table)
        with self.transaction() as db:
            db.execute(f"DELETE FROM {table} WHERE rowid = ?", (rowid,))

    def default_platform(self):
        """Name of the platform of an empty Délégation cell, or None."""
        return self.scalar("SELECT name FROM list_plateformes WHERE is_default = 1")

    def set_default_platform(self, name=None):
        """Make a platform the default one; without a name, only make sure there is one (the first platform)."""
        with self.transaction() as db:
            if name is None:
                if self.default_platform() is None:
                    # No default platform yet: the first one becomes it (a single row has is_default = 1).
                    db.execute("UPDATE list_plateformes SET is_default = (id = (SELECT MIN(id) FROM list_plateformes))")
            elif not db.execute("SELECT 1 FROM list_plateformes WHERE name = ?", (name,)).fetchone():
                raise ValueError(f"plateforme par défaut inconnue : « {name} »")
            else:
                db.execute("UPDATE list_plateformes SET is_default = (name = ?)", (name,))

    def file_releases(self):
        """Release of each file: {file number: [the release of the base]} — one release for the whole base
        (configuration general.release, kept in self.release by Configuration)."""
        if not self.release:
            return {}
        return {r["id"]: [self.release] for r in self.query("SELECT id FROM files")}

    def tome_id(self, name):
        """Id of a tome, from its name; None if unknown."""
        return self.scalar("SELECT id FROM list_tomes WHERE name = ?", (name,))

    def rename(self, renames):
        """Rename values of the referentials: {kind: {old name: new name}}. One UPDATE of the referential: what refers
        to it by id follows."""
        with self.lock:
            for kind, names in renames.items():
                table, key, _ = self._table(kind)
                for old, new in names.items():
                    # Every rename is checked first (old value known, new value free), then all are applied in one
                    # transaction.
                    current = self.scalar(f"SELECT id FROM {table} WHERE {key} = ?", (old,))
                    if current is None:
                        raise ValueError(f"valeur inconnue dans {kind} : « {old} »")
                    if self.scalar(f"SELECT id FROM {table} WHERE {key} = ?", (new,)) not in (None, current):
                        raise ValueError(f"« {new} » existe déjà dans {kind}")
            with self.transaction() as db:
                for kind, names in renames.items():
                    table, key, _ = self._table(kind)
                    for old, new in names.items():
                        db.execute(f"UPDATE {table} SET {key} = ? WHERE {key} = ?", (new, old))

    def ensure(self, kind, names, db=None):
        """Add to a list referential (tomes) the names it does not have yet, at the end."""
        table, key, label = self._table(kind)
        db = db or self.db
        for name in names:
            if name and not db.execute(f"SELECT 1 FROM {table} WHERE {key} = ?", (name,)).fetchone():
                db.execute(f"INSERT INTO {table} ({key}) VALUES (?)", (name,))
