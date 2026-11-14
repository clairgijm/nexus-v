"""Base SQL de NEXUS-V : un fichier SQLite, nexus-v.db, dans le dossier base de NEXUS-V.

Les lectures et les écritures vont directement dans ce fichier ; une transaction validée y est enregistrée. dump()
donne une sauvegarde de toute la base en SQL (un fichier .sql), à partir de laquelle elle peut être reconstruite.

Les tables (colonnes, clés, liens, index) et la version de la base sont définies dans nexus-v-schema.sql, à côté de
la base et de config.json.

Référentiels communs (gérés depuis la page Configuration, voir configuration.py) : tomes, plateformes, iadt, mocs,
avec leur couleur ; la release, une seule, est un réglage de config.json. Leur ordre est celui de leur id : une
nouvelle valeur vient en dernier. Les autres tables y renvoient par id : renommer une valeur la change partout sans
toucher à ces tables. Une valeur encore utilisée ne peut pas être retirée de la configuration. Le référentiel des
tests (fe_tests) est rempli par les STDHL et les fiches ; un test que plus rien ne cite est retiré (déclencheurs
SQL).

Liens (clés étrangères) : retirer un fichier retire ce qui en vient (exigences, lignes de plan, couverture des STDHL
et des fiches), par cascade. La table meta contient la version du schéma et le prochain numéro de fichier. Le module
sqlite3 fait partie de Python.
"""
import colorsys
import contextlib
import random
import re
import sqlite3
import threading
from pathlib import Path

# Fichiers propres à NEXUS-V, dans un dossier de nom fixe, relatif au dossier de NEXUS-V : réglages, définition des
# tables, base. Il se trouve dans le dossier des fichiers importés, à côté de leurs sous-dossiers.
BASE_FOLDER = "fichiers-references/base"
CONFIG_NAME = "config.json"
SCHEMA_NAME = "nexus-v-schema.sql"
DATABASE_NAME = "nexus-v.db"


def read_schema(path):
    """(version de la base, noms des tables dans l'ordre du fichier, SQL qui les crée) d'après un fichier de schéma.
    La version est le commentaire « -- version : x » du fichier ; elle ne change que lorsqu'une migration est
    décidée. ValueError quand le fichier manque ou est invalide."""
    try:
        sql = Path(path).read_text(encoding="utf-8")
    except OSError:
        raise ValueError(f"définition des tables introuvable : {path}") from None
    version = re.search(r"^--\s*version\s*:\s*(\S+)", sql, re.M)
    try:
        # Le schéma est essayé sur une base vide en mémoire : une erreur de syntaxe est signalée avant de toucher la
        # vraie base, et les noms des tables sortent dans l'ordre du fichier.
        check = sqlite3.connect(":memory:")
        check.executescript(sql)
        tables = tuple(r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY rowid"))
        check.close()
    except sqlite3.Error as e:
        raise ValueError(f"{path} : définition des tables invalide ({e})") from None
    if version is None or not tables:
        raise ValueError(f"{path} : version ou tables absentes")
    return version[1], tables, sql


# Teintes (degrés) des couleurs réservées, jamais données à une plateforme : rouge (non couvert), vert (couvert) ; les
# gris sont évités par une saturation d'au moins 55 %.
RESERVED_HUES = ((0, 25), (330, 360), (62, 175))




def random_platform_color(taken=()):
    """Une couleur tirée au hasard pour une plateforme : saturée, de luminosité moyenne, hors des teintes rouge et
    verte, et aussi loin que possible des couleurs déjà prises."""
    # Teinte (0-360°) d'une couleur #rrggbb, et distance entre deux teintes sur le cercle chromatique.
    def hue_of(color):
        r, g, b = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))
        return colorsys.rgb_to_hls(r, g, b)[0] * 360

    def gap(h1, h2):
        return min(abs(h1 - h2), 360 - abs(h1 - h2))

    used = [hue_of(c) for c in taken if re.fullmatch(r"#[0-9a-fA-F]{6}", c or "")]
    best, best_gap = None, -1
    for _ in range(40):   # la plus éloignée de 40 tirages des teintes déjà utilisées
        hue = random.uniform(0, 360)
        if any(low <= hue < high for low, high in RESERVED_HUES):
            continue
        distance = min((gap(hue, h) for h in used), default=360)
        if distance > best_gap:
            best, best_gap = hue, distance
    # Luminosité moyenne et forte saturation : jamais un gris, lisible sur la page claire et la page sombre.
    rgb = colorsys.hls_to_rgb(best / 360, random.uniform(0.40, 0.55), random.uniform(0.55, 0.80))
    return "#" + "".join(f"{round(v * 255):02x}" for v in rgb)


# Première rangée d'onglets de la page Tables de la base : les fichiers et le contenu des documents, le long de la
# chaîne.
FIRST_TABLES = ("files", "requirements", "plan_acceptation", "stdhl_exi", "fe_tests", "fe_test_case", "stdhl_test_case",
                "fiches_essai")

# Nom du référentiel (comme dans la configuration) -> table, colonne clé, colonne libellé (None : une simple liste).
REFERENTIALS = {
    "tomes": ("list_tomes", "name", None),
    "platforms": ("list_plateformes", "name", None),
    "iadt": ("list_iadt", "code", "label"),
    "mocs": ("list_moc", "code", "description"),
}

# Lignes qui utilisent une valeur de référentiel : (libellé, SQL qui compte les utilisations de chaque valeur donnée par
# ?).
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
    """Id d'une valeur d'un référentiel, dans une transaction. Une valeur qu'il n'a pas est ajoutée (ligne de
    commande seulement : la page fait régler chaque valeur à l'import)."""
    row = db.execute(f"SELECT id FROM {table} WHERE {key} = ?", (value,)).fetchone()
    return row[0] if row else db.execute(f"INSERT INTO {table} ({key}) VALUES (?)", (value,)).lastrowid


def test_id(db, name):
    """Id d'un test du référentiel fe_tests (dans une transaction), ajouté s'il est nouveau."""
    db.execute("INSERT OR IGNORE INTO fe_tests (name) VALUES (?)", (name,))
    return db.execute("SELECT id FROM fe_tests WHERE name = ?", (name,)).fetchone()[0]


# Migrations des bases d'une version plus ancienne : {version: fonction(connexion) -> version suivante}.
MIGRATIONS = {}


class NexusDatabase:
    """Base SQLite de NEXUS-V. Chaque module de documents (file_database, requirements_db, plans_db, stdhl_db,
    fiches_db, traceability_db) et la configuration travaillent dessus."""

    def __init__(self, path):
        """path : le fichier de la base (nexus-v.db), créé s'il manque ; ses tables sont définies par
        nexus-v-schema.sql dans le même dossier. ValueError quand ce fichier manque ou que la base est d'une autre
        version."""
        self.path = Path(path)
        self.release = ""   # la release de la base, fixée par Configuration (general.release)
        self.version, self.tables, schema = read_schema(self.path.parent / SCHEMA_NAME)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.lock = threading.RLock()
        self._depth = 0           # imbrication de transaction() : seule la plus extérieure valide
        # Une connexion partagée par les fils du serveur, protégée par le verrou.
        self.db = sqlite3.connect(self.path, check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        with self.lock:
            self.db.execute("PRAGMA foreign_keys = ON")
            # Journal d'écriture anticipée (WAL) : les écritures et les petites requêtes sont bien plus rapides. Il faut
            # un disque local ; après une coupure de courant, les dernières transactions peuvent être perdues, le
            # fichier reste valide. Les fichiers -wal et -shm existent tant que la base est ouverte.
            self.db.execute("PRAGMA journal_mode = WAL")
            self.db.execute("PRAGMA synchronous = NORMAL")
            # CREATE TABLE IF NOT EXISTS : une nouvelle base reçoit ses tables, une base existante est laissée telle
            # quelle (un changement de table demande une migration).
            # Une base plus ancienne d'une version connue est d'abord migrée, pas à pas, jusqu'à la version du schéma :
            # ses tables n'ont pas les colonnes dont les index et déclencheurs du schéma ont besoin.
            stored = self.db.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'meta'").fetchone()
            version = self.db.execute("SELECT value FROM meta WHERE key = 'schema_version'").fetchone() if stored else None
            version = version[0] if version else None
            while version in MIGRATIONS and version != self.version:
                version = MIGRATIONS[version](self.db, schema)
            self.db.executescript(schema)
            self.created = version is None   # une nouvelle base : Configuration y écrit les référentiels par défaut
            if version not in (None, self.version):
                self.db.close()
                raise ValueError(f"{self.path} : base en version {version}, version attendue : {self.version} ; "
                                 "une migration est nécessaire")
            with self.db:
                # Nettoyage à chaque ouverture et après chaque transaction (voir transaction).
                self._drop_unused(self.db)
                self.set_meta("schema_version", self.version)

    @staticmethod
    def _drop_unused(db):
        """Retire les valeurs STDHL qu'aucune exigence ne nomme plus."""
        db.execute("DELETE FROM list_stdhl WHERE id NOT IN (SELECT stdhl_id FROM requirements WHERE stdhl_id IS NOT NULL)")

    def checkpoint(self):
        """Fait tout passer dans le fichier principal, pour qu'il puisse être copié seul."""
        with self.lock:
            self.db.execute("PRAGMA wal_checkpoint(TRUNCATE)")

    def dump(self):
        """Sauvegarde de toute la base, en instructions SQL."""
        with self.lock:
            return "\n".join(self.db.iterdump()) + "\n"

    def close(self):
        """Ferme le fichier de la base."""
        with self.lock:
            self.db.close()

    @contextlib.contextmanager
    def transaction(self):
        """Verrou + transaction : tout ce qui est dedans est écrit, ou rien (exception)."""
        with self.lock:
            # Les transactions peuvent être imbriquées (une méthode qui en appelle une autre) : seule la plus extérieure
            # valide, ou annule tout sur une exception. Le verrou est réentrant (RLock) pour la même raison.
            outer = self._depth == 0
            self._depth += 1
            try:
                if outer:
                    with self.db:
                        yield self.db
                        self._drop_unused(self.db)
                else:
                    yield self.db
            # Une contrainte non respectée (valeur unique, lien vers une ligne absente) devient un message pour la page.
            except sqlite3.IntegrityError as e:
                raise ValueError(f"base SQL : contrainte non respectée ({e})") from None
            finally:
                self._depth -= 1

    def query(self, sql, params=()):
        """Lignes d'un SELECT, sous forme de dicts."""
        with self.lock:
            return [dict(r) for r in self.db.execute(sql, params)]

    def scalar(self, sql, params=()):
        """Première valeur de la première ligne d'une requête, ou None."""
        with self.lock:
            row = self.db.execute(sql, params).fetchone()
            return row[0] if row else None

    # ------------------------------------------------------------------ #
    # meta
    # ------------------------------------------------------------------ #

    def rename_test(self, old, new):
        """Renomme un test du référentiel fe_tests : chaque STDHL et chaque fiche qui le cite montre le nouveau nom.
        Refusé (ValueError) pour un nom vide, un test inconnu ou un nom déjà pris par un autre test. Les documents ne
        sont pas changés : relus, ils redonnent le nom qu'ils contiennent."""
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
        """Réglage interne de la table meta, ou `default`."""
        value = self.scalar("SELECT value FROM meta WHERE key = ?", (key,))
        return default if value is None else value

    def set_meta(self, key, value, db=None):
        """Écrit un réglage interne, dans la transaction `db` quand elle est donnée."""
        (db or self.db).execute("INSERT INTO meta (key, value) VALUES (?, ?) "
                                "ON CONFLICT (key) DO UPDATE SET value = excluded.value", (key, str(value)))

    # ------------------------------------------------------------------ #
    # Référentiels
    # ------------------------------------------------------------------ #

    @staticmethod
    def _table(kind):
        """(table, colonne clé, colonne libellé) d'un référentiel. KeyError s'il est inconnu."""
        if kind not in REFERENTIALS:
            raise KeyError(kind)
        return REFERENTIALS[kind]

    def get(self, kind):
        """Contenu d'un référentiel sous la forme de la configuration : une liste de noms, ou un dict code ->
        libellé."""
        table, key, label = self._table(kind)
        if label is None:
            return [r[key] for r in self.query(f"SELECT {key} FROM {table} ORDER BY id")]
        return {r[key]: r[label] for r in self.query(f"SELECT {key}, {label} FROM {table} ORDER BY id")}

    def set_many(self, values):
        """Remplace le contenu de plusieurs référentiels en une transaction : {type: liste ou dict}.

        Les valeurs doivent déjà être validées (voir Configuration._check) : noms uniques, codes en majuscules. Une
        valeur encore utilisée (tome avec des exigences ou des fichiers, plateforme ou code d'une ligne de plan) ne
        peut pas être retirée.
        """
        with self.lock:
            # Chaque référentiel est vérifié avant toute écriture : un refus laisse la base inchangée.
            for kind, value in values.items():
                self._check_removal(kind, value)
            with self.transaction() as db:
                for kind, value in values.items():
                    table, key, label = self._table(kind)
                    # Les valeurs retirées de la liste sont supprimées (celles encore utilisées ont été refusées plus
                    # haut) ; les autres sont insérées, ou leur libellé mis à jour : une valeur existante garde son id,
                    # les lignes qui l'utilisent ne sont donc pas touchées.
                    items = [(name, "") for name in value] if label is None else list(value.items())
                    kept = [name for name, _ in items]
                    others = f"WHERE {key} NOT IN ({','.join('?' * len(kept))})"
                    db.execute(f"DELETE FROM {table} {others}", kept)   # valeurs utilisées : refusées avant
                    changes = [f"{label} = excluded.{label}"] if label else []
                    action = f"DO UPDATE SET {', '.join(changes)}" if changes else "DO NOTHING"
                    for name, text in items:
                        columns, values = ((key, label), (name, text)) if label else ((key,), (name,))
                        db.execute(f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({', '.join('?' * len(values))}) "
                                   f"ON CONFLICT ({key}) {action}", values)

    def _check_removal(self, kind, value):
        """ValueError qui nomme les valeurs qui seraient retirées alors qu'elles sont encore utilisées."""
        if kind not in USES:
            return
        kept = {str(v).casefold() for v in value}
        removed = [v for v in self.get(kind) if v.casefold() not in kept]
        problems = []
        # Pour chaque valeur retirée : combien de lignes l'utilisent encore, par type d'utilisation (voir USES).
        for name in removed:
            uses = [f"{n} {several if n > 1 else one}" for one, several, sql in USES[kind] if (n := self.scalar(sql, (name,)))]
            if uses:
                problems.append(f"« {name} » ({', '.join(uses)})")
        if problems:
            # Nom (singulier, pluriel) et accord de chaque référentiel.
            noun, nouns, feminine = {"tomes": ("tome", "tomes", False), "platforms": ("plateforme", "plateformes", True),
                                     "iadt": ("code IADT", "codes IADT", False), "mocs": ("code MOC", "codes MOC", False)}[kind]
            many = len(problems) > 1
            used = "utilisé" + ("e" if feminine else "") + ("s" if many else "")
            raise ValueError(f"{nouns if many else noun} encore {used}, suppression impossible : {', '.join(problems)}")

    def colors(self, kind):
        """{nom ou code: couleur} des valeurs d'un référentiel (plateformes, tomes, iadt, mocs) ; une valeur sans
        couleur en reçoit une tirée au hasard (gardée dans la base)."""
        table, key, _ = self._table(kind)
        with self.lock:
            rows = self.query(f"SELECT id, {key} AS name, color FROM {table} ORDER BY id")
            taken = {r["color"].lower() for r in rows if r["color"]}
            # Une valeur sans couleur (ajoutée par la ligne de commande ou un import) en reçoit une maintenant, gardée
            # dans la base pour qu'elle soit partout la même.
            missing = [r for r in rows if not r["color"]]
            if missing:
                with self.transaction() as db:
                    for r in missing:
                        r["color"] = random_platform_color(taken)
                        taken.add(r["color"])
                        db.execute(f"UPDATE {table} SET color = ? WHERE id = ?", (r["color"], r["id"]))
            return {r["name"]: r["color"] for r in rows}

    def tome_descriptions(self):
        """{tome: description} des tomes (ce que chacun contient, donné par l'utilisateur)."""
        return {r["name"]: r["description"] for r in self.query("SELECT name, description FROM list_tomes ORDER BY id")}

    def set_tome_descriptions(self, descriptions):
        """Descriptions des tomes : {nom: texte}."""
        with self.transaction() as db:
            for name, text in descriptions.items():
                db.execute("UPDATE list_tomes SET description = ? WHERE name = ?", (str(text).strip(), name))

    def set_colors(self, colors):
        """Couleurs des valeurs des référentiels : {type: {nom ou code: "#rrggbb"}}."""
        with self.transaction() as db:
            for kind, values in colors.items():
                table, key, _ = self._table(kind)
                for name, color in values.items():
                    db.execute(f"UPDATE {table} SET color = ? WHERE {key} = ?", (color, name))

    def set(self, kind, value):
        """Remplace le contenu d'un référentiel (voir set_many)."""
        self.set_many({kind: value})

    # ------------------------------------------------------------------ #
    # Accès direct aux tables (page Tables de la base)
    # ------------------------------------------------------------------ #

    def _columns(self, table):
        """Colonnes d'une table : [{name, type, required, key}]. KeyError pour une table inconnue."""
        if table not in self.tables:
            raise KeyError(table)
        return [{"name": c["name"], "type": c["type"].upper(), "required": bool(c["notnull"]), "key": bool(c["pk"])}
                for c in self.query(f"PRAGMA table_info({table})")]

    def table_list(self):
        """Les tables avec leur nombre de lignes et leur groupe, les rangées d'onglets de la page Tables de la base :
        1 les fichiers et le contenu des documents (FIRST_TABLES, dans cet ordre), 2 les autres : meta d'abord, puis
        les listes de la configuration (list_*) et toute autre table, dans l'ordre du fichier de schéma."""
        # Deux rangées d'onglets : le contenu le long de la chaîne d'abord, puis meta et les référentiels.
        group = lambda t: 1 if t in FIRST_TABLES else 2
        order = [t for t in FIRST_TABLES if t in self.tables] + sorted(
            (t for t in self.tables if t not in FIRST_TABLES),
            key=lambda t: (t != "meta", not t.startswith("list_"), self.tables.index(t)))
        return [{"name": t, "rows": self.scalar(f"SELECT COUNT(*) FROM {t}"), "group": group(t)} for t in order]

    def table(self, table):
        """Contenu d'une table : {name, columns, rows: [{rowid, values: {colonne: valeur}}]}."""
        columns = self._columns(table)
        # rowid repère une ligne même dans une table dont la clé n'est pas un nombre (meta).
        rows = self.query(f"SELECT rowid AS rowid_, * FROM {table} ORDER BY rowid")
        return {"name": table, "columns": columns,
                "rows": [{"rowid": r["rowid_"], "values": {c["name"]: r[c["name"]] for c in columns}} for r in rows]}

    def _cell_values(self, table, values):
        """Valeurs saisies pour une ligne, converties : un nombre entier pour une colonne INTEGER, NULL pour une
        valeur facultative vide."""
        columns = {c["name"]: c for c in self._columns(table)}
        clean = {}
        for name, value in values.items():
            if name not in columns:
                raise ValueError(f"colonne inconnue : {name}")
            text = "" if value is None else str(value)
            # La page envoie chaque valeur en texte : une colonne INTEGER prend un nombre entier, une valeur facultative
            # vide devient NULL.
            if columns[name]["type"] == "INTEGER":
                try:
                    clean[name] = int(text) if text.strip() else None
                except ValueError:
                    raise ValueError(f"{name} : nombre entier attendu") from None
            else:
                clean[name] = None if not text and not columns[name]["required"] else text
        return clean

    def insert_row(self, table, values):
        """Ajoute une ligne ; les colonnes laissées vides prennent leur valeur par défaut."""
        if table == "files":
            raise ValueError("un fichier se charge depuis la page Vérification : il faut aussi sa copie stockée")
        clean = {k: v for k, v in self._cell_values(table, values).items() if v not in (None, "")}
        if not clean:
            raise ValueError("ligne vide")
        with self.transaction() as db:
            db.execute(f"INSERT INTO {table} ({', '.join(clean)}) VALUES ({', '.join('?' * len(clean))})", list(clean.values()))

    def update_row(self, table, rowid, values):
        """Change des valeurs d'une ligne. Les liens entre tables restent vérifiés par la base."""
        clean = self._cell_values(table, values)
        if clean:
            with self.transaction() as db:
                db.execute(f"UPDATE {table} SET {', '.join(f'{k} = ?' for k in clean)} WHERE rowid = ?",
                           [*clean.values(), rowid])

    def delete_row(self, table, rowid):
        """Retire une ligne (et ce qui en dépend par cascade) ; refusé quand d'autres lignes l'utilisent encore."""
        self._columns(table)
        with self.transaction() as db:
            db.execute(f"DELETE FROM {table} WHERE rowid = ?", (rowid,))

    def requirement_platforms(self):
        """Les plateformes auxquelles les exigences (non supprimées) sont déléguées, dans l'ordre du référentiel."""
        return [r["name"] for r in self.query(
            "SELECT DISTINCT p.id, p.name FROM requirements q JOIN list_plateformes p ON p.id = q.delegation_id "
            "WHERE q.deleted = 0 ORDER BY p.id")]

    def file_releases(self):
        """Release de chaque fichier : {numéro de fichier: [la release de la base]} — une release pour toute la base
        (configuration general.release, gardée dans self.release par Configuration)."""
        if not self.release:
            return {}
        return {r["id"]: [self.release] for r in self.query("SELECT id FROM files")}

    def tome_id(self, name):
        """Id d'un tome, d'après son nom ; None s'il est inconnu."""
        return self.scalar("SELECT id FROM list_tomes WHERE name = ?", (name,))

    def rename(self, renames):
        """Renomme des valeurs des référentiels : {type: {ancien nom: nouveau nom}}. Une seule mise à jour du
        référentiel : ce qui y renvoie par id suit."""
        with self.lock:
            for kind, names in renames.items():
                table, key, _ = self._table(kind)
                for old, new in names.items():
                    # Chaque renommage est vérifié d'abord (ancienne valeur connue, nouvelle valeur libre), puis tous
                    # sont appliqués en une transaction.
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
        """Ajoute à un référentiel en liste (tomes) les noms qu'il n'a pas encore, à la fin."""
        table, key, label = self._table(kind)
        db = db or self.db
        for name in names:
            if name and not db.execute(f"SELECT 1 FROM {table} WHERE {key} = ?", (name,)).fetchone():
                db.execute(f"INSERT INTO {table} ({key}) VALUES (?)", (name,))
