"""Exigences de NEXUS-V, dans la base SQL (table requirements). Seules les exigences de la release de la base
(configuration general.release) sont importées.

Une exigence appartient à un tome (lien vers la table list_tomes) ; elle vient d'un fichier d'exigences importé
(file_id : elle est retirée avec le fichier, par cascade).

Un import dans un tome qui existe déjà est soit un remplacement (le tome reçoit exactement le nouveau contenu), soit
un complément (les nouvelles exigences sont ajoutées, celles déjà présentes sont mises à jour).

Lecture d'un fichier d'exigences (read_excel) : une ligne d'en-tête, puis les colonnes repérées par leur en-tête (ID,
titre, Specification releases, Validation releases, STDHL, IADT, MOC, délégation). Le fichier est la référence de
chaque exigence : un code IADT, un code MOC et une délégation (plateforme), valeurs de la configuration ; la
délégation est obligatoire. Une ligne avec une délégation vide, un code IADT ou MOC inconnu de la configuration,
plusieurs valeurs, ou un MOC que les règles IADT / MOC n'attendent pas pour l'IADT refuse le fichier ; une délégation
vers une plateforme absente de la configuration est une nouvelle plateforme, ajoutée au référentiel à l'import.
La colonne « Specification releases » donne les releases où l'exigence s'applique (seules celles de la release de la
base sont importées) ; « Validation releases » les releases où elle est testée, séparées par des virgules, et/ou
QUALIFIE. Elle donne le champ validation de l'exigence, pour la release de la base : qualified quand QUALIFIE est
écrit (seul ce point compte : les releases écrites avec lui sont celles de tests faits avant), to_qualify quand la
release de la base est dans la liste, out_of_scope sinon (cellule vide comprise : l'exigence est hors du scope de la
validation de cette release). Elle ne refuse jamais le fichier. L'import se fait en deux temps : analyse (analyse :
feuilles, exigences, exigences hors release, différences avec le tome), puis, une fois approuvé dans la page,
enregistrement du fichier et de ses exigences (import_file). Les lignes sans la release de la base (autre release, ou
aucune) sont montrées à l'utilisateur groupées par releases : elles sont rejetées à l'import (jamais forcées dans la
release de la base).
Avant l'import, l'utilisateur règle les conflits : la ligne gardée pour un numéro présent sur plusieurs lignes du
fichier, la version gardée pour un numéro déjà dans un autre tome de la base, la valeur gardée pour plusieurs valeurs
STDHL.

Une exigence peut être modifiée, ajoutée ou supprimée à la main (edit) : elle est alors marquée (manual = « added » /
« modified », deleted = 1 : montrée barrée et laissée de côté par la vérification, elle peut être rétablie). Le
prochain import du tome écrase ces changements : le fichier reste la référence.

Pour la page, les exigences sont données par tome (to_dict) :

    {"tomes": {"fct1": {"description": "...", "files": [3, 7],
                        "requirements": [{"id", "title", "releases", "stdhl", "iadt", "moc", "delegation", "file_id",
                                          "manual", "deleted"}, ...]}}}
"""
import io
import re

import openpyxl

from configuration import DEFAULT_CONFIG, fold, match_header, plural

MODES = ("replace", "complement")

# Champs comparés entre les exigences d'un tome et celles d'un fichier.
DIFF_FIELDS = ("title", "releases", "validation", "stdhl", "iadt", "moc", "delegation")
QUALIFIED = "QUALIFIE"   # valeur de « Validation releases » : l'exigence est qualifiée
# Validation d'une exigence dans la release de la base (champ validation), et son libellé.
VALIDATION_STATES = {"to_qualify": "à qualifier", "qualified": "qualifiée", "out_of_scope": "hors scope"}
# Référence d'une exigence : (champ, table, colonne de la valeur).
REFERENCE = (("iadt", "list_iadt", "code"), ("moc", "list_moc", "code"), ("delegation", "list_plateformes", "name"))

# Caractères que Word / Excel mettent à la place des espaces et des tirets ordinaires.
# Un numéro d'exigence collé depuis Word ou une page web porte souvent de tels caractères : il ne correspondrait pas.
NORMALIZATION = str.maketrans({"\u00a0": " ", "\u202f": " ", "\u2011": "-", "\u2010": "-", "\u2013": "-", "\u00ad": ""})


class RequirementsDatabase:
    """Exigences, par tome, dans la base SQL.

    Une exigence est un dict : id (str), title (str), releases (liste de str), stdhl (str), iadt, moc, delegation
    (str), validation (str), file_id (int ou None : numéro du fichier importé d'où elle vient).
    """

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    # ------------------------------------------------------------------ #
    # Lecture / écriture
    # ------------------------------------------------------------------ #

    @staticmethod
    def _clean(req, file_id=None):
        """Copie validée d'une exigence."""
        req_id = str(req.get("id") or "").strip()
        if not req_id:
            raise ValueError("exigence sans ID")
        releases = req.get("releases") or []
        # La page peut envoyer les releases en un texte « A, B » ou en liste.
        if isinstance(releases, str):
            releases = releases.split(",")
        return {
            "id": req_id,
            "title": str(req.get("title") or "").strip(),
            "releases": list(dict.fromkeys(str(r).strip().upper() for r in releases if str(r).strip())),
            "stdhl": str(req.get("stdhl") or "").strip(),
            # Validation dans la release de la base (voir validation_state).
            "validation": req.get("validation") if req.get("validation") in VALIDATION_STATES else "out_of_scope",
            **{f: str(req.get(f) or "").strip() for f, _, _ in REFERENCE},
            "file_id": file_id if file_id is not None else (int(req["file_id"]) if req.get("file_id") is not None else None),
        }

    @staticmethod
    def validation_values(value, separator=","):
        """Valeurs de « Validation releases » (texte ou liste) : les releases en majuscules, dans l'ordre, puis
        QUALIFIE quand il est écrit (quels que soient sa casse et son accent)."""
        if isinstance(value, str):
            value = re.split(separator, value)
        values = [str(v).strip().upper() for v in value or () if str(v).strip()]
        qualified = any(fold(v) == fold(QUALIFIED) for v in values)
        return list(dict.fromkeys(v for v in values if fold(v) != fold(QUALIFIED))) + ([QUALIFIED] if qualified else [])

    @staticmethod
    def validation_state(values, release):
        """Validation d'une exigence dans `release` d'après ses « Validation releases » : qualified (QUALIFIE écrit,
        quelles que soient les releases), to_qualify (la release est dans la liste) ou out_of_scope."""
        if QUALIFIED in values:
            return "qualified"
        return "to_qualify" if release and release.upper() in values else "out_of_scope"

    @staticmethod
    def _stdhl_id(db, name):
        """Id d'une valeur de la colonne STDHL, ajoutée à la table list_stdhl si elle est nouvelle ; None pour une
        cellule vide."""
        if not name:
            return None
        row = db.execute("SELECT id FROM list_stdhl WHERE name = ?", (name,)).fetchone()
        return row[0] if row else db.execute("INSERT INTO list_stdhl (name) VALUES (?)", (name,)).lastrowid

    @staticmethod
    def _write_tome(db, tome, reqs):
        """Remplace les exigences d'un tome (dans une transaction)."""
        tome_id = db.execute("SELECT id FROM list_tomes WHERE name = ?", (tome,)).fetchone()[0]
        # Supprimées puis réécrites : les lignes des plans, STDHL et fiches qui pointaient vers elles sont retirées par
        # la cascade et reconstruites par le serveur, qui relit ces documents après un import.
        db.execute("DELETE FROM requirements WHERE tome_id = ?", (tome_id,))
        for r in reqs:
            db.execute(
                "INSERT INTO requirements (req_id, file_id, stdhl_id, tome_id, title, iadt_id, moc_id, delegation_id, "
                "validation) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (r["id"], r["file_id"], RequirementsDatabase._stdhl_id(db, r["stdhl"]), tome_id, r["title"],
                 *RequirementsDatabase._reference_ids(db, r), r["validation"]))

    @staticmethod
    def _reference_ids(db, req):
        """(iadt_id, moc_id, delegation_id) de la référence d'une exigence (None pour une valeur vide ou inconnue)."""
        ids = []
        for field, table, column in REFERENCE:
            row = db.execute(f"SELECT id FROM {table} WHERE {column} = ? COLLATE NOCASE", (req.get(field) or "",)).fetchone()
            ids.append(row[0] if row else None)
        return ids

    def _requirements(self, where="", params=()):
        """Exigences (avec leur tome), dans l'ordre des tomes puis des fichiers."""
        # Chaque exigence reçoit la release de la base : elle n'est pas enregistrée, toute la base travaille sur une
        # release.
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
    # Accès
    # ------------------------------------------------------------------ #

    def tomes(self):
        """Noms des tomes qui ont des exigences, dans l'ordre du référentiel."""
        return [r["name"] for r in self.sql.query(
            "SELECT name FROM list_tomes t WHERE EXISTS (SELECT 1 FROM requirements r WHERE r.tome_id = t.id) ORDER BY t.id")]

    def get(self, tome):
        """Un tome : {"description", "files", "requirements"}, ou None s'il n'a pas d'exigence."""
        found = self.to_dict(tome)["tomes"]
        return next(iter(found.values()), None)

    def to_dict(self, only=None):
        """Les exigences par tome (les tomes sans exigence sont laissés de côté)."""
        result = {}
        where, params = ("WHERE t.name = ?", (only,)) if only is not None else ("", ())
        for r in self._requirements(where, params):
            tome = r.pop("tome")
            result.setdefault(tome, []).append(r)
        tomes = {}
        # Pour chaque tome : sa description (Configuration > Tomes), les fichiers importés dedans et ses exigences.
        for name, reqs in result.items():
            tomes[name] = {
                "description": self.sql.scalar("SELECT description FROM list_tomes WHERE name = ?", (name,)) or "",
                "files": [f["id"] for f in self.sql.query(
                    "SELECT f.id FROM files f JOIN list_tomes t ON t.id = f.tome_id "
                    "WHERE f.source = 'requirements' AND t.name = ? ORDER BY f.id", (name,))],
                "requirements": reqs,
            }
        # « stdhl_values » : chaque valeur de la table list_stdhl, celles que peut prendre une exigence ajoutée ou
        # modifiée à la main ;
        # « validation_states » : les valeurs du champ validation et leur libellé.
        return {"tomes": tomes, "stdhl_values": self.stdhl_names(), "validation_states": VALIDATION_STATES}

    def distribution(self, colors=None):
        """Part des exigences (non supprimées) dans chaque tome, pour la barre cumulée de la page Exigences : [{tome,
        count, percent, color}] dans l'ordre des tomes ; `colors` : {tome: couleur} (configuration)."""
        # Pourcentages à une décimale, pour que les petits tomes montrent encore une part.
        rows = self.sql.query("SELECT t.name AS tome, COUNT(q.id) AS count FROM list_tomes t "
                              "JOIN requirements q ON q.tome_id = t.id AND q.deleted = 0 GROUP BY t.id ORDER BY t.id")
        total = sum(r["count"] for r in rows)
        return [{**r, "percent": round(100 * r["count"] / total, 1) if total else 0,
                 "color": (colors or {}).get(r["tome"], "")} for r in rows]

    def delegation_shares(self, colors=None):
        """Part de chaque plateforme dans les délégations des exigences (non supprimées) de chaque tome, pour les
        barres de la page Statistiques : [{tome, total, parts: [{platform, count, percent, color}] de la plus
        grande}] ; `colors` : {plateforme: couleur} (configuration)."""
        counts = {}
        for r in self.sql.query("SELECT t.name AS tome, p.name AS platform, COUNT(*) AS n FROM requirements q "
                                "JOIN list_tomes t ON t.id = q.tome_id JOIN list_plateformes p ON p.id = q.delegation_id "
                                "WHERE q.deleted = 0 GROUP BY t.id, p.id"):
            counts.setdefault(r["tome"], {})[r["platform"]] = r["n"]
        shares = []
        for tome, by in counts.items():
            total = sum(by.values())
            shares.append({"tome": tome, "total": total, "parts": [
                {"platform": name, "count": n, "percent": round(100 * n / total, 1), "color": (colors or {}).get(name, "")}
                for name, n in sorted(by.items(), key=lambda x: -x[1])]})
        return shares

    def stdhl_names(self):
        """Valeurs de la table list_stdhl (colonne STDHL des fichiers importés), triées."""
        return sorted((r["name"] for r in self.sql.query("SELECT name FROM list_stdhl")), key=str.casefold)

    def requirements(self, tomes=None):
        """Exigences des tomes donnés (tous par défaut), chacune avec sa clé « tome »."""
        return [r for r in self._requirements() if tomes is None or r["tome"] in tomes]

    # ------------------------------------------------------------------ #
    # Lecture d'un fichier d'exigences
    # ------------------------------------------------------------------ #

    @staticmethod
    def _cell_text(value):
        """Texte d'une cellule Excel (un nombre entier est écrit sans décimales)."""
        if value is None:
            return ""
        # Excel garde chaque nombre en flottant : 12 se lirait « 12.0 ».
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        return str(value).translate(NORMALIZATION)

    @staticmethod
    def _open(content):
        """Ouvre un classeur à partir de ses octets. ValueError quand ce n'est pas un fichier Excel lisible."""
        try:
            return openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except Exception as e:   # pas un classeur .xlsx / .xlsm (un ancien .xls par exemple)
            raise ValueError(f"classeur Excel illisible (formats .xlsx et .xlsm) : {e}") from None

    @classmethod
    def sheet_names(cls, content):
        """Noms des feuilles d'un classeur."""
        workbook = cls._open(content)
        try:
            return workbook.sheetnames
        finally:
            workbook.close()

    @classmethod
    def read_excel(cls, content, sheet=None, config=None, rows=None):
        """Lit un classeur d'exigences (octets). Renvoie {"sheets" : noms, "sheet" : feuille lue, "requirements" :
        [{id, title, releases, stdhl, iadt, moc, delegation, row}], "releases" : releases nommées dans le fichier,
        "repeated" : [{id, rows: [{row, title, releases, stdhl}], chosen}] numéros trouvés sur plusieurs lignes,
        "rejected_numbers" : numéros hors du format des exigences (requirement_regex), laissés de côté}.

        La ligne 1 est l'en-tête : elle doit contenir chaque texte attendu (configuration headers_requirements, texte
        -> rôle) ; les colonnes des rôles id, title, releases, validation, stdhl, iadt, moc, delegation y sont
        repérées (stdhl lue seulement avec use_requirement_stdhl : sinon le STDHL de chaque exigence est vide).
        Sinon, le fichier est refusé (ValueError). La référence des lignes importées est vérifiée après le choix des
        releases (voir check_reference, dans analyse et import_file).
        Les releases sont mises en majuscules, leur séparateur est dans la configuration. Un numéro trouvé sur
        plusieurs lignes garde la ligne choisie par l'utilisateur, `rows` = {numéro: numéro de ligne} ; tant qu'elle
        n'est pas choisie, sa première ligne la remplace et son « chosen » est None. La première feuille est lue
        quand `sheet` n'est pas donnée.
        """
        get = config.get if config is not None else DEFAULT_CONFIG.get
        headers = get("headers_requirements")
        workbook = cls._open(content)
        try:
            names = workbook.sheetnames
            if sheet and sheet not in names:
                raise ValueError(f"feuille « {sheet} » introuvable : {', '.join(names)}")
            name = sheet or names[0]
            # La ligne 1 est l'en-tête : les colonnes sont trouvées par leur texte (headers_requirements), pas par leur
            # position.
            first = next(workbook[name].iter_rows(max_row=1, values_only=True), ())
            cells = [cls._cell_text(v) for v in first]
            cols, missing = match_header(cells, headers)
            # La colonne STDHL n'est obligatoire que lorsque les STDHL sont vérifiés par rapport à elle
            # (use_requirement_stdhl).
            optional = set() if get("use_requirement_stdhl") else {h for h, role in headers.items() if role == "stdhl"}
            missing = [h for h in missing if h not in optional]
            if optional:   # l'option désactivée : la colonne STDHL n'est pas lue, même quand elle est là
                cols.pop("stdhl", None)
            if missing:
                # Un seul message : l'en-tête n'est pas celui attendu (les colonnes sont trouvées par leur nom,
                # n'importe où).
                raise ValueError(f"en-tête non conforme à l'attendu (ligne 1 de la feuille « {name} » : colonnes "
                                 + ", ".join(f"« {h} »" for h in headers if headers[h] not in ("check", "ignore")
                                             and h not in optional)
                                 + ")")
            found, releases, rejected = {}, {}, []
            number_regex = re.compile(get("requirement_regex"))
            # Chaque ligne après l'en-tête ; `number` est le numéro de ligne montré à l'utilisateur (numérotation
            # Excel).
            for number, row in enumerate(workbook[name].iter_rows(min_row=2, values_only=True), start=2):
                value = lambda field: cls._cell_text(row[cols[field]] if field in cols and cols[field] < len(row)
                                                     else None)
                req_id, title, rels, validation, stdhl, iadt, moc, delegation = (
                    value(f) for f in ("id", "title", "releases", "validation", "stdhl", "iadt", "moc", "delegation"))
                rels = [r.strip().upper() for r in re.split(get("release_separator_regex"), rels) if r.strip()]
                releases.update(dict.fromkeys(rels))
                req_id = req_id.strip()
                # Une cellule STDHL peut contenir plusieurs valeurs ; une valeur écrite deux fois avec des casses
                # différentes compte une fois.
                values = cls.stdhl_values(stdhl, config)
                stdhl = values[0] if len(values) == 1 else stdhl   # « F1, f1 » est une seule valeur
                # Un numéro hors du format des exigences est rejeté : les plans et les STDHL ne pourraient pas le citer.
                if req_id and not number_regex.fullmatch(req_id):
                    rejected.append(req_id)
                elif req_id:
                    found.setdefault(req_id, []).append(
                        {"id": req_id, "title": title.strip(), "releases": list(dict.fromkeys(rels)),
                         "stdhl": stdhl.strip(), "stdhl_values": values, "iadt": iadt.strip(), "moc": moc.strip(),
                         "delegation": delegation.strip(), "row": number,
                         # « Validation releases » telle qu'elle est écrite ; le champ validation est fixé par
                         # check_reference.
                         "validation_releases": cls.validation_values(validation, get("release_separator_regex"))})
            # Un numéro trouvé sur plusieurs lignes : la ligne choisie par l'utilisateur (`rows`), sinon la première
            # ligne pour l'instant, avec « chosen » à None pour que la page demande.
            rows = {str(k): int(v) for k, v in (rows or {}).items()}
            requirements, repeated = [], []
            for req_id, versions in found.items():
                chosen = next((v for v in versions if v["row"] == rows.get(req_id)), None)
                requirements.append(chosen or versions[0])
                if len(versions) > 1:
                    repeated.append({"id": req_id, "chosen": chosen and chosen["row"],
                                     "rows": [{f: v[f] for f in ("row", "title", "releases", "stdhl")} for v in versions]})
            return {"sheets": names, "sheet": name, "requirements": requirements, "releases": list(releases),
                    "rejected_numbers": list(dict.fromkeys(rejected)),
                    "repeated": repeated}
        finally:
            workbook.close()

    @staticmethod
    def check_reference(requirements, config=None):
        """La référence de chaque exigence lue : un code IADT, un code MOC et une délégation, valeurs de la
        configuration (écrites comme là-bas), la délégation obligatoire, le MOC parmi ceux que les règles IADT / MOC
        attendent ; son champ validation fixé pour la release de la base d'après ses « Validation releases » (jamais
        une erreur). ValueError qui liste les lignes sinon. Une délégation vers une plateforme absente de la
        configuration n'est pas une erreur : renvoie ces nouvelles plateformes, {nom: [numéros]}."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        known = {"iadt": list(get("iadt") or {}), "moc": list(get("mocs") or {}), "delegation": list(get("platforms") or [])}
        names = {"iadt": "IADT", "moc": "MOC", "delegation": "délégation"}
        separator = get("code_separator_regex") or r"[,;+\s]+"
        problems, new = [], {}
        for r in requirements:
            # « Validation releases » ne refuse jamais le fichier : elle donne seulement la validation dans la release
            # de la base.
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
                    # Une nouvelle plateforme : ajoutée au référentiel à l'import (un seul nom quelle que soit sa
                    # casse).
                    key = next((k for k in new if fold(k) == fold(value)), value)
                    new.setdefault(key, []).append(r["id"])
                    r[field] = key
                elif len([v for v in re.split(separator if field != "delegation" else r"[,;/+\n]+", value) if v]) > 1:
                    problems.append(f"ligne {r['row']} ({r['id']}) : plusieurs valeurs {names[field]} « {value} »")
                else:
                    problems.append(f"ligne {r['row']} ({r['id']}) : {names[field]} « {value} » absent de la configuration")
            # Les règles IADT / MOC de la configuration : le MOC doit être l'un de ceux attendus pour l'IADT.
            expected = (get("iadt_moc_rules") or {}).get(r["iadt"])
            if expected and r["moc"] and r["moc"] in known["moc"] and r["moc"] not in expected:
                problems.append(f"ligne {r['row']} ({r['id']}) : MOC « {r['moc']} » non attendu pour l'IADT "
                                f"« {r['iadt']} » (attendu : {' ou '.join(expected)})")
        if problems:
            more = len(problems) - 10
            raise ValueError(f"{plural(len(problems), 'erreur')} dans la référence des exigences (IADT, MOC, délégation) : "
                             + " ; ".join(problems[:10]) + (f" … et {more} autre{'s' if more > 1 else ''}" if more > 0 else "")
                             + " : fichier non importé")
        return new

    @staticmethod
    def filter_releases(requirements, selected):
        """(exigences gardées, groupes rejetés) pour les releases sélectionnées (None : toutes). Une release non
        sélectionnée est retirée d'une exigence. Une exigence sans aucune release sélectionnée est rejetée à l'import
        : groupée par ses releases ("" pour aucune), [{"key", "count", "requirements"}]."""
        if selected is None:
            return requirements, []
        # Une exigence qui a au moins une release sélectionnée est gardée avec ces seules releases ; les autres sont
        # groupées par leur liste exacte de releases (« B, C », "" pour aucune).
        chosen, kept, outside = [s.strip().upper() for s in selected], [], {}
        for r in requirements:
            releases = [x for x in r["releases"] if x in chosen]
            if releases:
                kept.append({**r, "releases": releases})
                continue
            outside.setdefault(", ".join(r["releases"]), []).append(r)
        groups = [{"key": key, "count": len(rs), "requirements": rs}
                  for key, rs in sorted(outside.items(), key=lambda g: (g[0] == "", g[0]))]
        return kept, groups

    def diff(self, tome, requirements, mode):
        """Différences entre les exigences d'un tome (celles supprimées à la main laissées de côté) et celles d'un
        fichier : {"added" : [exigence], "updated" : [{"id", "changes": [{"field", "before", "after"}]}], "removed" :
        [exigence] (remplacement seulement)}."""
        current = self.get(tome)
        # Comparées champ par champ (les releases en texte) ; dans un complément, rien n'est retiré.
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
                       for f in self.changed_fields(old[req_id], r)]
            if changes:
                updated.append({"id": req_id, "changes": changes})
        removed = [r for i, r in old.items() if i not in fresh] if mode == "replace" else []
        return {"added": added, "updated": updated, "removed": removed}

    @staticmethod
    def stdhl_values(value, config=None):
        """Valeurs d'une cellule STDHL, sans doublon : une exigence ne peut en avoir qu'une."""
        separator = (config.get if config is not None else DEFAULT_CONFIG.get)("stdhl_separator_regex")
        # Lues à l'envers pour que, entre des valeurs égales à la casse près, la première écrite l'emporte.
        values = {v.strip().casefold(): v.strip() for v in reversed(re.split(separator, value or "")) if v.strip()}
        return list(reversed(values.values()))

    def conflicts(self, requirements, tome, mode, existing=None):
        """Ce que l'utilisateur doit régler avant l'import : {"duplicates" : exigences du fichier déjà dans un autre
        tome de la base, quel que soit le mode (un numéro, un tome : garder la base ou prendre le fichier) ;
        "multi_stdhl" : exigences du fichier à plusieurs valeurs STDHL}. Une exigence déjà dans le tome importé n'est
        pas un conflit : un remplacement ou un complément la met à jour. `existing` : les exigences de la base (avec
        leur tome), lues quand elles ne sont pas données (la vérification avant import donne sa base en mémoire)."""
        # Un numéro déjà dans un autre tome de la base est un conflit ; dans le tome importé, le fichier le met à jour.
        existing = {r["id"]: r for r in (self._requirements() if existing is None else existing)}
        names = {r["id"]: r["name"] for r in self.sql.query("SELECT id, name FROM files")}
        fields = DIFF_FIELDS
        duplicates = []
        for r in requirements:
            e = existing.get(r["id"])
            if e and e["tome"].casefold() != (tome or "").casefold():
                duplicates.append({"id": r["id"], "same": all(e[f] == r[f] for f in fields),
                                   "existing": {**{f: e[f] for f in fields}, "tome": e["tome"], "file": names.get(e["file_id"], "")},
                                   "new": {f: r[f] for f in fields}})
        multi = [{"id": r["id"], "title": r["title"], "values": r["stdhl_values"]} for r in requirements
                 if len(r["stdhl_values"]) > 1]
        return {"duplicates": duplicates, "multi_stdhl": multi}

    @staticmethod
    def _resolve(requirements, found, choices, strict):
        """(exigences à importer, numéros pris à un autre tome) avec les choix de l'utilisateur : {"row" : {numéro:
        ligne du fichier}, "keep" : {numéro: "existing" ou "new"}, "stdhl" : {numéro: valeur}} ; les lignes sont
        appliquées à la lecture (read_excel). Avec `strict`, ValueError quand quelque chose n'est pas réglé."""
        # Compte ce qui n'est pas encore réglé : doublons sans choix, cellules STDHL sans valeur choisie, numéros
        # répétés sans ligne.
        keep, stdhl = (choices or {}).get("keep") or {}, (choices or {}).get("stdhl") or {}
        duplicates = [d for d in found["duplicates"] if keep.get(d["id"]) in ("existing", "new")]
        settled = [m for m in found["multi_stdhl"] if stdhl.get(m["id"]) in m["values"]]
        repeated = sum(r["chosen"] is None for r in found.get("repeated", ()))
        left = len(found["duplicates"]) - len(duplicates) + len(found["multi_stdhl"]) - len(settled) + repeated
        if strict and left:
            raise ValueError(f"{plural(left, 'conflit')} à régler avant l'import : choisir la ligne d'une exigence "
                             "présente plusieurs fois dans le fichier, l'exigence à conserver ou la valeur STDHL")
        # « existing » : l'exigence de la base reste, la ligne du fichier est abandonnée ; « new » : la ligne du fichier
        # est importée et l'exigence quitte son autre tome (`moved`).
        kept = {d["id"] for d in duplicates if keep[d["id"]] == "existing"}
        values = {m["id"]: stdhl[m["id"]] for m in settled}
        result = [{**r, "stdhl": values.get(r["id"], r["stdhl"])} for r in requirements if r["id"] not in kept]
        return result, [d["id"] for d in duplicates if keep[d["id"]] == "new"]

    def analyse(self, content, sheet=None, tome=None, mode="replace", config=None, releases=None, choices=None):
        """Ce que ferait l'import d'un fichier d'exigences, pour les fenêtres d'approbation : feuilles, releases
        trouvées (avec leur nombre d'exigences), exigences gardées pour les `releases` sélectionnées (None : toutes),
        groupes d'exigences hors de celles-ci (« outside », voir filter_releases ; rejetés à l'import), conflits à
        régler et, dans un tome qui a déjà des exigences, les différences une fois les `choices` appliqués.
        Sans feuille, un classeur de plusieurs feuilles ne donne que les noms de ses feuilles (la feuille est choisie
        d'abord)."""
        # Premier appel, sans feuille : un classeur de plusieurs feuilles ne renvoie que les noms de ses feuilles, la
        # page demande laquelle lire.
        if sheet is None:
            names = self.sheet_names(content)
            if len(names) > 1:
                return {"sheets": names, "sheet": None, "count": None}
        data = self.read_excel(content, sheet, config, (choices or {}).get("row"))
        # Nombre d'exigences par release nommée dans le fichier, avant le filtrage.
        counts = {x: sum(x in r["releases"] for r in data["requirements"]) for x in data["releases"]}
        reqs, outside = self.filter_releases(data["requirements"], releases)
        # Les lignes importées : une référence fausse refuse le fichier ; une nouvelle plateforme est ajoutée à
        # l'import.
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
        """Tome nommé dans le nom d'un fichier d'exigences (configuration requirements_tome_regex) : un tome existant
        (quelle que soit la casse), ou le nom trouvé pour un nouveau tome ; "" quand le nom ne correspond pas."""
        # Le tome nommé dans le nom du fichier, écrit comme le tome existant si l'un correspond, quelle que soit la
        # casse.
        match = re.search(config.get("requirements_tome_regex"), name or "")
        found = (match.group(match.lastindex or 0) if match else "").strip()
        if not found:
            return ""
        return next((r["name"] for r in self.sql.query("SELECT name FROM list_tomes") if fold(r["name"]) == fold(found)), found)

    def import_file(self, files, content, name, tome, mode="replace", sheet=None, config=None, releases=None,
                    choices=None):
        """Importe un fichier d'exigences dans un tome : le fichier est enregistré (FileDatabase `files`), puis ses
        exigences ; si l'import échoue, le fichier enregistré est retiré. Un remplacement retire les fichiers
        importés auparavant dans le tome. Les exigences hors des `releases` sélectionnées sont rejetées. Les conflits
        doivent être réglés par `choices` (voir _resolve), sinon ValueError. Renvoie le résumé de
        import_requirements, plus « removed_files », « rejected » (hors release), « rejected_numbers » (hors format),
        « kept » (déjà dans la base, gardées telles quelles) et « releases » (nommées dans le fichier)."""
        data = self.read_excel(content, sheet, config, (choices or {}).get("row"))
        selected, outside = self.filter_releases(data["requirements"], releases)
        new_platforms = self.check_reference(selected, config)
        found = {**self.conflicts(selected, tome, mode), "repeated": data["repeated"]}
        reqs, moved = self._resolve(selected, found, choices, True)
        if not reqs:
            raise ValueError("aucune exigence à importer" + (f" dans la feuille « {data['sheet']} »" if sheet else ""))
        # Le tome avant l'import : ses fichiers deviennent inutiles après un remplacement.
        previous = self.get(tome)
        record = files.add("requirements", name, content, tome=tome)
        try:
            # Nouvelles plateformes des délégations : ajoutées d'abord au référentiel (leur couleur tirée à la lecture).
            if new_platforms:
                with self.sql.transaction() as db:
                    self.sql.ensure("platforms", list(new_platforms), db)
            summary = self.import_requirements(tome, reqs, mode, record["id"], moved)
        except Exception:
            files.remove(record["id"])   # rien d'importé : le fichier n'est pas gardé
            raise
        summary["removed_files"] = []
        if previous and summary["mode"] == "replace":
            # Un remplacement rend inutiles les fichiers importés auparavant dans le tome.
            for file_id in previous["files"]:
                old = files.get(file_id)
                if file_id != record["id"] and old and files.remove(file_id):
                    summary["removed_files"].append(old["name"])
        return {**summary, "file_id": record["id"], "kept": len(selected) - len(reqs),
                "rejected": sum(g["count"] for g in outside),
                "new_platforms": list(new_platforms), "rejected_numbers": data["rejected_numbers"],
                "releases": data["releases"]}

    # ------------------------------------------------------------------ #
    # Enregistrement des exigences
    # ------------------------------------------------------------------ #

    def import_requirements(self, tome, requirements, mode="replace", file_id=None, moved=()):
        """Enregistre dans un tome les exigences lues dans un fichier. Renvoie un résumé des changements.

        mode « replace » : le tome reçoit exactement ces exigences ;
        mode « complement » : les nouvelles exigences sont ajoutées, celles déjà dans le tome sont mises à jour.
        Un nouveau tome est créé (ajouté au référentiel) quel que soit le mode.
        moved : numéros pris à un autre tome, d'où ils sont retirés.
        """
        tome = str(tome or "").strip()
        if not tome:
            raise ValueError("nom de tome manquant")
        if not requirements:
            raise ValueError("aucune exigence à importer")
        if mode not in MODES:
            raise ValueError(f"mode d'import inconnu : {mode} (attendu : {', '.join(MODES)})")
        with self.sql.lock:
            current = self.get(tome)
            plan = self.merge(current["requirements"] if current else None, requirements, mode, file_id)
            merged = plan["merged"]
            with self.sql.transaction() as db:
                self.sql.ensure("tomes", [tome], db)
                # Les numéros que l'utilisateur a pris à un autre tome le quittent avant l'écriture du tome (un numéro,
                # un tome).
                db.executemany("DELETE FROM requirements WHERE req_id = ? AND tome_id != (SELECT id FROM list_tomes WHERE name = ?)",
                               [(number, tome) for number in moved])
                self._write_tome(db, tome, merged)
            return {"tome": tome, "mode": "replace" if current is None else mode, "created": current is None,
                    "added": len(plan["added"]), "updated": len(plan["updated"]), "removed": len(plan["removed"]),
                    "total": len(merged)}

    @staticmethod
    def changed_fields(old, new):
        """Champs de DIFF_FIELDS dont la valeur diffère entre deux versions d'une exigence (listes comparées en
        texte)."""
        value = lambda r, f: ", ".join(r.get(f) or []) if isinstance(r.get(f), list) else (r.get(f) or "")
        return [f for f in DIFF_FIELDS if value(old, f) != value(new, f)]

    @classmethod
    def merge(cls, current, requirements, mode, file_id=None):
        """Ce que fait l'import de `requirements` dans un tome, sans la base : `current` les exigences du tome (None
        : un nouveau tome). Remplacement : le tome devient exactement le fichier ; complément : les exigences du
        fichier remplacent celles de même ID, les autres sont ajoutées à la fin ; les exigences supprimées à la main
        sont abandonnées dans les deux cas (le fichier est la référence). Renvoie {"merged" : les exigences du tome
        après import, "added", "updated" (même ID, un champ changé : {numéro: [champs]}), "removed" : numéros}.
        Utilisé par l'import et par la vérification avant import."""
        new = {}
        # Chaque exigence validée ; un ID présent deux fois garde sa première ligne, comme à la lecture du fichier.
        for r in requirements:
            r = cls._clean(r, file_id)
            new.setdefault(r["id"], r)
        old = {r["id"]: r for r in current or () if not r.get("deleted")}
        if mode == "replace" or current is None:
            merged = list(new.values())
            removed = [i for i in old if i not in new]
        else:
            merged = [new.get(i, r) for i, r in old.items()] + [r for i, r in new.items() if i not in old]
            removed = []
        updated = {i: cls.changed_fields(old[i], r) for i, r in new.items() if i in old}
        return {"merged": merged, "added": [i for i in new if i not in old], "removed": removed,
                "updated": {i: f for i, f in updated.items() if f}}

    def edit(self, tome, action, old_id=None, requirement=None, config=None):
        """Modification à la main d'une exigence d'un tome (le prochain import du tome l'écrase).

        action « add » : nouvelle exigence (file_id None : « saisie manuelle »), à la fin du tome, marquée « added »
        ;
        action « update » : l'exigence old_id remplacée par `requirement` (son ID peut changer, son fichier est
        gardé), marquée « modified » (ou reste « added ») ;
        action « delete » : l'exigence old_id marquée supprimée (barrée, non vérifiée) — retirée pour de bon quand
        elle avait été ajoutée à la main ;
        action « restore » : la marque de suppression est retirée.
        Renvoie l'exigence enregistrée (None pour une suppression). ValueError quand le changement est refusé.
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
                    # Une exigence ajoutée à la main disparaît ; une exigence venue d'un fichier est seulement marquée
                    # supprimée, et peut être rétablie.
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
            req["releases"] = [self.sql.release] if self.sql.release else []   # la release de la base
            if existing is not None:
                # Un champ de la référence non envoyé par la page garde sa valeur.
                before = self._requirements("WHERE r.id = ?", (existing["id"],))[0]
                for field in [f for f, _, _ in REFERENCE] + ["validation"]:
                    if field not in (requirement or {}) or (field == "validation"
                                                            and requirement[field] not in VALIDATION_STATES):
                        req[field] = before[field]
            self.check_reference([{**req, "row": "saisie"}], config)
            if existing is None:
                req["file_id"] = None
            else:
                # Enregistrée sans aucun changement (releases dans n'importe quel ordre) : rien n'est écrit, pour qu'une
                # exigence importée ne devienne pas « modified ».
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
            # Les releases et le STDHL sont choisis parmi les valeurs existantes, pas saisis.
            declared = set((config.get if config is not None else DEFAULT_CONFIG.get)("releases") or {})
            unknown = [r for r in req["releases"] if r not in declared]
            if unknown:
                raise ValueError(f"{'releases inconnues' if len(unknown) > 1 else 'release inconnue'} : {', '.join(unknown)} "
                                 "(choisir parmi les releases de la configuration)")
            if req["stdhl"] and req["stdhl"] not in self.stdhl_names():
                raise ValueError(f"valeur STDHL inconnue : « {req['stdhl']} » (choisir parmi les valeurs existantes)")
            # Un ID nouveau ou renommé ne doit exister nulle part dans la base (un numéro, une exigence).
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
        """Nombre d'exigences venues d'un fichier, par tome : {tome: nombre} (seulement les tomes > 0)."""
        return {r["tome"]: r["n"] for r in self.sql.query(
            "SELECT t.name AS tome, COUNT(*) AS n FROM requirements r JOIN list_tomes t ON t.id = r.tome_id "
            "WHERE r.file_id = ? GROUP BY t.id", (file_id,))}

    def remove_file(self, file_id):
        """Retire les exigences venues d'un fichier. Renvoie {tome: nombre d'exigences retirées}."""
        return self.prune_files(lambda f: f != file_id)

    def prune_files(self, keep):
        """Retire les exigences dont le file_id ne satisfait pas keep(file_id) (les exigences sans file_id sont
        gardées). Renvoie {tome: nombre retiré}."""
        with self.sql.lock:
            # Compte par tome avant la suppression, pour les messages.
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
        """Retire les exigences d'un tome (le tome reste dans le référentiel). Renvoie False s'il n'y en avait
        aucune."""
        with self.sql.transaction() as db:
            return db.execute("DELETE FROM requirements WHERE tome_id = ?", (self.sql.tome_id(tome),)).rowcount > 0
