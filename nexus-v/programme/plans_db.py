"""Plans d'acceptation de NEXUS-V, dans la base SQL (tables files et plan_acceptation).

Quand un plan d'acceptation est chargé, son contenu utile (lu dans le document Word) est enregistré ici. Le plan est
repéré par le numéro de son fichier (table files) : son nom et son tome sont ceux du fichier ; il s'applique à la
release de la base ; retirer le fichier retire son contenu (cascade).

Tout renvoie à l'exigence : une ligne de plan est une exigence du tome du plan, avec ses codes IADT et MOC et sa
délégation (plusieurs lignes du document pour une exigence sont fusionnées). Le titre et la colonne Méthode ne sont
pas enregistrés. À l'import, un numéro inconnu du tome, une valeur vide ou inconnue de la configuration, un MOC hors
des règles IADT / MOC, une délégation vide ou plusieurs délégations sont réglés par l'utilisateur : valeur attendue
choisie, ou ligne ignorée (un numéro inconnu ne peut qu'être ignoré).

La classe réunit tout ce qui concerne les plans : lecture du document Word (read_docx : tableaux dont l'en-tête
correspond au réglage headers_plan), règles d'une ligne (check_row : IADT, MOC, délégation), vérification d'un tome
(check_tome : chaque exigence du tome doit être dans un plan de ce tome), chargement d'un fichier (analyse puis
load_file) et enregistrement.

Pour la page, les plans sont donnés par numéro de fichier (to_dict) :

    {"plans": {"12": {"file_id": 12, "name": "...docx", "tome": "fct1", "releases": ["A"],
                      "rows": [{"key": ID, "requirements": [ID], "iadt", "moc", "delegation", "manual",
                                "deleted"}, ...]}}}

Une ligne corrigée à l'import, ajoutée, modifiée ou supprimée à la main (edit) est marquée (manual = added, modified,
deleted ; deleted = 1 : barrée, sans effet). Quand le plan est relu depuis son document, seules les lignes non
marquées sont réécrites : une exigence qui a une ligne marquée la garde.
"""
import re

from configuration import DEFAULT_CONFIG, fold, header_column, header_text, match_header, plural, table_left_out
from file_database import FileDatabase, WordDocument
from nexus_db import value_id

# Colonnes d'un tableau de plan (leurs textes d'en-tête attendus sont dans la configuration : headers_plan).
PLAN_COLUMNS = ("exigence", "iadt", "moc", "delegation")


def natural_key(text):
    """Clé de tri qui met VS2 avant VS10."""
    return [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", str(text))]


def unique(values):
    """Valeurs non vides sans doublon, dans leur ordre."""
    return list(dict.fromkeys(v for v in values if v))


class PlansDatabase:
    """Contenu des plans d'acceptation, dans la base SQL."""

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def _write(db, file_id, entries):
        """Remplace les lignes d'un plan lues dans son document (dans une transaction), d'après ses entrées (voir
        entries). Les lignes marquées manual restent, et les exigences qui en ont une ne sont pas réécrites ; un
        numéro qui n'est pas une exigence du tome du plan n'a pas de ligne, ni une exigence active dans un autre plan
        du tome (une exigence n'est que dans un plan de son tome : l'autre a été gardé)."""
        # Les lignes lues dans le document sont réécrites ; les lignes marquées par une correction ou une modification à
        # la main restent.
        db.execute("DELETE FROM plan_acceptation WHERE file_id = ? AND manual IS NULL", (file_id,))
        # Exigences à ne pas écrire : celles qui ont encore une ligne marquée dans ce plan, et celles actives dans un
        # autre plan du même tome (une exigence n'est que dans un plan ; un plan de même nom est celui qui est
        # remplacé).
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
        """Lignes d'une exigence d'un plan : une par code IADT et MOC (dans une transaction)."""
        delegation = (delegation or "").strip()
        platforms = {fold(r[1]): r[0] for r in db.execute("SELECT id, name FROM list_plateformes")}
        # La plateforme telle qu'elle est écrite dans la configuration, quelle que soit la casse ; un nom qu'elle ne
        # connaît pas n'est possible qu'en ligne de commande (la page les fait régler d'abord).
        delegation_id = (platforms.get(fold(delegation)) or value_id(db, "list_plateformes", "name", delegation)) if delegation else None
        iadt_ids = [value_id(db, "list_iadt", "code", c) for c in PlansDatabase.split_codes(iadt)] or [None]
        moc_ids = [value_id(db, "list_moc", "code", c) for c in PlansDatabase.split_codes(moc)] or [None]
        # Une ligne par paire de codes : une cellule « A » avec le MOC « MOC2, MOC3 » donne deux lignes de la même
        # exigence.
        db.executemany("INSERT INTO plan_acceptation (requirement_id, iadt_id, moc_id, delegation_id, file_id, manual, "
                       "deleted) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       [(requirement_id, i, m, delegation_id, file_id, manual, deleted) for i in iadt_ids for m in moc_ids])

    @classmethod
    def entries(cls, rows):
        """Les lignes lues dans un document de plan, par exigence : une entrée par numéro d'exigence, avec ses codes
        IADT et MOC et ses délégations joints (plusieurs lignes d'une exigence sont fusionnées), et les emplacements
        de ses lignes."""
        found = {}
        for r in rows:
            for number in r["requirements"]:
                # Une exigence citée sur plusieurs lignes (ou une ligne qui cite plusieurs exigences) : ses codes et
                # délégations sont fusionnés, et ses emplacements (tableau, ligne) gardés pour les messages.
                e = found.setdefault(number, {"requirement": number, "iadt": [], "moc": [], "delegation": [], "places": []})
                e["iadt"] += cls.split_codes(r["iadt"])
                e["moc"] += cls.split_codes(r["moc"])
                if r["delegation"].strip():
                    e["delegation"].append(r["delegation"].strip())
                e["places"].append([r["table"], r["row"]])
        return [{**e, "iadt": ", ".join(unique(e["iadt"])), "moc": ", ".join(unique(e["moc"])),
                 "delegation": ", ".join(unique(e["delegation"]))} for e in found.values()]

    # ------------------------------------------------------------------ #
    # Accès
    # ------------------------------------------------------------------ #

    def plans(self, tome=None, file_id=None, config=None, with_deleted=False):
        """Plans avec leurs lignes, éventuellement ceux d'un tome ou d'un fichier : une ligne par exigence, ses codes
        IADT et MOC joints par « , ». Les lignes supprimées (à la main ou ignorées à l'import) sont laissées de côté
        sauf avec `with_deleted` (page). Avec une configuration, chaque ligne reçoit aussi le résultat de check_row
        (codes, délégation, erreurs)."""
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
            # Les lignes enregistrées d'une exigence dans un plan (une par paire de codes) sont regroupées en une ligne,
            # codes joints par « , ».
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
        # Avec une configuration, chaque ligne reçoit aussi sa délégation et ses erreurs.
        check = (lambda row: self.check_row(row, config)) if config is not None else (lambda row: row)
        return [{"file_id": h["file_id"], "name": h["name"], "tome": h["tome"], "releases": releases.get(h["file_id"], []),
                 "rows": [check(r) for r in rows.get(h["file_id"], [])]} for h in heads]

    def get(self, file_id):
        """Le plan d'un fichier avec ses lignes, ou None."""
        found = self.plans(file_id=file_id)
        return found[0] if found else None

    def to_dict(self, config=None):
        """Les plans par numéro de fichier, pour la page ; les lignes sont vérifiées quand une configuration est
        donnée."""
        plans = self.plans(config=config, with_deleted=True)
        return {"plans": {str(p["file_id"]): p for p in plans}}

    # ------------------------------------------------------------------ #
    # Lecture d'un document de plan
    # ------------------------------------------------------------------ #

    @staticmethod
    def _settings(config):
        """En-têtes attendus et format des numéros d'exigence, d'après une configuration (valeurs par défaut sans
        elle)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        return get("headers_plan"), re.compile(get("requirement_regex"))

    @staticmethod
    def _row(table, row, exigence, iadt, moc, delegation, regex, separator):
        """Forme enregistrée d'une ligne de plan : les numéros d'exigence de la cellule Exigence (pas le titre), les
        codes IADT et MOC (en majuscules, séparés par « , » quel que soit le séparateur du document) et la
        délégation."""
        # Des codes écrits « A/T », « A ; T » ou « a, t » sont tous enregistrés « A, T ».
        codes = lambda cell: ", ".join(c for c in re.split(separator, cell.upper()) if c)
        return {"table": table, "row": row, "requirements": unique(regex.findall(exigence)),
                "iadt": codes(iadt), "moc": codes(moc), "delegation": delegation.strip()}

    @staticmethod
    def _find_header(grid, headers):
        """({rôle: index de colonne}, [textes attendus manquants]) de l'en-tête d'un tableau, toujours sa première
        ligne, ou None. Un tableau de plan est reconnu à l'en-tête de sa colonne exigence et d'au moins une autre
        colonne lue (configuration headers_plan) ; chaque en-tête attendu doit alors être là."""
        exigence = header_text(headers, "exigence")
        others = [header_text(headers, role) for role in PLAN_COLUMNS if role != "exigence"]
        cells = grid[0] if grid else []
        # Un tableau est un tableau de plan quand sa première ligne a la colonne des exigences et au moins une autre
        # colonne lue ; chaque en-tête attendu doit alors être là, sinon le tableau est signalé et laissé de côté.
        if header_column(cells, exigence) is not None and any(header_column(cells, t) is not None for t in others):
            return match_header(cells, headers)
        return None

    @classmethod
    def read_docx(cls, content, config=None):
        """Lit un plan d'acceptation (octets .docx). Renvoie {"rows" : [lignes enregistrées], "ids" : numéros
        d'exigence cités n'importe où dans le document, "platforms" : plateformes nommées dans les délégations,
        "skipped" : tableaux de plan laissés de côté parce qu'il manque une colonne attendue à leur en-tête}. Un tel
        tableau n'arrête pas la lecture. Les numéros de tableau et de ligne commencent à 1, la ligne d'en-tête
        comprise."""
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
            # Chaque ligne après l'en-tête, numérotée comme dans Word (l'en-tête est la ligne 1).
            for r_index, cells in enumerate(grid[1:], 2):
                def cell(field):
                    """Texte de la cellule d'une colonne lue dans cette ligne, vide quand la ligne est trop courte."""
                    k = cols[field]
                    return cells[k] if k is not None and k < len(cells) else ""
                # Saute les lignes vides et les lignes d'en-tête répétées dans le tableau.
                if not any(cells) or (header_column([cell("exigence")], header_text(headers, "exigence")) is not None
                                      and header_column([cell("iadt")], header_text(headers, "iadt")) is not None):
                    continue
                rows.append(cls._row(t_index, r_index, *(cell(f) for f in PLAN_COLUMNS), regex, separator))
        platforms = unique([r["delegation"] for r in rows])
        # ids : chaque numéro cité n'importe où dans le document (corps, en-têtes, notes), pas seulement dans les
        # tableaux de plan.
        return {"rows": rows, "ids": unique(regex.findall(document.text())), "platforms": platforms,
                "skipped": skipped}

    @classmethod
    def _readable(cls, content, config=None):
        """Contenu d'un plan chargeable : refusé (ValueError) seulement quand aucun tableau de plan valide n'est
        trouvé ; les tableaux à l'en-tête incomplet sont sautés."""
        headers, _ = cls._settings(config)
        data = cls.read_docx(content, config)
        if not data["rows"]:
            raise ValueError(f"aucun tableau de plan trouvé (en-tête « {header_text(headers, 'exigence')} » et colonnes "
                             "stockées, voir Configuration > En-têtes des tableaux) : fichier non chargé")
        return data

    # ------------------------------------------------------------------ #
    # Règles
    # ------------------------------------------------------------------ #

    @staticmethod
    def split_codes(value):
        """Codes d'une valeur IADT ou MOC enregistrée."""
        return [c for c in (value or "").split(", ") if c]

    @classmethod
    def check_row(cls, row, config):
        """Une ligne de plan avec ses codes IADT et MOC, sa délégation et ses erreurs (règles de la configuration)."""
        errors = []
        iadt = cls.split_codes(row["iadt"])
        if not iadt:
            errors.append("IADT vide")
        # Erreurs par rapport à la configuration.
        errors += [f"IADT inconnu : « {c} »" for c in iadt if c not in config.get("iadt")]
        moc = cls.split_codes(row["moc"])
        if not moc:
            errors.append("MOC vide")
        errors += [f"MOC inconnu : « {c} »" for c in moc if c not in config.get("mocs")]
        # Une délégation est toujours écrite : il n'y a pas de plateforme par défaut.
        delegation = config.delegation(row["delegation"])
        if not delegation:
            errors.append("délégation vide")
        elif not config.is_platform(delegation):
            errors.append(f"plateforme de délégation inconnue : « {delegation} »")
        return {**row, "iadt_codes": iadt, "moc_codes": moc, "effective_delegation": delegation, "errors": errors}

    @classmethod
    def rule_issues(cls, row, config):
        """Règles d'import d'une ligne de plan (configuration iadt_moc_rules) : pour chaque IADT qui a une règle, au
        moins un de ses MOC ; aucun MOC hors des règles de l'IADT de la ligne. Messages des règles non respectées."""
        rules = config.get("iadt_moc_rules")
        iadt, moc = cls.split_codes(row["iadt"]), cls.split_codes(row["moc"])
        # Chaque IADT qui a une règle demande au moins un de ses MOC ; et quand chaque IADT de la ligne a une règle,
        # aucun MOC hors de ces règles n'est admis.
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
        """Plateformes nommées dans la cellule de délégation d'une ligne (séparateur delegation_separator_regex),
        sans doublon."""
        values = {}
        for v in re.split(config.get("delegation_separator_regex"), row["delegation"] or ""):
            if v.strip():
                values.setdefault(fold(v), v.strip())
        return list(values.values())

    @staticmethod
    def expected_mocs(iadt_cell, config):
        """Valeurs MOC que peut prendre une ligne avec ces codes IADT, d'après les règles (iadt_moc_rules) : chaque
        MOC attendu, et tous ensemble quand il y en a plusieurs ; chaque MOC de la configuration quand un IADT n'a
        pas de règle."""
        rules = config.get("iadt_moc_rules")
        iadt = PlansDatabase.split_codes(iadt_cell)
        if not iadt or any(i not in rules for i in iadt):
            return list(config.get("mocs"))
        allowed = unique(m for i in iadt for m in rules[i])
        # Les choix proposés : chaque MOC attendu seul, puis tous ensemble (« MOC2, MOC3 »).
        return allowed + ([", ".join(allowed)] if len(allowed) > 1 else [])

    @classmethod
    def to_confirm(cls, entries, config, known=None, elsewhere=None, releases=None, applies=None):
        """Entrées d'un plan (voir entries) que l'utilisateur doit régler à l'import, chacune avec sa clé (le numéro
        d'exigence), ses problèmes et « choose » : les cellules à remplacer par une valeur de la configuration (iadt,
        moc, delegation). Rien n'est importé tel quel : un IADT ou un MOC vide ou inconnu, un MOC qui ne respecte pas
        les règles IADT/MOC (le MOC attendu est proposé, voir expected_mocs), une délégation vide, plusieurs
        délégations ou une délégation inconnue sont remplacés, ou l'entrée est ignorée. Avec `known` (numéros des
        exigences du tome), un numéro qui n'en fait pas partie ne peut qu'être ignoré. Avec `elsewhere` (voir
        PlansDatabase.elsewhere), une exigence déjà dans un autre plan de la release est réglée aussi : l'autre
        gardée (cette entrée n'est pas enregistrée), ou celle-ci prise (la ligne de l'autre plan est retirée). Avec
        les `releases` du plan et `applies` ({numéro: releases de l'exigence}), une exigence qui ne s'applique à
        aucune d'elles est une erreur d'import : elle ne peut qu'être ignorée."""
        wanted = {r.strip().upper() for r in releases or () if r.strip()}
        elsewhere = elsewhere or {}
        found = []
        for r in entries:
            base = {"key": r["requirement"], "requirements": [r["requirement"]], "places": r["places"], "iadt": r["iadt"],
                    "moc": r["moc"], "delegation": r["delegation"]}
            # Un numéro qui n'est pas une exigence du tome ne peut qu'être ignoré.
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
            # Puis, cellule par cellule, ce que l'utilisateur doit choisir : l'IADT (et son MOC avec lui), le MOC, la
            # délégation.
            if r["requirement"] in elsewhere:
                issues.append(f"déjà dans un autre plan de la release : {', '.join(elsewhere[r['requirement']])} — "
                              "garder l'une des deux lignes")
            if not iadt or any(c not in config.get("iadt") for c in iadt):
                issues.append("IADT vide" if not iadt else
                              "IADT inconnu : " + ", ".join(c for c in iadt if c not in config.get("iadt")))
                choose += ["iadt", "moc"]   # le MOC attendu dépend de l'IADT choisi
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
            if not delegations:
                issues.append("délégation vide : en choisir une")
                choose.append("delegation")
            elif len(delegations) > 1:
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
        """Choix de l'utilisateur pour les entrées à régler, sous la forme {numéro d'exigence: (ignorée, iadt, moc,
        délégation)}. `decisions` : {"rows": {numéro: {"ignore": true} | {"iadt": code, "moc": codes, "delegation":
        plateforme}}} — une valeur de la configuration pour chaque cellule à choisir, l'entrée corrigée respectant
        les règles IADT/MOC ; un numéro inconnu ne peut qu'être ignoré —, ou {"accept_all": true} (ligne de commande
        : entrées chargées telles quelles, numéros inconnus ignorés). ValueError quand une entrée n'est pas réglée."""
        decisions = decisions or {}
        chosen, result, left = decisions.get("rows") or {}, {}, 0
        for c in found:
            d = chosen.get(c["key"]) or {}
            # (1, None, None, None) = ignorée : la ligne sera enregistrée barrée (manual = deleted).
            if d.get("ignore") or (c["unknown"] and decisions.get("accept_all")):
                result[c["key"]] = (1, None, None, None)
                continue
            values = {f: str(d.get(f) or "").strip() for f in c["choose"]}
            row = {"iadt": values.get("iadt", c["iadt"]), "moc": values.get("moc", c["moc"])}
            iadt, moc = cls.split_codes(row["iadt"]), cls.split_codes(row["moc"])
            # Chaque cellule à choisir a une valeur ; l'IADT et le MOC corrigés sont connus et respectent les règles.
            valid = not c["unknown"] and all(values.values())
            valid = valid and bool(iadt) and all(x in config.get("iadt") for x in iadt)
            valid = valid and bool(moc) and all(x in config.get("mocs") for x in moc)
            valid = valid and not cls.rule_issues(row, config)
            valid = valid and ("delegation" not in values or config.is_platform(values["delegation"]))
            if valid:
                # (0, iadt, moc, délégation) = corrigée : None pour une cellule que l'utilisateur n'avait pas à choisir.
                if "delegation" in values:   # le nom tel qu'il est écrit dans la configuration
                    values["delegation"] = next(p for p in config.get("platforms") if fold(p) == fold(values["delegation"]))
                result[c["key"]] = (0, values.get("iadt"), values.get("moc"), values.get("delegation"))
            elif not decisions.get("accept_all"):
                left += 1
        if left:
            raise ValueError(f"{plural(left, 'ligne')} du plan à régler avant le chargement : choisir les valeurs "
                             "attendues ou ignorer la ligne")
        return result

    def edit(self, file_id, action, requirement=None, values=None, config=None):
        """Modification à la main de la ligne d'une exigence dans un plan. `values` : {"iadt": code, "moc": codes,
        "delegation": plateforme}, pris dans la configuration et respectant les règles IADT/MOC.
        action « add » : ligne d'une exigence du tome pas encore dans le plan (manual = added) ; « update » : valeurs
        remplacées (manual = modified, ou reste added ; une ligne supprimée est rétablie en même temps ; rien n'est
        écrit quand elles ne changent pas) ; « delete » : ligne barrée et ignorée (manual = deleted sauf si elle a
        déjà été modifiée), retirée pour de bon quand elle avait été ajoutée à la main ; « restore » : suppression
        annulée, quand la ligne respecte les règles.
        ValueError quand le changement est refusé."""
        if action not in ("add", "update", "delete", "restore"):
            raise ValueError(f"action inconnue : {action}")
        with self.sql.lock:
            # L'exigence doit appartenir au tome du plan.
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
            # Ajout ou retour d'une ligne : l'exigence ne doit pas être active dans un autre plan du tome.
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
            # La ligne telle que la page la montre (codes joints), pour la comparer aux nouvelles valeurs.
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
            # Ajout ou mise à jour : les valeurs doivent venir de la configuration et respecter les règles IADT/MOC,
            # comme à l'import.
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
                return current   # enregistrée sans aucun changement : rien n'est écrit, la ligne garde sa marque
            with self.sql.transaction() as db:
                # Les lignes de l'exigence sont réécrites (une par paire de codes), marquées added ou modified.
                db.execute("DELETE FROM plan_acceptation WHERE file_id = ? AND requirement_id = ?", (file_id, requirement_id))
                self._insert(db, file_id, requirement_id, iadt, ", ".join(codes), delegation,
                             "added" if action == "add" or manual == "added" else "modified")
            return {"key": requirement}

    def check_tome(self, tome, requirements, config, selected=()):
        """Vérification des plans d'un tome par rapport à ses exigences (RequirementsDatabase `requirements`),
        release par release : une exigence doit se trouver dans un plan du tome qui s'applique à l'une de ses
        releases. `selected` : releases choisies dans l'interface (aucune = toutes). Les exigences supprimées à la
        main sont ignorées."""
        selected = {s.upper() for s in selected}
        in_selection = lambda releases: not selected or bool(selected & set(releases))
        current = requirements.get(tome)
        reqs = [r for r in (current["requirements"] if current else [])
                if not r["deleted"] and in_selection(r["releases"])]
        plans = [p for p in self.plans(tome=tome, config=config) if in_selection(p["releases"])]
        by_id = {r["id"]: r for r in reqs}
        # Couverte : dans un plan du tome ; anomalies : lignes avec des erreurs de configuration (montrées dans la
        # vérification du tome).
        covered, misplaced, anomalies = set(), {}, []
        for plan in plans:
            for row in plan["rows"]:
                for req_id in row["requirements"]:
                    r = by_id.get(req_id)
                    if r is None:   # exigence supprimée à la main, ou hors des releases sélectionnées
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
        # Par release : les plans qui s'y appliquent, les exigences attendues et combien manquent.
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
    # Chargement d'un fichier de plan
    # ------------------------------------------------------------------ #

    def analyse(self, content, config=None, tome=None, releases=(), name=None):
        """Ce que ferait le chargement d'un fichier de plan : le nombre d'entrées (exigences) lues, les plateformes
        nommées, les tableaux sautés et, avec une configuration, les plateformes nommées seules dans une cellule
        Délégation et qu'elle ne connaît pas (proposées à l'utilisateur pour ajout avant le règlement des lignes) et
        les entrées à régler (avec le `tome` : les numéros qu'il ne connaît pas aussi).
        Refusé (ValueError) quand le document est illisible ou n'a aucun tableau de plan valide."""
        data = self._readable(content, config)
        entries = self.entries(data["rows"])
        # Le tome nommé dans le nom du fichier (configuration plan_tome_regex), quand c'est un tome de la base :
        # présélectionné.
        tome_name = ""
        if config is not None and name:
            match = re.search(config.get("plan_tome_regex"), name)
            found = fold(match.group(match.lastindex or 0)) if match else ""
            tome_name = next((r["name"] for r in self.sql.query("SELECT name FROM list_tomes") if fold(r["name"]) == found), "")
        # Plateformes nommées seules dans une cellule Délégation et inconnues de la configuration : la page propose de
        # les ajouter avant le règlement des lignes (une cellule à plusieurs plateformes est réglée comme une ligne).
        unknown = unique([d[0] for e in entries for d in [self.delegation_values(e, config)]
                          if len(d) == 1 and not config.is_platform(d[0])]) if config is not None else []
        return {"rows": len(entries), "platforms": data["platforms"], "skipped": data["skipped"],
                "unknown_platforms": unknown, "tome_name": tome_name,
                "to_confirm": self.to_confirm(entries, config, self._known(tome), self.elsewhere(tome, exclude_name=name),
                                              releases, self._releases_of(tome) if tome else None)
                if config is not None else []}

    def elsewhere(self, tome, exclude_file=None, exclude_name=None):
        """{numéro d'exigence: [noms de plans]} : les exigences présentes (non supprimées) dans les autres plans du
        tome (une release pour la base) — une exigence n'est que dans un plan. Le plan `exclude_file` et les plans
        nommés `exclude_name` (remplacés par un nouveau chargement) ne comptent pas."""
        if not tome:
            return {}
        found = {}
        # Chaque exigence de chaque autre plan du tome, avec les noms des plans qui la contiennent.
        for p in self.plans(tome=tome):   # une release pour la base : chaque plan du tome
            if p["file_id"] == exclude_file or p["name"] == exclude_name:
                continue
            for row in p["rows"]:
                found.setdefault(row["key"], []).append(p["name"])
        return found

    def _releases_of(self, tome):
        """{numéro d'exigence: ensemble de releases} des exigences d'un tome : la release de la base pour toutes."""
        release = {self.sql.release} if self.sql.release else set()
        return {r["req_id"]: set(release) for r in self.sql.query(
            "SELECT q.req_id FROM requirements q JOIN list_tomes t ON t.id = q.tome_id WHERE t.name = ?", (tome,))}

    def _known(self, tome):
        """Numéros des exigences d'un tome (None sans tome)."""
        if not tome:
            return None
        tome_id = self.sql.tome_id(tome)
        return set(FileDatabase.tome_requirements(self.sql.db, tome_id)) if tome_id is not None else set()

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None):
        """Charge un plan d'acceptation : rattaché à un tome qui a des exigences et à des releases (au moins une). Le
        fichier est enregistré (FileDatabase `files`) puis son contenu ; si l'enregistrement du contenu échoue, le
        fichier est retiré. Un plan de même nom déjà chargé dans le tome est remplacé. Les entrées à régler
        (to_confirm) demandent une décision de l'utilisateur (voir _decide) : corrigée (ligne marquée modified), ou
        ignorée (ligne barrée ; rien pour un numéro inconnu)."""
        files.check_attachment("Plan d'acceptation", requirements, tome)
        data = self._readable(content, config)
        entries = self.entries(data["rows"])
        found = self.to_confirm(entries, config, self._known(tome), self.elsewhere(tome, exclude_name=name),
                                releases, self._releases_of(tome)) if config is not None else []
        # Numéros inconnus et exigences d'autres releases : ignorés, rien n'est enregistré.
        # Les décisions sont vérifiées avant toute écriture : une ligne non réglée refuse le chargement.
        only_ignore = {c["key"] for c in found if c["unknown"]}
        decided = self._decide(found, decisions, config) if config is not None else {}
        # Exigences prises dans ce fichier alors qu'elles sont dans un autre plan de la release : l'autre ligne est
        # barrée.
        taken = {c["key"]: c["duplicate"] for c in found if c.get("duplicate") and not decided.get(c["key"], (1,))[0]}
        duplicates = {c["key"] for c in found if c.get("duplicate")}
        by_number = {e["requirement"]: e for e in entries}
        record = files.add("plan", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                known = FileDatabase.tome_requirements(db, db.execute("SELECT tome_id FROM files WHERE id = ?",
                                                                      (record["id"],)).fetchone()[0])
                # Exigences prises dans ce fichier alors qu'elles sont dans un autre plan de la release : la ligne de
                # l'autre plan est retirée — une seule ligne pour une exigence et une release (avant l'écriture des
                # lignes de ce plan).
                for number, names in taken.items():
                    if number in known:
                        db.execute("DELETE FROM plan_acceptation "
                                   "WHERE requirement_id = ? AND file_id IN (SELECT id FROM files WHERE source = 'plan' "
                                   f"AND name IN ({','.join('?' * len(names))}) AND id != ?)",
                                   [known[number], *names, record["id"]])
                # Entrées réglées : lignes corrigées (modified) ou ignorées (deleted), gardées quand le plan est relu ;
                # une exigence gardée dans l'autre plan n'est pas du tout enregistrée ici.
                for number, (ignored, iadt, moc, delegation) in decided.items():
                    if not ignored and not (iadt or moc or delegation):
                        continue   # prise telle quelle : écrite avec les lignes du document
                    if ignored and (number in duplicates or number in only_ignore):
                        continue   # l'autre plan la garde, ou ce n'est pas une exigence de ces releases
                    if number in known:
                        e = by_number[number]
                        self._insert(db, record["id"], known[number], iadt or e["iadt"], moc or e["moc"],
                                     delegation or e["delegation"], "deleted" if ignored else "modified", ignored)
                self._write(db, record["id"], entries)
        except Exception:
            files.remove(record["id"])   # contenu non enregistré : le fichier n'est pas gardé
            raise
        # Résumé pour la page : lignes enregistrées et ignorées, plateformes utilisées, tableaux laissés de côté,
        # avertissements STDHL.
        ignored = sum(values[0] for values in decided.values())
        stdhl_warnings = self.stdhl_requests(record["id"], config) if config is not None else []
        # Plateformes des lignes enregistrées, avec les délégations choisies.
        platforms = unique([(decided.get(e["requirement"]) or (0, None, None, None))[3] or e["delegation"]
                            for e in entries if not (decided.get(e["requirement"]) or (0,))[0]])
        return {"file_id": record["id"], "tome": tome, "rows": len(entries) - ignored, "ignored": ignored,
                "platforms": platforms, "skipped": data["skipped"], "replaced": files.remove_older(record),
                "stdhl_warnings": stdhl_warnings}

    def stdhl_requests(self, file_id, config):
        """Lignes d'un plan déléguées à une autre plateforme que celle de niveau système alors que la colonne STDHL
        de l'exigence demande un STDHL (seules les exigences déléguées à la plateforme de niveau système sont dans un
        STDHL ; la colonne est facultative) : messages pour la page."""
        plan = self.get(file_id)
        if plan is None:
            return []
        # Valeur de la colonne STDHL de chaque exigence du tome du plan.
        default = config.get("system_platform")   # la plateforme de niveau système, à laquelle s'appliquent les STDHL
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
        """Lit et enregistre le contenu des fichiers de plan qui n'en ont pas encore, et de tous avec `again` : les
        lignes d'un plan pointent vers les exigences, elles sont donc relues quand les exigences changent. Renvoie
        les noms lus ; les fichiers illisibles sont laissés tels quels."""
        done = []
        for record in files.list(slot="plan"):
            # Les lignes marquées (corrections, modifications à la main) restent : _write ne réécrit que les lignes lues
            # dans le document.
            if again or not (self.get(record["id"]) or {}).get("rows"):   # aucune ligne encore enregistrée
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
    # Retrait
    # ------------------------------------------------------------------ #

    def remove(self, file_id):
        """Retire le contenu du plan d'un fichier. Renvoie False s'il est inconnu."""
        with self.sql.transaction() as db:
            return db.execute("DELETE FROM plan_acceptation WHERE file_id = ?", (file_id,)).rowcount > 0

    def prune_files(self, keep):
        """Retire les plans dont le file_id ne satisfait pas keep(file_id). Renvoie les noms retirés."""
        gone = [p for p in self.sql.query("SELECT id, name FROM files WHERE source = 'plan'")
                if not keep(p["id"])]
        for p in gone:
            self.remove(p["id"])
        return [p["name"] for p in gone]
