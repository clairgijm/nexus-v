"""Plans d'acceptation de NEXUS-V : vérifiés par rapport aux exigences, sans contenu enregistré (table files).

La référence de chaque exigence (IADT, MOC, délégation, validation) est le fichier d'exigences. Un plan d'acceptation
est seulement vérifié : chaque exigence de son tome doit y être, avec les mêmes valeurs. Toute différence refuse le
plan (Bloquant) : une exigence du tome absente du plan, un numéro qui n'est pas une exigence du tome, un IADT, un MOC,
une délégation ou une validation différents de ceux de l'exigence (plusieurs valeurs dans une cellule comprises).

Un plan accepté garde son fichier et le rapport de son import (FileDatabase) ; il remplace les autres plans de son
tome. Il ne garde aucune ligne : le tome est alors au moins à l'étape « plan d'acceptation vérifié ». Quand les
exigences du tome changent (import, modification à la main, fichier retiré), ses plans sont marqués à recharger
(files.stale) : le tome revient à l'étape des exigences jusqu'au chargement d'un plan cohérent.

Lecture du document Word (read_docx) : les tableaux dont l'en-tête correspond au réglage headers_plan (Exigence,
IADT, MOC, Délégation, Validation) ; plusieurs lignes pour une exigence sont réunies (entries). La cellule Validation
prend l'une des valeurs réglées dans la configuration (plan_validation_to_qualify, plan_validation_qualified,
plan_validation_out_of_scope : à qualifier, déjà qualifiée, hors scope), casse et accents ignorés.

La comparaison (differences) sert à l'import (load_file) et à la vérification avant import (FileCheck) : le même
code, pour que les deux ne divergent pas.
"""
import re
from typing import Any

from configuration import DEFAULT_CONFIG, fold, header_column, header_text, match_header, plural, table_left_out
from file_database import FileDatabase, WordDocument

# Champs d'un tableau de plan (leurs textes d'en-tête attendus sont dans la configuration : headers_plan).
PLAN_COLUMNS = ("exigence", "iadt", "moc", "delegation", "validation")
# Champs comparés à la référence de l'exigence, avec leur libellé dans les messages.
COMPARED = (("iadt", "IADT"), ("moc", "MOC"), ("delegation", "délégation"), ("validation", "validation"))
# Validation d'une exigence -> réglage de la configuration qui donne son texte dans un plan.
VALIDATION_SETTINGS = {"to_qualify": "plan_validation_to_qualify", "qualified": "plan_validation_qualified",
                       "out_of_scope": "plan_validation_out_of_scope"}
LIMIT = 30   # numéros cités dans le message d'un refus, au-delà : « … et N autres »


def natural_key(text):
    """Clé de tri qui met VS2 avant VS10."""
    return [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", str(text))]


def unique(values):
    """Valeurs non vides sans doublon, dans leur ordre."""
    return list(dict.fromkeys(v for v in values if v))


class PlansDatabase:
    """Plans d'acceptation : lecture, comparaison avec les exigences, fichiers chargés."""

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    # ------------------------------------------------------------------ #
    # Accès
    # ------------------------------------------------------------------ #

    def plans(self, tome=None, file_id=None, config=None, with_deleted=False) -> list[dict[str, Any]]:
        """Plans chargés, éventuellement ceux d'un tome ou d'un fichier : [{file_id, name, tome, releases, stale}] ;
        `stale` : les exigences du tome ont changé depuis le chargement (plan à recharger)."""
        conditions, params = ["f.source = 'plan'"], []
        if tome is not None:
            conditions.append("tm.name = ?")
            params.append(tome)
        if file_id is not None:
            conditions.append("f.id = ?")
            params.append(file_id)
        releases = self.sql.file_releases()
        return [{"file_id": r["file_id"], "name": r["name"], "tome": r["tome"], "stale": bool(r["stale"]),
                 "releases": releases.get(r["file_id"], [])}
                for r in self.sql.query(f"""SELECT f.id AS file_id, f.name, f.stale, tm.name AS tome FROM files f
                                            LEFT JOIN list_tomes tm ON tm.id = f.tome_id
                                            WHERE {" AND ".join(conditions)} ORDER BY f.id""", params)]

    def get(self, file_id):
        """Le plan d'un fichier, ou None."""
        found = self.plans(file_id=file_id)
        return found[0] if found else None

    def to_dict(self, config=None) -> dict[str, Any]:
        """Les plans par numéro de fichier, pour la page."""
        return {"plans": {str(p["file_id"]): p for p in self.plans()}}

    def mark_stale(self, tomes):
        """Les plans des tomes donnés sont à recharger : leurs exigences ont changé."""
        tomes = [t for t in tomes if t]
        if tomes:
            with self.sql.transaction() as db:
                db.execute(f"""UPDATE files SET stale = 1 WHERE source = 'plan' AND tome_id IN
                               (SELECT id FROM list_tomes WHERE name IN ({",".join("?" * len(tomes))}))""", tomes)

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
    def _row(table, row, cells, regex, separator):
        """Une ligne de plan : les numéros d'exigence de la cellule Exigence (pas le titre), les codes IADT et MOC
        (en majuscules, séparés par « , » quel que soit le séparateur du document), la délégation et la validation
        telles qu'écrites."""
        codes = lambda cell: ", ".join(c for c in re.split(separator, cell.upper()) if c)
        return {"table": table, "row": row, "requirements": unique(regex.findall(cells["exigence"])),
                "iadt": codes(cells["iadt"]), "moc": codes(cells["moc"]), "delegation": cells["delegation"].strip(),
                "validation": cells["validation"].strip()}

    @staticmethod
    def _find_header(grid, headers) -> tuple[dict[str, int], list[str]] | None:
        """({champ: index de colonne}, [textes attendus manquants]) de l'en-tête d'un tableau, toujours sa première
        ligne, ou None. Un tableau de plan est reconnu à l'en-tête de sa colonne exigence et d'au moins une autre
        colonne lue (configuration headers_plan) ; chaque en-tête attendu doit alors être là."""
        exigence = header_text(headers, "exigence")
        others = [header_text(headers, field) for field in PLAN_COLUMNS if field != "exigence"]
        cells = grid[0] if grid else []
        if header_column(cells, exigence) is not None and any(header_column(cells, t) is not None for t in others):
            return match_header(cells, headers)
        return None

    @classmethod
    def read_docx(cls, content, config=None) -> dict[str, Any]:
        """Lit un plan d'acceptation (octets .docx). Renvoie {"rows" : [lignes], "platforms" : plateformes nommées
        dans les délégations, "skipped" : tableaux de plan laissés de côté parce qu'il manque une colonne attendue à
        leur en-tête}. Un tel tableau n'arrête pas la lecture. Les numéros de tableau et de ligne commencent à 1, la
        ligne d'en-tête comprise."""
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
            for r_index, line in enumerate(grid[1:], 2):
                cells = {f: (line[cols[f]] if cols.get(f) is not None and cols[f] < len(line) else "")
                         for f in PLAN_COLUMNS}
                # Saute les lignes vides et les lignes d'en-tête répétées dans le tableau.
                if not any(line) or (header_column([cells["exigence"]], header_text(headers, "exigence")) is not None
                                     and header_column([cells["iadt"]], header_text(headers, "iadt")) is not None):
                    continue
                rows.append(cls._row(t_index, r_index, cells, regex, separator))
        return {"rows": rows, "platforms": unique([r["delegation"] for r in rows]), "skipped": skipped}

    @classmethod
    def readable(cls, content, config=None) -> dict[str, Any]:
        """Contenu d'un plan lisible : refusé (ValueError) seulement quand aucun tableau de plan valide n'est trouvé ;
        les tableaux à l'en-tête incomplet sont sautés."""
        headers, _ = cls._settings(config)
        data = cls.read_docx(content, config)
        if not data["rows"]:
            raise ValueError(f"aucun tableau de plan trouvé (en-têtes « {header_text(headers, 'exigence')} », IADT, "
                             "MOC, Délégation et Validation, voir Configuration > Plan d'acceptation)")
        return data

    @staticmethod
    def split_codes(value):
        """Codes d'une valeur IADT ou MOC lue."""
        return [c for c in (value or "").split(", ") if c]

    @classmethod
    def entries(cls, rows):
        """Les lignes d'un document de plan, par exigence : une entrée par numéro, ses valeurs réunies quand elle est
        sur plusieurs lignes (ou qu'une ligne en cite plusieurs), et les emplacements de ses lignes."""
        found = {}
        for r in rows:
            for number in r["requirements"]:
                e = found.setdefault(number, {"requirement": number, "iadt": [], "moc": [], "delegation": [],
                                              "validation": [], "places": []})
                e["iadt"] += cls.split_codes(r["iadt"])
                e["moc"] += cls.split_codes(r["moc"])
                for field in ("delegation", "validation"):
                    if r[field]:
                        e[field].append(r[field])
                e["places"].append([r["table"], r["row"]])
        return [{**e, **{f: unique(e[f] if f in ("iadt", "moc") else [v for v in e[f]]) for f in
                         ("iadt", "moc", "delegation", "validation")}} for e in found.values()]

    # ------------------------------------------------------------------ #
    # Comparaison avec les exigences
    # ------------------------------------------------------------------ #

    @staticmethod
    def validation_of(text, config):
        """Validation (to_qualify, qualified, out_of_scope) écrite dans une cellule Validation d'un plan, d'après les
        valeurs de la configuration (casse et accents ignorés) ; None pour une valeur inconnue."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        return next((state for state, name in VALIDATION_SETTINGS.items() if fold(get(name)) == fold(text)), None)

    @classmethod
    def differences(cls, entries, requirements: list[dict[str, Any]], config) -> list[dict[str, Any]]:
        """Différences entre les entrées d'un plan (entries) et les exigences de son tome ([{id, iadt, moc,
        delegation, validation}], celles supprimées à la main laissées de côté) : [{"kind", "requirement",
        "field", "plan", "expected", "places"}]. kind « unknown » : numéro qui n'est pas une exigence du tome ;
        « missing » : exigence du tome absente du plan ; « field » : un champ différent (plusieurs valeurs comprises).
        Chaque différence refuse le plan."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        reference = {r["id"]: r for r in requirements}
        found, cited = [], set()
        for e in entries:
            number = e["requirement"]
            cited.add(number)
            r = reference.get(number)
            if r is None:
                found.append({"kind": "unknown", "requirement": number, "field": "", "plan": "", "expected": "",
                              "places": e["places"]})
                continue
            for field, _ in COMPARED:
                values = e[field]
                if field == "validation":
                    states = [cls.validation_of(v, config) for v in values]
                    same = len(values) == 1 and states[0] == r["validation"]
                    expected = get(VALIDATION_SETTINGS.get(r["validation"], "")) if r["validation"] in VALIDATION_SETTINGS else ""
                elif field == "delegation":
                    same = len(values) == 1 and fold(values[0]) == fold(r["delegation"])
                    expected = r["delegation"]
                else:
                    same = len(values) == 1 and values[0] == (r[field] or "").upper()
                    expected = r[field]
                if not same:
                    found.append({"kind": "field", "requirement": number, "field": field,
                                  "plan": ", ".join(values) or "vide", "expected": expected or "vide",
                                  "places": e["places"]})
        for r in requirements:
            if r["id"] not in cited:
                found.append({"kind": "missing", "requirement": r["id"], "field": "", "plan": "", "expected": "",
                              "places": []})
        return found

    @staticmethod
    def difference_text(d):
        """Une différence en clair : « SSS_X_001 : MOC MOC2 dans le plan, MOC3 attendu »."""
        if d["kind"] == "unknown":
            return f"{d['requirement']} : pas une exigence du tome"
        if d["kind"] == "missing":
            return f"{d['requirement']} : exigence du tome absente du plan"
        label = dict(COMPARED)[d["field"]]
        return f"{d['requirement']} : {label} « {d['plan']} » dans le plan, « {d['expected']} » attendu"

    def tome_requirements(self, tome) -> list[dict[str, Any]]:
        """Exigences (non supprimées) d'un tome, avec leur référence : [{id, iadt, moc, delegation, validation}]."""
        return [{"id": r["req_id"], "iadt": r["iadt"] or "", "moc": r["moc"] or "", "delegation": r["delegation"] or "",
                 "validation": r["validation"]} for r in self.sql.query(
            """SELECT q.req_id, i.code AS iadt, m.code AS moc, p.name AS delegation, q.validation
               FROM requirements q JOIN list_tomes t ON t.id = q.tome_id
               LEFT JOIN list_iadt i ON i.id = q.iadt_id LEFT JOIN list_moc m ON m.id = q.moc_id
               LEFT JOIN list_plateformes p ON p.id = q.delegation_id
               WHERE t.name = ? AND q.deleted = 0 ORDER BY q.id""", (tome,))]

    # ------------------------------------------------------------------ #
    # Chargement d'un fichier de plan
    # ------------------------------------------------------------------ #

    def tome_from_name(self, name, config):
        """Le tome nommé dans le nom du fichier (configuration plan_tome_regex), quand c'est un tome de la base."""
        if config is None or not name:
            return ""
        match = re.search(config.get("plan_tome_regex"), name)
        found = fold(match.group(match.lastindex or 0)) if match else ""
        return next((r["name"] for r in self.sql.query("SELECT name FROM list_tomes") if fold(r["name"]) == found), "")

    def analyse(self, content, config=None, tome=None, releases=(), name=None):
        """Ce que ferait le chargement d'un fichier de plan : le nombre d'exigences lues, les plateformes nommées, les
        tableaux sautés, le tome nommé dans le nom du fichier et, avec un `tome`, les différences avec ses exigences.
        Refusé (ValueError) quand le document est illisible ou n'a aucun tableau de plan valide."""
        data = self.readable(content, config)
        entries = self.entries(data["rows"])
        diffs = self.differences(entries, self.tome_requirements(tome), config) if tome else []
        return {"rows": len(entries), "platforms": data["platforms"], "skipped": data["skipped"],
                "tome_name": self.tome_from_name(name, config), "differences": [self.difference_text(d) for d in diffs],
                "unknown_platforms": [], "to_confirm": []}

    @classmethod
    def refusal(cls, diffs):
        """Message du refus d'un plan qui diffère des exigences de son tome."""
        texts = [cls.difference_text(d) for d in diffs]
        more = f" … et {len(texts) - LIMIT} autres" if len(texts) > LIMIT else ""
        return (f"{plural(len(diffs), 'différence')} avec les exigences du tome (voir le rapport) : "
                + " ; ".join(texts[:LIMIT]) + more)

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None):
        """Charge un plan d'acceptation rattaché à un tome qui a des exigences : il est comparé aux exigences du tome
        (differences) et refusé (ValueError) à la moindre différence. Accepté, son fichier est enregistré
        (FileDatabase `files`) et remplace les autres plans du tome ; rien d'autre n'est gardé."""
        files.check_attachment("Plan d'acceptation", requirements, tome)
        data = self.readable(content, config)
        entries = self.entries(data["rows"])
        diffs = self.differences(entries, self.tome_requirements(tome), config)
        # Avec skip_unknown_req, un numéro qui n'est pas une exigence du tome ne refuse pas le plan : il est ignoré.
        ignored = [d["requirement"] for d in diffs if d["kind"] == "unknown"] if FileDatabase.skip_unknown(config) else []
        diffs = [d for d in diffs if not (ignored and d["kind"] == "unknown")]
        if diffs:
            raise ValueError(self.refusal(diffs))
        record = files.add("plan", name, content, tome=tome, releases=releases)
        # Le nouveau plan remplace tous les autres plans du tome (de même nom ou non).
        replaced = [p["file_id"] for p in self.plans(tome=tome) if p["file_id"] != record["id"]]
        for file_id in replaced:
            files.remove(file_id)
        return {"file_id": record["id"], "tome": tome, "rows": len(entries), "ignored": len(ignored),
                "platforms": data["platforms"], "skipped": data["skipped"], "replaced": replaced,
                "stdhl_warnings": []}
