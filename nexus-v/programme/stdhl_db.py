"""STDHL de NEXUS-V, dans la base SQL (tables files, stdhl_exi et fe_tests).

Quand un STDHL est chargé, son contenu utile (lu dans le document Word) est enregistré ici. Il est repéré par le
numéro de son fichier (table files) ; retirer le fichier retire son contenu (cascade). Au chargement, l'utilisateur
rattache le STDHL à un tome, ou à aucun : il s'applique alors à tous les tomes (un seul STDHL pour toute la
spécification).

Le nom du STDHL vient du nom de son fichier : [date]_[texte]_STDHL_[nom du STDHL]_rev[texte].docx. Les fiches (tests)
sont les paragraphes SYS_[type]_[fonctions]_[N°] ; chacune est suivie de tableaux dont l'en-tête correspond au
réglage headers_stdhl. Seuls les numéros d'exigence de la colonne « Exigences associées » sont enregistrés, chacun
avec sa fiche (le test demandé) ; la vérification s'assure ensuite que chaque test demandé couvre l'exigence dans les
fiches.

Un tableau sans les colonnes attendues est laissé de côté sans arrêter la lecture ; le fichier n'est refusé que si
aucun tableau n'a un en-tête correct. Les anomalies du document (nom de fichier, tableau sans fiche, tableau laissé
de côté, aucun numéro d'exigence reconnu) sont enregistrées avec lui et montrées dans la page STDHL.

Pour la page (to_dict) :

    {"stdhl": {"12": {"file_id": 12, "name": "...docx", "releases": ["A"], "stdhl_name": "FCT1",
                      "tome": "fct1" (None : tous les tomes),
                      "rows": [{"requirements": [ID], "test": fiche SYS_}, ...] (une ligne par exigence et test),
                      "issues": [texte]}}}
"""
import re

from configuration import DEFAULT_CONFIG, fold, header_column, header_text, match_header, plural, table_left_out
from file_database import WordDocument


def found(match):
    """Texte reconnu par une expression réglée : son groupe 1 s'il en a un, sinon toute la correspondance."""
    return (match[1] if match.re.groups else match[0]).strip()


# Formats de STDHL reconnus à l'import, gardés dans files.format : 1 = tableaux (un STDHL par tome), 2 = un document en
# texte.
FORMATS = {1: "par tome", 2: "unique"}


def unique(values):
    """Valeurs non vides sans doublon, dans leur ordre."""
    return list(dict.fromkeys(v for v in values if v))


class StdhlDatabase:
    """Contenu des documents STDHL, dans la base SQL."""

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    # ------------------------------------------------------------------ #
    # Lecture d'un document STDHL
    # ------------------------------------------------------------------ #

    @staticmethod
    def _settings(config):
        """En-têtes attendus et format des numéros d'exigence, d'après une configuration (valeurs par défaut sans
        elle)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        return get("headers_stdhl"), re.compile(get("requirement_regex"))

    @staticmethod
    def _find_header(grid, headers):
        """(colonne des exigences, [textes attendus manquants]) de l'en-tête d'un tableau, toujours sa première
        ligne, ou None. Un tableau STDHL est reconnu à l'en-tête de sa colonne des exigences."""
        text = header_text(headers, "requirements")
        cells = grid[0] if grid else []
        # Seule la colonne des exigences identifie un tableau STDHL ; les autres colonnes attendues sont ensuite
        # vérifiées, et un tableau à qui il en manque une est signalé et laissé de côté.
        if header_column(cells, text) is not None:
            cols, missing = match_header(cells, headers)
            return cols["requirements"], missing
        return None

    @classmethod
    def read_docx(cls, content, file_name, config=None):
        """Lit un STDHL (octets .docx). Renvoie {"stdhl_name", "rows": [{fiche, requirements: [ID]}], "skipped" :
        tableaux laissés de côté, "issues" : anomalies du document}. Une ligne est la cellule d'un tableau sous
        l'en-tête des exigences ; sa fiche (le test) est la dernière fiche de test SYS_ avant le tableau. Les
        tableaux sont numérotés au premier niveau."""
        headers, regex = cls._settings(config)
        get = config.get if config is not None else DEFAULT_CONFIG.get
        fiche_regex = re.compile(get("fiche_regex"))
        match = re.search(get("stdhl_name_regex"), file_name)
        rows, skipped, orphans = [], [], []
        tables = 0   # tableaux à l'en-tête correct, quel que soit le contenu de leurs cellules
        fiche = ""
        # Le document est parcouru dans l'ordre de lecture : un paragraphe qui correspond à fiche_regex ouvre une fiche
        # (le test), et chaque tableau qui suit appartient à cette fiche jusqu'à la suivante.
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
            # Un tableau valide avant toute fiche : ses exigences sont gardées, sans test, et le document est signalé.
            if not fiche:
                orphans.append(number)
            # Chaque ligne du tableau sous l'en-tête : les numéros d'exigence trouvés dans sa cellule des exigences.
            for cells in value[1:]:
                ids = unique(regex.findall(cells[col])) if col < len(cells) else []
                if ids:
                    rows.append({"fiche": fiche, "requirements": ids})
        label = header_text(headers, "requirements")
        # Anomalies du document, gardées avec lui et montrées dans la page STDHL : nom de fichier, tableaux laissés de
        # côté, tableaux sans fiche, aucun numéro d'exigence.
        issues = ([] if match else ["nom de fichier hors format [date]_[texte]_STDHL_[Nom du STDHL]_rev[texte].docx"]) \
            + skipped + [f"tableau {t} « {label} » sans fiche SYS_ avant lui" for t in orphans] \
            + ([f"aucun numéro d'exigence reconnu dans {'les ' + str(tables) + ' tableaux' if tables > 1 else 'le tableau'}"
                f" : vérifier la colonne « {label} » et le format des numéros (Configuration)"] if tables and not rows else [])
        if not tables:
            # Aucun tableau STDHL : le document est peut-être un STDHL de format 2 (tous les STDHL en un document, en
            # texte).
            single = cls.read_single(content, file_name, config)
            if single["tables"]:
                return single
        return {"stdhl_name": found(match) if match else file_name.rsplit(".", 1)[0], "rows": rows, "tables": tables,
                "skipped": skipped, "issues": issues, "format": FORMATS[1], "gaps": [], "refusals": []}

    @classmethod
    def _readable(cls, content, file_name, config=None):
        """Contenu d'un STDHL chargeable : refusé (ValueError) quand aucun tableau n'a un en-tête correct et qu'aucun
        test du format 2 n'est trouvé, ou quand un test du format 2 est d'une autre release que la base. Des tableaux
        sans aucun numéro d'exigence reconnu sont une anomalie du document, pas un refus."""
        data = cls.read_docx(content, file_name, config)
        if not data["tables"]:
            raise ValueError("aucun tableau STDHL valide (format par tome) ni test du format unique trouvé (voir "
                             "Configuration > STDHL) : fichier non chargé")
        refusals = data["refusals"] + [
            f"test {g['test']} : {g['requirement']} " + ("citée dans les test cases mais absente des exigences associées"
                                                         if g["side"] == "test case" else
                                                         "dans les exigences associées mais citée dans aucun test case")
            for g in data["gaps"]]
        if refusals:
            raise ValueError(" ; ".join(refusals) + " : fichier non chargé")
        return data

    @classmethod
    def read_single(cls, content, file_name, config=None):
        """Un STDHL de format 2 (« unique ») sous la forme de read_docx : une ligne par test et exigence, "tables"
        étant le nombre de tests. La couverture d'un test est l'ensemble des exigences de ses test cases ; sa liste «
        Exigences associées » doit contenir exactement celles-là : une exigence dans une seule des deux parties
        refuse le fichier ("gaps" : [{test, requirement, side « liste » ou « test case »}]), comme une release d'un
        test qui n'est pas celle de la base ("refusals"). Un test sans release, un nom de test répété, un test sans
        exigence sont des anomalies du document ("issues")."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        data = cls.read_text_format(content, file_name, config)
        release = (get("release") or "").upper()
        tests, rows, issues, refusals, gaps = data["tests"], [], [], [], []
        names = [t["name"] for t in tests]
        issues += [f"test {n} présent {names.count(n)} fois" for n in sorted({n for n in names if names.count(n) > 1})]
        for t in tests:
            if t["releases"] is None:
                issues.append(f"test {t['name']} : applicabilité absente")
            elif release and release not in t["releases"]:
                refusals.append(f"test {t['name']} : applicabilité {', '.join(t['releases']) or 'illisible'}, la base "
                                f"est en release {release}")
            covered = unique(i for c in t["cases"] for i in c["requirements"])
            listed = t["listed"]
            if not covered and not listed:
                issues.append(f"test {t['name']} : aucune exigence")
            gaps += [{"test": t["name"], "requirement": i, "side": "test case"} for i in covered if i not in listed]
            gaps += [{"test": t["name"], "requirement": i, "side": "liste"} for i in listed if i not in covered]
            rows += [{"fiche": t["name"], "requirements": [i]} for i in covered]
        return {"stdhl_name": data["stdhl_name"], "rows": rows, "tables": len(tests), "skipped": [], "issues": issues,
                "format": FORMATS[2], "gaps": gaps, "refusals": refusals,
                "cases": sum(len(t["cases"]) for t in tests),
                # Les test cases demandés, par leur nom complet : une fiche d'un autre test case est ignorée.
                "test_cases": list(dict.fromkeys((t["name"], c["name"]) for t in tests for c in t["cases"]))}

    @classmethod
    def read_text_format(cls, content, file_name, config=None):
        """Lit un STDHL de format 2 (sans tableau STDHL) : ses tests, tels qu'ils sont écrits (voir read_single). Un
        test est un titre « <nom> : <titre> » (stdhl2_test_regex ; une ligne qui commence par un numéro d'exigence
        n'est jamais un test ; la table des matières, styles TM* / TOC*, est laissée de côté). Viennent ensuite des
        parties étiquetées : sa release (stdhl2_release_label, lue par stdhl2_release_regex), la liste de ses
        exigences associées (stdhl2_requirements_label, puis les éléments de liste qui suivent : la fin de la liste
        la termine), puis ses test cases (stdhl2_test_case_regex), dont les phases citent des exigences. Renvoie
        {"stdhl_name", "tests": [{name, releases (None quand elle est absente), listed: [numéros], cases: [{name,
        requirements}]}]}."""
        _, regex = cls._settings(config)
        get = config.get if config is not None else DEFAULT_CONFIG.get
        test_regex = re.compile(get("stdhl2_test_regex"))
        case_regex = re.compile(get("stdhl2_test_case_regex"))
        release_regex = re.compile(get("stdhl2_release_regex"))
        requirements_label, release_label = fold(get("stdhl2_requirements_label")), fold(get("stdhl2_release_label"))
        tests, part = [], None   # part : "list" dans les exigences associées, "case" dans un test case, None ailleurs
        for style, text, item in WordDocument(content).paragraphs():
            if not text.strip() or style.upper().startswith(("TM", "TOC")):
                continue
            if part == "list" and not item:
                part = None   # la fin de la liste termine les exigences associées
            starts_with_number = (m := regex.search(text)) is not None and not text[:m.start()].strip()
            if not starts_with_number and (t := test_regex.search(text)):
                tests.append({"name": found(t), "releases": None, "listed": [], "cases": []})
                part = None
                continue
            if not tests:
                continue   # avant le premier test : introduction du document
            test = tests[-1]
            label, _, value = text.partition(":")
            label = fold(label)
            if label == release_label:
                test["releases"] = [r.upper() for r in release_regex.findall(value)]
                part = None
            elif label == requirements_label:
                part = "list"
                test["listed"] += regex.findall(value)
            elif case_regex.search(text):
                test["cases"].append({"name": text.strip(), "requirements": regex.findall(text)})
                part = "case"
            elif part == "list":
                test["listed"] += regex.findall(text)
            elif part == "case":
                test["cases"][-1]["requirements"] += regex.findall(text)
        for t in tests:
            t["listed"] = unique(t["listed"])
            for c in t["cases"]:
                c["requirements"] = unique(c["requirements"])
        match = re.search(get("stdhl_name_regex"), file_name)
        return {"stdhl_name": found(match) if match else file_name.rsplit(".", 1)[0], "tests": tests}

    @classmethod
    def analyse(cls, content, file_name, config=None):
        """Ce que ferait le chargement d'un STDHL. Refusé (ValueError) quand il est illisible ou sans tableau valide."""
        data = cls._readable(content, file_name, config)
        return {"tables": data["tables"], "rows": len(data["rows"]), "stdhl_name": data["stdhl_name"],
                "skipped": data["skipped"], "issues": data["issues"], "format": data["format"],
                "gaps": data["gaps"]}

    # ------------------------------------------------------------------ #
    # Enregistrement
    # ------------------------------------------------------------------ #

    @staticmethod
    def _write(db, file_id, data):
        """Remplace le contenu enregistré d'un STDHL (dans une transaction) : une ligne de stdhl_exi par exigence
        qu'il cite et test (fiche SYS_) qu'il demande pour elle, le test pris dans le référentiel fe_tests. Un numéro
        absent des exigences de la base n'est pas enregistré : c'est une anomalie du document. Une ligne supprimée
        par l'utilisateur (page STDHL) reste supprimée quand le fichier est relu."""
        from nexus_db import test_id as test_of
        deleted = {f"{r[0] or ''}|{r[1]}" for r in db.execute(
            "SELECT t.name, q.req_id FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
            "LEFT JOIN fe_tests t ON t.id = x.test_id WHERE x.file_id = ? AND x.deleted = 1", (file_id,))}
        db.execute("DELETE FROM stdhl_exi WHERE file_id = ?", (file_id,))
        # Numéro d'exigence -> id de la base, pour enregistrer la couverture par id.
        known = {r[0]: r[1] for r in db.execute("SELECT req_id, id FROM requirements")}
        unknown = []
        for r in data["rows"]:
            for number in r["requirements"]:
                if number in known:
                    # Le test de la ligne dans le référentiel fe_tests (ajouté s'il est nouveau) ; une ligne sans fiche
                    # n'a pas de test.
                    # La même exigence et le même test peuvent être cités sur plusieurs lignes : enregistrés une fois
                    # (UNIQUE ne couvre pas un test NULL, d'où la vérification).
                    test = test_of(db, r["fiche"]) if r["fiche"] else None
                    if not db.execute("SELECT 1 FROM stdhl_exi WHERE file_id = ? AND requirement_id = ? "
                                      "AND test_id IS ?", (file_id, known[number], test)).fetchone():
                        key = f"{r['fiche']}|{number}"   # supprimée par l'utilisateur : reste supprimée
                        db.execute("INSERT INTO stdhl_exi (file_id, requirement_id, test_id, deleted) VALUES (?, ?, ?, ?)",
                                   (file_id, known[number], test, int(key in deleted)))
                elif number not in unknown:
                    unknown.append(number)
        issues = data["issues"] + ([plural(len(unknown), "numéro d'exigence absent de la base, non enregistré",
                                           "numéros d'exigence absents de la base, non enregistrés")
                                    + " : " + ", ".join(unknown)] if unknown else [])
        # Les test cases demandés par un STDHL « unique », dans le référentiel fe_test_case (ajoutés s'ils sont
        # nouveaux).
        db.execute("DELETE FROM stdhl_test_case WHERE file_id = ?", (file_id,))
        for test, case in data.get("test_cases", []):
            test = test_of(db, test)
            db.execute("INSERT OR IGNORE INTO fe_test_case (test_id, test_case) VALUES (?, ?)", (test, case))
            db.execute("INSERT OR IGNORE INTO stdhl_test_case (file_id, test_case_id) SELECT ?, id FROM fe_test_case "
                       "WHERE test_id = ? AND test_case = ?", (file_id, test, case))
        # Le nom du STDHL et les anomalies sont gardés sur la ligne du fichier : la page STDHL les montre.
        db.execute("UPDATE files SET stdhl_name = ?, issues = ?, format = ? WHERE id = ?",
                   (data["stdhl_name"], "\n".join(issues), data.get("format", ""), file_id))

    def load_file(self, files, requirements, content, name, tome, releases, config=None):
        """Charge un STDHL : il appartient au `tome` donné, ou à aucun ("" : tous les tomes). Le fichier est
        enregistré, puis son contenu (le fichier est retiré en cas d'échec) ; un STDHL de même nom est remplacé, quel
        que soit son tome. Un STDHL « unique » dont les exigences associées et les test cases ne correspondent pas
        est refusé (voir _readable)."""
        tome = files.check_tome("STDHL", requirements, tome)
        data = self._readable(content, name, config)
        # La copie est enregistrée d'abord : si l'écriture du contenu échoue, la copie est retirée, rien ne reste à
        # moitié chargé.
        record = files.add("stdhl", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data)
        except Exception:
            files.remove(record["id"])
            raise
        return {"file_id": record["id"], "rows": len(data["rows"]), "skipped": data["skipped"],
                "format": data["format"], "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Lit et enregistre le contenu des fichiers STDHL qui n'en ont pas encore, et de tous avec `again` (réglages
        de lecture, plans ou exigences changés). Renvoie les noms lus ; les fichiers illisibles sont laissés tels
        quels."""
        # Un STDHL sans aucune ligne enregistrée n'a jamais été lu (ou ses lignes ont été vidées) : il est relu.
        stored = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM stdhl_exi")}   # déjà lus
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
    # Accès
    # ------------------------------------------------------------------ #

    def documents(self, with_deleted=False):
        """Les documents STDHL avec leur tome (None : tous les tomes), leurs lignes et leurs anomalies. Les lignes
        supprimées par l'utilisateur sont laissées de côté, sauf avec `with_deleted` (page STDHL : montrées barrées,
        "deleted" vrai)."""
        heads = self.sql.query("SELECT f.id AS file_id, f.name, f.stdhl_name, f.issues, t.name AS tome FROM files f "
                               "LEFT JOIN list_tomes t ON t.id = f.tome_id WHERE f.source = 'stdhl' ORDER BY f.id")
        rows, releases = {}, self.sql.file_releases()
        # Une ligne par exigence et test demandé, dans l'ordre du document.
        for r in self.sql.query("SELECT x.file_id, q.req_id, t.name AS test, x.deleted FROM stdhl_exi x "
                                "JOIN requirements q ON q.id = x.requirement_id LEFT JOIN fe_tests t ON t.id = x.test_id "
                                f"{'' if with_deleted else 'WHERE x.deleted = 0 '}ORDER BY x.file_id, x.id"):
            # Les lignes d'un fichier supprimé entre-temps sont ignorées grâce aux en-têtes ci-dessous.
            rows.setdefault(r["file_id"], []).append({"requirements": [r["req_id"]], "test": r["test"] or "",
                                                     "deleted": bool(r["deleted"])})
        return [{"file_id": h["file_id"], "name": h["name"], "releases": releases.get(h["file_id"], []),
                 "stdhl_name": h["stdhl_name"], "tome": h["tome"], "rows": rows.get(h["file_id"], []),
                 "issues": [i for i in h["issues"].split("\n") if i]} for h in heads]

    def to_dict(self):
        """Les STDHL par numéro de fichier, pour la page."""
        return {"stdhl": {str(d["file_id"]): d for d in self.documents(with_deleted=True)}}

    def edit(self, file_id, action, requirement, test=""):
        """Ligne d'une exigence et d'un test dans un STDHL, supprimée par l'utilisateur (« delete » : gardée, barrée,
        laissée de côté par la vérification) ou rétablie (« restore »). ValueError quand la ligne est inconnue."""
        if action not in ("delete", "restore"):
            raise ValueError(f"action inconnue : {action}")
        with self.sql.transaction() as db:
            row = db.execute(
                "SELECT x.id FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
                "LEFT JOIN fe_tests t ON t.id = x.test_id WHERE x.file_id = ? AND q.req_id = ? AND COALESCE(t.name, '') = ?",
                (file_id, str(requirement or ""), str(test or ""))).fetchone()
            if row is None:
                raise ValueError(f"ligne inconnue dans ce STDHL : {requirement} / {test or 'sans test'}")
            db.execute("UPDATE stdhl_exi SET deleted = ? WHERE id = ?", (int(action == "delete"), row[0]))
        return {"deleted": action == "delete"}
