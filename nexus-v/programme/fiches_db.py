"""Fiches d'essai de NEXUS-V, dans la base SQL (tables files, fe_tests et fiches_essai).

Les fiches XML et les fichiers Excel des autres plateformes (voir traceability_db.py) vont dans la même table
fiches_essai. Chaque fichier garde la plateforme d'où il vient, choisie à son chargement (files.platform_id).

Comme un STDHL, un fichier de fiches (XML ou Excel) appartient à un tome, ou à aucun : il s'applique alors à tous les
tomes. Seules les exigences de son tome (de toute la base sans tome) sont enregistrées ; un autre numéro est montré à
l'import, où il ne peut qu'être ignoré. Retirer le fichier retire son contenu (cascade).

Le XML est lu au moyen de trois balises réglables (test_tag, test_case_tag, requirement_tag) : un test contient des
test cases, qui contiennent les exigences testées, séparées par des virgules dans une balise. Chaque valeur est le
texte entre les balises ouvrante et fermante (sans le texte des balises qu'elle contient) ; les attributs ne sont pas
lus. La table fiches_essai a une ligne par exigence de chaque test case : le test (référentiel fe_tests), le test
case et l'exigence. Un fichier où aucun numéro n'est trouvé ainsi est refusé.

Un second format XML, « campagne » (export de l'outil de campagne), est reconnu à ses éléments TestSuite : un
TestSuite est un test (attribut name, sinon sa propriété Title), ses TestCases sont les test cases (propriété Title)
; les exigences d'un test case sont la propriété VerifiedRequirements de ses TestSteps, de lui-même et de son
TestSuite ; chacune prend le résultat du test case, d'après le RunResult de ses steps (FAILED si l'un a échoué,
PASSED si tous ceux exécutés sont passés, NOT_RUN sinon). Le nom de la campagne (release) et la configuration de
l'outil ne sont pas lus. Le format reconnu est gardé dans files.format (« balises » ou « campagne »).

Pour la page (to_dict) :

    {"fiches": {"12": {"file_id": 12, "name": "...xml", "releases": ["A"], "tome": "fct1" (None : tous),
                       "ids": [ID], "rows": [{"test_id", "test", "test_case", "requirements": [ID]}, ...]}}}
"""
import re
import xml.etree.ElementTree as ET

from configuration import DEFAULT_CONFIG, plural
from file_database import FileDatabase, WordDocument


def local_name(el):
    """Nom d'un élément sans son espace de noms, en minuscules."""
    # ElementTree écrit une balise avec espace de noms sous la forme {uri}nom : seul le nom compte, quelle que soit sa
    # casse.
    return el.tag.rsplit("}", 1)[-1].lower()


# Formats de fiches XML reconnus à l'import, gardés dans files.format.
FORMATS = {"tags": "balises", "campaign": "campagne"}


def properties(el):
    """{nom: texte} des propriétés <Properties><Property name="...">texte</Property></Properties> d'un élément (les
    siennes, pas celles des éléments qu'il contient)."""
    found = {}
    for block in el:
        if local_name(block) == "properties":
            for prop in block:
                if local_name(prop) == "property" and prop.get("name"):
                    found[prop.get("name")] = (prop.text or "").strip()
    return found


def own_text(el):
    """Texte entre les balises ouvrante et fermante d'un élément, sans le texte des éléments qu'il contient."""
    # Le texte d'un élément est son texte de tête plus le texte qui suit chaque enfant (tail) ; le texte propre des
    # enfants est laissé de côté. Les suites d'espaces et de sauts de ligne deviennent une espace.
    return " ".join(" ".join([el.text or "", *(child.tail or "" for child in el)]).split())


class FichesDatabase:
    """Contenu des fichiers de Fiches d'Essai (tests, test cases, exigences), dans la base SQL."""

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def read_xml(content, config=None):
        """Lit un fichier de Fiches d'Essai (octets XML). Renvoie {"rows": [{test, test_case, requirements: [ID]}],
        "ids": numéros d'exigence testés}, au moyen des trois balises réglées (un test contient des test cases, qui
        contiennent les exigences)."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        regex = re.compile(get("requirement_regex"))
        try:
            root = ET.fromstring(content)
        except ET.ParseError as e:
            raise ValueError(f"XML mal formé : {e}") from None

        def children(parent, tag):
            """Éléments sous `parent` (lui compris) qui ont cette balise, quels que soient leur espace de noms et
            leur casse."""
            tag = tag.lower()
            return [el for el in parent.iter() if local_name(el) == tag]

        if children(root, "testsuite"):
            return FichesDatabase.read_campaign(root, regex, children)
        rows = []
        # Une ligne par test case de chaque test, avec les numéros d'exigence de ses balises d'exigence (séparés par des
        # virgules dans une balise, ou dans plusieurs balises). WordDocument.normalize remplace les tirets et espaces
        # spéciaux d'un copier-coller par des ordinaires, pour que les numéros correspondent.
        for test in children(root, get("test_tag")):
            for case in children(test, get("test_case_tag")):
                text = [own_text(el) for el in children(case, get("requirement_tag"))]
                ids = list(dict.fromkeys(regex.findall(WordDocument.normalize("\n".join(text)))))
                if ids:
                    rows.append({"test": own_text(test), "test_case": own_text(case), "requirements": ids})
        # ids : chaque numéro d'exigence du fichier, une fois, dans l'ordre de lecture.
        return {"rows": rows, "ids": list(dict.fromkeys(i for row in rows for i in row["requirements"])),
                "format": FORMATS["tags"]}

    @staticmethod
    def read_campaign(root, regex, children):
        """Un fichier de fiches au format « campagne » : une ligne par TestCase qui a des exigences, {test,
        test_case, requirements, result PASSED | FAILED | NOT_RUN du test case} (voir le docstring du module)."""
        numbers = lambda text: regex.findall(WordDocument.normalize(text or ""))
        rows = []
        for suite in children(root, "testsuite"):
            suite_props = properties(suite)
            test = (suite.get("name") or suite_props.get("Title") or "").strip()
            for case in children(suite, "testcase"):
                case_props = properties(case)
                title = case_props.get("Title") or ""   # sans titre : ignoré à l'import
                found, runs = [], []   # exigences du test case ; résultats de ses steps
                for step in children(case, "teststep"):
                    step_props = properties(step)
                    found += numbers(step_props.get("VerifiedRequirements"))
                    if (step_props.get("RunResult") or "").upper() in ("PASSED", "FAILED"):
                        runs.append(step_props["RunResult"].upper())
                found += numbers(case_props.get("VerifiedRequirements")) + numbers(suite_props.get("VerifiedRequirements"))
                # Le résultat du test case : FAILED si un step a échoué, PASSED si tous les steps exécutés sont passés,
                # sinon NOT_RUN.
                result = "FAILED" if "FAILED" in runs else "PASSED" if runs else "NOT_RUN"
                if found:
                    rows.append({"test": test, "test_case": title, "requirements": list(dict.fromkeys(found)),
                                 "result": result})
        return {"rows": rows, "ids": list(dict.fromkeys(i for row in rows for i in row["requirements"])),
                "format": FORMATS["campaign"]}

    @classmethod
    def _readable(cls, content, config=None):
        """Contenu d'un fichier chargeable : ValueError quand aucun numéro d'exigence n'a été trouvé."""
        data = cls.read_xml(content, config)
        if not data["ids"]:
            raise ValueError("aucun numéro d'exigence trouvé dans les balises attendues (voir Configuration > "
                             "En-têtes des tableaux) : fichier non chargé")
        return data

    def analyse(self, content, config=None, releases=(), tome=None):
        """Ce que ferait le chargement d'un fichier de fiches : le nombre de numéros d'exigence et les numéros à
        ignorer (absents des exigences du `tome`, de toute la base sans tome). Refusé (ValueError) quand il est mal
        formé ou sans aucun numéro."""
        data = self._readable(content, config)
        found = self.screen(data, tome or None, config)
        return {"ids": len(data["ids"]), "format": data["format"], "to_confirm": [],
                **{k: v for k, v in found.items() if k != "rows"}}

    def screen(self, data, tome=None, config=None, platform=""):
        """Les règles de l'import sur les lignes lues dans un fichier de fiches XML : {"rows" : les lignes à
        enregistrer [(test, test case, [numéros], résultat)], "unknown" : numéros absents de la base (du `tome`),
        ignorés ; "untitled" : test cases sans titre, ignorés ; "ignored_tests" : tests qu'aucun STDHL ne demande,
        leurs lignes ignorées ; "not_specified" : exigences dans aucun STDHL de leur tome (cela refuse le fichier) ;
        "delegated" : exigences non déléguées à la `platform` du fichier (obligatoire), chargées sans leur résultat ;
        "ignored_cases" : test cases que le STDHL « unique » de leur test ne demande pas, ignorés}. Les règles du
        STDHL (tests et test cases demandés, exigence dans aucun STDHL) s'appliquent à une fiche de la plateforme de
        niveau système (system_platform), sur les exigences qui lui sont déléguées."""
        # Les STDHL spécifient les tests de leur plateforme (configuration system_platform) : ses fiches sont vérifiées
        # par rapport à eux.
        system_platform = (config.get("system_platform") if config is not None else "") or ""
        with_stdhl = bool(platform) and platform.casefold() == system_platform.casefold()
        tome_of = {r["req_id"]: r["tome"] for r in self.sql.query(
            "SELECT q.req_id, t.name AS tome FROM requirements q JOIN list_tomes t ON t.id = q.tome_id"
            + (" WHERE t.name = ?" if tome else ""), (tome,) if tome else ())}
        asked = {r["name"] for r in self.sql.query(
            "SELECT DISTINCT t.name FROM stdhl_exi x JOIN fe_tests t ON t.id = x.test_id WHERE x.deleted = 0")}
        # Test cases demandés par les STDHL « unique » (un STDHL « par tome » n'en a pas : ses tests sont vérifiés
        # seuls).
        cases = {(r["test"], r["test_case"]) for r in self.sql.query(
            "SELECT t.name AS test, c.test_case FROM stdhl_test_case s JOIN fe_test_case c ON c.id = s.test_case_id "
            "JOIN fe_tests t ON t.id = c.test_id")}
        with_cases = {t for t, _ in cases}
        specified = {(r["req_id"], r["tome"]) for r in self.sql.query(
            "SELECT q.req_id, t.name AS tome FROM stdhl_exi x JOIN requirements q ON q.id = x.requirement_id "
            "JOIN files f ON f.id = x.file_id LEFT JOIN list_tomes t ON t.id = f.tome_id WHERE x.deleted = 0")}
        # Déléguée ailleurs : aucune des lignes de plan de l'exigence n'est déléguée à la plateforme du fichier.
        own = {r["req_id"] for r in self.sql.query(
            "SELECT DISTINCT q.req_id FROM plan_acceptation p JOIN requirements q ON q.id = p.requirement_id "
            "JOIN list_plateformes l ON l.id = p.delegation_id WHERE p.deleted = 0 AND LOWER(l.name) = LOWER(?)",
            (platform,))}
        planned = {r["req_id"] for r in self.sql.query(
            "SELECT DISTINCT q.req_id FROM plan_acceptation p JOIN requirements q ON q.id = p.requirement_id "
            "WHERE p.deleted = 0")}
        rows, unknown, untitled, ignored_tests, ignored_cases = [], [], [], [], []
        for r in data["rows"]:
            if not (r["test_case"] or "").strip():
                untitled.append(r["test"])
                continue
            if with_stdhl and r["test"] not in asked:
                if r["test"] not in ignored_tests:
                    ignored_tests.append(r["test"])
                continue
            if with_stdhl and r["test"] in with_cases and (r["test"], r["test_case"].strip()) not in cases:
                ignored_cases.append(f"{r['test']} / {r['test_case']}")
                continue
            keep = [i for i in r["requirements"] if i in tome_of]
            unknown += [i for i in r["requirements"] if i not in tome_of and i not in unknown]
            if keep:
                rows.append((r["test"], r["test_case"], keep, r.get("result")))
        cited = list(dict.fromkeys(i for _, _, ids, _ in rows for i in ids))
        delegated = [i for i in cited if i in planned and i not in own]
        # Dans aucun STDHL de son tome : seulement une exigence déléguée à la plateforme de niveau système (les autres
        # n'ont pas de STDHL).
        not_specified = [i for i in cited if with_stdhl and i not in delegated
                         and (i, tome_of[i]) not in specified and (i, None) not in specified]
        # Une exigence déléguée ailleurs est chargée sans son résultat : une ligne à part, NOT_RUN.
        split = []
        for test, case, ids, result in rows:
            split.append((test, case, [i for i in ids if i not in delegated], result))
            if any(i in delegated for i in ids):
                split.append((test, case, [i for i in ids if i in delegated], "NOT_RUN"))
        return {"rows": [x for x in split if x[2]], "unknown": unknown, "untitled": untitled,
                "ignored_tests": ignored_tests, "ignored_cases": ignored_cases,
                "not_specified": not_specified, "delegated": delegated}

    @staticmethod
    def refusal(found):
        """Le message qui refuse un fichier de fiches dont des exigences ne sont dans aucun STDHL, ou None."""
        parts = []
        if found["not_specified"]:
            parts.append(f"{plural(len(found['not_specified']), 'exigence')} dans aucun STDHL de son tome : "
                         + ", ".join(found["not_specified"]) + " (les spécifier d'abord dans le STDHL)")
        return " ; ".join(parts) + " : fichier non chargé" if parts else None

    @staticmethod
    def _write(db, file_id, data, rows):
        """Remplace le contenu enregistré d'un fichier (dans une transaction) : ses `rows` gardées par screen, seules
        les exigences de son tome (de toute la base pour un fichier sans tome)."""
        # Le tome du fichier (None : tous les tomes) limite les exigences enregistrées.
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        FichesDatabase.store_tests(db, file_id, rows, tome_id)
        db.execute("UPDATE files SET format = ? WHERE id = ?", (data.get("format", ""), file_id))

    @staticmethod
    def store_tests(db, file_id, tests, tome_id=None, by_case=True):
        """Remplace la couverture d'un fichier de fiches (dans une transaction) : une ligne de fiches_essai par
        exigence d'un test case [(nom du test, test case, [ID], éventuellement le résultat du test case)], le
        résultat PASSED ou FAILED quand le fichier en donne un (RunResult), NOT_RUN sinon ; le test pris dans le
        référentiel fe_tests et le test case dans fe_test_case (ajoutés s'ils sont nouveaux) ; seules les exigences
        de la base (du tome `tome_id` s'il est donné) sont enregistrées. Une ligne pointe toujours vers son test ;
        avec `by_case` (fiches XML), aussi vers son test case."""
        from nexus_db import test_id as test_of
        db.execute("DELETE FROM fiches_essai WHERE file_id = ?", (file_id,))   # tests plus cités : retirés
        # Numéro d'exigence -> id, du tome si le fichier en a un, de toute la base sinon.
        known = {r[0]: r[1] for r in db.execute(
            "SELECT req_id, id FROM requirements" + (" WHERE tome_id = ?" if tome_id is not None else ""),
            (tome_id,) if tome_id is not None else ())}
        for name, test_case, ids, *rest in tests:
            result = rest[0] if rest and rest[0] in ("PASSED", "FAILED") else "NOT_RUN"
            # Le test dans le référentiel fe_tests (partagé avec les STDHL) ; la même exigence dans le même test case
            # est enregistrée une fois (INSERT OR IGNORE sur la clé UNIQUE).
            test_id = test_of(db, name)
            if by_case:
                db.execute("INSERT OR IGNORE INTO fe_test_case (test_id, test_case) VALUES (?, ?)", (test_id, test_case))
                case_id = db.execute("SELECT id FROM fe_test_case WHERE test_id = ? AND test_case = ?",
                                     (test_id, test_case)).fetchone()[0]
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_id, test_case_id, result) "
                               "VALUES (?, ?, ?, ?, ?)", [(file_id, known[i], test_id, case_id, result) for i in ids if i in known])
            else:
                db.executemany("INSERT OR IGNORE INTO fiches_essai (file_id, requirement_id, test_id, result) "
                               "VALUES (?, ?, ?, ?)", [(file_id, known[i], test_id, result) for i in ids if i in known])

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None, platform=""):
        """Charge un fichier de fiches (XML) de la `platform` choisie : il appartient au `tome` donné, ou à aucun (""
        : tous les tomes). Les numéros absents de son tome doivent être ignorés par l'utilisateur (`decisions`). Le
        fichier est enregistré, puis son contenu (le fichier est retiré en cas d'échec) ; un fichier de même nom est
        remplacé, quel que soit son tome."""
        tome = files.check_tome("de fiches d'essai", requirements, tome)
        data = self._readable(content, config)
        # Les règles de l'import (voir screen) : refusé pour une exigence dans aucun STDHL.
        # La plateforme d'où vient le fichier : obligatoire, une plateforme de la configuration.
        if not platform:
            raise ValueError("indiquer la plateforme d'où vient ce fichier de fiches d'essai")
        if config is not None and not config.is_platform(platform):
            raise ValueError(f"plateforme {platform} absente de la configuration")
        found = self.screen(data, tome, config, platform)
        if (message := self.refusal(found)) is not None:
            raise ValueError(message)
        # La copie est enregistrée d'abord, le contenu ensuite, dans une transaction : en cas d'échec, la copie est
        # retirée.
        record = files.add("xml", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data, found["rows"])
                # La plateforme d'où vient la fiche.
                db.execute("UPDATE files SET platform_id = (SELECT id FROM list_plateformes WHERE LOWER(name) = LOWER(?)) "
                           "WHERE id = ?", (platform, record["id"]))
        except Exception:
            files.remove(record["id"])
            raise
        stored = len({i for _, _, ids, _ in found["rows"] for i in ids})
        return {"file_id": record["id"], "ids": stored, "ignored": len(found["unknown"]), "format": data["format"],
                "unknown": found["unknown"], "untitled": found["untitled"], "ignored_tests": found["ignored_tests"],
                "ignored_cases": found["ignored_cases"],
                "delegated": found["delegated"], "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Lit et enregistre le contenu des fichiers de fiches qui n'ont encore aucune ligne, et de tous avec `again`
        (réglages de lecture ou exigences changés). Renvoie les noms lus ; un fichier illisible garde son contenu
        enregistré."""
        known = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM fiches_essai")}   # déjà lus
        done = []
        for record in files.list(slot="xml"):
            if record["id"] not in known or again:
                try:
                    data = self._readable(files.content(record["id"]) or b"", config)
                except ValueError:
                    continue
                # Relu : les mêmes règles, une exigence désormais dans aucun STDHL laissée de côté au lieu de refuser.
                found = self.screen(data, record["tome"], config, self.sql.scalar(
                    "SELECT l.name FROM files f JOIN list_plateformes l ON l.id = f.platform_id WHERE f.id = ?",
                    (record["id"],)) or "")
                blocked = set(found["not_specified"])
                rows = [(t, c, [i for i in ids if i not in blocked], r) for t, c, ids, r in found["rows"]]
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data, [x for x in rows if x[2]])
                done.append(record["name"])
        return done

    def documents(self, source="xml"):
        """Les fichiers de fiches d'un type (xml : Fiches d'essai XML ; traceability : fichiers Excel des autres
        plateformes) avec leur plateforme, leur tome (None : tous les tomes), leurs releases, leurs numéros
        d'exigence et leurs lignes (test, test case, exigences)."""
        heads = self.sql.query("SELECT f.id AS file_id, f.name, p.name AS platform, t.name AS tome "
                               "FROM files f LEFT JOIN list_plateformes p ON p.id = f.platform_id "
                               "LEFT JOIN list_tomes t ON t.id = f.tome_id "
                               "WHERE f.source = ? ORDER BY f.id", (source,))
        releases = self.sql.file_releases()
        wanted = {h["file_id"] for h in heads}
        rows = {}
        # Une ligne pointe vers son test, et pour les fiches XML aussi vers son test case (aucun pour les autres
        # plateformes).
        for r in self.sql.query("SELECT x.file_id, x.test_id, t.name AS test, COALESCE(c.test_case, '') AS test_case, "
                                "q.req_id, x.result FROM fiches_essai x LEFT JOIN fe_test_case c ON c.id = x.test_case_id "
                                "JOIN fe_tests t ON t.id = x.test_id "
                                "JOIN requirements q ON q.id = x.requirement_id ORDER BY x.file_id, x.id"):
            if r["file_id"] not in wanted:
                continue
            cases = rows.setdefault(r["file_id"], [])
            if not cases or (cases[-1]["test_id"], cases[-1]["test_case"], cases[-1]["result"]) != \
                    (r["test_id"], r["test_case"], r["result"]):
                cases.append({"test_id": r["test_id"], "test": r["test"], "test_case": r["test_case"],
                              "result": r["result"], "requirements": []})
            cases[-1]["requirements"].append(r["req_id"])
        return [{"file_id": h["file_id"], "name": h["name"], "source": source, "platform": h["platform"] or "",
                 "tome": h["tome"], "releases": releases.get(h["file_id"], []),
                 "ids": list(dict.fromkeys(i for c in rows.get(h["file_id"], []) for i in c["requirements"])),
                 "rows": rows.get(h["file_id"], [])} for h in heads]

    def to_dict(self):
        """Les fichiers de fiches par numéro de fichier, pour la page."""
        return {"fiches": {str(d["file_id"]): d for d in self.documents()}}
