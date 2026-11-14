"""Fiches d'essai des autres plateformes : fichiers Excel, dans la base SQL (tables files et fiches_essai).

Ces fichiers Excel accompagnent les fiches XML : ils tracent les exigences déléguées aux autres plateformes. Comme un
fichier de fiches XML, un fichier appartient à un tome ou à aucun (il s'applique alors à tous les tomes) ; le retirer
retire son contenu (cascade).

Leur format n'est pas encore décrit : seuls les numéros d'exigence de son tome trouvés dans ses cellules sont
enregistrés, dans la table fiches_essai (un test sans nom, sans test case). Le fichier garde la plateforme d'où il
vient, choisie au chargement. Un numéro inconnu de son tome est montré à l'import, où il ne peut qu'être ignoré. Un
fichier sans aucun numéro est refusé.

Pour la page (to_dict) :

    {"traceability": {"12": {"file_id": 12, "name": "...xlsx", "tome": "fct1" (None : tous), "releases": ["A"],
                             "platform": "PLT1_N-1", "ids": [ID]}}}
"""
import io
import re

import openpyxl

from configuration import DEFAULT_CONFIG
from file_database import FileDatabase


class TraceabilityDatabase:
    """Numéros d'exigence cités par les fichiers de traçabilité des autres plateformes, dans la base SQL."""

    def __init__(self, sql):
        """sql : la NexusDatabase partagée."""
        self.sql = sql
        self.path = sql.path

    @staticmethod
    def read_excel(content, config=None):
        """Lit un classeur de traçabilité (octets). Renvoie {"ids": numéros d'exigence cités dans une cellule
        quelconque}."""
        get = config.get if config is not None else DEFAULT_CONFIG.get
        regex = re.compile(get("requirement_regex"))
        try:
            # Lecture seule et valeurs seules (pas les formules) : le format de ces fichiers n'est pas encore défini,
            # toutes les cellules de toutes les feuilles sont donc parcourues à la recherche de numéros d'exigence.
            workbook = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except Exception as e:   # pas un classeur .xlsx / .xlsm
            raise ValueError(f"classeur Excel illisible (formats .xlsx et .xlsm) : {e}") from None
        try:
            ids = {}
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for value in row:
                        if value is not None:
                            # Un dictionnaire garde chaque numéro une fois, dans l'ordre où il est rencontré.
                            ids.update(dict.fromkeys(regex.findall(str(value))))
            return {"ids": list(ids)}
        finally:
            workbook.close()

    @classmethod
    def _readable(cls, content, config=None):
        """Contenu d'un fichier chargeable : ValueError quand aucun numéro d'exigence n'a été trouvé."""
        data = cls.read_excel(content, config)
        if not data["ids"]:
            raise ValueError("aucun numéro d'exigence trouvé : fichier non chargé")
        return data

    def _known(self, tome):
        """Numéros des exigences d'un tome (None sans tome)."""
        # Sans tome, les numéros sont vérifiés par rapport à toute la base (FileDatabase.to_ignore).
        if not tome:
            return None
        tome_id = self.sql.tome_id(tome)
        return set(FileDatabase.tome_requirements(self.sql.db, tome_id)) if tome_id is not None else set()

    def analyse(self, content, config=None, tome=None, releases=()):
        """Ce que ferait le chargement d'un fichier de traçabilité : le nombre de numéros d'exigence et les numéros
        inconnus du `tome` (de toute la base sans tome), à ignorer (même forme que les entrées de plan à régler).
        Refusé (ValueError) quand il est illisible ou sans aucun numéro."""
        ids = self._readable(content, config)["ids"]
        return {"ids": len(ids), "to_confirm": FileDatabase.to_ignore(self.sql.db, ids, tome or None)}

    @staticmethod
    def _write(db, file_id, ids):
        """Remplace les exigences d'un fichier dans fiches_essai (dans une transaction) : les exigences de son tome
        (de toute la base sans tome), sous un test sans nom (fe_tests) ni test case."""
        from fiches_db import FichesDatabase
        # Enregistré comme une fiche : un test sans nom et sans test case, les exigences de son tome (toutes sans tome).
        tome_id = db.execute("SELECT tome_id FROM files WHERE id = ?", (file_id,)).fetchone()[0]
        FichesDatabase.store_tests(db, file_id, [("", "", list(dict.fromkeys(ids)))], tome_id, by_case=False)

    def load_file(self, files, requirements, content, name, tome, releases, config=None, decisions=None, platform=""):
        """Charge un fichier de fiches d'une autre plateforme (Excel) : rattaché à un tome qui a des exigences ou à
        aucun ("" : tous les tomes), et à la `platform` d'où il vient. Les numéros inconnus de son tome doivent être
        ignorés par l'utilisateur (`decisions` {"rows": {numéro: {"ignore": true}}}, ou {"accept_all": true} en ligne
        de commande). Le fichier est enregistré, puis son contenu (le fichier est retiré en cas d'échec) ; un fichier
        de même nom est remplacé, quel que soit son tome."""
        tome = files.check_tome("de fiches d'une autre plateforme", requirements, tome)
        # La plateforme d'où vient le fichier est obligatoire : la couverture est jugée par rapport à la délégation des
        # plans.
        if config is not None and not config.is_platform(platform or ""):
            raise ValueError("indiquer la plateforme d'où vient ce fichier (une plateforme de la configuration)")
        data = self._readable(content, config)
        # Numéros absents du tome (de la base sans tome) : ignorés par l'utilisateur, non enregistrés.
        unknown = FileDatabase.to_ignore(self.sql.db, data["ids"], tome)
        FileDatabase.check_ignored(unknown, decisions)
        record = files.add("traceability", name, content, tome=tome, releases=releases)
        try:
            with self.sql.transaction() as db:
                self._write(db, record["id"], data["ids"])
                if config is not None:
                    # La plateforme telle qu'elle est écrite dans la configuration (la page peut l'envoyer dans une
                    # autre casse).
                    name_ok = next(p for p in config.get("platforms") if p.casefold() == platform.casefold())
                    db.execute("UPDATE files SET platform_id = (SELECT id FROM list_plateformes WHERE name = ?) "
                               "WHERE id = ?", (name_ok, record["id"]))
        except Exception:
            files.remove(record["id"])
            raise
        return {"file_id": record["id"], "tome": tome, "ids": len(data["ids"]) - len(unknown), "ignored": len(unknown),
                "replaced": files.remove_older(record)}

    def read_missing(self, files, config=None, again=False):
        """Lit et enregistre le contenu des fichiers de traçabilité qui n'en ont pas encore, et de tous avec `again`
        (les exigences ont changé). Renvoie les noms lus."""
        # Les fichiers sans aucune ligne dans fiches_essai n'ont pas encore été lus (nouvelle base, tables vidées).
        known = {r["file_id"] for r in self.sql.query("SELECT DISTINCT file_id FROM fiches_essai")}   # déjà lus
        done = []
        for record in files.list(slot="traceability"):
            if again or record["id"] not in known:
                try:
                    data = self.read_excel(files.content(record["id"]) or b"", config)
                except ValueError:
                    continue
                with self.sql.transaction() as db:
                    self._write(db, record["id"], data["ids"])
                done.append(record["name"])
        return done

    def documents(self):
        """Les fichiers de fiches des autres plateformes, sous la même forme que les Fiches d'essai (FichesDatabase)."""
        from fiches_db import FichesDatabase
        return FichesDatabase(self.sql).documents("traceability")

    def to_dict(self):
        """Les fichiers de traçabilité par numéro de fichier, pour la page."""
        return {"traceability": {str(d["file_id"]): d for d in self.documents()}}
