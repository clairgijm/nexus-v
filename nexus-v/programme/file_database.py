"""Fichiers chargés dans la page, dans la base SQL (table files). Chaque fichier appartient à la release de la base.

Chaque fichier chargé est copié dans le dossier des fichiers sous le nom « <n°>_<nom du fichier> », le numéro
croissant à chaque chargement (jamais réutilisé, même après un retrait). Un enregistrement (dict) a les champs :

    id, slot, name, stored_name, tome, releases

Le rapport de l'import d'un fichier (le même que la vérification avant import) est gardé dans le sous-dossier
« rapport » du dossier des fichiers, sous le nom « <n°>_rapport.html » ; il est supprimé avec le fichier. Un fichier
rejeté à l'import laisse un rapport de rejet pour son type et son tome, « rejet_<type>_<tome>.html » (et
« rejet_<type>.html », le dernier rejet du type), remplacé au rejet suivant et supprimé quand un fichier de ce type est
chargé dans ce tome.

Le tome est celui d'un fichier d'exigences ou d'un plan, et pour un fichier STDHL ou de fiches celui choisi au
chargement (None : tous les tomes). Le module lit aussi les documents Word (WordDocument), avec la seule bibliothèque
standard.
"""
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Optional

from configuration import DEFAULT_CONFIG, fold, plural

SLOTS = ("requirements", "plan", "stdhl", "xml", "traceability")
# Réglage donnant le sous-dossier du dossier des fichiers pour chaque type de fichier.
FOLDER_SETTINGS = {"requirements": "requirements_folder", "plan": "plan_folder", "stdhl": "stdhl_folder",
                   "xml": "fiches_folder", "traceability": "traceability_folder"}

COLUMNS = "id, source, name, stored_name"


# Lecture d'un document Word (.docx), avec la seule bibliothèque standard.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Caractères que Word insère à la place des espaces et des tirets ordinaires.
# Les espaces insécables et fines deviennent des espaces, les divers tirets de Word deviennent « - », le trait d'union
# conditionnel disparaît : un numéro d'exigence saisi dans Word correspond alors à requirement_regex.
NORMALIZATION = str.maketrans({" ": " ", " ": " ", "‑": "-", "‐": "-", "–": "-", "­": ""})


class WordDocument:
    """Un document .docx lu à partir de ses octets."""

    # Parties du document dont le texte compte : corps, en-têtes, pieds de page, notes, commentaires.
    TEXT_PARTS = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml")

    def __init__(self, content):
        """Ouvre un .docx à partir de ses octets. ValueError quand ce n'est pas un document Word lisible."""
        try:
            self._zip = zipfile.ZipFile(io.BytesIO(content))
            self.body = ET.fromstring(self._zip.read("word/document.xml"))
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
            raise ValueError(f"document Word illisible (format .docx attendu) : {e}") from None

    @staticmethod
    def normalize(text):
        """Texte où les espaces et tirets spéciaux de Word sont remplacés par des ordinaires."""
        return (text or "").translate(NORMALIZATION)

    @classmethod
    def paragraph_text(cls, p):
        """Texte d'un paragraphe, ses « runs » recousus (Word coupe souvent un même mot en plusieurs morceaux, par
        exemple « REQ- » + « 001 »)."""
        pieces = []
        # Morceaux de texte (w:t) dans l'ordre ; un trait d'union insécable, une tabulation ou un saut de ligne sont des
        # éléments à part dans le XML de Word.
        for el in p.iter():
            if el.tag == f"{W}t":
                pieces.append(el.text or "")
            elif el.tag == f"{W}noBreakHyphen":
                pieces.append("-")
            elif el.tag in (f"{W}tab", f"{W}br"):
                pieces.append(" ")
        return cls.normalize("".join(pieces))

    @classmethod
    def table_grid(cls, tbl):
        """Texte des cellules d'un tableau, ligne par ligne (paragraphes d'une cellule joints par des sauts de
        ligne)."""
        # Une liste par ligne (w:tr), une chaîne par cellule (w:tc) ; une cellule fusionnée apparaît une fois.
        return [["\n".join(cls.paragraph_text(p) for p in tc.iter(f"{W}p")).strip()
                 for tc in tr if tc.tag == f"{W}tc"]
                for tr in tbl if tr.tag == f"{W}tr"]

    def tables(self):
        """Tous les tableaux du corps (imbriqués compris), dans l'ordre du document : (numéro à partir de 1, grille)."""
        return [(i, self.table_grid(tbl)) for i, tbl in enumerate(self.body.iter(f"{W}tbl"), 1)]

    def blocks(self):
        """Paragraphes et tableaux de premier niveau du corps, dans l'ordre du document : ("p", None, texte) ou
        ("tbl", numéro à partir de 1 parmi les tableaux de premier niveau, grille). Ce qui est dans un tableau
        appartient au tableau."""
        number = 0

        # Parcours récursif : les paragraphes d'une zone de texte ou d'une section sont trouvés aussi, mais un tableau
        # est rendu entier (ses propres paragraphes ne sont pas répétés comme paragraphes).
        def walk(parent):
            """Paragraphes et tableaux sous `parent`, dans l'ordre, sans entrer dans un tableau."""
            nonlocal number
            for el in parent:
                if el.tag == f"{W}tbl":
                    number += 1
                    yield "tbl", number, self.table_grid(el)
                    continue
                if el.tag == f"{W}p":
                    yield "p", None, self.paragraph_text(el)
                yield from walk(el)

        yield from walk(self.body)

    def paragraphs(self):
        """Paragraphes du corps hors tableaux, dans l'ordre du document : [(identifiant de style, texte, élément de
        liste)]. L'identifiant de style est celui du fichier Word (« Titre3 », « TM1 »...), "" pour un paragraphe
        sans style ; un élément de liste est un paragraphe numéroté ou à puce par Word (numPr), ou d'un style de
        liste (« Paragraphedeliste », « ListParagraph »)."""
        found = []

        def walk(parent):
            """Paragraphes sous `parent`, dans l'ordre, sans entrer dans un tableau."""
            for el in parent:
                if el.tag == f"{W}tbl":
                    continue
                if el.tag == f"{W}p":
                    style = el.find(f"{W}pPr/{W}pStyle")
                    style = style.get(f"{W}val") if style is not None else ""
                    item = el.find(f"{W}pPr/{W}numPr") is not None or "list" in style.lower() or "liste" in style.lower()
                    found.append((style, self.paragraph_text(el), item))
                walk(el)

        walk(self.body)
        return found

    def text(self):
        """Texte complet : corps, tableaux, en-têtes, pieds de page, notes, commentaires."""
        paragraphs = []
        for part in self._zip.namelist():
            if self.TEXT_PARTS.fullmatch(part):
                root = ET.fromstring(self._zip.read(part))
                paragraphs.extend(self.paragraph_text(p) for p in root.iter(f"{W}p"))
        return "\n".join(paragraphs)


class FileDatabase:
    """Fichiers chargés : enregistrements dans la base SQL, copies dans le dossier des fichiers.

    Un enregistrement est un dict : id (int), slot (str), name (str), stored_name (str), tome (str, ou None pour un
    fichier STDHL ou de fiches qui s'applique à tous les tomes), releases (liste : la release de la base).
    """

    EDITABLE = ("slot", "name")

    def __init__(self, directory, sql, config=None):
        """directory : dossier des fichiers (copies des fichiers) ; sql : NexusDatabase ; config : donne le
        sous-dossier de chaque type de fichier (valeurs par défaut sans elle)."""
        self.directory = Path(directory)
        self.sql = sql
        self.path = sql.path
        self.config = config

    def folder(self, slot):
        """Sous-dossier du dossier des fichiers pour un type de fichier."""
        get = self.config.get if self.config is not None else DEFAULT_CONFIG.get
        return self.directory / get(FOLDER_SETTINGS[slot])

    @staticmethod
    def _insert(db, r):
        """Ajoute un enregistrement de fichier et ses releases (dans une transaction) ; r["tome"] est un nom de tome."""
        db.execute(f"INSERT INTO files ({COLUMNS}, tome_id) "
                   "VALUES (?, ?, ?, ?, (SELECT id FROM list_tomes WHERE name = ?))",
                   (r["id"], r["slot"], r["name"], r["stored_name"], r["tome"]))
        # Aucune release enregistrée : chaque fichier appartient à la release de la base.

    # ------------------------------------------------------------------ #
    # Accès
    # ------------------------------------------------------------------ #

    def _records(self, where="", params=()):
        """Enregistrements qui répondent à une condition SQL, chacun avec son nom de tome et ses releases."""
        # La colonne source de la table est le slot de l'enregistrement ; le tome est donné par son nom.
        rows = self.sql.query(f"SELECT {COLUMNS}, format, source AS slot, (SELECT t.name FROM list_tomes t WHERE t.id = tome_id) AS tome "
                              f"FROM files {where} ORDER BY id", params)
        releases = self.sql.file_releases()
        # « report » : le rapport de l'import du fichier existe.
        return [{**r, "releases": releases.get(r["id"], []), "report": self.report_path(r["id"]).is_file()}
                for r in rows]

    def report_path(self, file_id):
        """Chemin du rapport de l'import d'un fichier : <dossier des fichiers>/rapport/<n°>_rapport.html."""
        return self.directory / "rapport" / f"{file_id}_rapport.html"

    def save_report(self, file_id, html):
        """Écrit le rapport de l'import d'un fichier (le dossier rapport est créé au besoin)."""
        path = self.report_path(file_id)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(html, encoding="utf-8")

    def report(self, file_id):
        """Texte HTML du rapport de l'import d'un fichier, ou None s'il n'existe pas."""
        path = self.report_path(file_id)
        return path.read_text(encoding="utf-8") if path.is_file() else None

    def reject_path(self, kind, tome=None):
        """Chemin du rapport du dernier rejet d'un type de fichier (requirements, plans, stdhl, fiches,
        traceability) pour un tome : <dossier des fichiers>/rapport/rejet_<type>_<tome>.html ; sans tome, le dernier
        rejet du type, quel que soit le tome : rejet_<type>.html."""
        # Le nom du tome sans caractère hors d'un nom de fichier (pas de barre oblique inverse dans une f-string :
        # Python 3.9 la refuse).
        suffix = "_" + re.sub(r"[^\w.-]", "_", tome) if tome else ""
        return self.directory / "rapport" / f"rejet_{kind}{suffix}.html"

    def save_reject(self, kind, tome, html):
        """Écrit le rapport du rejet d'un fichier : comme dernier rejet de son type, et comme dernier rejet de son
        tome pour ce type quand le tome est connu (chacun remplace le précédent)."""
        for path in [self.reject_path(kind)] + ([self.reject_path(kind, tome)] if tome else []):
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(html, encoding="utf-8")

    def reject(self, kind, tome=None):
        """Texte HTML du rapport du dernier rejet d'un type (pour un tome), ou None."""
        path = self.reject_path(kind, tome)
        return path.read_text(encoding="utf-8") if path.is_file() else None

    def rejects(self):
        """Rejets gardés par tome : [{"kind", "tome"}] (les chevrons de ces tomes passent au rouge)."""
        folder = self.directory / "rapport"
        kinds = ("requirements", "plans", "stdhl", "fiches", "traceability")
        found = []
        for path in folder.glob("rejet_*_*.html") if folder.is_dir() else []:
            kind, _, tome = path.stem[len("rejet_"):].partition("_")
            if kind in kinds and tome:
                found.append({"kind": kind, "tome": tome})
        return found

    # Types de fichier dans l'ordre des étapes : un import réussi efface les rejets de son étape et des suivantes.
    REJECT_ORDER = ("requirements", "plans", "stdhl", "fiches", "traceability")

    def remove_reject(self, kind, tome=None):
        """Un fichier de ce type vient d'être chargé dans le tome : ses rejets et ceux des étapes suivantes du tome,
        à refaire de toute façon, sont retirés, ainsi que le dernier rejet du type."""
        self.reject_path(kind).unlink(missing_ok=True)
        if tome:
            for later in self.REJECT_ORDER[self.REJECT_ORDER.index(kind):]:
                self.reject_path(later, tome).unlink(missing_ok=True)

    def prune_reports(self, tomes):
        """Retire les rapports devenus sans objet : ceux des fichiers qui ne sont plus dans la base, et les rejets
        des tomes qui n'ont plus d'exigences (`tomes` : ceux qui en ont). Renvoie le nombre de rapports retirés."""
        folder = self.directory / "rapport"
        if not folder.is_dir():
            return 0
        known = {str(r["id"]) for r in self.list()}
        kept = {t.casefold() for t in tomes}
        gone = 0
        for path in folder.glob("*.html"):
            number = path.stem.partition("_")[0]
            stale = path.stem.endswith("_rapport") and number not in known
            if path.stem.startswith("rejet_"):
                kind, _, tome = path.stem[len("rejet_"):].partition("_")
                stale = bool(tome) and tome.casefold() not in kept
                # Sans aucune exigence dans la base, les derniers rejets de chaque type partent aussi.
                stale = stale or (not tome and not kept)
            if stale:
                path.unlink(missing_ok=True)
                gone += 1
        return gone

    def list(self, slot=None) -> list[dict[str, Any]]:
        """Enregistrements, éventuellement ceux d'un seul slot."""
        conditions, params = [], []
        if slot is not None:
            conditions.append("source = ?")
            params.append(slot)
        return self._records(("WHERE " + " AND ".join(conditions)) if conditions else "", params)

    def get(self, file_id) -> Optional[dict[str, Any]]:
        """Enregistrement de ce numéro, ou None."""
        found = self._records("WHERE id = ?", (file_id,))
        return found[0] if found else None

    def set(self, file_id, **fields):
        """Met à jour les champs modifiables d'un enregistrement. Renvoie l'enregistrement, ou None s'il est inconnu."""
        unknown = set(fields) - set(self.EDITABLE)
        if unknown:
            raise ValueError(f"champs non modifiables : {', '.join(sorted(unknown))}")
        if "slot" in fields and fields["slot"] not in SLOTS:
            raise ValueError(f"type de fichier inconnu : {fields['slot']}")
        with self.sql.transaction() as db:
            if not db.execute("SELECT 1 FROM files WHERE id = ?", (file_id,)).fetchone():
                return None
            if fields:
                # Seuls les champs modifiables arrivent dans ce SQL (vérifiés plus haut) : les noms de colonnes sont
                # sûrs.
                assignments = ', '.join([f"{k.replace('slot', 'source')} = ?" for k in fields])

                query = f"UPDATE files SET {assignments} WHERE id = ?"
                params = (*fields.values(), file_id)

                db.execute(query, params)

        return self.get(file_id)

    def path_of(self, file_id):
        """Chemin de la copie enregistrée d'un fichier, ou None s'il est inconnu."""
        record = self.get(file_id)
        return self.folder(record["slot"]) / record["stored_name"] if record else None

    def content(self, file_id):
        """Octets de la copie enregistrée d'un fichier, ou None s'il est inconnu ou absent du disque."""
        path = self.path_of(file_id)
        return path.read_bytes() if path and path.is_file() else None

    # ------------------------------------------------------------------ #
    # Ajout / retrait de fichiers
    # ------------------------------------------------------------------ #

    @staticmethod
    def safe_name(name):
        """Nom de fichier sans partie de dossier ni caractère interdit par Windows."""
        # Ne garde que la dernière partie d'un chemin (un navigateur peut envoyer C:\fakepath\nom), puis remplace les
        # caractères interdits par Windows.
        name = Path(name.replace("\\", "/")).name
        return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .") or "fichier"

    def add(self, slot, name, content, tome=None, releases=None):
        """Enregistre un nouveau fichier sous « <n°>_<nom> » et l'inscrit. Renvoie le nouvel enregistrement.

        tome : tome d'un fichier d'exigences (créé s'il est nouveau), d'un plan d'acceptation, ou d'un fichier STDHL
        ou de fiches (il doit exister, l'appelant le vérifie ; None : tous les tomes) ; releases : gardé pour les
        appelants, chaque fichier appartient à la release de la base.
        """
        if slot not in SLOTS:
            raise ValueError(f"type de fichier inconnu : {slot}")
        name = self.safe_name(name)
        releases = list(dict.fromkeys(r.strip().upper() for r in releases or [] if r.strip()))
        with self.sql.transaction() as db:
            # Le numéro vient de meta.next_file_id et n'est jamais réutilisé, même après une suppression : une copie
            # enregistrée a toujours un nouveau nom.
            file_id = int(self.sql.meta("next_file_id", 1))
            stored_name = f"{file_id:04d}_{name}"
            if tome:
                self.sql.ensure("tomes", [tome], db)
            self._insert(db, {"id": file_id, "slot": slot, "name": name, "stored_name": stored_name,
                              "tome": tome or None, "releases": releases})
            self.sql.set_meta("next_file_id", file_id + 1, db)
            # La copie est écrite dans la transaction : en cas d'échec, rien n'est inscrit.
            folder = self.folder(slot)
            folder.mkdir(parents=True, exist_ok=True)
            (folder / stored_name).write_bytes(content)
        return self.get(file_id)

    @staticmethod
    def to_ignore(db, ids, tome=None):
        """Numéros d'un fichier de fiches qui ne peuvent qu'être ignorés à l'import, sous la forme des lignes à
        régler : absents des exigences (du `tome` s'il est donné)."""
        # Les numéros connus : du tome s'il est donné, de toute la base sinon.
        own = {r[0] for r in db.execute("SELECT q.req_id FROM requirements q LEFT JOIN list_tomes t ON t.id = q.tome_id"
                                        + (" WHERE t.name = ?" if tome else ""), (tome,) if tome else ())}
        return [{"key": i, "requirements": [i], "places": [], "iadt": "", "moc": "", "delegation": "", "choose": [],
                 "unknown": True, "expected_mocs": [], "duplicate": [],
                 "issues": [f"exigence absente de la base des exigences{' du tome' if tome else ''} — seule possibilité, l'ignorer"]}
                for i in dict.fromkeys(ids) if i not in own]

    @staticmethod
    def check_ignored(found, decisions):
        """Chaque numéro de `found` (voir to_ignore) doit être ignoré par l'utilisateur (`decisions` {"rows":
        {numéro: {"ignore": true}}}), ou {"accept_all": true} en ligne de commande. ValueError sinon."""
        chosen = (decisions or {}).get("rows") or {}
        # Chaque numéro inconnu doit avoir été ignoré par l'utilisateur ; en ligne de commande, accept_all les ignore
        # tous.
        left = [c["key"] for c in found if not (chosen.get(c["key"]) or {}).get("ignore")]
        if left and not (decisions or {}).get("accept_all"):
            raise ValueError(f"{plural(len(left), 'numéro', 'numéros')} à ignorer avant le chargement (absents de la base "
                             f"des exigences) : {', '.join(left)}")

    @staticmethod
    def tome_requirements(db, tome_id):
        """{numéro d'exigence: id} des exigences d'un tome."""
        return {r["req_id"]: r["id"] for r in db.execute("SELECT id, req_id FROM requirements WHERE tome_id = ?", (tome_id,))}

    @staticmethod
    def skip_unknown(config, platform):
        """Vrai quand les exigences inconnues d'un fichier de fiches de `platform` sont ignorées au lieu de le refuser :
        seulement avec skip_unknown_req, et pour une autre plateforme que celle de niveau système. Un plan, un STDHL
        ou des fiches de la plateforme de niveau système qui citent une exigence inconnue sont toujours refusés."""
        if config is None or not config.get("skip_unknown_req"):
            return False
        return fold(platform or "") != fold(config.get("system_platform") or "")

    @staticmethod
    def unknown_refusal(unknown, tome=None):
        """Message du refus d'un fichier qui cite des numéros qui ne sont pas des exigences de son tome (de la base
        sans tome) : il ne serait pas utilisable tel quel ; None quand il n'y en a pas."""
        if not unknown:
            return None
        more = f" … et {len(unknown) - 30} autres" if len(unknown) > 30 else ""
        where = f"du tome {tome}" if tome else "de la base"
        what = "ne sont pas des exigences" if len(unknown) > 1 else "n'est pas une exigence"
        return (f"{plural(len(unknown), 'numéro')} qui {what} {where} : " + ", ".join(unknown[:30]) + more
                + " — corriger les numéros, ou importer d'abord ces exigences : fichier non chargé")

    @staticmethod
    def check_requirements(requirements):
        """Un plan, un STDHL ou un fichier de fiches cite des exigences : elles doivent être importées d'abord.
        ValueError sinon. Le tome ne se choisit qu'à l'import des exigences : un STDHL ou des fiches n'en ont pas, un
        plan prend celui des exigences qu'il cite."""
        if not requirements.tomes():
            raise ValueError("importer d'abord les Exigences")

    def group(self, file_id):
        """Numéros des fichiers supprimés ensemble avec `file_id` : pour un fichier d'exigences, tous les fichiers
        d'exigences de son tome (le premier import et ses compléments, un import en remplacement retirant les
        précédents) ; pour un autre fichier, lui seul."""
        record = self.get(file_id)
        if not record or record["slot"] != "requirements" or not record["tome"]:
            return [file_id]
        tome = record["tome"].casefold()
        return [r["id"] for r in self.list(slot="requirements") if (r["tome"] or "").casefold() == tome]

    def remove_older(self, record):
        """Retire les autres fichiers de même type et de même nom que `record` (et de même tome pour un fichier
        d'exigences ou un plan ; un fichier STDHL ou de fiches chargé à nouveau remplace l'ancien quel que soit son
        tome). Renvoie leurs numéros."""
        # Un fichier d'exigences ou un plan de même nom dans un autre tome est un autre document ; un fichier STDHL ou
        # de fiches chargé à nouveau remplace l'ancien même si son tome a changé.
        same_tome = record["slot"] in ("requirements", "plan")
        older = [r["id"] for r in self.list(slot=record["slot"]) if r["id"] != record["id"] and r["name"] == record["name"]
                 and (not same_tome or (r["tome"] or "").casefold() == (record["tome"] or "").casefold())]
        for file_id in older:
            self.remove(file_id)
        return older

    def remove(self, file_id, delete_file=True):
        """Retire un enregistrement (et par défaut sa copie enregistrée). Renvoie False s'il est inconnu.

        Les exigences importées du fichier et le contenu d'un plan sont retirés avec lui (cascade).
        """
        record = self.get(file_id)
        if record is None:
            return False
        with self.sql.transaction() as db:
            db.execute("DELETE FROM files WHERE id = ?", (file_id,))
        if delete_file:
            (self.folder(record["slot"]) / record["stored_name"]).unlink(missing_ok=True)
        # Le rapport de l'import part avec le fichier.
        self.report_path(file_id).unlink(missing_ok=True)
        return True
