"""Configuration de NEXUS-V, gardée dans un fichier JSON.

Le fichier est un objet JSON découpé en parties, une pour les paramètres généraux et une par type de fichier importé
:

    {
      "general": {"release": "A", "system_platform": "TOP", "server_port": 8000, ...},
      "requirements": {"requirement_regex": "...", "iadt_moc_rules": {...}, "headers_requirements": {...}},
      "plan": {"headers_plan": {...}},
      "stdhl": {"fiche_regex": "...", "stdhl_name_regex": "...", "headers_stdhl": {...}},
      "fiches": {"test_tag": "test", "test_case_tag": "test_case", "requirement_tag": "requirements"}
    }

Types : texte, entier, décimal, booléen, liste de textes, dictionnaire texte -> texte.

Le référentiel commun (tomes, plateformes, iadt, mocs), leurs couleurs et les descriptions des tomes ne sont pas dans
ce fichier : ils sont gardés dans la base SQL (nexus_db.py, fichier nexus-v.db) quand une base est donnée, et restent
accessibles ici comme les autres variables (get, update, to_dict). Une nouvelle base reçoit le référentiel par
défaut, une fois. La release, une seule, est le réglage general.release : « releases » ne fait que la donner. Une
variable absente du fichier prend sa valeur par défaut (DEFAULT_CONFIG), puis y est ajoutée.

Variables :
  slogan                          slogan affiché sous le nom NEXUS-V
  server_port, open_browser       serveur.py
  chart_max_parts                 parts au plus d'un camembert de la page Vérification, Non couvert
                                  non compté (toujours montré) ; au-delà : « Autres »
  chart_max_parts_platforms       la même chose pour les camemberts à une part par plateforme
  requirement_regex               forme d'un numéro d'exigence
  requirements_tome_regex         tome dans le nom d'un fichier d'exigences : proposé à l'import
  fiche_regex                     nom d'une fiche (test) dans un paragraphe d'un STDHL
  stdhl_name_regex                nom du STDHL dans le nom de son fichier
  stdhl2_*                        STDHL de format 2 (un document, sans tableau STDHL) :
    stdhl2_test_regex             titre d'un test, groupe 1 = nom du test (« TEST-001 : ZOULOU »)
    stdhl2_requirements_label     libellé qui ouvre la liste « Exigences associées » d'un test (la liste la termine)
    stdhl2_release_label          libellé de la release d'un test, lue par stdhl2_release_regex (groupe 1)
    stdhl2_test_case_regex        début d'un test case : les exigences de ses phases sont la couverture du test
  release_separator_regex         séparateur des releases dans une cellule d'un fichier d'exigences
  stdhl_separator_regex           séparateur de plusieurs valeurs dans une cellule STDHL (plusieurs valeurs : à
  régler)
  plan_tome_regex                 tome dans le nom d'un fichier de plan : présélectionné au chargement
  code_separator_regex            séparateur des codes IADT et MOC dans une cellule d'un plan
  delegation_separator_regex      séparateur de plusieurs plateformes dans la cellule Délégation (à régler à
  l'import)
  iadt_moc_rules                  MOC attendus pour chaque IADT, dans le fichier d'exigences (un autre MOC le refuse)
  et
                                  dans un plan : au moins un ; un autre MOC, ou aucun, est à régler à l'import
  test_tag                        balise XML d'un test dans les fiches ; elle contient
  test_case_tag                   la balise d'un test case, qui contient
  requirement_tag                 la balise des exigences testées
                                  (pour fiche_regex et stdhl_name_regex : groupe 1 s'il y en a un, sinon toute la
                                  correspondance)
  system_platform                 la plateforme de niveau système (paramètre général) : les STDHL s'appliquent à
  elle,
                                  seules les exigences qui lui sont déléguées sont attendues dans un STDHL, seules ses
                                  fiches sont vérifiées par rapport aux tests du STDHL. Non définie par défaut (aucune
                                  exigence attendue dans un STDHL) ; seulement une plateforme des délégations des
                                  exigences importées
  use_requirement_stdhl           prendre en compte la colonne STDHL des exigences : une exigence sans valeur n'est
  pas
                                  vérifiée à l'étape STDHL (NON DEMANDÉ) ; sans cette option, la colonne n'est pas lue
  release                         release unique de la base (ne change que tant qu'aucun fichier n'est chargé)
  tomes                           tomes auxquels appartiennent les fichiers d'exigences
  platforms                       plateformes de test (délégations)
  iadt                            codes IADT -> signification
  mocs                            codes MOC -> description
  headers_requirements            en-têtes attendus dans le fichier d'exigences (ligne 1)
  headers_plan                    en-têtes attendus dans les tableaux des plans d'acceptation
  headers_stdhl                   en-têtes attendus dans les tableaux des STDHL
                                  Chacun est une liste ordonnée « texte d'en-tête -> rôle » : le rôle dit quelle
                                  colonne NEXUS-V lit (id, title, releases, validation, stdhl, iadt, moc, delegation ;
                                  exigence, iadt, moc, delegation ; requirements) : ces colonnes sont obligatoires.
                                  Les autres colonnes sont « check » (vérifiées pour valider le tableau) ou « ignore »
                                  (seulement documentées) ; elles peuvent être ajoutées ou retirées. Un en-tête est
                                  reconnu quand il contient le texte attendu, casse et accents ignorés.
"""
import copy
import json
import os
import re
import threading
import unicodedata
from pathlib import Path

DEFAULT_CONFIG = {
    "slogan": "Le fil d’Ariane de la traçabilité, de la spécification à l'essai",
    "server_port": 8000,
    "open_browser": True,
    "chart_max_parts": 6,
    "chart_max_parts_platforms": 20,   # camemberts par plateforme : 20 plateformes au plus (hors parts réservées)
    # Deux formes, regroupées pour que \b borne les deux : SSS_[T<n°>_]<fonctionnalités>_REQ_<n°> (la partie T<n°>_
    # facultative), et HQ1/2_WF_<n°>[_REQ_<n°>] (la barre oblique fait partie du nom).
    "requirement_regex": r"\b(?:SSS_(?:T\d+_)?[A-Za-z0-9_-]+?_REQ_\d+|HQ1/2_WF_\d+(?:_REQ_\d+)?)\b",
    "requirements_tome_regex": r"(?i)^\d{4}-\d{2}-\d{2}[-_](.+?)(?:[-_]|\.xls)",
    # STDHL : fiche de test SYS_<type>_<fonctionnalités>_<n°> en début de paragraphe (éventuellement après une
    # numérotation), paragraphe TEST CASE, et nom du STDHL dans [date]_[texte]_STDHL_[Nom du STDHL]_rev[texte].docx.
    "fiche_regex": r"^\s*(?:[\d.]+\s+)?(SYS_[A-Za-z0-9-]+_[A-Za-z0-9_-]+?_\d+)(?![A-Za-z0-9_])",
    # STDHL de format 2 : pas de tableau, un titre par test, puis des parties étiquetées (libellés : casse et accents
    # ignorés).
    "stdhl2_test_regex": r"^\s*([A-Za-z][A-Za-z0-9_-]*?[_-]\d+)\s*:",
    "stdhl2_requirements_label": "Exigences associées",
    "stdhl2_release_label": "Applicabilité",
    "stdhl2_release_regex": r"(?i)release\s+(\S+)",
    "stdhl2_test_case_regex": r"(?i)^\s*TEST CASE\s*#\s*\d+",
    "stdhl_name_regex": r"(?i)STDHL_(.+?)_rev",
    # Dossier des fichiers importés (relatif au dossier de NEXUS-V) et son sous-dossier pour chaque type de fichier.
    "storage_dir": ".",
    "requirements_folder": "exigences",
    "plan_folder": "plans",
    "stdhl_folder": "stdhl",
    "fiches_folder": "fiches",
    "traceability_folder": "tracabilite",
    # Séparateurs dans une cellule, en expressions régulières.
    "release_separator_regex": r",",
    "stdhl_separator_regex": r"[,;/\n]+",
    "plan_tome_regex": r"(?i)_PA_(.+?)_",
    "code_separator_regex": r"[,;+\s]+",
    "delegation_separator_regex": r"[,;/+\n]+",
    # Règles IADT / MOC : pour chaque code IADT, les codes MOC attendus (au moins l'un d'eux, et aucun autre).
    "iadt_moc_rules": {"D": ["MOC4"], "T": ["MOC4"], "I": ["MOC7"], "A": ["MOC2", "MOC3"]},
    # Fiches d'Essai (XML) : un test contient des test cases, qui contiennent les exigences testées.
    "test_tag": "test",
    "test_case_tag": "test_case",
    "requirement_tag": "requirements",
    # La colonne STDHL des exigences nomme le STDHL où une exigence doit se trouver (True), ou n'est pas lue (False).
    "use_requirement_stdhl": True,
    "system_platform": "",
    # La release sur laquelle travaille la base : chaque exigence importée et chaque fichier chargé lui appartient.
    "release": "",
    "tomes": [],
    # Les méthodes de test s'écrivent « V&V [PLATEFORME] : ... ».
    "platforms": ["TOP"],
    "iadt": {"I": "Inspection", "A": "Analyse", "D": "Démonstration", "T": "Test"},
    "mocs": {f"MOC{i}": "" for i in range(10)},
    # En-têtes attendus des tableaux, dans l'ordre : texte d'en-tête -> rôle (« check » = colonne seulement vérifiée).
    # Une cellule d'en-tête doit contenir le texte (casse et accents ignorés).
    "headers_requirements": {"ID": "id", "Titre": "title", "Specification releases": "releases",
                             "Validation releases": "validation", "STDHL": "stdhl", "IADT": "iadt", "MOC": "moc",
                             "Délégation": "delegation"},
    "headers_plan": {"Exigence": "exigence", "Méthode": "ignore", "IADT": "iadt", "MOC": "moc", "Délégation": "delegation"},
    "headers_stdhl": {"Description du test": "check", "Résultat attendu": "check", "Exigences associées": "requirements"},
}

# Rôles des colonnes de chaque réglage d'en-têtes : chacun doit être donné exactement une fois.
HEADER_ROLES = {
    "headers_requirements": ("id", "title", "releases", "validation", "stdhl", "iadt", "moc", "delegation"),
    "headers_plan": ("exigence", "iadt", "moc", "delegation"),
    "headers_stdhl": ("requirements",),
}

# Parties du fichier et leurs variables, toujours écrites : le fichier montre chaque réglage, et toutes les règles pour
# décoder les fichiers importés. Les autres variables (réglages de l'installation) restent à la racine du fichier.
# Une partie fait correspondre les noms écrits dans le fichier aux variables : dans une partie, un nom n'a pas besoin de
# préfixe (« folder »).
SECTIONS = {
    "general": {n: n for n in ("release", "system_platform", "slogan", "server_port", "open_browser", "chart_max_parts",
                               "chart_max_parts_platforms", "storage_dir")},
    "requirements": {"folder": "requirements_folder",
                     **{n: n for n in ("requirement_regex", "requirements_tome_regex", "release_separator_regex",
                                       "stdhl_separator_regex", "iadt_moc_rules",
                                       "headers_requirements")}},
    "plan": {"folder": "plan_folder",
             **{n: n for n in ("plan_tome_regex", "code_separator_regex", "delegation_separator_regex",
                               "headers_plan")}},
    "stdhl": {"folder": "stdhl_folder",
              **{n: n for n in ("use_requirement_stdhl", "fiche_regex", "stdhl_name_regex", "headers_stdhl",
                                "stdhl2_test_regex", "stdhl2_requirements_label", "stdhl2_release_label",
                                "stdhl2_release_regex", "stdhl2_test_case_regex")}},
    "fiches": {"folder": "fiches_folder", **{n: n for n in ("test_tag", "test_case_tag", "requirement_tag")}},
    "traceability": {"folder": "traceability_folder"},
}
SECTION_VARIABLES = tuple(var for names in SECTIONS.values() for var in names.values())

# Variables qui contiennent le nom d'un sous-dossier du dossier des fichiers.
FOLDER_VARIABLES = ("requirements_folder", "plan_folder", "stdhl_folder", "fiches_folder", "traceability_folder")

# Variables qui contiennent le nom d'une balise XML.
TAG_VARIABLES = ("test_tag", "test_case_tag", "requirement_tag")

# Variables qui contiennent une expression régulière.
REGEX_VARIABLES = ("requirement_regex", "requirements_tome_regex", "fiche_regex", "stdhl_name_regex", "plan_tome_regex",
                   "release_separator_regex", "stdhl_separator_regex", "code_separator_regex",
                   "delegation_separator_regex", "stdhl2_test_regex", "stdhl2_release_regex", "stdhl2_test_case_regex")

# Variables gardées dans la base SQL (nexus_db.py) plutôt que dans le fichier JSON.
REFERENTIALS = ("tomes", "platforms", "iadt", "mocs")

_MISSING = object()


def fold(text):
    """Forme comparable d'un libellé : sans accents, en minuscules, espaces simples."""
    # NFD sépare une lettre accentuée en la lettre et son accent (un caractère combinant), qui est ensuite laissé de
    # côté.
    text = unicodedata.normalize("NFD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text).strip().lower()


METHOD_PLATFORM = re.compile(r"^\s*V\s*&\s*V\s*\[\s*([^\]]*?)\s*\]", re.I)


def storage_path(config, base):
    """Dossier des fichiers importés : general.storage_dir, relatif à `base` (le dossier qui contient le dossier
    base, fichiers-references pour NEXUS-V ; « . » est ce dossier lui-même) sauf s'il est absolu."""
    folder = Path(config.get("storage_dir"))
    return (folder if folder.is_absolute() else Path(base) / folder).resolve()


def plural(n, singular, several=None):
    """« 1 ligne », « 3 lignes » (singulier jusqu'à 1, comme en français)."""
    return f"{n} {(several or singular + 's') if n > 1 else singular}"


def header_column(cells, expected, taken=()):
    """Index de la cellule d'en-tête qui correspond au texte attendu (casse et accents ignorés), où qu'elle soit :
    une cellule égale à lui d'abord, sinon une cellule où il commence un mot (« Titre de l'exigence », « Releases » ;
    pas « ID » dans « Validation »), la plus courte ; les cellules `taken` par une autre colonne sont sautées. None
    quand il n'y en a pas."""
    key = fold(expected)
    folded = [fold(str(c or "")).strip() for c in cells]
    free = [i for i in range(len(folded)) if i not in taken]
    exact = next((i for i in free if folded[i] == key), None)
    if exact is not None or not key:
        return exact
    word = re.compile(r"(?<![a-z0-9])" + re.escape(key))
    found = [i for i in free if word.search(folded[i])]
    return min(found, key=lambda i: len(folded[i])) if found else None


def match_header(cells, headers):
    """Colonnes d'une ligne d'en-tête pour un réglage d'en-têtes (texte -> rôle) : ({rôle: index de colonne} pour les
    colonnes trouvées qui ont un rôle, [textes attendus non trouvés]). L'ordre des colonnes n'importe pas : les
    cellules égales à un texte attendu sont prises d'abord, puis les autres, chaque cellule pour une seule colonne."""
    cols, missing, taken = {}, [], set()
    wanted = [(text, role) for text, role in headers.items() if role != "ignore"]   # « ignore » : seulement documentée
    found = {}
    for text, role in sorted(wanted, key=lambda w: header_column(cells, w[0]) is None or
                             fold(str(cells[header_column(cells, w[0])] or "")).strip() != fold(w[0])):
        k = header_column(cells, text, taken)
        if k is not None:
            taken.add(k)
        found[text] = k
    for text, role in wanted:
        k = found[text]
        if k is None:
            missing.append(text)
        elif role not in ("check", ""):
            cols[role] = k
    return cols, missing


def table_left_out(number, missing):
    """Message d'un tableau d'un document Word laissé de côté parce qu'il manque des textes attendus à son en-tête."""
    texts = "en-têtes attendus introuvables" if len(missing) > 1 else "en-tête attendu introuvable"
    return f"tableau {number} : {texts} : " + ", ".join(f"« {m} »" for m in missing)


def header_text(headers, role):
    """Texte d'en-tête attendu d'un rôle dans un réglage d'en-têtes."""
    return next(text for text, r in headers.items() if r == role)


def method_platform(text):
    """Plateforme nommée dans une méthode de test « V&V [PLATEFORME] : ... », ou None quand la méthode a une autre
    forme."""
    m = METHOD_PLATFORM.match(text or "")
    return m[1] if m else None


def type_name(value):
    """Nom du type d'une valeur (bool est testé avant int : True est aussi un int)."""
    for t, name in ((bool, "bool"), (int, "int"), (float, "float"), (str, "str"), (list, "list"), (dict, "dict")):
        if isinstance(value, t):
            return name
    raise ValueError(f"type non géré : {type(value).__name__}")


class Configuration:
    """Lit, écrit, donne et fixe les variables de configuration gardées dans un fichier JSON."""

    def __init__(self, path, defaults=None, database=None):
        """database : la NexusDatabase qui contient le référentiel (tomes, plateformes, iadt, mocs) ; sans elle, ils
        sont gardés dans le fichier JSON comme les autres variables."""
        path = Path(path)
        self.database = database
        self.path = path.with_suffix(".json")
        self.defaults = copy.deepcopy(DEFAULT_CONFIG if defaults is None else defaults)
        self._values = {}             # variables lues dans le fichier ou fixées dedans
        self._lock = threading.RLock()
        self.read()

    # ------------------------------------------------------------------ #
    # Lecture / écriture du fichier JSON
    # ------------------------------------------------------------------ #

    def read(self):
        """(Re)charge depuis le disque ; un fichier absent veut dire que chaque variable a sa valeur par défaut."""
        with self._lock:
            self._values = {}
            changed = False
            if self.path.is_file():
                data = self._load_json()
                # Le fichier est lu partie par partie ; dans une partie, un nom peut être court (« folder ») et
                # correspond à sa variable (« requirements_folder »). Chaque valeur est validée : une erreur faite à la
                # main est signalée au démarrage.
                for name, value in data.items():
                    if name in SECTIONS:
                        for key, v in value.items():
                            var = SECTIONS[name].get(key, key)
                            self._values[var] = self._check(var, v)
                    else:
                        self._values[name] = self._check(name, value)
            self._default_referential()
            # Chaque variable d'une partie est toujours écrite dans le fichier, sa valeur par défaut comprise : le
            # fichier documente chaque règle utilisée pour lire les documents importés.
            for name in SECTION_VARIABLES:
                if name in self.defaults and name not in self._values:
                    self._values[name] = copy.deepcopy(self.defaults[name])
                    changed = True
            if changed:
                self.write()
            if self.database is not None:
                self.database.release = self.get("release")

    def _default_referential(self):
        """Une base juste créée à partir du schéma reçoit les référentiels par défaut, une fois ; ensuite, seul
        l'utilisateur les change (un référentiel vidé reste vide)."""
        if self.database is None:
            return
        if self.database.created:
            self.database.created = False
            for name in REFERENTIALS:
                if name in self.defaults:
                    self.database.set(name, self._check(name, self.defaults[name]))

    def write(self):
        """Enregistre sur le disque de façon atomique (fichier temporaire, puis renommage). Seules les variables
        fixées sont écrites."""
        with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # Écrit dans un fichier temporaire puis renommé : un arrêt brutal pendant l'écriture ne laisse jamais un
            # config.json à moitié écrit.
            tmp = self.path.with_suffix(".tmp")
            parts = {part: {key: self._values[var] for key, var in names.items() if var in self._values}
                     for part, names in SECTIONS.items()}
            others = {n: v for n, v in self._values.items() if n not in SECTION_VARIABLES}
            data = {**{part: values for part, values in parts.items() if values}, **others}
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, self.path)

    def _load_json(self):
        """Variables du fichier JSON. ValueError quand ce n'est pas un objet JSON valide."""
        try:
            values = json.loads(self.path.read_text(encoding="utf-8"))
        except ValueError as e:
            raise ValueError(f"{self.path} : JSON invalide ({e})") from None
        if not isinstance(values, dict):
            raise ValueError(f"{self.path} : un objet JSON est attendu")
        return values

    # ------------------------------------------------------------------ #
    # Accès
    # ------------------------------------------------------------------ #

    def names(self):
        """Chaque nom de variable : celles par défaut d'abord, puis celles en plus fixées dans le fichier."""
        with self._lock:
            return list(dict.fromkeys([*self.defaults, *self._values]))

    def __contains__(self, name):
        """True pour une variable fixée ou qui a une valeur par défaut."""
        return name in self._values or name in self.defaults

    def get(self, name, default=_MISSING):
        """Valeur d'une variable (copie) : valeur fixée, sinon valeur par défaut, sinon `default`, sinon KeyError.
        Les référentiels sont lus dans la base quand il y en a une."""
        # Valeurs gardées dans la base plutôt que dans le fichier JSON.
        if name == "tome_descriptions":   # ce que contient chaque tome, dans la base : {tome: texte}
            return self.database.tome_descriptions() if self.database is not None else {}
        if name == "colors":   # couleur de chaque valeur des référentiels, dans la base : {type: {nom: couleur}}
            return {kind: self.database.colors(kind) for kind in REFERENTIALS} if self.database is not None else {}
        if name == "releases":   # les releases déclarées : celle de la base
            release = self.get("release")
            return {release: ""} if release else {}
        if self.database is not None and name in REFERENTIALS:
            return self.database.get(name)
        with self._lock:
            if name in self._values:
                return copy.deepcopy(self._values[name])
            if name in self.defaults:
                return copy.deepcopy(self.defaults[name])
        if default is _MISSING:
            raise KeyError(name)
        return default

    def to_dict(self):
        """Chaque variable avec sa valeur actuelle, référentiels et leurs couleurs compris."""
        return {**{name: self.get(name) for name in self.names()}, "colors": self.get("colors"),
                "tome_descriptions": self.get("tome_descriptions"), "system_platforms": self.system_platforms()}

    def system_platforms(self):
        """Les plateformes que peut être la plateforme de niveau système : celles auxquelles les exigences importées
        sont déléguées."""
        return self.database.requirement_platforms() if self.database is not None else []

    def set(self, name, value):
        """Fixe une variable après validation et enregistre."""
        self.update({name: value})

    def update(self, values):
        """Fixe plusieurs variables après validation et enregistre une fois. ValueError si l'une est invalide.

        Une variable fixée à sa valeur par défaut n'est pas écrite dans le fichier, pour qu'un changement ultérieur
        de la valeur par défaut s'applique encore à elle.
        """
        values = dict(values)
        # Valeurs que la page envoie en plus des variables : les renommages, les couleurs et les descriptions vont dans
        # la base.
        renames = values.pop("renames", None) or {}   # {type: {ancien nom: nouveau nom}}, voir NexusDatabase.rename
        colors = values.pop("colors", None) or {}   # {type: {nom: "#rrggbb"}}
        descriptions = values.pop("tome_descriptions", None) or {}   # {tome: texte}
        if not isinstance(descriptions, dict):
            raise ValueError("descriptions des tomes : format invalide")
        if any(kind not in REFERENTIALS or not isinstance(v, dict) for kind, v in colors.items()):
            raise ValueError("couleurs : liste inconnue")
        bad = [c for v in colors.values() for c in v.values() if not re.fullmatch(r"#[0-9a-fA-F]{6}", str(c))]
        if bad:
            raise ValueError(f"couleur invalide : {bad[0]}")
        # Chaque valeur est validée avant toute écriture : une valeur invalide refuse toute la mise à jour.
        clean = {name: self._check(name, value) for name, value in values.items()}
        if "release" in clean and clean["release"] != self.get("release") and self.database is not None \
                and self.database.scalar("SELECT COUNT(*) FROM files"):
            raise ValueError("la release ne peut changer que si aucun fichier n'est chargé : supprimer d'abord les fichiers")
        # Un code IADT ou MOC renommé garde ses règles.
        # Les règles IADT -> MOC sont écrites avec des codes : un code renommé l'est aussi dans les règles.
        moved = {kind: {o.upper(): n.upper() for o, n in (renames.get(kind) or {}).items()} for kind in ("iadt", "mocs")}
        if any(moved.values()):
            rules = clean.get("iadt_moc_rules", self.get("iadt_moc_rules"))
            clean["iadt_moc_rules"] = {moved["iadt"].get(i, i): [moved["mocs"].get(m, m) for m in mocs]
                                       for i, mocs in rules.items()}
        # La plateforme de niveau système suit le renommage de sa plateforme ; quand elle est fixée, ce doit être l'une
        # des plateformes auxquelles les exigences importées sont déléguées.
        moved_platforms = renames.get("platforms") or {}
        system_platform = clean.get("system_platform", self.get("system_platform")) or ""
        system_platform = moved_platforms.get(system_platform, system_platform)
        if system_platform:
            delegated = [moved_platforms.get(p, p) for p in self.system_platforms()]
            if system_platform not in delegated:
                raise ValueError(f"plateforme de niveau système « {system_platform} » : ce n'est la délégation d'aucune "
                                 "exigence importée (choisir parmi " + (", ".join(delegated) or "aucune") + ")")
        clean["system_platform"] = system_platform
        # Deux types de fichiers dans le même sous-dossier mélangeraient leurs copies.
        folders = [clean.get(name, self.get(name)) for name in FOLDER_VARIABLES]
        same = sorted({f for f in folders if [x.casefold() for x in folders].count(f.casefold()) > 1})
        if same:
            raise ValueError("chaque type de fichier doit avoir son propre sous-dossier : « " + "», « ".join(same)
                             + " » est utilisé plusieurs fois")
        if self.database is not None:
            if renames:
                self.database.rename(renames)
            # Les référentiels vont dans la base, en une transaction, avant l'écriture du fichier JSON.
            referentials = {name: clean.pop(name) for name in REFERENTIALS if name in clean}
            if referentials:
                self.database.set_many(referentials)
            if colors:
                self.database.set_colors(colors)
            if descriptions:
                self.database.set_tome_descriptions(descriptions)
        with self._lock:
            # Les variables restantes vont dans le fichier JSON.
            for name, value in clean.items():
                if name in self.defaults and value == self.defaults[name] and name not in SECTION_VARIABLES:
                    self._values.pop(name, None)
                else:
                    self._values[name] = value
            self.write()
            if self.database is not None:
                self.database.release = self.get("release")
            return self.to_dict()

    def reset(self, name):
        """Oublie la valeur fixée d'une variable (elle reprend sa valeur par défaut) et enregistre."""
        with self._lock:
            if self._values.pop(name, _MISSING) is not _MISSING:
                self.write()

    @staticmethod
    def _check_rules(value):
        """Règles validées {code IADT: [codes MOC attendus]} : codes en majuscules, chaque IADT une fois, au moins un
        MOC."""
        if not isinstance(value, dict):
            raise ValueError("iadt_moc_rules doit associer chaque IADT à une liste de MOC")
        rules = {}
        for iadt, mocs in value.items():
            code = str(iadt).strip().upper()
            if not code:
                continue
            if code in rules:
                raise ValueError(f"règle en double pour l'IADT {code}")
            if not isinstance(mocs, list):
                raise ValueError(f"règle de l'IADT {code} : liste de MOC attendue")
            expected = list(dict.fromkeys(str(m).strip().upper() for m in mocs if str(m).strip()))
            if not expected:
                raise ValueError(f"règle de l'IADT {code} : choisir au moins un MOC")
            rules[code] = expected
        return rules

    def _check(self, name, value):
        """Copie validée d'une valeur : même type que la valeur par défaut, listes de plateformes et de codes
        nettoyées."""
        if not isinstance(name, str) or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
            raise ValueError(f"nom de variable invalide : « {name} »")
        # Le type doit être celui de la valeur par défaut (un int est accepté pour un float).
        kind = type_name(value)
        if name in self.defaults:
            expected = type_name(self.defaults[name])
            if expected == "float" and kind == "int":
                value, kind = float(value), "float"
            if kind != expected:
                raise ValueError(f"{name} doit être de type {expected}, pas {kind}")
        if name == "iadt_moc_rules":
            return self._check_rules(value)
        if kind in ("list", "dict"):
            items = value.values() if kind == "dict" else value
            if any(isinstance(v, (list, dict)) for v in items):
                raise ValueError(f"{name} : une liste ou un dictionnaire ne peut contenir que des valeurs simples")
        if kind == "list":
            value = [str(v).strip() for v in value if str(v).strip()]
        elif kind == "dict":
            value = {str(k).strip(): str(v or "").strip() for k, v in value.items() if str(k).strip()}
        elif kind == "str":
            value = value.strip()

        # Vérifications propres à certaines variables : doublons dans les référentiels, rôles d'en-tête présents une
        # fois chacun, noms de dossier et de balise valides, expressions régulières compilables, plages de nombres.
        if name in ("platforms", "tomes"):
            keys = [fold(n) for n in value]
            duplicates = sorted({n for n in value if keys.count(fold(n)) > 1})
            if duplicates:
                raise ValueError(f"{'plateformes' if name == 'platforms' else 'tomes'} en double : {', '.join(duplicates)}")
        elif name in ("iadt", "mocs"):
            codes = {}
            for code, label in value.items():
                code = code.upper()
                if code in codes:
                    raise ValueError(f"code en double dans {name} : {code}")
                codes[code] = label
            value = codes
        elif name in HEADER_ROLES:
            roles = HEADER_ROLES[name]
            # {texte d'en-tête: rôle} ; une colonne sans rôle lu est vérifiée (« check ») ou seulement documentée («
            # ignore »).
            unknown = sorted({r for r in value.values() if r not in roles and r not in ("check", "ignore")})
            if unknown:
                raise ValueError(f"{name} : {'rôles inconnus' if len(unknown) > 1 else 'rôle inconnu'} : {', '.join(unknown)}")
            for role in roles:
                count = list(value.values()).count(role)
                if count != 1:
                    raise ValueError(f"{name} : la colonne « {role} » doit apparaître une fois ({plural(count, 'trouvée')})")
            keys = [fold(t) for t in value]
            duplicates = sorted({t for t in value if keys.count(fold(t)) > 1})
            if duplicates:
                raise ValueError(f"{name} : {'en-têtes' if len(duplicates) > 1 else 'en-tête'} en double : {', '.join(duplicates)}")
        elif name in FOLDER_VARIABLES:
            if not re.fullmatch(r"[^\\/:*?\"<>|.][^\\/:*?\"<>|]*", value):
                raise ValueError(f"nom de dossier invalide ({name}) : « {value} »")
        elif name in TAG_VARIABLES:
            if not re.fullmatch(r"[^\W\d][\w.-]*", value):
                raise ValueError(f"nom de balise XML invalide ({name}) : « {value} »")
        elif name in REGEX_VARIABLES:
            try:
                re.compile(value)
            except re.error as e:
                raise ValueError(f"expression régulière invalide ({name}) : {e}") from None
        elif name == "server_port" and not 1 <= value <= 65535:
            raise ValueError("server_port doit être compris entre 1 et 65535")
        elif name == "release":
            value = value.upper()
            if value and not re.fullmatch(r"[^\s,;]+", value):
                raise ValueError(f"nom de release invalide : « {value} »")
        elif name in ("chart_max_parts", "chart_max_parts_platforms") and not 2 <= value <= 50:
            raise ValueError(f"{name} doit être compris entre 2 et 50")
        return copy.deepcopy(value)

    # ------------------------------------------------------------------ #
    # Vérifications utilisées sur les tableaux des plans d'acceptation
    # ------------------------------------------------------------------ #

    def find_platform(self, text):
        """Plateforme configurée d'une méthode de test, ou None.

        « V&V [PLATEFORME] : ... » nomme la plateforme entre crochets ; sinon la méthode doit commencer par un nom de
        plateforme (le plus long suivi d'un caractère non alphanumérique).
        """
        # « V&V [TOP] : ... » nomme la plateforme entre crochets ; sinon, le plus long nom de plateforme par lequel
        # commence la méthode (suivi d'un séparateur, pour que « TOP2 » ne soit pas lu « TOP »).
        named = method_platform(text)
        if named is not None:
            return next((p for p in self.get("platforms") if fold(p) == fold(named)), None)
        key = fold(text)
        best = None
        for name in self.get("platforms"):
            p = fold(name)
            if p and key.startswith(p) and (len(key) == len(p) or not key[len(p)].isalnum()):
                if best is None or len(p) > len(fold(best)):
                    best = name
        return best

    def is_platform(self, name):
        """True pour une plateforme de la configuration (casse et accents ignorés)."""
        return any(fold(name) == fold(p) for p in self.get("platforms"))

    def delegation(self, value):
        """Délégation écrite dans une cellule (vide quand la cellule l'est : il n'y a pas de plateforme par défaut)."""
        return (value or "").strip()
