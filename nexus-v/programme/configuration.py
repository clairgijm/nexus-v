"""Configuration of NEXUS-V, kept in a JSON file.

The file is a JSON object split into parts, one for the general settings and one per type of imported file:

    {
      "general": {"server_port": 8000, "open_browser": true, ...},
      "requirements": {"requirement_regex": "...", "headers_requirements": {...}},
      "plan": {"headers_plan": {...}},
      "stdhl": {"fiche_regex": "...", "stdhl_name_regex": "...", "headers_stdhl": {...}},
      "fiches": {"test_tag": "test", "test_case_tag": "test_case", "requirement_tag": "requirements"}
    }

Types: text, integer, decimal, boolean, list of texts, dictionary text -> text.

The common referential (tomes, platforms, iadt, mocs), their colours, the default platform and the descriptions of
the tomes are not in this file: they are kept in the SQL database (nexus_db.py, file nexus-v.db) when a database is
given, and stay reachable here like the other variables (get, update, to_dict). A new base gets the default
referential once. The release, a single one, is the setting general.release: "releases" only gives it. A variable
missing from the file takes its default value (DEFAULT_CONFIG), and is then added to it.

Variables:
  slogan                          slogan shown under the name NEXUS-V
  server_port, open_browser       serveur.py
  chart_max_parts                 most parts of a pie chart of the page Vérification, Non couvert, N/A and Absent
                                  not counted (always shown); beyond: « Autres »
  chart_max_parts_platforms       the same for the charts with one part per platform
  requirement_regex               form of a requirement number
  requirements_tome_regex         tome in the name of a requirements file: offered at import
  fiche_regex                     name of a fiche (test) in a paragraph of a STDHL
  stdhl_name_regex                name of the STDHL in the name of its file
  stdhl2_*                        STDHL of format 2 (one document, no STDHL table), read by the check before import:
    stdhl2_test_regex             title of a test, group 1 = name of the test (« TEST-001 : ZOULOU »)
    stdhl2_requirements_label     label opening the list « Exigences associées » of a test (the list ends with it)
    stdhl2_release_label          label of the release of a test, stdhl2_release_regex reading it (group 1)
    stdhl2_test_case_regex        start of a test case: the requirements of its phases are the coverage of the test
  release_separator_regex         separator of the releases in a cell of a requirements file
  stdhl_separator_regex           separator of several values in a STDHL cell (several values: to settle)
  plan_tome_regex                 tome in the name of a plan file: preselected at load
  code_separator_regex            separator of the IADT and MOC codes in a cell of a plan
  delegation_separator_regex      separator of several platforms in the Délégation cell (to settle at import)
  iadt_moc_rules                  MOC expected for each IADT of a plan: at least one; another MOC or none is settled
                                  at import
  test_tag                        XML tag of a test in the fiches; it contains
  test_case_tag                   the tag of a test case, which contains
  requirement_tag                 the tag of the requirements tested
                                  (for fiche_regex and stdhl_name_regex: group 1 if any, else the whole match)
  use_requirement_stdhl           take the column STDHL of the requirements into account: a requirement without a
                                  value is not checked at the STDHL stage (NON DEMANDÉ)
  release                         single release of the base (changes only while no file is loaded)
  tomes                           tomes the requirements files belong to
  platforms                       test platforms (delegation of the plan rows)
  iadt                            IADT codes -> meaning
  mocs                            MOC codes -> description
  default_delegation              delegation of an empty cell: the platform marked is_default in the base
  headers_requirements            headers expected in the requirements file (row 1)
  headers_plan                    headers expected in the tables of the acceptance plans
  headers_stdhl                   headers expected in the tables of the STDHL
                                  Each one is an ordered list « header text -> role »: the role says which column
                                  NEXUS-V reads and stores (id, title, releases, stdhl; exigence, iadt, moc,
                                  delegation; requirements): those columns are required. The other columns are
                                  « check » (checked to validate the table) or « ignore » (only documented); they
                                  may be added or removed. A header is recognised when it contains the expected
                                  text, case and accents ignored.
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
    "chart_max_parts_platforms": 20,   # pie charts by platform: up to 20 platforms (reserved parts not counted)
    # Two forms, grouped so that \b bounds both: SSS_[T<n°>_]<fonctionnalités>_REQ_<n°> (the T<n°>_ part optional), and
    # HQ1/2_WF_<n°>[_REQ_<n°>] (the slash is part of the name).
    "requirement_regex": r"\b(?:SSS_(?:T\d+_)?[A-Za-z0-9_-]+?_REQ_\d+|HQ1/2_WF_\d+(?:_REQ_\d+)?)\b",
    "requirements_tome_regex": r"(?i)^\d{4}-\d{2}-\d{2}[-_](.+?)(?:[-_]|\.xls)",
    # STDHL: test sheet SYS_<type>_<fonctionnalités>_<n°> at the start of a paragraph (possibly after a numbering),
    # TEST CASE paragraph, and STDHL name in [date]_[texte]_STDHL_[Nom du STDHL]_rev[texte].docx.
    "fiche_regex": r"^\s*(?:[\d.]+\s+)?(SYS_[A-Za-z0-9-]+_[A-Za-z0-9_-]+?_\d+)(?![A-Za-z0-9_])",
    # STDHL of format 2: no table, one title per test, then labelled sections (labels: case and accents ignored).
    "stdhl2_test_regex": r"^\s*([A-Za-z][A-Za-z0-9_-]*?[_-]\d+)\s*:",
    "stdhl2_requirements_label": "Exigences associées",
    "stdhl2_release_label": "Applicabilité",
    "stdhl2_release_regex": r"(?i)release\s+(\S+)",
    "stdhl2_test_case_regex": r"(?i)^\s*TEST CASE\s*#\s*\d+",
    "stdhl_name_regex": r"(?i)STDHL_(.+?)_rev",
    # Folder of the imported files (relative to the folder of NEXUS-V) and its sub-folder for each type of file.
    "storage_dir": ".",
    "requirements_folder": "exigences",
    "plan_folder": "plans",
    "stdhl_folder": "stdhl",
    "fiches_folder": "fiches",
    "traceability_folder": "tracabilite",
    # Separators inside a cell, as regular expressions.
    "release_separator_regex": r",",
    "stdhl_separator_regex": r"[,;/\n]+",
    "plan_tome_regex": r"(?i)_PA_(.+?)_",
    "code_separator_regex": r"[,;+\s]+",
    "delegation_separator_regex": r"[,;/+\n]+",
    # Plan import: for each IADT code, the MOC codes expected (at least one of them, and no other one).
    "iadt_moc_rules": {"D": ["MOC4"], "T": ["MOC4"], "I": ["MOC7"], "A": ["MOC2", "MOC3"], "N/A": ["N/A"]},
    # Fiches d'Essai (XML): a test contains test cases, which contain the requirements tested.
    "test_tag": "test",
    "test_case_tag": "test_case",
    "requirement_tag": "requirements",
    # Column D of the requirements names the STDHL a requirement must be found in (True), or is ignored (False).
    "use_requirement_stdhl": True,
    # The release the base works on: every requirement imported and every file loaded belongs to it.
    "release": "",
    "tomes": [],
    # Test methods are written "V&V [PLATFORM] : ..."; TOP is the platform of V&V itself.
    "platforms": ["TOP"],
    "iadt": {"I": "Inspection", "A": "Analyse", "D": "Démonstration", "T": "Test", "N/A": "Non applicable"},
    "mocs": {f"MOC{i}": "" for i in range(10)},
    "default_delegation": "TOP",
    # Expected headers of the tables, in order: header text -> role ("" = column only checked).
    # A header cell must contain the text (case and accents ignored).
    "headers_requirements": {"ID": "id", "Titre": "title", "Release": "releases", "STDHL": "stdhl"},
    "headers_plan": {"Exigence": "exigence", "Méthode": "ignore", "IADT": "iadt", "MOC": "moc", "Délégation": "delegation"},
    "headers_stdhl": {"Description du test": "check", "Résultat attendu": "check", "Exigences associées": "requirements"},
}

# Roles of the columns of each table header setting: each one must be given exactly once.
HEADER_ROLES = {
    "headers_requirements": ("id", "title", "releases", "stdhl"),
    "headers_plan": ("exigence", "iadt", "moc", "delegation"),
    "headers_stdhl": ("requirements",),
}

# Parts of the file and their variables, always written: the file shows every setting, and all the rules to decode the
# imported files. Other variables (settings of the installation) stay at the root of the file.
# A part maps the names written in the file to the variables: inside a part, a name needs no prefix (« folder »).
SECTIONS = {
    "general": {n: n for n in ("release", "slogan", "server_port", "open_browser", "chart_max_parts",
                               "chart_max_parts_platforms", "storage_dir")},
    "requirements": {"folder": "requirements_folder",
                     **{n: n for n in ("requirement_regex", "requirements_tome_regex", "release_separator_regex",
                                       "stdhl_separator_regex",
                                       "headers_requirements")}},
    "plan": {"folder": "plan_folder",
             **{n: n for n in ("plan_tome_regex", "code_separator_regex", "delegation_separator_regex", "iadt_moc_rules",
                               "headers_plan")}},
    "stdhl": {"folder": "stdhl_folder",
              **{n: n for n in ("use_requirement_stdhl", "fiche_regex", "stdhl_name_regex", "headers_stdhl",
                                "stdhl2_test_regex", "stdhl2_requirements_label", "stdhl2_release_label",
                                "stdhl2_release_regex", "stdhl2_test_case_regex")}},
    "fiches": {"folder": "fiches_folder", **{n: n for n in ("test_tag", "test_case_tag", "requirement_tag")}},
    "traceability": {"folder": "traceability_folder"},
}
SECTION_VARIABLES = tuple(var for names in SECTIONS.values() for var in names.values())

# Variables holding the name of a sub-folder of the storage folder.
FOLDER_VARIABLES = ("requirements_folder", "plan_folder", "stdhl_folder", "fiches_folder", "traceability_folder")

# Variables holding the name of an XML tag.
TAG_VARIABLES = ("test_tag", "test_case_tag", "requirement_tag")

# Variables holding a regular expression.
REGEX_VARIABLES = ("requirement_regex", "requirements_tome_regex", "fiche_regex", "stdhl_name_regex", "plan_tome_regex",
                   "release_separator_regex", "stdhl_separator_regex", "code_separator_regex",
                   "delegation_separator_regex", "stdhl2_test_regex", "stdhl2_release_regex", "stdhl2_test_case_regex")

# Variables kept in the SQL database (nexus_db.py) rather than in the JSON file.
REFERENTIALS = ("tomes", "platforms", "iadt", "mocs")

_MISSING = object()


def fold(text):
    """Comparable form of a label: no accents, lower case, single spaces."""
    # NFD splits an accented letter into the letter and its accent (a combining character), which is then dropped.
    text = unicodedata.normalize("NFD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text).strip().lower()


METHOD_PLATFORM = re.compile(r"^\s*V\s*&\s*V\s*\[\s*([^\]]*?)\s*\]", re.I)


def storage_path(config, base):
    """Folder of the imported files: general.storage_dir, relative to `base` (the folder that holds the folder base,
    fichiers-references for NEXUS-V; « . » is that folder itself) unless absolute."""
    folder = Path(config.get("storage_dir"))
    return (folder if folder.is_absolute() else Path(base) / folder).resolve()


def plural(n, singular, several=None):
    """« 1 ligne », « 3 lignes » (singular up to 1, as in French)."""
    return f"{n} {(several or singular + 's') if n > 1 else singular}"


def header_column(cells, expected):
    """Index of the header cell matching the expected text (case and accents ignored): a cell equal to it
    first, otherwise the first cell containing it; None when there is none."""
    key = fold(expected)
    folded = [fold(str(c or "")) for c in cells]
    # A cell equal to the expected text wins over a cell that only contains it ("ID" is also in "Titre IDentifié").
    exact = next((i for i, f in enumerate(folded) if f == key), None)
    return exact if exact is not None else next((i for i, f in enumerate(folded) if key and key in f), None)


def match_header(cells, headers):
    """Columns of a header row for a header setting (text -> role): ({role: column index} for the columns
    found that have a role, [expected texts not found])."""
    cols, missing = {}, []
    for text, role in headers.items():
        if role == "ignore":   # documented only: not used to validate the table
            continue
        k = header_column(cells, text)
        if k is None:
            missing.append(text)
        elif role not in ("check", ""):
            cols[role] = k
    return cols, missing


def table_left_out(number, missing):
    """Message of a table of a Word document left out because its header lacks expected texts."""
    texts = "en-têtes attendus introuvables" if len(missing) > 1 else "en-tête attendu introuvable"
    return f"tableau {number} : {texts} : " + ", ".join(f"« {m} »" for m in missing)


def header_text(headers, role):
    """Expected header text of a role in a header setting."""
    return next(text for text, r in headers.items() if r == role)


def method_platform(text):
    """Platform named in a test method "V&V [PLATFORM] : ...", or None when the method has another form."""
    m = METHOD_PLATFORM.match(text or "")
    return m[1] if m else None


def type_name(value):
    """Type name of a value (bool is tested before int: True is also an int)."""
    for t, name in ((bool, "bool"), (int, "int"), (float, "float"), (str, "str"), (list, "list"), (dict, "dict")):
        if isinstance(value, t):
            return name
    raise ValueError(f"type non géré : {type(value).__name__}")


class Configuration:
    """Read, write, get and set the configuration variables kept in a JSON file."""

    def __init__(self, path, defaults=None, database=None):
        """database: NexusDatabase holding the referential (tomes, platforms, iadt, mocs);
        without it, they are kept in the JSON file like the other variables."""
        path = Path(path)
        self.database = database
        self.path = path.with_suffix(".json")
        self.defaults = copy.deepcopy(DEFAULT_CONFIG if defaults is None else defaults)
        self._values = {}             # variables read from or set in the file
        self._lock = threading.RLock()
        self.read()

    # ------------------------------------------------------------------ #
    # Reading / writing the JSON file
    # ------------------------------------------------------------------ #

    def read(self):
        """(Re)load from disk; an absent file means every variable has its default value."""
        with self._lock:
            self._values = {}
            changed = False
            if self.path.is_file():
                data = self._load_json()
                # The file is read part by part; inside a part a name may be short ("folder") and is mapped to its
                # variable ("requirements_folder"). Every value is validated: a hand-edited mistake is reported at
                # start.
                for name, value in data.items():
                    if name in SECTIONS:
                        for key, v in value.items():
                            var = SECTIONS[name].get(key, key)
                            self._values[var] = self._check(var, v)
                    else:
                        self._values[name] = self._check(name, value)
            self._default_referential()
            # Every variable of a part is always written in the file, its default value included: the file documents
            # every rule used to read the imported documents.
            for name in SECTION_VARIABLES:
                if name in self.defaults and name not in self._values:
                    self._values[name] = copy.deepcopy(self.defaults[name])
                    changed = True
            if changed:
                self.write()
            if self.database is not None:
                self.database.release = self.get("release")

    def _default_referential(self):
        """A base just created from the schema gets the default referentials, once; afterward they are only changed
        by the user (a referential emptied stays empty)."""
        if self.database is None:
            return
        if self.database.created:
            self.database.created = False
            for name in REFERENTIALS:
                if name in self.defaults:
                    self.database.set(name, self._check(name, self.defaults[name]))
        # There is always a default platform: the one of the defaults if it exists, else the first platform.
        if self.database.default_platform() is None:
            wanted = self.defaults.get("default_delegation")
            known = wanted in self.database.get("platforms")
            self.database.set_default_platform(wanted if known else None)

    def write(self):
        """Save to disk atomically (temporary file, then rename). Only set variables are written."""
        with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # Written to a temporary file then renamed: a crash while writing never leaves a half-written config.json.
            tmp = self.path.with_suffix(".tmp")
            parts = {part: {key: self._values[var] for key, var in names.items() if var in self._values}
                     for part, names in SECTIONS.items()}
            others = {n: v for n, v in self._values.items() if n not in SECTION_VARIABLES}
            data = {**{part: values for part, values in parts.items() if values}, **others}
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, self.path)

    def _load_json(self):
        """Variables of the JSON file. ValueError when it is not a valid JSON object."""
        try:
            values = json.loads(self.path.read_text(encoding="utf-8"))
        except ValueError as e:
            raise ValueError(f"{self.path} : JSON invalide ({e})") from None
        if not isinstance(values, dict):
            raise ValueError(f"{self.path} : un objet JSON est attendu")
        return values

    # ------------------------------------------------------------------ #
    # Access
    # ------------------------------------------------------------------ #

    def names(self):
        """Every variable name: defaults first, then the extra ones set in the file."""
        with self._lock:
            return list(dict.fromkeys([*self.defaults, *self._values]))

    def __contains__(self, name):
        """True for a variable that is set or has a default value."""
        return name in self._values or name in self.defaults

    def get(self, name, default=_MISSING):
        """Value of a variable (copy): set value, else default value, else `default`, else KeyError.
        The referentials are read in the database when there is one."""
        # Values kept in the database rather than in the JSON file.
        if name == "tome_descriptions":   # what each tome holds, in the database: {tome: text}
            return self.database.tome_descriptions() if self.database is not None else {}
        if name == "colors":   # colour of each value of the referentials, in the database: {kind: {name: colour}}
            return {kind: self.database.colors(kind) for kind in REFERENTIALS} if self.database is not None else {}
        if name == "releases":   # the declared releases: the one of the base
            release = self.get("release")
            return {release: ""} if release else {}
        if self.database is not None and name in REFERENTIALS:
            return self.database.get(name)
        if self.database is not None and name == "default_delegation":
            return self.database.default_platform() or ""
        with self._lock:
            if name in self._values:
                return copy.deepcopy(self._values[name])
            if name in self.defaults:
                return copy.deepcopy(self.defaults[name])
        if default is _MISSING:
            raise KeyError(name)
        return default

    def to_dict(self):
        """Every variable with its current value, referentials and their colours included."""
        return {**{name: self.get(name) for name in self.names()}, "colors": self.get("colors"),
                "tome_descriptions": self.get("tome_descriptions")}

    def set(self, name, value):
        """Set one variable after validation and save."""
        self.update({name: value})

    def update(self, values):
        """Set several variables after validation and save once. Raises ValueError if one is invalid.

        A variable set to its default value is not written to the file, so that a later
        change of the default value still applies to it.
        """
        values = dict(values)
        # Values the page sends besides the variables: renames, colours and descriptions go to the database.
        renames = values.pop("renames", None) or {}   # {kind: {old name: new name}}, see NexusDatabase.rename
        colors = values.pop("colors", None) or {}   # {kind: {name: "#rrggbb"}}
        descriptions = values.pop("tome_descriptions", None) or {}   # {tome: text}
        if not isinstance(descriptions, dict):
            raise ValueError("descriptions des tomes : format invalide")
        if any(kind not in REFERENTIALS or not isinstance(v, dict) for kind, v in colors.items()):
            raise ValueError("couleurs : liste inconnue")
        bad = [c for v in colors.values() for c in v.values() if not re.fullmatch(r"#[0-9a-fA-F]{6}", str(c))]
        if bad:
            raise ValueError(f"couleur invalide : {bad[0]}")
        # Every value is validated before anything is written: one invalid value refuses the whole update.
        clean = {name: self._check(name, value) for name, value in values.items()}
        if "release" in clean and clean["release"] != self.get("release") and self.database is not None \
                and self.database.scalar("SELECT COUNT(*) FROM files"):
            raise ValueError("la release ne peut changer que si aucun fichier n'est chargé : supprimer d'abord les fichiers")
        # A renamed IADT or MOC code keeps its rules.
        # The rules IADT -> MOC are written with codes: a renamed code is renamed in the rules too.
        moved = {kind: {o.upper(): n.upper() for o, n in (renames.get(kind) or {}).items()} for kind in ("iadt", "mocs")}
        if any(moved.values()):
            rules = clean.get("iadt_moc_rules", self.get("iadt_moc_rules"))
            clean["iadt_moc_rules"] = {moved["iadt"].get(i, i): [moved["mocs"].get(m, m) for m in mocs]
                                       for i, mocs in rules.items()}
        # Two types of files in the same sub-folder would mix their copies.
        folders = [clean.get(name, self.get(name)) for name in FOLDER_VARIABLES]
        same = sorted({f for f in folders if [x.casefold() for x in folders].count(f.casefold()) > 1})
        if same:
            raise ValueError("chaque type de fichier doit avoir son propre sous-dossier : « " + "», « ".join(same)
                             + " » est utilisé plusieurs fois")
        if self.database is not None:
            default = clean.pop("default_delegation", None)   # a field of the table list_plateformes
            if renames:
                self.database.rename(renames)
            # Referentials go to the database, in one transaction, before the JSON file is written.
            referentials = {name: clean.pop(name) for name in REFERENTIALS if name in clean}
            if referentials:
                self.database.set_many(referentials)
            self.database.set_default_platform(default or None)
            if colors:
                self.database.set_colors(colors)
            if descriptions:
                self.database.set_tome_descriptions(descriptions)
        with self._lock:
            # The remaining variables go to the JSON file.
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
        """Forget the value set for a variable (it gets its default value back) and save."""
        with self._lock:
            if self._values.pop(name, _MISSING) is not _MISSING:
                self.write()

    @staticmethod
    def _check_rules(value):
        """Validated rules {IADT code: [expected MOC codes]}: codes upper-cased, each IADT once, at least one MOC."""
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
        """Validated copy of a value: same type as the default, cleaned lists of platforms and codes."""
        if not isinstance(name, str) or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
            raise ValueError(f"nom de variable invalide : « {name} »")
        # The type must be the one of the default value (an int is accepted for a float).
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

        # Checks specific to some variables: duplicates in the referentials, header roles each present once, valid
        # folder and tag names, compilable regular expressions, ranges of numbers.
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
            # {header text: role}; a column without a stored role is checked ("check") or only documented ("ignore").
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
    # Checks used on the acceptance plan tables
    # ------------------------------------------------------------------ #

    def find_platform(self, text):
        """Configured platform of a test method, or None.

        "V&V [PLATFORM] : ..." names the platform between brackets; otherwise the method
        must start with a platform name (the longest one followed by a non-alphanumeric character).
        """
        # "V&V [TOP] : ..." names the platform between brackets; otherwise the longest platform name the method starts
        # with (followed by a separator, so that "TOP2" is not read as "TOP").
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
        """True for a platform of the configuration (case and accents ignored)."""
        return any(fold(name) == fold(p) for p in self.get("platforms"))

    def delegation(self, value):
        """Delegation actually meant by a cell: the default one when empty."""
        return (value or "").strip() or self.get("default_delegation")
