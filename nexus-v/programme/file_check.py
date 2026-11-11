"""NEXUS-V — check of files before their import, without importing them.

Several files of one stage (requirements workbooks, or acceptance plans) are read with the same functions as the
import windows (RequirementsDatabase.analyse, PlansDatabase.analyse): each problem that would refuse a file, ask a
choice of the user or leave a requirement uncovered is listed with the correction to make. The files are also compared
with each other. Nothing is written: the base is only read, and the effect on the coverage is computed by running the
Verification on the requirements and plans of the base with, in memory, the content of the files checked.
The result is a self-contained HTML report (French, like the page).
"""
import html
import re
from datetime import datetime

from configuration import plural
from plans_db import PlansDatabase
from requirements_db import RequirementsDatabase
from verification import Verification

# Severity of a problem, from the most serious: a refused file, a choice asked at import, a point to look at, a note.
BLOCKING, SETTLE, CHECK, INFO = "bloquant", "regler", "verifier", "info"
SEVERITY_LABELS = {BLOCKING: "Bloquant", SETTLE: "À régler", CHECK: "À vérifier", INFO: "Information"}
# Verdict of a file, from its most serious problem.
VERDICTS = {BLOCKING: "Refusé à l'import", SETTLE: "Chargeable, avec des choix à faire", CHECK: "Chargeable, à vérifier",
            INFO: "Chargeable"}
STEP_TITLES = {"requirements": "Exigences", "plans": "Plans d'acceptation"}
MODE_LABELS = {"replace": "Si le fichier écrase la base (import en remplacement)",
               "complement": "Si le fichier modifie la base (import en complément)"}
MODE_SHORT = {"replace": "écrase la base", "complement": "modifie la base"}
FIELD_LABELS = {"title": "titre", "releases": "release", "stdhl": "STDHL"}
LIMIT = 30   # numbers listed in one problem, beyond: « … et N autres »


def problem(severity, text, correction="", place=""):
    """One problem of a file: its severity, what is wrong, the correction to make and where (table, row)."""
    return {"severity": severity, "text": text, "correction": correction, "place": place}


def listed(numbers):
    """Numbers joined, the first LIMIT of them, then how many others."""
    numbers = list(numbers)
    more = len(numbers) - LIMIT
    return ", ".join(numbers[:LIMIT]) + (f" … et {more} autre{'s' if more > 1 else ''}" if more > 0 else "")


class RequirementsInMemory:
    """The requirements of the base, some tomes replaced by the content of checked files: what Verification reads."""

    def __init__(self, requirements, replaced):
        """requirements: RequirementsDatabase; replaced: {tome: [requirements of the file]}."""
        self.requirements, self.replaced = requirements, replaced

    def to_dict(self, only=None):
        data = self.requirements.to_dict(only)
        tomes = dict(data["tomes"])
        for tome, reqs in self.replaced.items():
            tomes[tome] = {"description": "", "files": [], "requirements": reqs}
        return {**data, "tomes": tomes}


class PlansInMemory:
    """The plans of the base as left by checked plan files (rows replaced, removed or added): what Verification
    reads."""

    def __init__(self, plans):
        """plans: in the form of PlansDatabase.plans, rows checked."""
        self.list = plans

    def plans(self, tome=None, file_id=None, config=None, with_deleted=False):
        return [p for p in self.list if tome is None or p["tome"] == tome]


class FileCheck:
    """Check of requirements or plan files against the base, before their import."""

    def __init__(self, config, requirements, plans, stdhl, fiches, traceability, sql):
        self.config, self.requirements, self.plans = config, requirements, plans
        self.stdhl, self.fiches, self.traceability, self.sql = stdhl, fiches, traceability, sql

    # ------------------------------------------------------------------ #
    # Requirements
    # ------------------------------------------------------------------ #

    def check_requirements(self, files):
        """Check requirements workbooks [(name, bytes)] for both imports: what happens if each file overwrites the
        base (replacement) and if it modifies it (complement). The problems found in both are given once, those of
        one import apart; the tiles after import are computed for each."""
        runs = {mode: self._requirements_mode(files, mode) for mode in MODE_LABELS}
        return self._both_modes("requirements", runs)

    def _both_modes(self, step, runs):
        """Report of a check run for both imports ({mode: {results, between, requirements, plans, hypothesis}}): the
        problems found in both given once, those of one import apart, tiles after each import."""
        rep, comp = runs["replace"], runs["complement"]
        results = []
        for a, b in zip(rep["results"], comp["results"]):
            common = [x for x in a["problems"] if x in b["problems"]]
            results.append({**a, "problems": common, "facts": [f for f in a["facts"] if f in b["facts"]],
                            "by_mode": {"replace": [x for x in a["problems"] if x not in common],
                                        "complement": [x for x in b["problems"] if x not in common]},
                            "mode_facts": {"replace": [f for f in a["facts"] if f not in b["facts"]],
                                           "complement": [f for f in b["facts"] if f not in a["facts"]]}})
        between = [x for x in rep["between"] if x in comp["between"]]
        between_by_mode = {"replace": [x for x in rep["between"] if x not in between],
                           "complement": [x for x in comp["between"] if x not in between]}
        return self._result(step, results, between, None, None, {m: runs[m]["hypothesis"] for m in MODE_LABELS},
                            modes={m: (runs[m]["requirements"], runs[m]["plans"]) for m in MODE_LABELS},
                            between_by_mode=between_by_mode)

    def _requirements_mode(self, files, mode="replace"):
        """Check requirements workbooks [(name, bytes)] for an import in `mode`: "replace" (the requirements of the
        file replace those of the tomes they belong to) or "complement" (they are added or update those of the base,
        nothing is removed). The tome named in the file name is not used: each requirement of the file is looked up
        in the whole base, and the tomes concerned are those that hold them. One result per file, the problems between
        files and the tiles of the Verification before and after their import."""
        mode = mode if mode in MODE_LABELS else "replace"
        release = self.config.get("release")
        number = re.compile(self.config.get("requirement_regex"))
        # The base after the files checked so far, tome by tome: {tome: {number: requirement}}, filled as each file
        # is applied in turn; and where each number of the base is.
        tomes = {}
        for r in self.requirements.requirements():
            if not r["deleted"]:
                tomes.setdefault(r["tome"], {})[r["id"]] = {k: v for k, v in r.items() if k != "tome"}
        base = {i: (t, r) for t, rs in tomes.items() for i, r in rs.items()}
        value = lambda r, f: ", ".join(r[f]) if f == "releases" else (r[f] or "")
        results, touched_all = [], set()
        for name, content in files:
            found = {"name": name, "tome": "", "problems": [], "facts": [], "tomes": []}
            results.append(found)
            try:
                sheets = RequirementsDatabase.sheet_names(content)
                data = self.requirements.analyse(content, sheets[0] if len(sheets) > 1 else None, None, mode,
                                                 self.config, [release])
            except ValueError as e:
                found["problems"].append(problem(
                    BLOCKING, f"fichier illisible ou refusé : {e}",
                    "corriger le classeur : la ligne 1 doit porter les en-têtes ID, Titre, Release et STDHL (textes "
                    "réglés dans Configuration > Exigences)"))
                continue
            p = found["problems"]
            if len(sheets) > 1:
                p.append(problem(INFO, f"classeur de {len(sheets)} feuilles ({', '.join(sheets)}) : la première, "
                                       f"« {sheets[0]} », a été vérifiée",
                                 "à l'import, choisir la feuille à lire"))
            reqs = data["requirements"]
            outside = [r for g in data["outside"] for r in g["requirements"]]
            found["facts"] += [f"{plural(data['count'], 'exigence')} en release {release}",
                               f"releases lues : {', '.join(f'{r} ({n})' for r, n in data['release_counts'].items()) or 'aucune'}"]
            if not reqs:
                p.append(problem(BLOCKING, f"aucune exigence en release {release}",
                                 f"renseigner la colonne Release (la base est en release {release})"))
            for group in data["outside"]:
                label = (f"release{'s' if ',' in group['key'] else ''} {group['key']}" if group["key"]
                         else "sans release")
                p.append(problem(SETTLE, f"{plural(group['count'], 'exigence')} {label}, hors de la release {release} : "
                                         + listed(r["id"] for r in group["requirements"]),
                                 f"ignorées par défaut à l'import ; si elles concernent la release {release}, corriger "
                                 "la colonne Release ou les importer en forçant leur groupe"))
            for r in data["repeated"]:
                rows = ", ".join(str(v["row"]) for v in r["rows"])
                p.append(problem(SETTLE, f"{r['id']} présente sur plusieurs lignes ({rows})",
                                 "ne garder qu'une ligne dans le fichier (sinon, la choisir à l'import)",
                                 f"lignes {rows}"))
            for r in reqs:
                if len(r["stdhl_values"]) > 1:
                    p.append(problem(SETTLE, f"{r['id']} : plusieurs valeurs STDHL ({', '.join(r['stdhl_values'])})",
                                     "n'en garder qu'une dans la cellule (sinon, la choisir à l'import)"))
            bad = [r["id"] for r in reqs if not number.fullmatch(r["id"])]
            if bad:
                p.append(problem(CHECK, f"{plural(len(bad), 'numéro')} hors du format des exigences : {listed(bad)}",
                                 "corriger le numéro : les plans et les STDHL ne pourront pas le citer (format réglé "
                                 "dans Configuration > Expressions régulières)"))
            untitled = [r["id"] for r in reqs if not r["title"].strip()]
            if untitled:
                p.append(problem(CHECK, f"{plural(len(untitled), 'exigence')} sans titre : {listed(untitled)}",
                                 "renseigner la colonne Titre"))

            # Each requirement of the file against the base, whatever its tome: new, unchanged or modified.
            fresh, seen = {}, set()
            for r in reqs:
                if r["id"] not in seen:   # a repeated number: its first row, as the import before the choice
                    seen.add(r["id"])
                    fresh[r["id"]] = r
            added = [i for i in fresh if i not in base]
            changes = {}
            for i, r in fresh.items():
                if i in base:
                    old = base[i][1]
                    fields = [f for f in FIELD_LABELS if value(old, f) != value(r, f)]
                    if fields:
                        changes[i] = fields
            unchanged = len(fresh) - len(added) - len(changes)
            # Tomes concerned: those of the base holding requirements of the file (in or outside the release).
            touched = {}
            for i in list(fresh) + [r["id"] for r in outside]:
                if i in base:
                    touched.setdefault(base[i][0], []).append(i)
            found["tomes"] = sorted(touched)
            found["tome"] = ", ".join(found["tomes"]) or "aucun (exigences nouvelles)"
            found["facts"].append(f"par rapport à la base : {len(added)} nouvelle(s), {len(changes)} modifiée(s), "
                                  f"{unchanged} inchangée(s)")
            if touched:
                found["facts"].append("tomes de la base concernés : "
                                      + ", ".join(f"{t} ({len(ids)})" for t, ids in sorted(touched.items())))
            if len(touched) > 1:
                p.append(problem(CHECK, f"le fichier contient des exigences de {len(touched)} tomes de la base : "
                                        + ", ".join(sorted(touched)),
                                 "un fichier d'exigences ne porte qu'un tome : le scinder par tome"
                                 + (" ; en remplacement, l'import ne remplace qu'un tome" if mode == "replace" else "")))
            if added:
                target = next(iter(touched)) if len(touched) == 1 else ""
                p.append(problem(INFO, f"{plural(len(added), 'exigence nouvelle', 'exigences nouvelles')}, absente"
                                       f"{'s' if len(added) > 1 else ''} de la base : " + listed(added),
                                 f"elles rejoindront le tome {target}" if target else
                                 "à l'import, choisir le tome où elles vont"))
            for field, label in FIELD_LABELS.items():
                changed = [i for i, fields in changes.items() if field in fields]
                if changed:
                    where = sorted({base[i][0] for i in changed})
                    p.append(problem(INFO, f"{plural(len(changed), 'exigence')} de la base au {label} modifié "
                                           f"(tome{'s' if len(where) > 1 else ''} {', '.join(where)}) : " + listed(changed)))
            left = [r["id"] for r in outside if r["id"] in base]
            if left:
                p.append(problem(CHECK, f"{plural(len(left), 'exigence')} de la base hors de la release {release} dans "
                                        f"le fichier : " + listed(left),
                                 "elles seront retirées de la base avec leur couverture : corriger leur release si "
                                 "elles doivent rester" if mode == "replace" else "la base les garde telles quelles"))
            removed = []
            if mode == "replace":
                # In replacement, the requirements of the tomes concerned that the file does not keep leave the base.
                for t in sorted(touched):
                    gone = [i for i in tomes.get(t, {}) if i not in fresh]
                    removed += gone
                    if gone:
                        p.append(problem(CHECK, f"{plural(len(gone), 'exigence')} du tome {t} absente"
                                                f"{'s' if len(gone) > 1 else ''} du fichier : " + listed(gone),
                                         "l'import remplace le tome : elles seront retirées de la base, avec leur "
                                         "couverture ; les ajouter au fichier si elles doivent rester"))
            found["facts"][-1 if not touched else -2] += f", {len(removed)} retirée(s)" if mode == "replace" else ""

            # The base after this file: removed ones out, modified ones updated in their tome, new ones added to the
            # tome concerned (a tome of their own when the file touches none, or several).
            for i in removed:
                tomes[base[i][0]].pop(i, None)
            new_tome = next(iter(touched)) if len(touched) == 1 else f"(nouveau : {name})"
            for i, r in fresh.items():
                tome = base[i][0] if i in base else new_tome
                tomes.setdefault(tome, {})[i] = {"id": i, "title": r["title"], "releases": r["releases"],
                                                 "stdhl": r["stdhl"], "file_id": None, "manual": None, "deleted": False}
            touched_all |= set(touched) | ({new_tome} if added else set())
        replaced = {t: list(tomes.get(t, {}).values()) for t in touched_all}
        between = self._between_requirements(results, files, mode)
        return {"results": results, "between": between, "plans": self.plans,
                "requirements": RequirementsInMemory(self.requirements, replaced),
                "hypothesis": ("chaque exigence du fichier remplace celle de même numéro de la base, quel que soit son "
                             "tome ; les exigences des tomes concernés absentes du fichier sont retirées"
                             if mode == "replace" else
                             "chaque exigence du fichier remplace celle de même numéro de la base, quel que soit son "
                             "tome, ou s'ajoute ; rien n'est retiré")
                            + " ; le nom du fichier n'est pas utilisé"}

    def _between_requirements(self, results, files, mode="replace"):
        """Problems between requirements files: in replacement, one tome of the base concerned by several files (the
        last one imported would win); one number in several files."""
        found = []
        by_tome = {}
        for r in results:
            for t in r.get("tomes", []):
                by_tome.setdefault(t, []).append(r["name"])
        for tome, names in sorted(by_tome.items()):
            if len(names) > 1 and mode == "replace":
                found.append(problem(SETTLE, f"tome {tome} concerné par {len(names)} fichiers : {', '.join(names)}",
                                     "un import en remplacement remplace le tome : seul le dernier importé resterait ; "
                                     "les réunir en un fichier, ou importer les suivants en complément"))
        numbers = {}
        for name, content in files:
            try:
                sheets = RequirementsDatabase.sheet_names(content)
                data = RequirementsDatabase.read_excel(content, sheets[0], self.config)
            except ValueError:
                continue
            for r in data["requirements"]:
                numbers.setdefault(r["id"], []).append(name)
        shared = {n: names for n, names in numbers.items() if len(set(names)) > 1}
        if shared:
            pairs = {}
            for n, names in shared.items():
                pairs.setdefault(" et ".join(sorted(set(names))), []).append(n)
            for names, ids in pairs.items():
                found.append(problem(SETTLE, f"{plural(len(ids), 'exigence')} dans plusieurs fichiers ({names}) : "
                                             + listed(ids),
                                     "une exigence n'appartient qu'à un tome : la retirer de l'un des fichiers"))
        return found

    # ------------------------------------------------------------------ #
    # Plans
    # ------------------------------------------------------------------ #

    def check_plans(self, files):
        """Check acceptance plan files [(name, bytes)] for both imports: what happens if each file overwrites the
        base (it replaces the plans of the tomes it concerns) and if it modifies it (its rows replace those of the
        same requirements, the others stay). The tome named in the file name is not used: each requirement of the plan
        is looked up in the whole base, and its row compared with the one of the plans of the base."""
        runs = {mode: self._plans_mode(files, mode) for mode in MODE_LABELS}
        return self._both_modes("plans", runs)

    def _plans_mode(self, files, mode):
        """Check of plan files for an import in `mode` (see check_plans): results, problems between files, and the
        plans of the base as each file in turn would leave them."""
        release = self.config.get("release")
        tome_of = {r["id"]: r["tome"] for r in self.requirements.requirements() if not r["deleted"]}
        # The plans of the base, copied: each file in turn replaces or removes rows in them.
        plans = [{**q, "rows": list(q["rows"])} for q in self.plans.plans(config=self.config)]
        codes = lambda v: sorted(PlansDatabase.split_codes(v or ""))
        same = {"iadt": lambda a, b: codes(a) == codes(b), "moc": lambda a, b: codes(a) == codes(b),
                "delegation": lambda a, b: self.config.delegation(a or "") == self.config.delegation(b or "")}
        labels = {"iadt": "IADT", "moc": "MOC", "delegation": "délégation"}
        results, contents = [], []
        for k, (name, content) in enumerate(files, 1):
            found = {"name": name, "tome": "", "problems": [], "facts": [], "tomes": []}
            results.append(found)
            p = found["problems"]
            try:
                data = self.plans.analyse(content, self.config, "", [release], name)
            except ValueError as e:
                p.append(problem(BLOCKING, f"fichier illisible ou refusé : {e}",
                                 "le plan doit avoir au moins un tableau dont la première ligne porte les en-têtes "
                                 "Exigence, IADT, MOC et Délégation (textes réglés dans Configuration > Plan "
                                 "d'acceptation)"))
                continue
            entries = self._entries(content)
            contents.append((name, entries))
            found["facts"].append(f"{plural(len(entries), 'exigence')} lue{'s' if len(entries) > 1 else ''} dans le plan")
            for skipped in data["skipped"]:
                p.append(problem(CHECK, f"tableau mis de côté : {skipped}",
                                 "compléter les en-têtes du tableau s'il fait partie du plan"))
            # Rows against the configuration (IADT, MOC, delegation): one problem per issue, the requirements that
            # have it listed together, with the places of their rows.
            settle = {c["key"]: c for c in data["to_confirm"]}
            unknown = [e for e in entries if e["requirement"] not in tome_of]
            grouped = {}
            for c in data["to_confirm"]:
                for issue in c["issues"]:
                    entry = grouped.setdefault((issue, self._plan_correction(c, issue)), {"ids": [], "places": []})
                    entry["ids"].append(c["key"])
                    entry["places"] += c["places"]
            if unknown:
                grouped[("exigence inconnue de la base : la ligne sera ignorée",
                         "corriger le numéro, ou importer d'abord l'exigence")] = {
                    "ids": [e["requirement"] for e in unknown], "places": [x for e in unknown for x in e["places"]]}
            for (issue, correction), entry in grouped.items():
                ids = entry["ids"]
                text = f"{ids[0]} : {issue}" if len(ids) == 1 else f"{plural(len(ids), 'exigence')} : {issue} — {listed(ids)}"
                p.append(problem(SETTLE, text, correction, self._places(entry["places"])))

            # Each known requirement against the plans of the base, whatever the plan: a new row, a row modified (IADT,
            # MOC, delegation) or unchanged. The tomes concerned are those of these requirements in the base.
            known = [e for e in entries if e["requirement"] in tome_of]
            if entries and not known:
                p.append(problem(BLOCKING, "aucune exigence du plan n'est dans la base : rien ne serait chargé",
                                 "importer d'abord les exigences du plan, ou corriger les numéros"))
            touched = {}
            for e in known:
                touched.setdefault(tome_of[e["requirement"]], []).append(e["requirement"])
            found["tomes"] = sorted(touched)
            found["tome"] = ", ".join(found["tomes"]) or "aucun"
            rows_of = {}
            for q in plans:
                for row in q["rows"]:
                    rows_of.setdefault(row["key"], []).append(row)
            added, changes = [], {}
            for e in known:
                old = rows_of.get(e["requirement"])
                if not old:
                    added.append(e["requirement"])
                    continue
                fields = [f for f in labels if not same[f](old[0][f], e[f])]
                if fields:
                    changes[e["requirement"]] = fields
            found["facts"].append(f"par rapport aux plans de la base : {len(added)} ligne(s) nouvelle(s), "
                                  f"{len(changes)} modifiée(s), {len(known) - len(added) - len(changes)} inchangée(s)")
            if touched:
                found["facts"].append("tomes de la base concernés : "
                                      + ", ".join(f"{t} ({len(ids)})" for t, ids in sorted(touched.items())))
            if len(touched) > 1:
                p.append(problem(CHECK, f"le plan contient des exigences de {len(touched)} tomes de la base : "
                                        + ", ".join(sorted(touched)),
                                 "un plan ne porte qu'un tome : le scinder par tome"))
            if added:
                p.append(problem(INFO, f"{plural(len(added), 'exigence')} sans ligne de plan dans la base, couverte"
                                       f"{'s' if len(added) > 1 else ''} par ce plan : " + listed(added)))
            for field, label in labels.items():
                changed = [i for i, fields in changes.items() if field in fields]
                if changed:
                    p.append(problem(INFO, f"{plural(len(changed), 'ligne')} de plan de la base à la {label} modifiée : "
                                           + listed(changed)) if field == "delegation" else
                             problem(INFO, f"{plural(len(changed), 'ligne')} de plan de la base "
                                           f"{'à l' + chr(39) if label == 'IADT' else 'au '}{label} modifié : "
                                           + listed(changed)))

            # The plans after this file. Overwrite: the plans of the tomes concerned give way to the file. Modify: the
            # rows of the requirements of the file give way to its rows, the others stay.
            keys = {e["requirement"] for e in known}
            if mode == "replace":
                plans = [q for q in plans if q["tome"] not in touched]
            else:
                plans = [{**q, "rows": [row for row in q["rows"] if row["key"] not in keys]} for q in plans]
            for t, ids in sorted(touched.items()):
                rows = []
                for e in known:
                    if tome_of[e["requirement"]] != t:
                        continue
                    row = PlansDatabase.check_row({"key": e["requirement"], "requirements": [e["requirement"]],
                                                   "iadt": e["iadt"], "moc": e["moc"], "delegation": e["delegation"],
                                                   "manual": "", "deleted": False}, self.config)
                    rows.append({**row, "errors": []} if e["requirement"] in settle else row)
                plans.append({"file_id": -k, "name": name, "tome": t, "releases": [release], "rows": rows})
            # Requirements of the tomes concerned left without plan row once the file imported.
            covered = {row["key"] for q in plans for row in q["rows"]}
            for t in sorted(touched):
                left = [i for i, tome in tome_of.items() if tome == t and i not in covered]
                if left:
                    p.append(problem(CHECK, f"{plural(len(left), 'exigence')} du tome {t} sans ligne de plan après "
                                            f"import : " + listed(left),
                                     "elles perdent leur ligne de plan : les ajouter au fichier si elles doivent "
                                     "rester couvertes" if mode == "replace" else
                                     "ajouter une ligne pour chacune dans ce plan ou dans un autre plan"))
        between = self._between_plans(results, contents, mode)
        return {"results": results, "between": between, "requirements": self.requirements,
                "plans": PlansInMemory(plans),
                "hypothesis": ("les plans des tomes concernés sont remplacés par le fichier" if mode == "replace" else
                               "chaque ligne du fichier remplace celle de la même exigence dans les plans de la base, "
                               "ou s'ajoute ; les autres lignes restent")
                + " ; les lignes à régler sont supposées corrigées, les exigences inconnues de la base ignorées ; le "
                  "nom du fichier n'est pas utilisé"}

    @staticmethod
    def _places(places):
        """Places of plan rows, table by table: « tableau 1, lignes 3, 4 ; tableau 2, ligne 7 »."""
        tables = {}
        for t, r in places:
            tables.setdefault(t, []).append(str(r))
        return " ; ".join(f"tableau {t}, ligne{'s' if len(rows) > 1 else ''} {listed(rows)}" for t, rows in tables.items())

    def _entries(self, content):
        """Entries (one per requirement) of a readable plan document, [] otherwise."""
        try:
            return PlansDatabase.entries(PlansDatabase.read_docx(content, self.config)["rows"])
        except ValueError:
            return []

    def _plan_in_memory(self, file_id, name, tome, release, content, settle):
        """A checked plan file as PlansDatabase.plans gives a stored plan: rows checked against the configuration,
        the rows to settle supposed corrected and the unknown or non applicable numbers left out."""
        rows = []
        data = PlansDatabase.read_docx(content, self.config)
        for e in PlansDatabase.entries(data["rows"]):
            c = settle.get(e["requirement"])
            if c and c["unknown"]:
                continue
            row = PlansDatabase.check_row({"key": e["requirement"], "requirements": [e["requirement"]],
                                           "iadt": e["iadt"], "moc": e["moc"], "delegation": e["delegation"],
                                           "manual": "", "deleted": False}, self.config)
            if c:
                row = {**row, "errors": []}
            rows.append(row)
        return {"file_id": file_id, "name": name, "tome": tome, "releases": [release], "rows": rows}

    @staticmethod
    def _plan_correction(entry, issue):
        """Correction proposed for an issue of a plan row (see PlansDatabase.to_confirm)."""
        if entry["unknown"]:
            return ("corriger le numéro, ou importer d'abord l'exigence dans son tome ; sinon la ligne sera ignorée"
                    if "inconnue" in issue else "corriger la release de l'exigence, ou retirer la ligne du plan")
        if issue.startswith("déjà dans un autre plan"):
            return "ne garder l'exigence que dans un plan du tome (le choix sera demandé à l'import)"
        if issue.startswith(("IADT vide", "IADT inconnu")):
            return "mettre un IADT de la configuration (I, A, D, T ou N/A) et le MOC attendu pour lui"
        if issue.startswith(("MOC vide", "MOC inconnu")) or "MOC" in issue:
            expected = " ou ".join(entry["expected_mocs"]) if entry["expected_mocs"] else ""
            return f"mettre le MOC attendu{' : ' + expected if expected else ''} (règles IADT / MOC de la configuration)"
        if issue.startswith("plusieurs délégations"):
            return "ne garder qu'une plateforme dans la cellule Délégation"
        if issue.startswith("plateforme de délégation inconnue"):
            return "corriger la cellule Délégation, ou ajouter la plateforme dans Configuration > Plateformes"
        return "corriger la ligne du plan"

    def _between_plans(self, results, contents, mode):
        """Problems between plan files: same file twice, one requirement in several files, and when they overwrite the
        base, one tome concerned by several files (the last one imported would win)."""
        found = []
        names = [n for n, _ in contents]
        for n in sorted({n for n in names if names.count(n) > 1}):
            found.append(problem(SETTLE, f"fichier {n} sélectionné plusieurs fois", "n'en garder qu'un"))
        numbers = {}
        for name, entries in contents:
            for e in entries:
                numbers.setdefault(e["requirement"], []).append(name)
        pairs = {}
        for number, files in numbers.items():
            if len(set(files)) > 1:
                pairs.setdefault(" et ".join(sorted(set(files))), []).append(number)
        for files, ids in pairs.items():
            found.append(problem(SETTLE, f"{plural(len(ids), 'exigence')} dans plusieurs plans ({files}) : {listed(ids)}",
                                 "une exigence n'a qu'une ligne de plan : la retirer de l'un d'eux (sinon, la ligne du "
                                 "dernier importé l'emporte)"))
        if mode == "replace":
            by_tome = {}
            for r in results:
                for t in r.get("tomes", []):
                    by_tome.setdefault(t, []).append(r["name"])
            for t, files in sorted(by_tome.items()):
                if len(files) > 1:
                    found.append(problem(SETTLE, f"tome {t} concerné par {len(files)} plans : {', '.join(files)}",
                                         "un import qui écrase la base remplace les plans du tome : seul le dernier "
                                         "importé resterait ; les réunir, ou importer les suivants en complément"))
        return found

    # ------------------------------------------------------------------ #
    # Result and report
    # ------------------------------------------------------------------ #

    @staticmethod
    def _verdict(problems):
        """(verdict, counts by severity) of a list of problems: its most serious severity."""
        order = [BLOCKING, SETTLE, CHECK, INFO]
        severities = [x["severity"] for x in problems]
        return next((s for s in order if s in severities), INFO), {s: severities.count(s) for s in order}

    @staticmethod
    def _counts_text(results, key):
        """« 1 refusé, 2 avec des choix à faire, 3 chargeables » for the verdicts results[i][key]."""
        refused = sum(r[key] == BLOCKING for r in results)
        settle = sum(r[key] == SETTLE for r in results)
        loadable = sum(r[key] in (CHECK, INFO) for r in results)
        return (f"{refused} refusé{'s' if refused > 1 else ''}, {settle} avec des choix à faire, "
                f"{loadable} chargeable{'s' if loadable > 1 else ''}")

    def _result(self, step, results, between, requirements, plans, hypothesis, modes=None, between_by_mode=None):
        """Verdict of each file, tiles before and after, and the HTML report. With `modes` ({mode: (requirements,
        plans) in memory}), the files are judged for each import: verdict and tiles per mode, `hypothesis` being
        then {mode: text}."""
        for r in results:
            r["verdict"], r["counts"] = self._verdict(r["problems"])
            for m, extra in (r.get("by_mode") or {}).items():
                r.setdefault("verdicts", {})[m] = self._verdict(r["problems"] + extra)[0]
        before = Verification(self.config, self.requirements, self.plans, self.stdhl, self.fiches,
                              self.traceability).run(())
        afters = {m: Verification(self.config, q, pl, self.stdhl, self.fiches, self.traceability).run(())
                  for m, (q, pl) in (modes or {None: (requirements, plans)}).items()}
        labels = [t["label"] for t in next(iter(afters.values())).get("tiles", []) or before.get("tiles", [])]
        tiles = []
        for k, label in enumerate(labels):
            b = before.get("tiles", [])
            row = {"label": label, "before": self._tile_text(b[k]) if k < len(b) else "—",
                   "after": {m: self._tile_text(a["tiles"][k]) if k < len(a.get("tiles", [])) else "—"
                             for m, a in afters.items()}}
            row["changed"] = any(v != row["before"] for v in row["after"].values())
            # A stage without file (« à charger » on the page) before and after import: its row is left out.
            if row["before"] == "—" and all(v == "—" for v in row["after"].values()):
                continue
            tiles.append(row)
        summary = {"files": len(results), "between": len(between) + sum(len(v) for v in (between_by_mode or {}).values())}
        if modes:
            summary["text"] = (f"{plural(len(results), 'fichier vérifié', 'fichiers vérifiés')}. "
                               + " ; ".join(f"S'il {MODE_SHORT[m]} : "
                                            + self._counts_text([{"v": r["verdicts"][m]} for r in results if "verdicts" in r]
                                                                + [{"v": r["verdict"]} for r in results if "verdicts" not in r], "v")
                                            for m in modes)
                               + (f" ; {plural(summary['between'], 'problème')} entre les fichiers" if summary["between"] else "")
                               + ". Rien n'a été importé.")
        else:
            summary["text"] = (f"{plural(len(results), 'fichier vérifié', 'fichiers vérifiés')} : "
                               + self._counts_text(results, "verdict")
                               + (f" ; {plural(len(between), 'problème')} entre les fichiers" if between else "")
                               + ". Rien n'a été importé.")
        stamp = datetime.now()
        return {"summary": summary, "file_name": f"verification-{'exigences' if step == 'requirements' else 'plans'}-{stamp:%Y-%m-%d-%H%M}.html",
                "html": self._html(step, results, between, tiles, hypothesis, stamp, list(modes or []),
                                   between_by_mode or {})}

    @staticmethod
    def _tile_text(tile):
        """Value of a tile as the page shows it; a stage without file (« à charger » on the page) is « — »."""
        if tile.get("waiting"):
            return "—"
        if tile.get("expected") is None:
            return str(tile["value"])
        return f"{tile['value']} / {tile['expected']} ({tile['percent']} %)"

    def _html(self, step, results, between, tiles, hypothesis, stamp, modes, between_by_mode):
        """The self-contained report."""
        e = lambda text: html.escape(str(text))
        badge = lambda s: f'<span class="badge {s}">{SEVERITY_LABELS[s]}</span>'
        verdict = lambda s: f'<span class="verdict {s}">{VERDICTS[s]}</span>'
        title = f"Vérification avant import — {STEP_TITLES[step]}"
        out = [f"""<!doctype html><html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>{e(title)}</title><style>
:root {{ --text:#1c2330; --muted:#5e6878; --border:#d8dce3; --bg:#ffffff; --head:#eef0f4;
  --ko:#b42318; --ko-bg:#fde8e6; --warn:#8a5a00; --warn-bg:#fdf1d6; --ok:#1d7a46; --ok-bg:#e2f4ea;
  --info:#1f5fbf; --info-bg:#e7f1fb; }}
body {{ font-family: "Segoe UI", system-ui, sans-serif; color: var(--text); background: var(--bg); margin: 24px auto;
  max-width: 1100px; padding: 0 16px; font-size: 14px; line-height: 1.45; }}
h1 {{ font-size: 1.6rem; margin: 0 0 4px; color: #0569a3; padding-bottom: 6px; border-bottom: 3px solid #10aed6; }}
h2 {{ font-size: 1.15rem; margin: 30px 0 8px; padding: 7px 12px; color: #0569a3; background: #e8f4fb;
  border-left: 5px solid #0569a3; border-radius: 4px; }}
h3 {{ font-size: 1rem; margin: 16px 0 6px; padding: 2px 0 2px 10px; color: #1c2330; border-left: 4px solid #10aed6; }}
.mode {{ margin-left: 10px; }} .mode:first-child {{ margin-left: 0; }}
.meta {{ color: var(--muted); margin: 0 0 16px; }}
table {{ border-collapse: collapse; width: 100%; margin: 8px 0 12px; }}
th, td {{ border: 1px solid var(--border); padding: 6px 8px; text-align: left; vertical-align: top; }}
th {{ background: var(--head); font-weight: 600; }}
td.n {{ text-align: right; font-variant-numeric: tabular-nums; }}
.badge, .verdict {{ display: inline-block; border-radius: 10px; padding: 1px 8px; font-size: .85em; font-weight: 600;
  white-space: nowrap; }}
.bloquant {{ background: var(--ko-bg); color: var(--ko); }} .regler {{ background: var(--warn-bg); color: var(--warn); }}
.verifier {{ background: var(--info-bg); color: var(--info); }} .info {{ background: var(--ok-bg); color: var(--ok); }}
tr.changed td {{ font-weight: 600; }} .facts {{ color: var(--muted); margin: 0 0 6px; }}
.hint {{ color: var(--muted); font-size: .9em; }}
@media print {{ body {{ margin: 0; max-width: none; }} h2 {{ break-after: avoid; }} tr {{ break-inside: avoid; }} }}
</style></head><body>
<h1>{e(title)}</h1>
<p class="meta">{stamp:%d/%m/%Y %H:%M} · base en release {e(self.config.get("release"))} · {plural(len(results), "fichier vérifié", "fichiers vérifiés")} ·
rien n'a été importé</p>"""]
        mode_head = "".join(f"<th>{e(MODE_SHORT[m].capitalize())}</th>" for m in modes)
        out.append("<h2>Synthèse</h2>")
        if modes:
            out.append("<p>Chaque fichier est vérifié pour les deux imports : s'il <b>écrase la base</b> (import en "
                       + ("remplacement : les exigences des tomes concernés absentes du fichier sont retirées) et s'il "
                          "<b>modifie la base</b> (import en complément : rien n'est retiré).</p>" if step == "requirements"
                          else "remplacement : les plans des tomes concernés sont remplacés par le fichier) et s'il "
                          "<b>modifie la base</b> (import en complément : ses lignes remplacent celles des mêmes "
                          "exigences, les autres restent).</p>"))
        out.append("<table><thead><tr><th>Fichier</th><th>Tomes de la base concernés</th>"
                   + (mode_head if modes else "<th>Résultat</th>")
                   + "<th>Bloquants</th><th>À régler</th><th>À vérifier</th></tr></thead><tbody>")
        for k, r in enumerate(results):
            cells = ("".join(f"<td>{verdict(r.get('verdicts', {}).get(m, r['verdict']))}</td>" for m in modes)
                     if modes else f"<td>{verdict(r['verdict'])}</td>")
            out.append(f'<tr><td><a href="#f{k}">{e(r["name"])}</a></td><td>{e(r["tome"] or "—")}</td>{cells}'
                       f'<td class="n">{r["counts"][BLOCKING]}</td><td class="n">{r["counts"][SETTLE]}</td>'
                       f'<td class="n">{r["counts"][CHECK]}</td></tr>')
        out.append("</tbody></table>")
        if modes:
            out.append('<p class="hint">Bloquants, À régler et À vérifier : les problèmes communs aux deux imports ; '
                       "ceux d'un seul import sont détaillés fichier par fichier.</p>")
        if tiles:
            after_head = "".join(f"<th>Après import, s'il {e(MODE_SHORT[m])}</th>" for m in modes) or "<th>Après import</th>"
            out.append(f"<h2>Couverture avant et après import</h2><table><thead><tr><th>Étape</th><th>Avant</th>"
                       f"{after_head}</tr></thead><tbody>")
            for t in tiles:
                out.append(f'<tr class="{"changed" if t["changed"] else ""}"><td>{e(t["label"])}</td>'
                           f'<td class="n">{e(t["before"])}</td>'
                           + "".join(f'<td class="n">{e(v)}</td>' for v in t["after"].values()) + "</tr>")
            out.append("</tbody></table>")
            if modes:
                out += [f'<p class="hint">Calcul de la page Vérification, sans rien importer. S\'il '
                        f'{e(MODE_SHORT[m])} : {e(hypothesis[m])}.</p>' for m in modes]
            else:
                out.append(f'<p class="hint">Calcul de la page Vérification, sans rien importer : {e(hypothesis)}.</p>')
        if between or any(between_by_mode.values()):
            out.append("<h2>Entre les fichiers</h2>")
            if between:
                out.append(self._problem_table(between, e, badge))
            for m, found in between_by_mode.items():
                if found:
                    out.append(f"<h3>{e(MODE_LABELS[m])}</h3>" + self._problem_table(found, e, badge))
        for k, r in enumerate(results):
            title = (" ".join(f'<span class="mode">s\'il {e(MODE_SHORT[m])} :</span> {verdict(r["verdicts"][m])}'
                              for m in modes) if "verdicts" in r else verdict(r["verdict"]))
            out.append(f'<h2 id="f{k}">{e(r["name"])}</h2><p class="facts">{title}</p>')
            facts = r["facts"]
            if facts:
                out.append(f'<p class="facts">{e(" · ".join(facts))}</p>')
            if modes and "by_mode" in r:
                out.append("<h3>Dans les deux cas</h3>")
            out.append(self._problem_table(r["problems"], e, badge) if r["problems"] else "<p>Aucun problème.</p>")
            for m, found in (r.get("by_mode") or {}).items():
                out.append(f"<h3>{e(MODE_LABELS[m])}</h3>")
                extra = r.get("mode_facts", {}).get(m)
                if extra:
                    out.append(f'<p class="facts">{e(" · ".join(extra))}</p>')
                out.append(self._problem_table(found, e, badge) if found else "<p>Rien de plus.</p>")
        out.append("</body></html>")
        return "\n".join(out)

    @staticmethod
    def _problem_table(problems, e, badge):
        """Table of problems, the most serious first."""
        order = {BLOCKING: 0, SETTLE: 1, CHECK: 2, INFO: 3}
        rows = sorted(problems, key=lambda x: order[x["severity"]])
        with_place = any(x["place"] for x in rows)
        head = "<th>Criticité</th><th>Problème</th>" + ("<th>Emplacement</th>" if with_place else "") + "<th>Correction</th>"
        body = "".join(f'<tr><td>{badge(x["severity"])}</td><td>{e(x["text"])}</td>'
                       + (f'<td>{e(x["place"])}</td>' if with_place else "") + f'<td>{e(x["correction"])}</td></tr>'
                       for x in rows)
        return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"
