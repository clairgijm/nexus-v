"""NEXUS-V — check of files before their import, without importing them.

Several files of one stage (requirements workbooks, acceptance plans, or STDHL of format 1) are read with the same
functions as the import windows (RequirementsDatabase.analyse, PlansDatabase.analyse, StdhlDatabase): each problem that would refuse a file, ask a
choice of the user or leave a requirement uncovered is listed with the correction to make. The files are also compared
with each other. Nothing is written: the base is only read, and the effect on the coverage is computed by running the
Verification on the requirements, plans and STDHL of the base with, in memory, the content of the files checked.
The result is a self-contained HTML report (French, like the page).
"""
import html
import re
from datetime import datetime

from configuration import plural
from plans_db import PlansDatabase
from requirements_db import RequirementsDatabase
from stdhl_db import StdhlDatabase, unique
from verification import Verification

# Severity of a problem, from the most serious: a refused file, a choice asked at import, a point to look at, a note.
BLOCKING, SETTLE, CHECK, INFO = "bloquant", "regler", "verifier", "info"
SEVERITY_LABELS = {BLOCKING: "Bloquant", SETTLE: "À régler", CHECK: "À vérifier", INFO: "Information"}
# Verdict of a file, from its most serious problem.
VERDICTS = {BLOCKING: "Refusé à l'import", SETTLE: "Chargeable, avec des choix à faire", CHECK: "Chargeable, à vérifier",
            INFO: "Chargeable"}
STEP_TITLES = {"requirements": "Exigences", "plans": "Plans d'acceptation", "stdhl": "STDHL", "fiches": "Fiches d'essai"}
FILE_NAMES = {"requirements": "exigences", "plans": "plans", "stdhl": "stdhl", "fiches": "fiches"}
MODE_LABELS = {"replace": "Si le fichier écrase la base (import en remplacement)",
               "complement": "Si le fichier modifie la base (import en complément)"}
MODE_SHORT = {"replace": "écrase la base", "complement": "modifie la base"}
FIELD_LABELS = {"title": "titre", "releases": "release", "stdhl": "STDHL", "iadt": "IADT", "moc": "MOC", "delegation": "délégation"}
LIMIT = 30   # numbers listed in one problem, beyond: « … et N autres »


def problem(severity, text, correction="", place=""):
    """One problem of a file: its severity, what is wrong, the correction to make and where (table, row)."""
    return {"severity": severity, "text": text, "correction": correction, "place": place}


def listed(numbers):
    """Numbers joined, the first LIMIT of them, then how many others."""
    numbers = list(numbers)
    more = len(numbers) - LIMIT
    return ", ".join(numbers[:LIMIT]) + (f" … et {more} autre{'s' if more > 1 else ''}" if more > 0 else "")


listed_ = listed   # for the methods where `listed` names a list


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


class StdhlInMemory:
    """The STDHL of the base as left by checked STDHL files: what Verification reads."""

    def __init__(self, documents):
        """documents: in the form of StdhlDatabase.documents."""
        self.list = documents

    def documents(self):
        return self.list


class FichesInMemory:
    """The fiches XML of the base as left by checked fiches files: what Verification reads."""

    def __init__(self, documents):
        """documents: in the form of FichesDatabase.documents."""
        self.list = documents

    def documents(self, source="xml"):
        return self.list


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
                            modes={m: (runs[m]["requirements"], runs[m]["plans"], runs[m].get("stdhl", self.stdhl))
                                   for m in MODE_LABELS},
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
                    "corriger le classeur : la ligne 1 doit porter les en-têtes attendus, dans n'importe quel ordre "
                    "(textes réglés dans Configuration > Exigences)"
                    if "en-tête" in str(e) else "corriger la référence des lignes citées (IADT, MOC, délégation)"))
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
            for platform, ids in (data.get("new_platforms") or {}).items():
                p.append(problem(CHECK, f"nouvelle plateforme à ajouter : {platform} ({plural(len(ids), 'exigence')} "
                                        "déléguée" + ("s" if len(ids) > 1 else "") + " : " + listed(ids) + ")",
                                 "elle sera ajoutée aux plateformes à l'import ; sinon, corriger la délégation"))
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
                    what = "à la délégation modifiée" if field == "delegation" else f"au {label} modifié"
                    p.append(problem(INFO, f"{plural(len(changed), 'exigence')} de la base {what} "
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
    # STDHL (format 1, one per tome)
    # ------------------------------------------------------------------ #

    def check_stdhl(self, files):
        """Check STDHL files [(name, bytes)] of format 1 for both imports: what happens if each file overwrites the
        base (it replaces the STDHL of the tomes it concerns) and if it modifies it (its rows replace those of the same
        requirements, the others stay). The tome named in the file name is not used: each requirement of the STDHL is
        looked up in the whole base, and the tests asked for it compared with those of the STDHL of the base."""
        runs = {mode: self._stdhl_mode(files, mode) for mode in MODE_LABELS}
        return self._both_modes("stdhl", runs)

    def _stdhl_mode(self, files, mode):
        """Check of STDHL files for an import in `mode` (see check_stdhl): results, problems between files, and the
        STDHL of the base as each file in turn would leave them."""
        release = self.config.get("release")
        default = self.config.get("default_delegation")
        tome_of = {r["id"]: r["tome"] for r in self.requirements.requirements() if not r["deleted"]}
        # From the plans: the requirements N/A (left out of a STDHL) and those delegated to another platform than the
        # default one (not specified in a STDHL of the default platform).
        hits = {}
        for q in self.plans.plans(config=self.config):
            for row in q["rows"]:
                hits.setdefault(row["key"], []).append(row)
        ended = {i for i, rows in hits.items() if all(r["ended"] for r in rows)}
        other = {i for i, rows in hits.items() if i not in ended
                 and default not in {r["effective_delegation"] for r in rows}}
        # The STDHL of the base, copied: each file in turn replaces or removes rows in them.
        docs = [{**d, "rows": list(d["rows"])} for d in self.stdhl.documents()]
        results, contents = [], []
        for k, (name, content) in enumerate(files, 1):
            found = {"name": name, "tome": "", "problems": [], "facts": [], "tomes": []}
            results.append(found)
            p = found["problems"]
            try:
                data = StdhlDatabase.read_docx(content, name, self.config)
            except ValueError as e:
                data = {"tables": 0, "error": str(e)}
            if not data["tables"]:
                p.append(problem(BLOCKING, "fichier illisible ou refusé : " + data.get("error", "aucun tableau STDHL "
                                           "valide (format par tome) ni test du format unique"),
                                 "format par tome : au moins un tableau dont la première ligne porte les en-têtes "
                                 "Description du test, Résultat attendu et Exigences associées ; format unique : des "
                                 "titres de test « <nom>_<n°> : <titre> » (Configuration > STDHL)"))
                continue
            for refusal in data["refusals"]:
                p.append(problem(BLOCKING, refusal, f"corriger l'applicabilité du test (RELEASE "
                                                    f"{self.config.get('release')}) : le fichier n'est pas importable"))
            for issue in data["issues"]:
                p.append(problem(CHECK, issue, "corriger le document"))
            # Format « unique »: the associated requirements of a test must be those of its test cases; a gap
            # refuses the file at import.
            gaps = {}
            for x in data["gaps"]:
                gaps.setdefault((x["test"], x["side"]), []).append(x["requirement"])
            for (test, side), ids in gaps.items():
                where = ("citée{s} dans les test cases mais absente{s} des exigences associées" if side == "test case"
                         else "associée{s} citée{s} dans aucun test case").format(s="s" if len(ids) > 1 else "")
                p.append(problem(BLOCKING, f"test {test} : {plural(len(ids), 'exigence')} {where} : " + listed(ids),
                                 "mettre dans les exigences associées exactement les exigences des test cases : le "
                                 "fichier n'est pas importable"))
            # Pairs requirement / test asked, once each, in the order of the document.
            pairs = list(dict.fromkeys((i, r["fiche"]) for r in data["rows"] for i in r["requirements"]))
            contents.append((name, data["stdhl_name"], pairs))
            found["facts"].append(f"STDHL « {data['stdhl_name']} » · "
                                  + (f"format unique : {plural(data['tables'], 'test')}, {plural(data['cases'], 'test case')} · "
                                     if data["format"] == "unique" else "format par tome : " +
                                     f"{plural(data['tables'], 'tableau valide', 'tableaux valides')} · ")
                                  + f"{plural(len({i for i, _ in pairs}), 'exigence')}, "
                                  f"{plural(len({t for _, t in pairs if t}), 'test')}")
            unknown = list(dict.fromkeys(i for i, _ in pairs if i not in tome_of))
            if unknown:
                p.append(problem(SETTLE, f"{plural(len(unknown), 'exigence inconnue', 'exigences inconnues')} de la "
                                         f"base, non enregistrée{'s' if len(unknown) > 1 else ''} : " + listed(unknown),
                                 "corriger le numéro, ou importer d'abord l'exigence"))
            na = list(dict.fromkeys(i for i, _ in pairs if i in ended))
            if na:
                p.append(problem(INFO, f"{plural(len(na), 'exigence')} N/A dans les plans, ignorée"
                                       f"{'s' if len(na) > 1 else ''} : " + listed(na),
                                 "une exigence N/A n'est pas testée : la retirer du STDHL"))
            elsewhere = list(dict.fromkeys(i for i, _ in pairs if i in other))
            if elsewhere:
                p.append(problem(CHECK, f"{plural(len(elsewhere), 'exigence')} déléguée"
                                        f"{'s' if len(elsewhere) > 1 else ''} à une autre plateforme que {default} dans "
                                        f"les plans : " + listed(elsewhere),
                                 f"un STDHL ne spécifie que les exigences de {default} : les retirer, ou corriger la "
                                 "délégation dans le plan"))
            no_test = list(dict.fromkeys(i for i, t in pairs if not t))
            if no_test:
                p.append(problem(CHECK, f"{plural(len(no_test), 'exigence')} sans test SYS_ : " + listed(no_test),
                                 "mettre le nom du test (paragraphe SYS_…) avant son tableau"))
            if any(d["stdhl_name"] == data["stdhl_name"] for d in docs if d["file_id"] > 0):
                p.append(problem(INFO, f"un STDHL de même nom (« {data['stdhl_name']} ») est déjà chargé : l'import "
                                       "le remplacera"))

            # Each pair against the STDHL of the base that apply to the tome of the requirement: a requirement newly
            # specified, a requirement whose tests change, or unchanged.
            kept = [(i, t) for i, t in pairs if i in tome_of and i not in ended]
            touched = {}
            for i, _ in kept:
                touched.setdefault(tome_of[i], set()).add(i)
            found["tomes"] = sorted(touched)
            found["tome"] = ", ".join(found["tomes"]) or "aucun"
            if len(touched) > 1 and data["format"] != "unique":   # a STDHL « unique » joins the tomes on purpose
                p.append(problem(CHECK, f"le STDHL contient des exigences de {len(touched)} tomes de la base : "
                                        + ", ".join(sorted(touched)),
                                 "un STDHL par tome : le scinder, ou le charger sans tome (il vaut alors pour tous)"))
            asked = {}
            for d in docs:
                for row in d["rows"]:
                    for i in row["requirements"]:
                        if d["tome"] is None or d["tome"] == tome_of.get(i):
                            asked.setdefault(i, set()).add(row["test"])
            tests = {}
            for i, t in kept:
                tests.setdefault(i, set()).add(t)
            new = [i for i in tests if i not in asked]
            changed = [i for i in tests if i in asked and asked[i] != tests[i]]
            found["facts"].append(f"par rapport aux STDHL de la base : {len(new)} exigence(s) nouvellement spécifiée(s), "
                                  f"{len(changed)} au test modifié, {len(tests) - len(new) - len(changed)} inchangée(s)")
            if touched:
                found["facts"].append("tomes de la base concernés : "
                                      + ", ".join(f"{t} ({len(ids)})" for t, ids in sorted(touched.items())))
            if new:
                p.append(problem(INFO, f"{plural(len(new), 'exigence')} sans STDHL dans la base, spécifiée"
                                       f"{'s' if len(new) > 1 else ''} par ce fichier : " + listed(new)))
            if changed:
                p.append(problem(INFO, f"{plural(len(changed), 'exigence')} de la base au test demandé modifié : "
                                       + listed(changed)))

            specified_before = {i for d in docs for row in d["rows"] for i in row["requirements"]}
            # The STDHL after this file. Overwrite: the STDHL of the tomes concerned give way to the file. Modify: the
            # rows of the requirements of the file give way to its rows, the others stay.
            if mode == "replace":
                docs = [d for d in docs if d["tome"] not in touched]
            else:
                keys = set(tests)
                docs = [{**d, "rows": [row for row in d["rows"] if not set(row["requirements"]) & keys]} for d in docs]
            for t in sorted(touched):
                docs.append({"file_id": -k, "name": name, "stdhl_name": data["stdhl_name"], "tome": t,
                             "releases": [release], "issues": [],
                             "rows": [{"requirements": [i], "test": test} for i, test in kept if tome_of[i] == t]})
            # Requirements of the default platform of the tomes concerned without STDHL: those that had none already
            # (in both imports), and those that lose theirs with the file.
            specified = {i for d in docs for row in d["rows"] for i in row["requirements"]}
            for t in sorted(touched):
                top = [i for i, tome in tome_of.items() if tome == t and i in hits and i not in ended and i not in other]
                before = [i for i in top if i not in specified_before]
                lost = [i for i in top if i in specified_before and i not in specified]
                still = [i for i in before if i not in specified]
                if still:
                    p.append(problem(CHECK, f"{plural(len(still), 'exigence')} de {default} du tome {t} déjà sans STDHL "
                                            f"dans la base, et absente{'s' if len(still) > 1 else ''} du fichier : "
                                            + listed(still),
                                     "les spécifier dans ce STDHL ou dans un autre STDHL du tome"))
                if lost:
                    p.append(problem(CHECK, f"{plural(len(lost), 'exigence')} de {default} du tome {t} perd"
                                            f"{'ent leur' if len(lost) > 1 else ' son'} STDHL avec l'import : " + listed(lost),
                                     "les ajouter au fichier si elles doivent rester spécifiées"))
        between = self._between_stdhl(results, contents, mode)
        return {"results": results, "between": between, "requirements": self.requirements, "plans": self.plans,
                "stdhl": StdhlInMemory(docs),
                "hypothesis": ("les STDHL des tomes concernés sont remplacés par le fichier" if mode == "replace" else
                               "les lignes du fichier remplacent celles des mêmes exigences dans les STDHL de la base, "
                               "ou s'ajoutent ; les autres lignes restent")
                + " ; les exigences inconnues de la base et les exigences N/A sont ignorées ; format unique : la "
                  "couverture d'un test est l'ensemble des exigences de ses test cases ; le nom du fichier n'est pas utilisé"}

    def _between_stdhl(self, results, contents, mode):
        """Problems between STDHL files: same file or same STDHL name twice, one requirement in several files, and when
        they overwrite the base, one tome concerned by several files."""
        found = []
        names = [n for n, _, _ in contents]
        for n in sorted({n for n in names if names.count(n) > 1}):
            found.append(problem(SETTLE, f"fichier {n} sélectionné plusieurs fois", "n'en garder qu'un"))
        by_name = {}
        for n, stdhl_name, _ in contents:
            by_name.setdefault(stdhl_name, []).append(n)
        for stdhl_name, files in sorted(by_name.items()):
            if len(set(files)) > 1:
                found.append(problem(SETTLE, f"même nom de STDHL « {stdhl_name} » dans {len(files)} fichiers : "
                                             + ", ".join(files),
                                     "un STDHL de même nom remplace le précédent : seul le dernier importé resterait"))
        numbers = {}
        for n, _, pairs in contents:
            for i, _ in pairs:
                numbers.setdefault(i, []).append(n)
        groups = {}
        for i, files in numbers.items():
            if len(set(files)) > 1:
                groups.setdefault(" et ".join(sorted(set(files))), []).append(i)
        for files, ids in groups.items():
            found.append(problem(SETTLE, f"{plural(len(ids), 'exigence')} dans plusieurs STDHL ({files}) : {listed(ids)}",
                                 "une exigence n'est spécifiée que dans un STDHL : la retirer de l'un d'eux"))
        if mode == "replace":
            by_tome = {}
            for r in results:
                for t in r.get("tomes", []):
                    by_tome.setdefault(t, []).append(r["name"])
            for t, files in sorted(by_tome.items()):
                if len(files) > 1:
                    found.append(problem(SETTLE, f"tome {t} concerné par {len(files)} STDHL : {', '.join(files)}",
                                         "un import qui écrase la base remplace les STDHL du tome : seul le dernier "
                                         "importé resterait ; les réunir, ou importer les suivants en complément"))
        return found

    # ------------------------------------------------------------------ #
    # Fiches d'essai
    # ------------------------------------------------------------------ #

    def check_fiches(self, files):
        """Check fiches d'essai XML [(name, bytes, platform)], formats « balises » and « campagne », each one coming
        from its platform (the default one if empty), with the rules of the import
        (FichesDatabase.screen). A fiche does not modify the base: one result per file, the file loaded as it is (a
        file of the same name replaced). The tome named in the file name is not used: each requirement is looked up in
        the whole base. The files are taken in turn, each one after the previous ones."""
        from fiches_db import FichesDatabase
        release = self.config.get("release")
        default = self.config.get("default_delegation")
        tome_of = {r["id"]: r["tome"] for r in self.requirements.requirements() if not r["deleted"]}
        docs = self.fiches.documents()
        covered_before = {i for d in docs for i in d["ids"]}
        results, names = [], [f[0] for f in files]
        for k, (name, content, *rest) in enumerate(files, start=1):
            p, facts = [], []
            platform = (rest[0] if rest else "") or default
            if not self.config.is_platform(platform):
                raise ValueError(f"{name} : plateforme {platform} absente de la configuration")
            try:
                data = FichesDatabase._readable(content, self.config)
            except Exception as e:   # malformed XML, no requirement number
                results.append({"name": name, "tome": "", "platform": platform, "problems": [problem(BLOCKING, str(e),
                                "corriger le fichier (XML bien formé, balises ou format campagne)")], "facts": []})
                continue
            found = self.fiches.screen(data, None, self.config, platform)
            if found["na"]:
                p.append(problem(BLOCKING, f"{plural(len(found['na']), 'exigence N/A', 'exigences N/A')} dans les plans, "
                                           "citée" + ("s" if len(found["na"]) > 1 else "") + " par la fiche : "
                                           + listed(found["na"]),
                                 "changer d'abord leur état (IADT, MOC) dans le plan d'acceptation et le STDHL, ou les "
                                 "retirer de la fiche"))
            if found["not_specified"]:
                p.append(problem(BLOCKING, f"{plural(len(found['not_specified']), 'exigence')} dans aucun STDHL de son "
                                           "tome : " + listed(found["not_specified"]),
                                 "les spécifier d'abord dans le STDHL de leur tome, ou les retirer de la fiche"))
            if found["unknown"]:
                p.append(problem(INFO, f"{plural(len(found['unknown']), 'numéro')} absent"
                                       f"{'s' if len(found['unknown']) > 1 else ''} de la base des exigences, "
                                       "ignoré" + ("s" if len(found["unknown"]) > 1 else "") + " : " + listed(found["unknown"])))
            if found["delegated"]:
                p.append(problem(INFO, f"{plural(len(found['delegated']), 'exigence')} déléguée"
                                       f"{'s' if len(found['delegated']) > 1 else ''} à une autre plateforme que "
                                       f"{platform} : chargée{'s' if len(found['delegated']) > 1 else ''}, résultat non "
                                       "conservé : " + listed(found["delegated"])))
            if found["ignored_tests"]:
                p.append(problem(INFO, f"{plural(len(found['ignored_tests']), 'test')} qu'aucun STDHL ne demande, "
                                       "ignoré" + ("s" if len(found["ignored_tests"]) > 1 else "") + " : "
                                       + listed(found["ignored_tests"])))
            if found["ignored_cases"]:
                p.append(problem(INFO, f"{plural(len(found['ignored_cases']), 'test case')} absent"
                                       f"{'s' if len(found['ignored_cases']) > 1 else ''} du STDHL de leur test, "
                                       "ignoré" + ("s" if len(found["ignored_cases"]) > 1 else "") + " : "
                                       + listed(found["ignored_cases"])))
            if found["untitled"]:
                p.append(problem(INFO, f"{plural(len(found['untitled']), 'test case')} sans titre, ignoré"
                                       f"{'s' if len(found['untitled']) > 1 else ''} (test"
                                       f"{'s' if len(set(found['untitled'])) > 1 else ''} "
                                       + listed(unique(found["untitled"])) + ")"))
            # Results of the test cases kept (format campagne): those failed, those not run.
            cases = {}
            for test, case, ids, result in found["rows"]:
                if result != "NOT_RUN" or (test, case) not in cases:
                    cases[(test, case)] = result
            if data["format"] == "campagne":
                failed = [f"{t} / {c}" for (t, c), r in cases.items() if r == "FAILED"]
                not_run = [f"{t} / {c}" for (t, c), r in cases.items() if r == "NOT_RUN"]
                if failed:
                    p.append(problem(INFO, f"{plural(len(failed), 'test case')} FAILED : " + listed(failed)))
                if not_run:
                    p.append(problem(INFO, f"{plural(len(not_run), 'test case')} non exécuté"
                                           f"{'s' if len(not_run) > 1 else ''} : " + listed(not_run)))
            kept = unique(i for _, _, ids, _ in found["rows"] for i in ids
                          if i not in found["na"] and i not in found["not_specified"])
            touched = sorted({tome_of[i] for i in kept if i in tome_of})
            new = [i for i in kept if i not in covered_before]
            if new:
                p.append(problem(INFO, f"{plural(len(new), 'exigence')} sans fiche dans la base, couverte"
                                       f"{'s' if len(new) > 1 else ''} par ce fichier : " + listed(new)))
            if any(d["name"] == name for d in docs):
                p.append(problem(INFO, f"une fiche de même nom est déjà chargée : elle sera remplacée"))
            facts = [f"plateforme {platform}", f"format {data['format']}", plural(len({t for t, _ in cases}), "test"),
                     plural(len(cases), "test case"), plural(len(kept), "exigence chargée", "exigences chargées")]
            # The fiches after this file: a file of the same name gives way to it; refused, the base stays as it was.
            blocked = bool(found["na"] or found["not_specified"])
            if not blocked:
                docs = [d for d in docs if d["name"] != name]
                rows = [{"test_id": 0, "test": t, "test_case": c, "result": r, "requirements": list(ids)}
                        for t, c, ids, r in found["rows"]]
                docs.append({"file_id": -k, "name": name, "source": "xml", "platform": platform,
                             "tome": touched[0] if len(touched) == 1 else None, "releases": [release],
                             "ids": unique(i for row in rows for i in row["requirements"]), "rows": rows})
            results.append({"name": name, "tome": ", ".join(touched), "tomes": touched, "platform": platform,
                            "problems": p, "facts": facts,
                            "kept": kept})
        # Tests asked by the STDHL and absent from every fiche after the import, for the requirements of each file.
        after = Verification(self.config, self.requirements, self.plans, self.stdhl, FichesInMemory(docs),
                             self.traceability).run(())
        missing = {r["id"]: r.get("tests_absent") or [] for r in after.get("rows", []) if r["statuses"][2] == "test"}
        for r in results:
            absent = [f"{i} ({', '.join(missing[i])})" for i in r.pop("kept", []) if i in missing]
            if absent:
                r["problems"].append(problem(CHECK, f"{plural(len(absent), 'exigence')} sans tous les tests demandés par "
                                                    "le STDHL après import (TEST MANQUANT) : " + listed(absent),
                                             "ajouter les tests manquants dans une fiche"))
        between = [problem(SETTLE, f"fichier {n} sélectionné plusieurs fois", "n'en garder qu'un")
                   for n in sorted({n for n in names if names.count(n) > 1})]
        return self._result("fiches", results, between, self.requirements, self.plans,
                            "le fichier est chargé tel quel (une fiche de même nom est remplacée) ; les exigences "
                            "inconnues, les tests qu'aucun STDHL ne demande, les test cases absents du STDHL et les test "
                            "cases sans titre sont ignorés ; un fichier refusé laisse la base telle quelle ; le nom du "
                            "fichier n'est pas utilisé", fiches=FichesInMemory(docs))

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

    def _result(self, step, results, between, requirements, plans, hypothesis, modes=None, between_by_mode=None,
                fiches=None):
        """Verdict of each file, tiles before and after, and the HTML report. With `modes` ({mode: (requirements,
        plans) in memory}), the files are judged for each import: verdict and tiles per mode, `hypothesis` being
        then {mode: text}. `fiches`: the fiches after import (check of fiches files), those of the base otherwise."""
        for r in results:
            r["verdict"], r["counts"] = self._verdict(r["problems"])
            for m, extra in (r.get("by_mode") or {}).items():
                r.setdefault("verdicts", {})[m] = self._verdict(r["problems"] + extra)[0]
        before = Verification(self.config, self.requirements, self.plans, self.stdhl, self.fiches,
                              self.traceability).run(())
        afters = {m: Verification(self.config, q, pl, st, fiches or self.fiches, self.traceability).run(())
                  for m, (q, pl, st) in (modes or {None: (requirements, plans, self.stdhl)}).items()}
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
        return {"summary": summary, "file_name": f"verification-{FILE_NAMES[step]}-{stamp:%Y-%m-%d-%H%M}.html",
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
                          "exigences, les autres restent).</p>" if step == "plans"
                          else "remplacement : les STDHL des tomes concernés sont remplacés par le fichier) et s'il "
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
