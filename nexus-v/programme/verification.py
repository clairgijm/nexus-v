"""Vérification de la traçabilité de NEXUS-V : toute la chaîne, des exigences aux fiches d'essai, et son rapport.

Pour chaque exigence, run() donne un statut à chaque étape (plan d'acceptation, STDHL, fiches), l'étape où elle est
perdue et sa couverture, puis le verdict, les tuiles et les camemberts que la page montre tels quels. report() écrit
le même résultat dans un classeur Excel. Les cas douteux sont réglés à l'import des fichiers : il n'y a pas ici de
liste d'anomalies.

Règles :

  - un plan d'acceptation n'est chargé que s'il est identique aux exigences de son tome (plans_db.py) : un tome qui a
    un plan a toutes ses exigences à l'étape Plan ; un plan dont les exigences ont changé depuis est à recharger
    (PLAN À RECHARGER) ;
  - un STDHL ou un fichier de fiches couvre une exigence de son tome, ou de tous les tomes quand il n'appartient à
    aucun ;
  - une exigence qualifiée ou hors scope (champ validation, d'après « Validation releases ») n'est attendue ni dans
    les STDHL ni dans les fiches (QUALIFIÉE, HORS SCOPE) ;
  - seules les exigences déléguées à la plateforme de niveau système (configuration system_platform) sont spécifiées
    dans un STDHL : pour les autres, l'étape STDHL ne s'applique pas (AUTRE PLATEFORME) ;
  - une exigence est couverte par une fiche de la plateforme à laquelle elle est déléguée ; pour la plateforme de
    niveau système, chaque test que le STDHL lui associe doit la citer dans les fiches (sinon TEST MANQUANT) ;
  - une étape sans aucun fichier chargé est à charger : ce n'est pas un défaut ;
  - seules les étapes du processus retenu pour la release sont vérifiées (configuration process_plan, process_stdhl :
    Plan d'acceptation et STDHL facultatifs) ; sans STDHL, les fiches ne sont pas comparées aux tests demandés.
"""
import io
import re
import time
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill

from configuration import fold, plural

SOURCE_NAME = "Exigences"
STAGE_PLAN, STAGE_STDHL, STAGE_FICHES = "Plan d'acceptation", "STDHL", "Fiches d'Essai"
STAGES = [STAGE_PLAN, STAGE_STDHL, STAGE_FICHES]

# Statuts d'une exigence à une étape ; les codes sont aussi les noms de styles de la page.
OK, MISSING, NO_DOC, ERROR, PENDING = "ok", "ko", "nodoc", "err", "wait"
QUALIFIED, OUT_OF_SCOPE = "qualified", "scope"   # étapes STDHL et fiches : non attendue (champ validation)
NOT_EXPECTED = (QUALIFIED, OUT_OF_SCOPE)
OTHER = "other"   # étape STDHL d'une exigence d'une autre plateforme que celle de niveau système
TEST_MISSING = "test"   # étape Fiches : couverte par les fiches, mais pas par chaque test demandé par le STDHL
STALE = "stale"   # étape Plan : le plan du tome a été chargé avant le dernier changement de ses exigences
STATUS_LABELS = {OK: "OK", MISSING: "MANQUANT", NO_DOC: "DOC NON CHARGÉ", ERROR: "ERREUR", PENDING: "À CHARGER",
                 QUALIFIED: "QUALIFIÉE", OUT_OF_SCOPE: "HORS SCOPE", OTHER: "AUTRE PLATEFORME",
                 TEST_MISSING: "TEST MANQUANT", STALE: "PLAN À RECHARGER"}

# Camembert de couverture des Fiches d'essai, par rapport à la délégation des exigences : les exigences couvertes par
# les fiches de leur plateforme de délégation, et les autres (NOT RUN, voir plus bas). Libellés courts (les aides des
# camemberts les expliquent).
NO_PLAN = "Absent"
# Couleurs réservées (styles de la page), les mêmes dans tous les camemberts : rouge vif pour absent des plans et les
# tests FAILED ; vert pour la part couverte ; gris pour NOT RUN. Aucune plateforme ne les prend.
COVERED_COLOR = "var(--ok)"
COVERED = "Couvert"
# Résultats des fiches XML : vert passé, rouge échoué, gris non exécuté (ou autre valeur).
# NOT RUN : un test non exécuté, ou (couverture par les fiches) une exigence qu'aucune fiche de sa plateforme ne couvre
# encore ; en gris, le rouge est réservé aux tests FAILED.
PASSED, FAILED, NOT_RUN = "PASSED", "FAILED", "NOT RUN"
NEUTRAL = (NO_PLAN, NOT_RUN)   # parts « rien trouvé » des camemberts
RESERVED = {NO_PLAN: "var(--not-covered)", COVERED: COVERED_COLOR,
            PASSED: "var(--ok)", FAILED: "var(--not-covered)", NOT_RUN: "var(--muted)"}
REPORT_COLOURS = {OK: "C6EFCE", MISSING: "FFC7CE", STALE: "FFC7CE", NO_DOC: "FFEB9C", TEST_MISSING: "FFC7CE"}   # cellules de statut du rapport


def unique(values):
    """Valeurs non vides sans doublon, dans leur ordre."""
    return list(dict.fromkeys(v for v in values if v))


def by_name(text):
    """Clé de tri d'un libellé : accents et casse ignorés, nombres dans leur ordre numérique."""
    return [int(p) if i % 2 else p for i, p in enumerate(re.split(r"(\d+)", fold(text)))]


def percent(part, total):
    """Pourcentage entier, une moitié étant arrondie au-dessus."""
    return (200 * part + total) // (2 * total) if total else 0


class Verification:
    """Vérification de toute la chaîne, d'après les données des classes de documents."""

    def __init__(self, config, requirements, plans, stdhl, fiches, traceability):
        self.config = config
        self.requirements = requirements
        self.plans = plans
        self.stdhl = stdhl
        self.fiches = fiches
        self.traceability = traceability

    # ------------------------------------------------------------------ #
    # Portée des documents
    # ------------------------------------------------------------------ #

    @staticmethod
    def _in_release(doc, requirement, selected):
        """Le document s'applique à l'une des releases de l'exigence (limitées à la sélection, s'il y en a une)."""
        # Sans sélection, chaque release de l'exigence compte ; avec une sélection, seules les releases sélectionnées
        # qu'elle a.
        wanted = [r for r in requirement["releases"] if r in selected] if selected else requirement["releases"]
        return any(r in wanted for r in doc["releases"])

    @staticmethod
    def _of_tome(docs, tome):
        """Fichiers STDHL ou de fiches qui s'appliquent à un tome : ceux qui lui sont rattachés, et ceux sans tome
        (tous les tomes)."""
        return [d for d in docs if d["tome"] is None or d["tome"] == tome]

    # ------------------------------------------------------------------ #
    # Vérification
    # ------------------------------------------------------------------ #

    def run(self, releases=()) -> dict[str, Any]:
        """Résultat de la vérification pour les releases sélectionnées (celles non déclarées sont laissées de côté),
        pour la page. Quand il n'y a rien à vérifier : {"message": pourquoi, vide sans aucune exigence}."""
        # Seules les releases déclarées peuvent être sélectionnées (une release pour la base : la sélection est cette
        # release ou rien).
        declared = self.config.get("releases")
        releases = [r for r in releases if r in declared]
        selected = set(releases)
        tomes = self.requirements.to_dict()["tomes"]
        # Chaque exigence de la base, mise à plat avec son tome ; celles supprimées à la main ne sont pas vérifiées.
        every = [{**r, "tome": tome} for tome, t in tomes.items() for r in t["requirements"] if not r["deleted"]]
        if not every:
            return {"message": ""}
        # Les exigences vérifiées : celles des releases sélectionnées (toutes sans sélection).
        requirements = [r for r in every if not selected or selected & set(r["releases"])]
        if not requirements:
            return {"message": f"Aucune exigence pour {'les releases' if len(releases) > 1 else 'la release'} "
                               f"{', '.join(releases)}."}

        # Contenu de chaque document, lu une fois dans la base ; chaque ligne y cherche ensuite son exigence.
        plans = self.plans.plans(config=self.config)
        stdhl = self.stdhl.documents()
        # Fiches d'essai : les XML et les Excel, chacune de la plateforme choisie à son import.
        fiches = self.fiches.documents() + self.traceability.documents()
        rows = self._rows(requirements, plans, stdhl, fiches, selected)
        # Étapes qui n'ont encore aucun fichier : nommées dans le verdict, et leurs tuiles disent « à charger ».
        stages = self.stages()
        waiting = [name for k, name in enumerate(stages) if any(r["statuses"][k] == PENDING for r in rows)]
        return {"names": stages, "rows": rows, "releases": releases, "waiting": waiting,
                "status_labels": STATUS_LABELS, "defects": sum(1 for r in rows if r["lost_at"]),
                "lost": [[name, sum(1 for r in rows if r["lost_at"] == name)] for name in stages],
                "verdict": self._verdict(rows, waiting, releases), "tiles": self._tiles(rows, waiting, releases),
                "charts": self._charts(rows, plans, fiches), "tome_charts": self._tome_charts(rows, fiches),
                "platform_charts": self._platform_charts(rows, fiches),
                "result_charts": self._result_charts(rows, fiches),
                "result_tome_charts": self._result_tome_charts(rows, fiches)}

    @staticmethod
    def tome_steps(result):
        """Étapes validées de chaque tome d'après le résultat de run() : {tome: [étapes]}, parmi plan, stdhl et
        fiches (celles du processus retenu). Chaque étape est jugée seule, d'après les exigences du tome : plan quand
        toutes sont dans un plan à jour, stdhl quand toutes celles de la plateforme de niveau système sont citées par
        un STDHL (les autres ne sont pas attendues), fiches quand toutes celles à qualifier sont couvertes par les
        fiches. Une étape n'attend pas la précédente."""
        names = result.get("names") or STAGES   # les étapes du processus retenu, dans l'ordre
        passed = {STAGE_PLAN: (OK,), STAGE_STDHL: (OK, *NOT_EXPECTED, OTHER), STAGE_FICHES: (OK, *NOT_EXPECTED)}
        keys = {STAGE_PLAN: "plan", STAGE_STDHL: "stdhl", STAGE_FICHES: "fiches"}
        by_tome = {}
        for r in result.get("rows") or []:
            by_tome.setdefault(r["tome"], []).append(dict(zip(names, r["statuses"])))
        return {tome: [keys[name] for name in names if all(s[name] in passed[name] for s in rows)]
                for tome, rows in by_tome.items()}

    def stages(self):
        """Étapes du processus retenu pour la release, dans l'ordre : Plan d'acceptation et STDHL quand ils sont
        retenus (configuration process_plan, process_stdhl), puis Fiches d'essai."""
        return ([STAGE_PLAN] if self.config.get("process_plan") else [])             + ([STAGE_STDHL] if self.config.get("process_stdhl") else []) + [STAGE_FICHES]

    def _sheets(self, fiches, requirement, selected):
        """Fichiers de fiches qui peuvent couvrir une exigence : l'une de ses releases, et son tome (ou aucun tome)."""
        return [x for x in self._of_tome(fiches, requirement["tome"]) if self._in_release(x, requirement, selected)]

    def _rows(self, requirements, plans, stdhl, fiches, selected):
        """Une ligne par exigence : ses statuts à chaque étape et sa couverture."""
        # Plans chargés de chaque tome : un plan chargé est identique aux exigences de son tome ; il est à recharger
        # quand elles ont changé depuis (stale).
        plans_of = {}
        for plan in plans:
            plans_of.setdefault(plan["tome"], []).append(plan)
        # Index des lignes de STDHL par fichier puis par numéro d'exigence : une exigence ne regarde que les STDHL qui
        # s'appliquent à son tome.
        stdhl_hits = {}
        for doc in stdhl:
            for row in doc["rows"]:
                for req in row["requirements"]:
                    stdhl_hits.setdefault(doc["file_id"], {}).setdefault(req, []).append(row)
        rows = []
        stages = self.stages()
        use_stdhl = STAGE_STDHL in stages   # sans STDHL, les fiches ne sont pas comparées aux tests qu'il demande
        for r in requirements:
            req_id = r["id"]
            # Plans de son tome : à jour, ou seulement à recharger.
            own_plans = plans_of.get(r["tome"], [])
            planned = any(not p["stale"] for p in own_plans)
            # STDHL qui s'appliquent à l'exigence (son tome, ou aucun tome) et fichiers de fiches qui la citent.
            docs = [d for d in self._of_tome(stdhl, r["tome"]) if self._in_release(d, r, selected)] if use_stdhl else []
            sheets = [x for x in self._sheets(fiches, r, selected) if req_id in x["ids"]]
            # Plateforme à laquelle l'exigence est déléguée (sa référence) ; une fiche de cette plateforme la couvre.
            platforms = {r["delegation"]} if r.get("delegation") else set()
            # Lignes de ces STDHL qui citent l'exigence, chacune avec le test (fiche SYS_) qu'elle demande.
            found = [row for d in docs for row in stdhl_hits.get(d["file_id"], {}).get(req_id, [])]
            # Qualifiée ou hors scope (champ validation) : attendue ni dans les STDHL ni dans les fiches.
            ended = {"qualified": QUALIFIED, "out_of_scope": OUT_OF_SCOPE}.get(r.get("validation"))
            # Seules les exigences déléguées à la plateforme de niveau système (leur référence) sont spécifiées dans un
            # STDHL.
            system_platform = self.config.get("system_platform")
            other = not ended and fold(r.get("delegation") or "") != fold(system_platform)
            # Tests demandés par le STDHL pour l'exigence, et tests des fiches de la plateforme de niveau système qui la
            # couvrent : chaque test demandé doit y être (davantage de tests dans les fiches ne gêne pas).
            default = system_platform
            asked = unique(row.get("test", "") for row in found)
            tested = unique(c["test"] for x in sheets if x["platform"] == default for c in x["rows"]
                            if req_id in c["requirements"])
            absent = [t for t in asked if t not in tested]
            # La couverture est jugée par rapport à la délégation de l'exigence : seules les fiches d'une plateforme de
            # délégation comptent, et celles de la plateforme de niveau système seulement avec chaque test demandé par
            # le STDHL.
            covering = unique(x["platform"] for x in sheets if x["platform"] in platforms
                              and not (x["platform"] == default and absent))
            own = bool(covering)
            # Couverte par des fiches de la plateforme de niveau système, mais pas par chaque test demandé par le STDHL
            # : TEST MANQUANT, pas MANQUANT.
            partial = not own and default in platforms and bool(tested) and bool(absent)
            # Un statut par étape, dans l'ordre de STAGES. Plan : à charger, manquant, à recharger, ou OK. STDHL
            # : qualifiée ou hors scope, autre plateforme, non demandé, à charger, trouvée, manquante dans les STDHL de
            # son tome, ou aucun STDHL chargé pour son tome. Fiches : qualifiée ou hors scope, à charger, couverte, un
            # test manquant, ou manquante.
            every = {
                STAGE_PLAN: PENDING if not plans else OK if planned else STALE if own_plans else MISSING,
                STAGE_STDHL: ended if ended else OTHER if other else PENDING if not stdhl
                else OK if found else MISSING if docs
                else NO_DOC,
                STAGE_FICHES: ended if ended else PENDING if not fiches else OK if own else TEST_MISSING if partial
                else MISSING,
            }
            # Seules les étapes du processus retenu.
            statuses = [every[name] for name in stages]
            # L'étape où l'exigence est perdue : la première qui est un défaut (OK, à charger, qualifiée, hors scope,
            # autre plateforme et non demandé ne sont pas des défauts).
            lost = next((name for name, s in zip(stages, statuses)
                         if s not in (OK, PENDING, *NOT_EXPECTED, OTHER)), "")
            # Résultat des tests de l'exigence, par plateforme qui la couvre : FAILED si l'un a échoué, PASSED si tous
            # sont passés, NOT RUN sinon (les fichiers Excel des autres plateformes n'ont pas de résultat).
            def result_of(values):
                return FAILED if "FAILED" in values else PASSED if values and all(v == "PASSED" for v in values) \
                    else NOT_RUN
            results = {p: result_of([c.get("result") for x in sheets if x["platform"] == p for c in x["rows"]
                                     if req_id in c["requirements"]]) for p in ([] if ended else covering)}
            fiche_result = result_of(list(results.values())) if results else None
            rows.append({**r, "statuses": statuses, "lost_at": lost, "fiche_result": fiche_result,
                         "results_by": results,
                         # Plateformes de délégation dont les fiches couvrent l'exigence (None : qualifiée ou hors
                         # scope).
                         "covered_by": None if ended else covering, "delegated_to": sorted(platforms),
                         # Tests demandés par le STDHL, trouvés dans les fiches (plateforme de niveau système), demandés
                         # mais pas trouvés.
                         "tests_asked": asked, "tests_found": tested, "tests_absent": absent})
        return rows

    # ------------------------------------------------------------------ #
    # Chiffres pour la page
    # ------------------------------------------------------------------ #

    @staticmethod
    def _scope(releases):
        """Releases sélectionnées, telles qu'affichées après un titre."""
        return f" (release{'s' if len(releases) > 1 else ''} {', '.join(releases)})" if releases else ""

    def _verdict(self, rows, waiting, releases):
        """{kind: style du verdict, text}."""
        # Trois cas : des défauts (rouge), rien de faux pour l'instant mais des fichiers à charger (attente), tout est
        # cohérent (vert).
        total, defects, scope = len(rows), sum(1 for r in rows if r["lost_at"]), self._scope(releases)
        to_load = f" Reste à charger : {', '.join(waiting)}." if waiting else ""
        if defects:
            return {"kind": "ko", "text": f"Incohérences détectées{scope} : {plural(defects, 'exigence')} sur {total} "
                                          f"en défaut.{to_load}"}
        if waiting:
            return {"kind": "wait", "text": f"Vérification partielle{scope} : {plural(total, 'exigence')}, "
                                            f"aucun défaut pour l'instant.{to_load}"}
        return {"kind": "ok", "text": f"Cohérent{scope} : les {total} {SOURCE_NAME} sont présentes et correctes dans "
                                      f"tous les documents."}

    def _tiles(self, rows, waiting, releases):
        """Nombre d'exigences, puis OK / attendues pour chaque étape ; les exigences qualifiées ou hors scope ne sont
        pas attendues, ni, dans les STDHL, celles déléguées à une autre plateforme."""
        tiles = [{"label": f"{SOURCE_NAME}{self._scope(releases)}", "value": len(rows)}]
        for k, name in enumerate(self.stages()):
            if name in waiting:
                tiles.append({"label": name, "waiting": True})
                continue
            # OK sur les exigences attendues à cette étape : celles qualifiées ou hors scope, d'une autre plateforme à
            # l'étape STDHL et celles sans STDHL demandé ne sont pas attendues.
            done = sum(1 for r in rows if r["statuses"][k] == OK)
            expected = sum(1 for r in rows if r["statuses"][k] not in (*NOT_EXPECTED, OTHER))
            tiles.append({"label": name, "value": done, "expected": expected, "percent": percent(done, expected)})
        return tiles

    def _chart(self, title, hint, parts, by_platform=False, colors=None, names=None):
        """Un camembert d'après [(libellé, [lignes d'exigence])] : les parts présentes (chacune avec les numéros de
        ses exigences, listés quand la part est cliquée), les plus petites regroupées dans « Autres » au-delà de
        chart_max_parts (configuration), ou de chart_max_parts_platforms pour un camembert à une part par plateforme.
        Les parts réservées (NOT RUN, Absent) sont toujours montrées et ne comptent pas dans le maximum."""
        # Un camembert par plateforme peut montrer jusqu'à chart_max_parts_platforms parts (une par plateforme), les
        # autres chart_max_parts.
        maximum = self.config.get("chart_max_parts_platforms" if by_platform else "chart_max_parts")
        # Chaque part garde les numéros de ses exigences (`ids`) : la page les liste quand la part est cliquée. Les
        # parts vides sont laissées de côté.
        shown = [(label, [r["id"] for r in rs]) for label, rs in parts if rs]
        # Au-delà du maximum, les plus grandes parts sont gardées (une place laissée pour « Autres ») et les plus
        # petites sont regroupées dans « Autres » ; les parts réservées (NOT RUN, Absent, Couvert) ne sont jamais
        # regroupées.
        categories = [(label, ids) for label, ids in shown if label not in RESERVED]
        if len(categories) > maximum:
            kept = {label for label, _ in sorted(categories, key=lambda p: -len(p[1]))[:maximum - 1]}
            rest = [i for label, ids in categories if label not in kept for i in ids]
            shown = [(label, ids) for label, ids in shown if label in kept or label in RESERVED]
            # « Autres » avant les parts réservées.
            at = next((i for i, (label, _) in enumerate(shown) if label in RESERVED), len(shown))
            shown.insert(at, ("Autres", rest))
        # Parts (et légende) de la plus grande à la plus petite.
        shown.sort(key=lambda p: -len(p[1]))
        total = sum(len(ids) for _, ids in shown)
        # Une part qui porte le nom d'une plateforme a la couleur de cette plateforme, la même dans tous les camemberts.
        # Une part qui porte le nom d'une valeur d'un référentiel (plateforme, IADT) a sa couleur, la même dans tous les
        # camemberts.
        colors = self.config.get("colors").get("platforms" if by_platform else colors, {}) if by_platform or colors \
            else {}
        return {"title": title, "hint": hint,
                # `names` : libellé montré pour un code (IADT : sa signification), la couleur restant celle du code.
                "parts": [{"label": label if label in RESERVED else (names or {}).get(label) or label,
                           "count": len(ids), "percent": percent(len(ids), total), "neutral": label in NEUTRAL,
                           "ids": ids,
                           **({"color": RESERVED.get(label) or colors[label]} if label in RESERVED or label in colors
                              else {})} for label, ids in shown]}

    def _platform_charts(self, rows, fiches):
        """Un camembert par plateforme à laquelle des exigences sont déléguées : couvertes par ses
        fiches ou non (page Fiches d'Essai). Les exigences qualifiées ou hors scope ne comptent pas."""
        if not fiches:
            return []
        # Les exigences qualifiées ou hors scope (covered_by None) ne sont attendues dans aucune fiche.
        expected = [r for r in rows if r["covered_by"] is not None]
        known = list(self.config.get("platforms"))
        # Les plateformes de la configuration d'abord, dans son ordre, puis toute autre plateforme de délégation.
        platforms = known + sorted({p for r in expected for p in r["delegated_to"]} - set(known), key=by_name)
        charts = []
        for platform in platforms:
            delegated = [r for r in expected if platform in r["delegated_to"]]
            if delegated:
                chart = self._chart(
                    f"Plateforme {platform}", f"Exigences déléguées à {platform}.",
                    [(COVERED, [r for r in delegated if platform in r["covered_by"]]),
                     (NOT_RUN, [r for r in delegated if platform not in r["covered_by"]])])
                for part in chart["parts"]:   # la part couverte en vert (couleur réservée), NOT RUN en gris
                    if part["label"] != NOT_RUN:
                        part["color"] = COVERED_COLOR
                charts.append(chart)
        return charts

    def _tome_charts(self, rows, fiches):
        """Une barre par tome (page Statistiques, onglet Fiches) : ses exigences à qualifier couvertes par les fiches de
        chaque plateforme (celle à laquelle elles sont déléguées), à la couleur de la plateforme, puis NOT RUN."""
        if not fiches:
            return []
        expected = [r for r in rows if r["covered_by"] is not None]
        tomes = list(self.config.get("tomes"))
        tomes += sorted({r["tome"] for r in expected} - set(tomes), key=by_name)
        known = list(self.config.get("platforms"))
        platforms = known + sorted({p for r in expected for p in r["covered_by"]} - set(known), key=by_name)
        charts = []
        for tome in tomes:
            own = [r for r in expected if r["tome"] == tome]
            if own:
                chart = self._chart(
                    f"Tome {tome}", f"Exigences à qualifier du tome {tome}, par plateforme dont les fiches les couvrent.",
                    [(p, [r for r in own if p in r["covered_by"]]) for p in platforms]
                    + [(NOT_RUN, [r for r in own if not r["covered_by"]])], by_platform=True)
                # NOT RUN en bout de barre : le reste à couvrir du tome.
                chart["parts"].sort(key=lambda part: part["label"] == NOT_RUN)
                charts.append(chart)
        return charts

    def _result_tome_charts(self, rows, fiches):
        """Une barre par tome (page Statistiques, onglet Résultats) : ses exigences à qualifier, toutes plateformes,
        PASSED, FAILED, puis NOT RUN (aucun test exécuté, ou pas encore couverte par une fiche de sa plateforme), comme
        le camembert des résultats de la page d'accueil."""
        if not fiches:
            return []
        expected = [r for r in rows if r["covered_by"] is not None]
        tomes = list(self.config.get("tomes"))
        tomes += sorted({r["tome"] for r in expected} - set(tomes), key=by_name)
        charts = []
        for tome in tomes:
            own = [r for r in expected if r["tome"] == tome]
            if own:
                chart = self._chart(
                    f"Tome {tome}", f"Exigences à qualifier du tome {tome} : état de leurs tests.",
                    [(PASSED, [r for r in own if r.get("fiche_result") == PASSED]),
                     (FAILED, [r for r in own if r.get("fiche_result") == FAILED]),
                     (NOT_RUN, [r for r in own if r.get("fiche_result") not in (PASSED, FAILED)])])
                # Toujours dans l'ordre PASSED, FAILED, NOT RUN : les barres des tomes se comparent.
                order = (PASSED, FAILED, NOT_RUN)
                chart["parts"].sort(key=lambda part: order.index(part["label"]))
                charts.append(chart)
        return charts

    def _result_charts(self, rows, fiches):
        """Un camembert par plateforme qui couvre des exigences avec ses fiches : l'état de leurs tests (PASSED,
        FAILED, NOT RUN), comme la couverture par plateforme (page Statistiques)."""
        if not fiches:
            return []
        known = list(self.config.get("platforms"))
        platforms = known + sorted({p for r in rows for p in r.get("results_by") or {}} - set(known), key=by_name)
        charts = []
        for platform in platforms:
            covered = [r for r in rows if platform in (r.get("results_by") or {})]
            if covered:
                charts.append(self._chart(
                    f"Plateforme {platform}", f"Exigences couvertes par les fiches de {platform} : état de leurs tests.",
                    [(label, [r for r in covered if r["results_by"][platform] == label])
                     for label in (PASSED, FAILED, NOT_RUN)]))
        return charts

    def _charts(self, rows, plans, fiches):
        """Exigences par plateforme, par IADT et par MOC (leur référence, dans les exigences), puis leur couverture.
        Une exigence a une valeur de chaque : elle compte dans une part. Ordre des parts : de la plus grande à la
        plus petite (voir _chart)."""

        def split(field, known):
            """Parts par valeur d'un champ de la référence : les valeurs de la configuration, puis les autres
            rencontrées."""
            labels = known + sorted({r[field] for r in rows if r[field]} - set(known), key=by_name)
            return [(label, [r for r in rows if r[field] == label]) for label in labels]

        charts = []
        if rows:
            charts.append(self._chart("Exigences par plateforme", "Toutes les exigences, par plateforme de délégation "
                                      "(fichier d'exigences).", split("delegation", list(self.config.get("platforms"))),
                                      by_platform=True))
            # IADT et MOC : la même colonne de la page (« column »), l'un sous l'autre.
            charts.append({**self._chart("Exigences par IADT", "Toutes les exigences, par IADT (fichier d'exigences).",
                                         split("iadt", list(self.config.get("iadt"))), colors="iadt",
                                         names=self.config.get("iadt")), "column": "reference"})
            # Le camembert des MOC peut être masqué (Configuration > MOC, show_moc_chart).
            if self.config.get("show_moc_chart"):
                charts.append({**self._chart("Exigences par MOC", "Toutes les exigences, par MOC (fichier d'exigences).",
                                             split("moc", list(self.config.get("mocs"))), colors="mocs",
                                             names=self.config.get("mocs")), "column": "reference"})
        # Couverture par les fiches : couvertes (vert) ou non (rouge), les exigences qualifiées ou hors scope laissées
        # de côté. Une exigence est couverte par une fiche de la plateforme à laquelle elle est déléguée.
        tested = [r for r in rows if r["covered_by"] is not None]
        # Le camembert de couverture n'a de sens qu'une fois des fiches chargées et des exigences à tester.
        if fiches and tested:
            # Une part par plateforme (ses exigences couvertes par ses fiches, à sa couleur), puis NOT RUN.
            known = list(self.config.get("platforms"))
            covering = known + sorted({p for r in tested for p in r["covered_by"]} - set(known), key=by_name)
            charts.append(self._chart(
                "Couverture par les fiches",
                "Exigences à qualifier, par plateforme dont les fiches les couvrent (celle à laquelle elles sont "
                "déléguées ; une fiche d'une autre plateforme ne compte pas), ou NOT RUN : aucune fiche de leur "
                "plateforme ne les couvre encore.",
                [(p, [r for r in tested if p in r["covered_by"]]) for p in covering]
                + [(NOT_RUN, [r for r in tested if not r["covered_by"]])], by_platform=True))
        # Résultats des fiches : toutes les exigences à qualifier, de toutes les plateformes ; une exigence qu'aucune
        # fiche de sa plateforme ne couvre encore, ou dont aucun test n'est exécuté, est NOT RUN.
        if fiches and tested:
            charts.append(self._chart(
                "Résultats des fiches d'essai",
                "Exigences à qualifier, toutes plateformes : PASSED si tous leurs tests sont passés, FAILED si l'un a "
                "échoué, NOT RUN si aucun n'est exécuté ou si aucune fiche de leur plateforme ne les couvre.",
                [(PASSED, [r for r in tested if r.get("fiche_result") == PASSED]),
                 (FAILED, [r for r in tested if r.get("fiche_result") == FAILED]),
                 (NOT_RUN, [r for r in tested if r.get("fiche_result") not in (PASSED, FAILED)])]))
        return charts

    # ------------------------------------------------------------------ #
    # Rapport Excel
    # ------------------------------------------------------------------ #

    def report(self, releases=()):
        """(nom de fichier, octets) du rapport complet : une ligne par exigence avec ses statuts et ses tests.
        ValueError quand il n'y a rien à rapporter."""
        result = self.run(releases)
        if "rows" not in result:
            raise ValueError(result["message"] or "aucune exigence à vérifier")
        workbook = openpyxl.Workbook()
        if (active := workbook.active) is not None:   # la feuille vide du classeur neuf
            workbook.remove(active)

        # Une feuille du rapport : ligne d'en-tête en gras et figée, chaque colonne aussi large que sa plus longue
        # valeur (de 12 à 60).
        def sheet(title, rows):
            ws = workbook.create_sheet(title)
            for row in rows:
                ws.append(row)
            for cell in ws[1]:
                cell.font = Font(bold=True)
            ws.freeze_panes = "A2"
            for col, column in enumerate(ws.columns):
                ws.column_dimensions[column[0].column_letter].width = min(
                    60, max(12, *(len(str(row[col] if row[col] is not None else "")) + 2 for row in rows)))

        labels = result["status_labels"]
        sheet("Traçabilité", [
            ["Exigence", "Titre", "Tome", "Releases", "IADT", "MOC", "Délégation",
             *result["names"], "Perdue à l'étape", "Tests demandés (STDHL)", "Tests des fiches", "Tests manquants"],
            *([r["id"], r["title"], r["tome"], ", ".join(r["releases"]), r["iadt"], r["moc"],
               r["delegation"], *(labels[s] for s in r["statuses"]), r["lost_at"], ", ".join(r["tests_asked"]),
               ", ".join(r["tests_found"]), ", ".join(r["tests_absent"])] for r in result["rows"])])
        # Cellules de statut colorées comme les badges de la page : vert OK, rouge manquant ou erreur, jaune document
        # non chargé.
        first = 8   # première colonne des statuts
        for line, r in enumerate(result["rows"], 2):
            for k, status in enumerate(r["statuses"]):
                if status in REPORT_COLOURS:
                    workbook["Traçabilité"].cell(line, first + k).fill = PatternFill("solid", fgColor=REPORT_COLOURS[status])
        buffer = io.BytesIO()
        workbook.save(buffer)
        return f"NEXUS-V_rapport_{time.strftime('%Y-%m-%d')}.xlsx", buffer.getvalue()
