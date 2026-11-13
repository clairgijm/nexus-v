"""Check of the traceability of NEXUS-V: the whole chain, from the requirements to the fiches d'essai, and its report.

For each requirement, run() gives a status at each stage (acceptance plan, STDHL, fiches), the stage where it is lost
and its coverage, then the verdict, the tiles and the pie charts the page shows as they are. report() writes the same
result into an Excel workbook. Doubtful cases are settled when the files are imported: there is no list of anomalies
here.

Rules:

  - a plan covers a requirement of its tome; a STDHL or a fiches file covers a requirement of its tome, or of every
    tome when it belongs to none;
  - IADT and MOC N/A on all its plan rows: the traceability of the requirement stops at the plan (TERMINÉE);
  - only the requirements delegated to the default platform (TOP) are specified in a STDHL: for the others the STDHL
    stage does not apply (AUTRE PLATEFORME); with use_requirement_stdhl, a requirement without value in the column
    STDHL is not checked there (NON DEMANDÉ);
  - a requirement is covered by a fiche of the platform its plan delegates it to; for TOP, each test the STDHL
    associates with it must cite it in the fiches (otherwise TEST MANQUANT);
  - a stage without any file loaded is to load: it is not a defect.
"""
import io
import re
import time

import openpyxl
from openpyxl.styles import Font, PatternFill

from configuration import fold, plural

SOURCE_NAME = "Exigences"
STAGE_PLAN, STAGE_STDHL, STAGE_FICHES = "Plan d'acceptation", "STDHL", "Fiches d'Essai"
STAGES = [STAGE_PLAN, STAGE_STDHL, STAGE_FICHES]

# Statuses of a requirement at a stage; the codes are also the style names of the page.
OK, MISSING, NO_DOC, ERROR, PENDING, ENDED = "ok", "ko", "nodoc", "err", "wait", "ended"
OTHER = "other"   # STDHL stage of a requirement delegated to another platform than the default one: not expected
UNTAGGED = "none"  # STDHL stage of a requirement without STDHL (column STDHL empty, taken into account): not checked
TEST_MISSING = "test"   # Fiches stage: covered by the fiches, but not by every test the STDHL asks for
STATUS_LABELS = {OK: "OK", MISSING: "MANQUANT", NO_DOC: "DOC NON CHARGÉ", ERROR: "ERREUR", PENDING: "À CHARGER",
                 ENDED: "TERMINÉE", OTHER: "AUTRE PLATEFORME", UNTAGGED: "NON DEMANDÉ", TEST_MISSING: "TEST MANQUANT"}

# Coverage chart of the Fiches d'essai, against the delegation of the plan: one part per platform (requirements
# delegated to it and covered by its fiches), and the requirements not covered by their delegated platform.
# Requirements whose plan rows are all N/A (IADT and MOC): not tested, a part of their own in the pie charts.
# Short labels (the hints of the charts explain them).
NOT_COVERED, NO_PLAN, ENDED_PLAN = "Non couvert", "Absent", "N/A"
NEUTRAL = (NO_PLAN, NOT_COVERED)   # « nothing found » parts of the charts
# Reserved colours (page styles), the same in every chart: bright red not covered and absent from the plans, grey N/A
# (not tested); green for the covered part of a chart of one platform (page Fiches d'Essai). No platform takes them.
COVERED_COLOR = "var(--ok)"
COVERED = "Couvert"
# Results of the fiches XML: green passed, red failed, grey not run (or another value).
PASSED, FAILED, NOT_RUN = "PASSED", "FAILED", "Non exécuté"
RESERVED = {NOT_COVERED: "var(--not-covered)", ENDED_PLAN: "var(--muted)", NO_PLAN: "var(--not-covered)", COVERED: COVERED_COLOR,
            PASSED: "var(--ok)", FAILED: "var(--not-covered)", NOT_RUN: "var(--muted)"}
REPORT_COLOURS = {OK: "C6EFCE", MISSING: "FFC7CE", ERROR: "FFC7CE", NO_DOC: "FFEB9C", TEST_MISSING: "FFC7CE"}   # status cells of the report


def unique(values):
    """Non-empty values without duplicates, in their order."""
    return list(dict.fromkeys(v for v in values if v))


def by_name(text):
    """Sort key of a label: accents and case ignored, numbers in their numeric order."""
    return [int(p) if i % 2 else p for i, p in enumerate(re.split(r"(\d+)", fold(text)))]


def percent(part, total):
    """Whole percentage, a half being rounded up."""
    return (200 * part + total) // (2 * total) if total else 0


class Verification:
    """Check of the whole chain, from the data of the document classes."""

    def __init__(self, config, requirements, plans, stdhl, fiches, traceability):
        self.config = config
        self.requirements = requirements
        self.plans = plans
        self.stdhl = stdhl
        self.fiches = fiches
        self.traceability = traceability

    # ------------------------------------------------------------------ #
    # Scope of the documents
    # ------------------------------------------------------------------ #

    @staticmethod
    def _in_release(doc, requirement, selected):
        """The document applies to one of the requirement's releases (restricted to the selection, if any)."""
        # Without a selection every release of the requirement counts; with one, only the selected releases it has.
        wanted = [r for r in requirement["releases"] if r in selected] if selected else requirement["releases"]
        return any(r in wanted for r in doc["releases"])

    @staticmethod
    def _of_tome(docs, tome):
        """STDHL or fiches files that apply to a tome: those attached to it, and those without tome (every tome)."""
        return [d for d in docs if d["tome"] is None or d["tome"] == tome]

    # ------------------------------------------------------------------ #
    # Check
    # ------------------------------------------------------------------ #

    def run(self, releases=()):
        """Result of the check for the selected releases (undeclared ones are left out), for the page. When there is
        nothing to check: {"message": why, empty without any requirement}."""
        # Only the declared releases can be selected (one release for the base: the selection is that release or none).
        declared = self.config.get("releases")
        releases = [r for r in releases if r in declared]
        selected = set(releases)
        tomes = self.requirements.to_dict()["tomes"]
        # Every requirement of the base, flattened with its tome; the ones deleted by hand are not checked.
        every = [{**r, "tome": tome} for tome, t in tomes.items() for r in t["requirements"] if not r["deleted"]]
        if not every:
            return {"message": ""}
        # The requirements checked: those of the selected releases (all of them without selection).
        requirements = [r for r in every if not selected or selected & set(r["releases"])]
        if not requirements:
            return {"message": f"Aucune exigence pour {'les releases' if len(releases) > 1 else 'la release'} "
                               f"{', '.join(releases)}."}

        # Content of every document, read once from the base; each row then looks its requirement up in them.
        plans = self.plans.plans(config=self.config)
        stdhl = self.stdhl.documents()
        # Fiches d'essai: the XML ones (default platform) and the Excel ones of the other platforms.
        fiches = self.fiches.documents() + self.traceability.documents()
        rows = self._rows(requirements, plans, stdhl, fiches, selected)
        # Stages that still have no file: named in the verdict, and their tiles say « à charger ».
        waiting = [name for k, name in enumerate(STAGES) if any(r["statuses"][k] == PENDING for r in rows)]
        return {"names": STAGES, "rows": rows, "releases": releases, "waiting": waiting,
                "status_labels": STATUS_LABELS, "defects": sum(1 for r in rows if r["lost_at"]),
                "lost": [[name, sum(1 for r in rows if r["lost_at"] == name)] for name in STAGES],
                "verdict": self._verdict(rows, waiting, releases), "tiles": self._tiles(rows, waiting, releases),
                "charts": self._charts(rows, plans, fiches), "platform_charts": self._platform_charts(rows, fiches),
                "result_charts": self._result_charts(rows, fiches)}

    def _sheets(self, fiches, requirement, selected):
        """Fiches files that may cover a requirement: one of its releases, and its tome (or no tome)."""
        return [x for x in self._of_tome(fiches, requirement["tome"]) if self._in_release(x, requirement, selected)]

    def _rows(self, requirements, plans, stdhl, fiches, selected):
        """One row per requirement: its statuses at each stage and its coverage."""
        # Index of the plan rows by requirement number, each row with its plan (tome, name), so that each requirement
        # finds its rows without scanning every plan.
        plan_hits = {}
        for plan in plans:
            for row in plan["rows"]:
                for req in row["requirements"]:
                    plan_hits.setdefault(req, []).append({**row, "plan": plan})
        # Index of the STDHL rows by file then by requirement number: a requirement only looks into the STDHL that apply
        # to its tome.
        stdhl_hits = {}
        for doc in stdhl:
            for row in doc["rows"]:
                for req in row["requirements"]:
                    stdhl_hits.setdefault(doc["file_id"], {}).setdefault(req, []).append(row)
        rows = []
        for r in requirements:
            req_id = r["id"]
            # Plan rows of the requirement, in a plan of its own tome (a number cited by a plan of another tome does not
            # count).
            hits = [h for h in plan_hits.get(req_id, [])
                    if h["plan"]["tome"] == r["tome"] and self._in_release(h["plan"], r, selected)]
            # STDHL that apply to the requirement (its tome, or no tome) and fiches files that cite it.
            docs = [d for d in self._of_tome(stdhl, r["tome"]) if self._in_release(d, r, selected)]
            sheets = [x for x in self._sheets(fiches, r, selected) if req_id in x["ids"]]
            # Platforms the requirement is delegated to (its plan rows); a fiche of one of them covers it.
            platforms = set(unique(h["effective_delegation"] for h in hits))
            # Rows of those STDHL that cite the requirement, each with the test (fiche SYS_) it asks for.
            found = [row for d in docs for row in stdhl_hits.get(d["file_id"], {}).get(req_id, [])]
            # N/A on every plan row (IADT and MOC): the requirement is not tested, its traceability stops at the plan.
            ended = bool(hits) and all(h["ended"] for h in hits)
            # Only the requirements delegated to the default platform are specified in a STDHL.
            other = bool(hits) and not ended and self.config.get("default_delegation") not in platforms
            # The column STDHL is optional: a requirement without STDHL is not checked against the STDHL.
            untagged = self.config.get("use_requirement_stdhl") and not r["stdhl"]
            # Tests asked by the STDHL for the requirement, and tests of the fiches of the default platform that
            # cover it: every test asked must be there (more tests in the fiches are fine).
            default = self.config.get("default_delegation")
            asked = unique(row.get("test", "") for row in found)
            tested = unique(c["test"] for x in sheets if x["platform"] == default for c in x["rows"]
                            if req_id in c["requirements"])
            absent = [t for t in asked if t not in tested]
            # Coverage is judged against the delegation of the plan: only the fiches of a delegated platform count,
            # and those of the default platform only with every test asked by the STDHL.
            covering = unique(x["platform"] for x in sheets if x["platform"] in platforms
                              and not (x["platform"] == default and absent))
            own = bool(covering)
            # Covered by fiches of the default platform, but not by every test the STDHL asks: TEST MANQUANT, not
            # MANQUANT.
            partial = not own and default in platforms and bool(tested) and bool(absent)
            # One status per stage, in the order of STAGES. Plan: to load, missing, a row with an error, or OK. STDHL:
            # ended, another platform, not asked, to load, found, missing in the STDHL of its tome, or no STDHL loaded
            # for its tome. Fiches: ended, to load, covered, a test missing, or missing.
            statuses = [
                PENDING if not plans else MISSING if not hits else ERROR if any(h["errors"] for h in hits) else OK,
                ENDED if ended else OTHER if other else UNTAGGED if untagged else PENDING if not stdhl
                else OK if found else MISSING if docs
                else NO_DOC,
                ENDED if ended else PENDING if not fiches else OK if own else TEST_MISSING if partial else MISSING,
            ]
            # The stage where the requirement is lost: the first one that is a defect (OK, to load, ended, another
            # platform and not asked are not defects).
            lost = next((name for name, s in zip(STAGES, statuses) if s not in (OK, PENDING, ENDED, OTHER, UNTAGGED)), "")
            # Result of the tests of the requirement, by platform covering it: FAILED if one failed, PASSED if all
            # passed, Non exécuté otherwise (the Excel files of the other platforms have no result).
            def result_of(values):
                return FAILED if "FAILED" in values else PASSED if values and all(v == "PASSED" for v in values) \
                    else NOT_RUN
            results = {p: result_of([c.get("result") for x in sheets if x["platform"] == p for c in x["rows"]
                                     if req_id in c["requirements"]]) for p in ([] if ended else covering)}
            fiche_result = result_of(list(results.values())) if results else None
            rows.append({**r, "statuses": statuses, "lost_at": lost, "fiche_result": fiche_result,
                         "results_by": results,
                         # Delegated platforms whose fiches cover the requirement (None: ended at the plan).
                         "covered_by": None if ended else covering, "delegated_to": sorted(platforms),
                         "iadt": ", ".join(unique(c for h in hits for c in h["iadt_codes"])),
                         "moc": unique(c for h in hits for c in h["moc_codes"]),
                         "delegation": ", ".join(unique(h["effective_delegation"] for h in hits)),
                         # Tests asked by the STDHL, found in the fiches (default platform), asked but not found.
                         "tests_asked": asked, "tests_found": tested, "tests_absent": absent})
        return rows

    # ------------------------------------------------------------------ #
    # Figures for the page
    # ------------------------------------------------------------------ #

    @staticmethod
    def _scope(releases):
        """Selected releases, as shown after a title."""
        return f" (release{'s' if len(releases) > 1 else ''} {', '.join(releases)})" if releases else ""

    def _verdict(self, rows, waiting, releases):
        """{kind: style of the verdict, text}."""
        # Three cases: defects (red), nothing wrong yet but files to load (waiting), everything consistent (green).
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
        """Number of requirements, then OK / expected for each stage; ended requirements are not expected, nor in
        the STDHL the requirements delegated to another platform."""
        tiles = [{"label": f"{SOURCE_NAME}{self._scope(releases)}", "value": len(rows)}]
        for k, name in enumerate(STAGES):
            if name in waiting:
                tiles.append({"label": name, "waiting": True})
                continue
            # OK out of the requirements expected at this stage: N/A requirements, those of another platform at the
            # STDHL stage and those without STDHL asked are not expected.
            done = sum(1 for r in rows if r["statuses"][k] == OK)
            expected = sum(1 for r in rows if r["statuses"][k] not in (ENDED, OTHER, UNTAGGED))
            tiles.append({"label": name, "value": done, "expected": expected, "percent": percent(done, expected)})
        return tiles

    def _chart(self, title, hint, parts, by_platform=False, colors=None, names=None):
        """A pie chart from [(label, [requirement rows])]: the parts present (each one with the numbers of its
        requirements, listed when the part is clicked), the smallest ones folded into « Autres » beyond chart_max_parts
        (configuration), or chart_max_parts_platforms for a chart with one part per platform. The reserved parts
        (Non couvert, N/A, Absent) are always shown and do not count in the maximum."""
        # A chart by platform may show up to chart_max_parts_platforms parts (one per platform), the others
        # chart_max_parts.
        maximum = self.config.get("chart_max_parts_platforms" if by_platform else "chart_max_parts")
        # Each part keeps the numbers of its requirements (`ids`): the page lists them when the part is clicked. Empty
        # parts are left out.
        shown = [(label, [r["id"] for r in rs]) for label, rs in parts if rs]
        # Beyond the maximum, the largest parts are kept (one place left for « Autres ») and the smallest ones are
        # merged into « Autres »; the reserved parts (Non couvert, N/A, Absent, Couvert) are never merged.
        categories = [(label, ids) for label, ids in shown if label not in RESERVED]
        if len(categories) > maximum:
            kept = {label for label, _ in sorted(categories, key=lambda p: -len(p[1]))[:maximum - 1]}
            rest = [i for label, ids in categories if label not in kept for i in ids]
            shown = [(label, ids) for label, ids in shown if label in kept or label in RESERVED]
            # « Autres » before the reserved parts.
            at = next((i for i, (label, _) in enumerate(shown) if label in RESERVED), len(shown))
            shown.insert(at, ("Autres", rest))
        # Parts (and legend) from the largest to the smallest.
        shown.sort(key=lambda p: -len(p[1]))
        total = sum(len(ids) for _, ids in shown)
        # A part named after a platform has the colour of that platform, the same in every chart.
        # A part named after a value of a referential (platform, IADT) has its colour, the same in every chart.
        colors = self.config.get("colors").get("platforms" if by_platform else colors, {}) if by_platform or colors \
            else {}
        return {"title": title, "hint": hint,
                # `names`: label shown for a code (IADT: its meaning), the colour staying the one of the code.
                "parts": [{"label": label if label in RESERVED else (names or {}).get(label) or label,
                           "count": len(ids), "percent": percent(len(ids), total), "neutral": label in NEUTRAL,
                           "ids": ids,
                           **({"color": RESERVED.get(label) or colors[label]} if label in RESERVED or label in colors
                              else {})} for label, ids in shown]}

    def _platform_charts(self, rows, fiches):
        """One pie chart per platform that has requirements delegated to it in the plans: covered by its fiches or
        not (page Fiches d'Essai). Requirements ended at the plan do not count."""
        if not fiches:
            return []
        # Requirements ended at the plan (covered_by None) are not expected in any fiche.
        expected = [r for r in rows if r["covered_by"] is not None]
        known = list(self.config.get("platforms"))
        # Platforms of the configuration first, in its order, then any other platform named by a plan.
        platforms = known + sorted({p for r in expected for p in r["delegated_to"]} - set(known), key=by_name)
        charts = []
        for platform in platforms:
            delegated = [r for r in expected if platform in r["delegated_to"]]
            if delegated:
                chart = self._chart(
                    f"Plateforme {platform}", f"Exigences déléguées à {platform} dans les plans d'acceptation.",
                    [(COVERED, [r for r in delegated if platform in r["covered_by"]]),
                     (NOT_COVERED, [r for r in delegated if platform not in r["covered_by"]])])
                for part in chart["parts"]:   # the covered part in green (reserved colour)
                    if part["label"] != NOT_COVERED:
                        part["color"] = COVERED_COLOR
                charts.append(chart)
        return charts

    def _result_charts(self, rows, fiches):
        """One pie chart per platform covering requirements with its fiches: the state of their tests (PASSED,
        FAILED, Non exécuté), like the coverage by platform (page Statistiques)."""
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
        """Requirements by platform and by IADT, then their coverage. A requirement is in one plan row only, with one
        delegation: each one counts in one part. Order of the parts: from the largest to the smallest (see _chart)."""
        # By IADT, the requirements N/A in the plans (not tested) have their own part; by platform, they count with
        # their platform like the others.
        tested = [r for r in rows if r["covered_by"] is not None]
        not_tested = [r for r in rows if r["covered_by"] is None]

        def split(values, known, with_na=True):
            """Parts by value: of the tested requirements then N/A, or of every requirement (with_na false); then
            absent from the plans."""
            counted = tested if with_na else rows
            # Values of the configuration first, then the other values met in the plans; a requirement without any value
            # is absent from the plans (part Absent).
            labels = known + sorted({v for r in counted for v in values(r)} - set(known), key=by_name)
            return [(label, [r for r in counted if label in values(r)]) for label in labels] \
                + ([(ENDED_PLAN, not_tested)] if with_na else []) + [(NO_PLAN, [r for r in counted if not values(r)])]

        charts = []
        if plans:
            hint = "N/A : exigence N/A dans les plans (non testée) ; Absent : exigence absente des plans."
            # By platform: every requirement, N/A included, with the platform it is delegated to.
            charts.append(self._chart("Exigences par plateforme", "Toutes les exigences, N/A compris, par plateforme "
                                      "de délégation. Absent : exigence absente des plans.", split(
                lambda r: r["delegation"].split(", ") if r["delegation"] else [], list(self.config.get("platforms")),
                with_na=False), by_platform=True))
            charts.append(self._chart("Exigences par IADT", hint, split(
                lambda r: r["iadt"].split(", ") if r["iadt"] else [], list(self.config.get("iadt"))), colors="iadt",
                names=self.config.get("iadt")))
        # Coverage by the fiches: covered (green) or not (red), the N/A requirements left out. A requirement is
        # covered by a fiche of the platform it is delegated to; one absent from the plans is not covered.
        tested = [r for r in rows if r["covered_by"] is not None]
        # The coverage chart only makes sense once fiches are loaded and some requirements are tested.
        if fiches and tested:
            charts.append(self._chart(
                "Couverture par les fiches",
                "Exigences testées (hors N/A) : couvertes par une fiche de la plateforme à laquelle elles sont déléguées "
                "dans le plan, ou non couvertes (une fiche d'une autre plateforme ne compte pas ; une exigence absente "
                "des plans n'est pas couverte).",
                [(COVERED, [r for r in tested if r["covered_by"]]),
                 (NOT_COVERED, [r for r in tested if not r["covered_by"]])]))
        # Results of the fiches XML: one part per requirement they cite.
        with_result = [r for r in rows if r.get("fiche_result")]
        if with_result:
            charts.append(self._chart(
                "Résultats des fiches d'essai",
                "Exigences couvertes par les fiches (hors N/A), toutes plateformes : PASSED si tous leurs tests sont "
                "passés, FAILED si l'un a échoué, Non exécuté sinon.",
                [(label, [r for r in with_result if r["fiche_result"] == label]) for label in (PASSED, FAILED, NOT_RUN)]))
        return charts

    # ------------------------------------------------------------------ #
    # Excel report
    # ------------------------------------------------------------------ #

    def report(self, releases=()):
        """(file name, bytes) of the complete report: one row per requirement with its statuses and its tests.
        ValueError when there is nothing to report."""
        result = self.run(releases)
        if "rows" not in result:
            raise ValueError(result["message"] or "aucune exigence à vérifier")
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

        # A sheet of the report: header row in bold and frozen, each column as wide as its longest value (12 to 60).
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
            ["Exigence", "Titre", "Tome", "Releases", "STDHL", "IADT", "MOC", "Délégation",
             *result["names"], "Perdue à l'étape", "Tests demandés (STDHL)", "Tests des fiches", "Tests manquants"],
            *([r["id"], r["title"], r["tome"], ", ".join(r["releases"]), r["stdhl"], r["iadt"], ", ".join(r["moc"]),
               r["delegation"], *(labels[s] for s in r["statuses"]), r["lost_at"], ", ".join(r["tests_asked"]),
               ", ".join(r["tests_found"]), ", ".join(r["tests_absent"])] for r in result["rows"])])
        # Status cells coloured like the badges of the page: green OK, red missing or error, yellow document not loaded.
        first = 9   # first column of the statuses
        for line, r in enumerate(result["rows"], 2):
            for k, status in enumerate(r["statuses"]):
                if status in REPORT_COLOURS:
                    workbook["Traçabilité"].cell(line, first + k).fill = PatternFill("solid", fgColor=REPORT_COLOURS[status])
        buffer = io.BytesIO()
        workbook.save(buffer)
        return f"NEXUS-V_rapport_{time.strftime('%Y-%m-%d')}.xlsx", buffer.getvalue()
