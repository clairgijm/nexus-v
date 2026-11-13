"""NEXUS-V — check of the traceability on the command line, with the same classes as the server (verification.py):
requirements (Excel) -> acceptance plan (Word) -> STDHL (Word) -> fiches d'essai (XML).

Two uses: without files, check the base of NEXUS-V (a copy of it: the server may be running); with files, check them
against each other, only the referentials and settings of the base being taken over. Exit code 0 when everything is
consistent, 1 otherwise. The help of the command line (HELP, in French) gives the usage.
"""
import argparse
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path

try:
    from configuration import Configuration, plural
    from fiches_db import FichesDatabase
    from file_database import FileDatabase
    from nexus_db import BASE_FOLDER, CONFIG_NAME, DATABASE_NAME, SCHEMA_NAME, NexusDatabase
    from plans_db import PlansDatabase
    from requirements_db import RequirementsDatabase
    from stdhl_db import StdhlDatabase
    from traceability_db import TraceabilityDatabase
    from verification import Verification
except ModuleNotFoundError as missing:
    sys.exit(f"Module Python manquant : {missing.name}\n"
             f"Installer les modules nécessaires : python -m pip install -r requirements.txt")

# Help of the command line, shown to the user (in French, like the interface).
HELP = """NEXUS-V — Le fil d’Ariane de la traçabilité, de la spécification à l'essai.

Vérification de la traçabilité en ligne de commande :
Exigences (Excel) -> Plan d'acceptation (Word) -> STDHL (Word) -> Fiches d'Essai (XML).

Deux usages :

  1. Sans fichier : vérifie la base de NEXUS-V, comme la page Vérification. La base est celle du dossier
     fichiers-references/base de NEXUS-V, ou --base. La base n'est pas modifiée : le travail se fait sur une copie, le serveur peut tourner.

  2. Avec des fichiers : les vérifie entre eux, sans toucher à la base dont seuls les référentiels et réglages sont
     repris. Toutes les exigences forment un seul tome, et chaque fichier s'applique à la release de la base et à
     tous les tomes.

Le code de sortie vaut 0 si tout est cohérent, 1 sinon (défaut, ou étape sans fichier).

Exemples :
    py programme/verif_coherence.py --rapport rapport.xlsx
    py programme/verif_coherence.py --base ../essai/base
    py programme/verif_coherence.py exigences.xlsx fiches_essai.xml --plan plan_A.docx plan_B.docx \\
        --stdhl 20261030_X_STDHL_ALARMES_rev02.docx 20261030_X_STDHL_COM_rev01.docx --release R2
"""

ROOT = Path(__file__).resolve().parent
BASE_DIR = ROOT.parent / BASE_FOLDER   # fichiers-references/base/ at the root of NEXUS-V
TOME = "exigences"   # single tome of the files given on the command line
SETTINGS = ("release", "platforms", "iadt", "mocs", "default_delegation")   # referentials taken over from the storage folder


class Workspace:
    """The classes of NEXUS-V on a temporary folder: what they write never reaches the storage folder."""

    def __init__(self, directory):
        self.sql = NexusDatabase(directory / DATABASE_NAME)
        self.config = Configuration(directory / "config.json", database=self.sql)
        self.files = FileDatabase(directory, self.sql, self.config)
        self.requirements = RequirementsDatabase(self.sql)
        self.plans = PlansDatabase(self.sql)
        self.stdhl = StdhlDatabase(self.sql)
        self.fiches = FichesDatabase(self.sql)
        self.traceability = TraceabilityDatabase(self.sql)
        self.verification = Verification(self.config, self.requirements, self.plans, self.stdhl, self.fiches,
                                         self.traceability)

    def close(self):
        """Release the temporary database."""
        self.sql.close()


def storage_folder(given):
    """Folder of the files of NEXUS-V: --base, otherwise fichiers-references/base of NEXUS-V."""
    return given or BASE_DIR


def copy_base(folder, target):
    """Copy the database and the settings of a storage folder, not the stored files."""
    if (folder / CONFIG_NAME).is_file():
        shutil.copy(folder / CONFIG_NAME, target / CONFIG_NAME)
    if (folder / SCHEMA_NAME).is_file():
        shutil.copy(folder / SCHEMA_NAME, target / SCHEMA_NAME)
    if (folder / DATABASE_NAME).is_file():
        # Through SQLite, not as a file: the copy is consistent even if the server is writing.
        # Opened read only: the real base is never written by the command line.
        source = sqlite3.connect(f"{(folder / DATABASE_NAME).resolve().as_uri()}?mode=ro", uri=True)
        copy = sqlite3.connect(target / DATABASE_NAME)
        source.backup(copy)
        copy.close()
        source.close()


def load_files(work, args, settings):
    """Load the files of the command line into an empty workspace."""
    work.config.update(settings)
    content = args.requirements.read_bytes()
    found = [work.config.get("release")]   # one release for the base
    work.requirements.import_file(work.files, content, args.requirements.name, TOME, sheet=args.sheet, config=work.config)
    for path in args.plan:
        # Nobody to ask: rows breaking the IADT/MOC rules are loaded as they are (several delegations stay an anomaly).
        work.plans.load_file(work.files, work.requirements, path.read_bytes(), path.name, TOME, found, config=work.config,
                             decisions={"accept_all": True})
    for path in args.stdhl:
        work.stdhl.load_file(work.files, work.requirements, path.read_bytes(), path.name, "", found, config=work.config)
    work.fiches.load_file(work.files, work.requirements, args.xml.read_bytes(), args.xml.name, "", found,
                          config=work.config)


def print_report(result):
    """Print the requirements with their status at each stage, then the summary and the anomalies."""
    rows, names, labels = result["rows"], result["names"], result["status_labels"]
    # Widths of the columns of the text table, from the longest values (bounded).
    w_id = max([len(r["id"]) for r in rows] + [10])
    w_rel = min(20, max([len(", ".join(r["releases"])) for r in rows] + [8]))
    w_doc = min(30, max([len(r["stdhl"]) for r in rows] + [5]))
    print(f"\n{plural(len(rows), 'exigence vérifiée', 'exigences vérifiées')}.\n")
    print("Exigence".ljust(w_id), "Releases".ljust(w_rel), "STDHL".ljust(w_doc),
          *[n[:18].center(18) for n in names], " Perdue à l'étape", sep=" | ")
    print("-" * (w_id + w_doc + w_rel + 21 * len(names) + 26))
    for r in rows:
        print(r["id"].ljust(w_id), ", ".join(r["releases"])[:w_rel].ljust(w_rel), r["stdhl"][:w_doc].ljust(w_doc),
              *[labels[s].center(18) for s in r["statuses"]], " " + r["lost_at"], sep=" | ")

    print("\nRésumé :")
    # One line per stage, from the tiles of the page (OK out of the requirements expected at that stage) and the
    # requirements lost at that stage.
    for k, (name, tile) in enumerate(zip(names, result["tiles"][1:])):
        if tile.get("waiting"):
            print(f"  {name} : à charger")
            continue
        faulty = [r["id"] for r in rows if r["lost_at"] == name]
        print(f"  {name} : {tile['value']}/{tile['expected']} présentes et correctes"
              + (f" — en défaut : {', '.join(faulty)}" if faulty else ""))


def arguments():
    """The options of the command line, checked."""
    parser = argparse.ArgumentParser(description=HELP, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("requirements", type=Path, nargs="?", help="Exigences (Excel) ; sans fichier, la base est vérifiée")
    parser.add_argument("xml", type=Path, nargs="?", help="Fiches d'Essai (XML)")
    parser.add_argument("--plan", type=Path, nargs="+", default=[], help="Un ou plusieurs Plans d'acceptation (Word)")
    parser.add_argument("--stdhl", type=Path, nargs="+", default=[], help="Un ou plusieurs documents STDHL (Word)")
    parser.add_argument("--base", dest="folder", type=Path,
                        help=f"Dossier de config.json, nexus-v-schema.sql et nexus-v.db (défaut : {BASE_DIR})")
    parser.add_argument("--release", help="Release des fichiers donnés (défaut : celle de la configuration)")
    parser.add_argument("--regex", help="Format d'un numéro d'exigence, avec des fichiers seulement (défaut : "
                                        "variable requirement_regex)")
    parser.add_argument("--feuille", dest="sheet", help="Feuille Excel à lire, avec des fichiers seulement")
    parser.add_argument("--rapport", dest="report", type=Path, help="Écrire aussi un rapport Excel (.xlsx)")
    args = parser.parse_args()
    given = [args.requirements, args.xml, *args.plan, *args.stdhl]
    # With files, the whole chain is needed; without any, the base is checked.
    if any(given) and not (args.xml and args.plan and args.stdhl):
        parser.error("avec des fichiers, il faut les Exigences, les Fiches d'Essai, --plan et --stdhl")
    if (args.regex or args.sheet) and not any(given):
        parser.error("--regex et --feuille s'utilisent avec des fichiers")
    for path in filter(None, given):
        if not path.is_file():
            parser.error(f"fichier introuvable : {path}")
    return args


def run(args, directory):
    """Run the check in a temporary folder. Returns True when everything is consistent."""
    folder = storage_folder(args.folder)
    copy_base(folder, directory)
    work = Workspace(directory)
    try:
        if args.requirements:
            # With files: the referentials and settings of the base are kept, then the copied base is replaced by an
            # empty one into which only the given files are loaded.
            settings = {name: work.config.get(name) for name in SETTINGS}
            if args.release:
                settings["release"] = args.release.upper()
            if args.regex:
                settings["requirement_regex"] = args.regex
            work.close()
            (directory / DATABASE_NAME).unlink(missing_ok=True)
            work = Workspace(directory)
            load_files(work, args, settings)
        releases = []   # one release for the base: everything is checked
        result = work.verification.run(releases)
        if "rows" not in result:
            sys.exit(result["message"] or f"Aucune exigence dans la base de {folder}.")
        print_report(result)
        if args.report:
            args.report.write_bytes(work.verification.report(releases)[1])
            print(f"\nRapport écrit : {args.report}")
        print(f"\n{result['verdict']['text']}")
        # Consistent: no requirement lost and no stage left without file.
        return not result["defects"] and not result["waiting"]
    finally:
        work.close()


def main():
    """Read the options, run the check and give its result as exit code."""
    args = arguments()
    with tempfile.TemporaryDirectory(prefix="nexus-v-") as directory:
        try:
            consistent = run(args, Path(directory))
        except ValueError as e:
            sys.exit(f"Erreur : {e}")
    print("RÉSULTAT :", "COHÉRENT" if consistent else "INCOHÉRENCES DÉTECTÉES")
    sys.exit(0 if consistent else 1)


if __name__ == "__main__":
    main()
