"""NEXUS-V — vérification de la traçabilité en ligne de commande, avec les mêmes classes que le serveur
(verification.py) : exigences (Excel) -> plan d'acceptation (Word) -> STDHL (Word) -> fiches d'essai (XML).

Deux usages : sans fichier, vérifier la base de NEXUS-V (une copie : le serveur peut tourner) ; avec des fichiers,
les vérifier entre eux, seuls les référentiels et les réglages de la base étant repris. Code de sortie 0 quand tout
est cohérent, 1 sinon. L'aide de la ligne de commande (HELP) donne l'usage.
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

# Aide de la ligne de commande, affichée à l'utilisateur.
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
BASE_DIR = ROOT.parent / BASE_FOLDER   # fichiers-references/base/ à la racine de NEXUS-V
TOME = "exigences"   # tome unique des fichiers donnés en ligne de commande
SETTINGS = ("release", "platforms", "iadt", "mocs")   # référentiels repris du dossier des fichiers


class Workspace:
    """Les classes de NEXUS-V sur un dossier temporaire : ce qu'elles écrivent n'atteint jamais le dossier des
    fichiers."""

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
        """Libère la base temporaire."""
        self.sql.close()


def storage_folder(given):
    """Dossier des fichiers de NEXUS-V : --base, sinon fichiers-references/base de NEXUS-V."""
    return given or BASE_DIR


def copy_base(folder, target):
    """Copie la base et les réglages d'un dossier des fichiers, pas les fichiers enregistrés."""
    if (folder / CONFIG_NAME).is_file():
        shutil.copy(folder / CONFIG_NAME, target / CONFIG_NAME)
    if (folder / SCHEMA_NAME).is_file():
        shutil.copy(folder / SCHEMA_NAME, target / SCHEMA_NAME)
    if (folder / DATABASE_NAME).is_file():
        # Par SQLite, pas comme un fichier : la copie est cohérente même si le serveur écrit.
        # Ouverte en lecture seule : la ligne de commande n'écrit jamais dans la vraie base.
        source = sqlite3.connect(f"{(folder / DATABASE_NAME).resolve().as_uri()}?mode=ro", uri=True)
        copy = sqlite3.connect(target / DATABASE_NAME)
        source.backup(copy)
        copy.close()
        source.close()


def load_files(work, args, settings):
    """Charge les fichiers de la ligne de commande dans un espace de travail vide."""
    work.config.update(settings)
    content = args.requirements.read_bytes()
    found = [work.config.get("release")]   # une release pour la base
    work.requirements.import_file(work.files, content, args.requirements.name, TOME, sheet=args.sheet, config=work.config)
    for path in args.plan:
        # Personne à qui demander : les lignes qui ne respectent pas les règles IADT/MOC sont chargées telles quelles
        # (plusieurs délégations restent une anomalie).
        work.plans.load_file(work.files, work.requirements, path.read_bytes(), path.name, TOME, found, config=work.config,
                             decisions={"accept_all": True})
    for path in args.stdhl:
        work.stdhl.load_file(work.files, work.requirements, path.read_bytes(), path.name, "", found, config=work.config)
    # Les fiches XML de la ligne de commande viennent de la plateforme de niveau système (configuration
    # system_platform).
    work.fiches.load_file(work.files, work.requirements, args.xml.read_bytes(), args.xml.name, "", found,
                          config=work.config, platform=work.config.get("system_platform"))


def print_report(result):
    """Affiche les exigences avec leur statut à chaque étape, puis la synthèse et les anomalies."""
    rows, names, labels = result["rows"], result["names"], result["status_labels"]
    # Largeurs des colonnes du tableau texte, d'après les valeurs les plus longues (bornées).
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
    # Une ligne par étape, d'après les tuiles de la page (OK sur les exigences attendues à cette étape) et les exigences
    # perdues à cette étape.
    for k, (name, tile) in enumerate(zip(names, result["tiles"][1:])):
        if tile.get("waiting"):
            print(f"  {name} : à charger")
            continue
        faulty = [r["id"] for r in rows if r["lost_at"] == name]
        print(f"  {name} : {tile['value']}/{tile['expected']} présentes et correctes"
              + (f" — en défaut : {', '.join(faulty)}" if faulty else ""))


def arguments():
    """Les options de la ligne de commande, vérifiées."""
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
    # Avec des fichiers, toute la chaîne est nécessaire ; sans fichier, c'est la base qui est vérifiée.
    if any(given) and not (args.xml and args.plan and args.stdhl):
        parser.error("avec des fichiers, il faut les Exigences, les Fiches d'Essai, --plan et --stdhl")
    if (args.regex or args.sheet) and not any(given):
        parser.error("--regex et --feuille s'utilisent avec des fichiers")
    for path in filter(None, given):
        if not path.is_file():
            parser.error(f"fichier introuvable : {path}")
    return args


def run(args, directory):
    """Lance la vérification dans un dossier temporaire. Renvoie True quand tout est cohérent."""
    folder = storage_folder(args.folder)
    copy_base(folder, directory)
    work = Workspace(directory)
    try:
        if args.requirements:
            # Avec des fichiers : les référentiels et réglages de la base sont gardés, puis la base copiée est remplacée
            # par une base vide où seuls les fichiers donnés sont chargés.
            settings = {name: work.config.get(name) for name in SETTINGS}
            if args.release:
                settings["release"] = args.release.upper()
            if args.regex:
                settings["requirement_regex"] = args.regex
            work.close()
            (directory / DATABASE_NAME).unlink(missing_ok=True)
            work = Workspace(directory)
            load_files(work, args, settings)
        releases = []   # une release pour la base : tout est vérifié
        result = work.verification.run(releases)
        if "rows" not in result:
            sys.exit(result["message"] or f"Aucune exigence dans la base de {folder}.")
        print_report(result)
        if args.report:
            args.report.write_bytes(work.verification.report(releases)[1])
            print(f"\nRapport écrit : {args.report}")
        print(f"\n{result['verdict']['text']}")
        # Cohérent : aucune exigence perdue et aucune étape sans fichier.
        return not result["defects"] and not result["waiting"]
    finally:
        work.close()


def main():
    """Lit les options, lance la vérification et donne son résultat comme code de sortie."""
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
