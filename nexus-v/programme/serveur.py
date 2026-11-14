"""NEXUS-V — serveur local de la page.

Sert la page et son API JSON. Lancé par nexus-v.py, à la racine de NEXUS-V ; ce module est dans le sous-dossier
programme. Les fichiers propres à NEXUS-V sont dans le dossier fichiers-references/base, à la racine ; son nom est
fixe :
  config.json         la configuration (voir configuration.py) ;
  nexus-v-schema.sql  la définition des tables et la version de la base ;
  nexus-v.db          la base SQLite : fichiers chargés, exigences, contenu des documents et référentiels.
Chaque fichier importé est gardé dans le dossier des fichiers importés, general.storage_dir de config.json (par
défaut « . » : le dossier qui contient base, c'est-à-dire fichiers-references), dans un sous-dossier par type, sous
le nom « <n°>_<nom du fichier> ». Ce dossier peut être changé dans la page.
L'aide de la ligne de commande (HELP) donne l'usage.
"""
import argparse
import json
import re
import shutil
import sys
import time
import webbrowser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from nexus_db import random_platform_color

try:
    from configuration import FOLDER_VARIABLES, Configuration, storage_path
    from fiches_db import FichesDatabase
    from file_check import FileCheck
    from file_database import FileDatabase
    from nexus_db import BASE_FOLDER, CONFIG_NAME, DATABASE_NAME, NexusDatabase
    from plans_db import PlansDatabase
    from requirements_db import RequirementsDatabase
    from stdhl_db import StdhlDatabase
    from traceability_db import TraceabilityDatabase
    from verification import Verification
except ModuleNotFoundError as missing:
    # En général, le serveur a été lancé avec un Python qui n'est pas celui du projet (.venv).
    sys.exit(f"Module Python manquant : {missing.name}\n"
             f"Python utilisé : {sys.executable}\n"
             f"Lancer NEXUS-V par nexus-v.py, qui utilise le Python du projet (.venv) s'il le faut,\n"
             f"ou installer les modules nécessaires : py -m pip install -r programme/requirements.txt")

# Aide de la ligne de commande, affichée à l'utilisateur.
HELP = """NEXUS-V — Le fil d’Ariane de la traçabilité, de la spécification à l'essai.

Serveur local de l'interface. Se lance par nexus-v.py, à la racine de NEXUS-V.

Usage (depuis la racine de NEXUS-V) :
    py nexus-v.py            puis ouvrir http://localhost:8000
    py nexus-v.py --port 8080 --sans-navigateur
    py nexus-v.py --base ../essai/base --port 8001     (base d'essai séparée)
"""

ROOT = Path(__file__).resolve().parent   # programme/ : le code
APP_DIR = ROOT.parent                    # dossier de nexus-v.py, le fichier que lance l'utilisateur
INTERFACE_DIR = ROOT / "interface"
INDEX_PAGE = "Verification_coherence.html"
BASE_DIR = APP_DIR / BASE_FOLDER   # config.json, nexus-v-schema.sql, nexus-v.db : fixes, non choisis par l'utilisateur
PIECE = 65536          # taille des morceaux d'une réponse, en octets

base_dir = None        # dossier des fichiers de NEXUS-V (--base pour les tests)
files_dir = None       # dossier des fichiers
database = None        # FileDatabase du dossier des fichiers
configuration = None   # Configuration du dossier des fichiers
requirements = None    # RequirementsDatabase du dossier des fichiers
plans = None           # PlansDatabase du dossier des fichiers
stdhl = None           # StdhlDatabase du dossier des fichiers
fiches = None          # FichesDatabase du dossier des fichiers
traceability = None    # TraceabilityDatabase du dossier des fichiers
verification = None    # Verification du dossier des fichiers
file_check = None      # FileCheck : fichiers vérifiés avant leur import
sql = None             # NexusDatabase du dossier des fichiers (fichier nexus-v.db)


# --------------------------------------------------------------------------- #
# Dossier des fichiers
# --------------------------------------------------------------------------- #

def open_storage(base):
    """Utilise `base`, le dossier des fichiers de NEXUS-V (config.json, nexus-v-schema.sql, nexus-v.db) ; les
    fichiers importés sont dans le dossier general.storage_dir de la configuration, un sous-dossier pour chaque type."""
    # Les objets ci-dessous sont des variables globales du module partagées par chaque fil de requête ; ils sont tous
    # reconstruits ici, aussi quand le dossier des fichiers change.
    global base_dir, files_dir, database, configuration, requirements, plans, stdhl, fiches, traceability, verification, sql
    global file_check
    base_dir = Path(base).resolve()
    # Une base précédente (dossier des fichiers changé) est fermée avant l'ouverture de la nouvelle.
    if sql is not None:
        sql.close()
    sql = NexusDatabase(base_dir / DATABASE_NAME)
    # L'ordre compte : les référentiels, puis les fichiers, puis ce qui y renvoie.
    configuration = Configuration(base_dir / CONFIG_NAME, database=sql)
    files_dir = storage_path(configuration, base_dir.parent)
    files_dir.mkdir(parents=True, exist_ok=True)
    database = FileDatabase(files_dir, sql, configuration)
    requirements = RequirementsDatabase(sql)
    plans = PlansDatabase(sql)
    stdhl = StdhlDatabase(sql)
    fiches = FichesDatabase(sql)
    traceability = TraceabilityDatabase(sql)
    verification = Verification(configuration, requirements, plans, stdhl, fiches, traceability)
    file_check = FileCheck(configuration, requirements, plans, stdhl, fiches, traceability, sql)
    # Les documents dont le contenu manque dans la base (nouvelle base, migration, tables vidées) sont relus depuis leur
    # copie enregistrée : la base peut toujours être reconstruite à partir des fichiers.
    for base in (plans, stdhl, fiches, traceability):
        for name in base.read_missing(database, configuration):
            print(f"Base : contenu de « {name} » lu et enregistré")
    # Les exigences et les plans dont le fichier n'est plus dans la base (supprimés avant que le lien soit gardé) sont
    # abandonnés.
    known = {r["id"] for r in database.list()}
    for tome, n in requirements.prune_files(lambda f: f in known).items():
        print(f"Base des exigences : {n} {'exigences' if n > 1 else 'exigence'} du tome « {tome} » sans fichier d'origine "
              f"{'supprimées' if n > 1 else 'supprimée'}")
    for name in plans.prune_files(lambda f: f in known):
        print(f"Base des plans : contenu de « {name} » sans fichier d'origine supprimé")


def storage_info():
    """Chemins montrés par la page (dossier des fichiers, base, configuration)."""
    return {"files_dir": str(files_dir), "database": str(database.path), "config": str(configuration.path),
            "requirements": str(requirements.path), "plans": str(plans.path), "sql": str(sql.path),
            "base": str(base_dir)}


def check_storage(path):
    """Le dossier en chemin résolu, créé si besoin et accessible en écriture. ValueError sinon."""
    path = (path or "").strip()
    if not path or not Path(path).is_absolute():
        raise ValueError("indiquer un chemin complet, par exemple D:\\NEXUS-V\\donnees")
    folder = Path(path).resolve()
    if folder.exists() and not folder.is_dir():
        raise ValueError("ce chemin désigne un fichier, pas un dossier")
    try:
        folder.mkdir(parents=True, exist_ok=True)
        # Écrire puis retirer un petit fichier est le seul moyen sûr de savoir que le dossier accepte l'écriture.
        probe = folder / ".nexus-v-test"
        probe.write_bytes(b"")
        probe.unlink()
    except OSError as e:
        raise ValueError(f"dossier inaccessible en écriture : {e.strerror or e}") from None
    return folder


def change_storage(path, copy_files):
    """Utilise un autre dossier pour les fichiers importés, en y copiant éventuellement leurs sous-dossiers ; il est
    gardé dans config.json (general.storage_dir)."""
    target = check_storage(path)
    if target == files_dir:
        raise ValueError("c'est déjà le dossier des fichiers importés")
    folders = [configuration.get(name) for name in FOLDER_VARIABLES]
    if copy_files:
        # La copie n'écrase jamais : un dossier qui contient déjà des sous-dossiers de NEXUS-V est refusé.
        if any((target / name).exists() for name in folders):
            raise ValueError("ce dossier contient déjà des fichiers NEXUS-V : copie refusée pour ne rien écraser")
        for name in folders:
            if (files_dir / name).is_dir():
                shutil.copytree(files_dir / name, target / name)
    # Le nouveau dossier est enregistré dans config.json, puis chaque objet est reconstruit dessus.
    configuration.update({"storage_dir": str(target)})
    open_storage(base_dir)


def choose_directory(initial):
    """Ouvre la fenêtre native de choix de dossier sur ce PC. Renvoie le dossier choisi, ou "" en cas d'annulation."""
    import tkinter
    from tkinter import filedialog
    # Une fenêtre Tk invisible, nécessaire seulement comme parent de la fenêtre native.
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)   # sinon la fenêtre peut s'ouvrir derrière le navigateur
    try:
        chosen = filedialog.askdirectory(parent=root, initialdir=initial, mustexist=False,
                                         title="NEXUS-V — dossier de stockage des données")
    finally:
        root.destroy()
    return str(Path(chosen)) if chosen else ""


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #

class Handler(SimpleHTTPRequestHandler):
    """Fichiers statiques de interface/ plus une petite API JSON sous /api/."""

    def __init__(self, *args, **kwargs):
        """Les fichiers statiques sont servis depuis le dossier interface."""
        super().__init__(*args, directory=str(INTERFACE_DIR), **kwargs)

    def end_headers(self):
        """Ajoute l'en-tête de cache à chaque réponse."""
        # Le navigateur doit vérifier à chaque fois qu'il a la version actuelle de la page (et des réponses de l'API) :
        # sinon une page ouverte avant une mise à jour continue d'exécuter l'ancien code.
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def copyfile(self, source, outputfile):
        """Les fichiers statiques aussi par morceaux, pour la même raison que dans send_body."""
        shutil.copyfileobj(source, outputfile, PIECE)

    def send_body(self, body, content_type, status=HTTPStatus.OK, file_name=None):
        """Envoie une réponse avec son type et sa longueur ; avec un nom de fichier, le navigateur l'enregistre."""
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        if file_name:
            self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        # Par morceaux : une seule grande écriture s'est parfois bloquée 19 s, puis a été coupée à 130560 octets.
        for start in range(0, len(body), PIECE):
            self.wfile.write(body[start:start + PIECE])

    def send_json(self, data, status=HTTPStatus.OK):
        """Envoie une valeur en JSON."""
        self.send_body(json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def send_text_error(self, status, message):
        """Envoie un message d'erreur, affiché tel quel par la page."""
        self.send_body(message.encode("utf-8"), "text/plain; charset=utf-8", status)

    def read_body(self):
        """Octets envoyés avec la requête."""
        return self.rfile.read(int(self.headers.get("Content-Length") or 0))

    def file_id(self, suffix=""):
        """Numéro de fichier d'une adresse /api/files/<n°><suffixe>, ou None."""
        m = re.fullmatch(rf"/api/files/(\d+){suffix}", urlparse(self.path).path)
        return int(m[1]) if m else None

    def do_GET(self):
        """Routes de lecture de l'API ; toute autre adresse est un fichier statique de l'interface."""
        path = urlparse(self.path).path
        # L'adresse racine sert la page ; les autres adresses sont des routes de l'API, ou des fichiers statiques de
        # l'interface.
        if path == "/":
            self.path = "/" + INDEX_PAGE
        elif path == "/api/files":
            return self.send_json(database.list())
        elif path == "/api/config":
            return self.send_json(configuration.to_dict())
        elif path == "/api/platform-color":   # couleur tirée pour une nouvelle plateforme, loin de taken
            # taken : les couleurs déjà utilisées par le brouillon de la page de configuration, pour que la nouvelle
            # couleur s'en distingue.
            taken = parse_qs(urlparse(self.path).query).get("taken", [""])[0].split(",")
            return self.send_json({"color": random_platform_color(taken)})
        elif path == "/api/requirements":   # avec la part de chaque tome et ses délégations
            colors = configuration.get("colors")
            return self.send_json({**requirements.to_dict(),
                                   "distribution": requirements.distribution(colors.get("tomes")),
                                   "delegation_shares": requirements.delegation_shares(colors.get("platforms"))})
        elif path == "/api/plans":
            # Lignes vérifiées par rapport à la configuration actuelle (plateformes, IADT, MOC, délégation).
            return self.send_json(plans.to_dict(configuration))
        elif path == "/api/stdhl":
            return self.send_json(stdhl.to_dict())
        elif path == "/api/fiches":
            return self.send_json(fiches.to_dict())
        elif path == "/api/traceability":
            return self.send_json(traceability.to_dict())
        elif path in ("/api/verification", "/api/verification/report"):
            selected = []   # une release pour la base : tout est vérifié
            if path == "/api/verification":
                return self.send_json(verification.run(selected))
            try:
                name, content = verification.report(selected)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_body(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  file_name=name)
        elif path == "/api/backup":
            # Toute la base en SQL, enregistrée par le navigateur.
            return self.send_body(sql.dump().encode("utf-8"), "application/sql; charset=utf-8",
                                  file_name=time.strftime("nexus-v_%Y-%m-%d_%H-%M.sql"))
        elif path == "/api/tables":
            return self.send_json(sql.table_list())
        elif m := re.fullmatch(r"/api/tables/(\w+)", path):
            try:
                return self.send_json(sql.table(m[1]))
            except KeyError:
                return self.send_text_error(HTTPStatus.NOT_FOUND, "table inconnue")
        elif path == "/api/plans/check":
            query = {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}
            selected = []   # une release pour la base
            return self.send_json(plans.check_tome(query.get("tome", ""), requirements, configuration, selected))
        elif path == "/api/info":
            return self.send_json(storage_info())
        # Octets bruts d'un fichier enregistré (la page y relit un fichier d'exigences).
        elif (file_id := self.file_id("/content")) is not None:
            content = database.content(file_id)
            if content is None:
                return self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
            return self.send_body(content, "application/octet-stream")
        super().do_GET()

    def do_POST(self):
        """Routes qui chargent, importent ou modifient des données."""
        url = urlparse(self.path)
        if url.path == "/api/storage":
            try:
                data = json.loads(self.read_body() or b"{}")
                # {path, copy} : le nouveau dossier des fichiers et si les fichiers y sont copiés.
                change_storage(data.get("path"), bool(data.get("copy")))
            except ValueError as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(storage_info())
        if m := re.fullmatch(r"/api/tables/(\w+)", url.path):
            return self.change_table(lambda values: sql.insert_row(m[1], values))
        if url.path == "/api/files/delete":
            # Plusieurs fichiers à la fois : {ids: [numéros de fichier]} ; les numéros déjà partis sont laissés de côté.
            try:
                ids = [int(i) for i in json.loads(self.read_body() or b"{}").get("ids", [])]
            except (ValueError, TypeError) as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            # Une réponse par fichier retiré, dans l'ordre demandé.
            return self.send_json([self.remove_file(i) for i in ids if database.get(i)])
        if url.path == "/api/plans/edit":
            # Modification à la main de la ligne d'une exigence dans un plan : {file_id, action:
            # add|update|delete|restore, requirement, values}.
            try:
                data = json.loads(self.read_body() or b"{}")
                row = plans.edit(int(data.get("file_id") or 0), data.get("action"), data.get("requirement"),
                                 data.get("values"), configuration)
            except ValueError as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json({"row": row})
        if url.path == "/api/stdhl/edit":
            # Ligne d'un STDHL supprimée ou rétablie par l'utilisateur : {file_id, action: delete|restore, requirement,
            # test}.
            try:
                data = json.loads(self.read_body() or b"{}")
                return self.send_json(stdhl.edit(int(data.get("file_id") or 0), data.get("action"),
                                                 data.get("requirement"), data.get("test")))
            except (ValueError, TypeError) as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        if url.path == "/api/tests/rename":
            # Nom d'un test changé dans la page STDHL : {old, new} ; chaque STDHL et chaque fiche qui le cite suit.
            try:
                data = json.loads(self.read_body() or b"{}")
                return self.send_json({"name": sql.rename_test(data.get("old"), data.get("new"))})
            except ValueError as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        if url.path == "/api/requirements/edit":
            # Modification à la main d'une exigence : {tome, action: add|update|delete, id, requirement}.
            try:
                data = json.loads(self.read_body() or b"{}")
                req = requirements.edit(str(data.get("tome") or ""), data.get("action"), data.get("id"),
                                        data.get("requirement"), configuration)
                if data.get("action") in ("add", "update"):   # un numéro cité peut être devenu connu
                    plans.read_missing(database, configuration, again=True)
                    stdhl.read_missing(database, configuration, again=True)
                    fiches.read_missing(database, configuration, again=True)
                    traceability.read_missing(database, configuration, again=True)
            except ValueError as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json({"requirement": req})
        # Rien ne peut être chargé avant que la release de la base soit fixée : chaque exigence et chaque fichier lui
        # appartient.
        if url.path.startswith("/api/") and url.path.endswith(("/analyse", "/import", "/load")) \
                and not configuration.get("release"):
            return self.send_text_error(HTTPStatus.BAD_REQUEST, "définir d'abord la release dans la configuration "
                                                                "(Paramètres généraux)")
        if url.path in ("/api/requirements/analyse", "/api/requirements/import"):
            # Fichier d'exigences envoyé tel quel (corps) ; sa lecture et l'import sont faits par RequirementsDatabase.
            query = {k: v[0] for k, v in parse_qs(url.query).items()}
            # Seules les exigences de la release de la base sont importées.
            selected = [configuration.get("release")]
            try:
                content = self.read_body()
                # Les choix de l'utilisateur, en JSON, viennent avant le fichier.
                # Le corps est le JSON des choix (choices_size octets : lignes, versions gardées, valeurs STDHL), suivi
                # des octets du fichier Excel.
                size = int(query.get("choices_size") or 0)
                choices = json.loads(content[:size]) if size else None
                content = content[size:]
                if url.path.endswith("/analyse"):
                    # Avec le nom du fichier : le tome qu'il nomme, proposé dans la fenêtre d'import.
                    return self.send_json({**requirements.analyse(content, query.get("sheet") or None,
                                                                  (query.get("tome") or "").strip() or None,
                                                                  query.get("mode", "replace"),
                                                                  configuration, selected, choices),
                                           "tome_name": requirements.tome_from_name(query.get("name"), configuration)})
                # Import : le fichier est enregistré et ses exigences écrites, les conflits étant réglés par les choix.
                if not query.get("name") or not (query.get("tome") or "").strip():
                    raise ValueError("paramètres name / tome manquants")
                summary = requirements.import_file(database, content, query["name"], query["tome"].strip(),
                                                   query.get("mode", "replace"), query.get("sheet") or None,
                                                   configuration, selected, choices)
                # Les plans, STDHL et fiches ne gardent que des exigences de la base : relus.
                plans.read_missing(database, configuration, again=True)
                stdhl.read_missing(database, configuration, again=True)
                fiches.read_missing(database, configuration, again=True)
                traceability.read_missing(database, configuration, again=True)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(summary)
        if m := re.fullmatch(r"/api/(plans|stdhl|fiches|traceability)/(analyse|load)", url.path):
            # Fichier de plan, STDHL, fiches ou traçabilité envoyé tel quel (corps) : la classe du sujet le lit et
            # l'enregistre.
            # Les deux mêmes temps pour chaque document : /analyse lit le fichier et renvoie ce que l'utilisateur doit
            # régler, rien n'est enregistré ; /load enregistre le fichier et son contenu avec les décisions de
            # l'utilisateur.
            base = {"plans": plans, "stdhl": stdhl, "fiches": fiches, "traceability": traceability}[m[1]]
            query = {k: v[0] for k, v in parse_qs(url.query).items()}
            try:
                content = self.read_body()
                if not query.get("name"):
                    raise ValueError("paramètre name manquant")
                if m[2] == "analyse":
                    if base is plans:   # avec tome et releases : inconnues et exigences d'un autre plan
                        return self.send_json(plans.analyse(content, configuration, (query.get("tome") or "").strip(),
                                                            [configuration.get("release")], query["name"]))
                    if base is traceability:   # avec le tome et la release : les numéros à ignorer
                        return self.send_json(base.analyse(content, configuration, (query.get("tome") or "").strip(),
                                                           [configuration.get("release")]))
                    if base is fiches:   # avec le tome (aucun : tous les tomes) : les numéros à ignorer
                        return self.send_json(fiches.analyse(content, configuration, [configuration.get("release")],
                                                             (query.get("tome") or "").strip()))
                    args = (content, query["name"], configuration) if base is stdhl else (content, configuration)
                    return self.send_json(base.analyse(*args))
                # Chargement : le tome choisi (aucun pour un fichier STDHL ou de fiches qui s'applique à tous les
                # tomes), les décisions sur les lignes à régler, et pour un fichier de fiches la plateforme d'où il
                # vient.
                extra = {}
                decisions = json.loads(query["decisions"]) if query.get("decisions") else None
                if base in (plans, traceability, fiches):   # décisions sur les lignes à régler, en JSON
                    extra = {**extra, "decisions": decisions}
                if base in (traceability, fiches):   # la plateforme d'où vient le fichier de fiches
                    extra["platform"] = (query.get("platform") or "").strip()
                result = base.load_file(database, requirements, content, query["name"], (query.get("tome") or "").strip(),
                                        [configuration.get("release")], configuration, **extra)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(result)
        if m := re.fullmatch(r"/api/check/(requirements|plans|stdhl|fiches)", url.path):
            # Fichiers d'une étape vérifiés avant leur import, rien n'est enregistré : le corps est la liste JSON des
            # fichiers [{name, size}] (manifest_size octets), puis les octets de chaque fichier dans cet ordre. La
            # réponse porte le rapport HTML écrit par FileCheck.
            if not configuration.get("release"):
                return self.send_text_error(HTTPStatus.BAD_REQUEST, "définir d'abord la release dans la configuration "
                                                                    "(Paramètres généraux)")
            query = {k: v[0] for k, v in parse_qs(url.query).items()}
            try:
                body = self.read_body()
                size = int(query.get("manifest_size") or 0)
                manifest = json.loads(body[:size] or b"[]")
                files, at = [], size
                for f in manifest:
                    # Un fichier de fiches vient avec la plateforme choisie pour lui.
                    files.append((str(f["name"]), body[at:at + int(f["size"])])
                                 + ((str(f.get("platform") or ""),) if m[1] == "fiches" else ()))
                    at += int(f["size"])
                if not files:
                    raise ValueError("aucun fichier à vérifier")
                # Chaque étape vérifiée pour les deux imports : remplacement et complément.
                check = {"requirements": file_check.check_requirements, "plans": file_check.check_plans,
                         "stdhl": file_check.check_stdhl, "fiches": file_check.check_fiches}[m[1]]
                return self.send_json(check(files))
            except (ValueError, KeyError, TypeError) as e:   # couvre aussi un JSON invalide
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        if url.path == "/api/storage/choose":
            try:
                return self.send_json({"path": choose_directory(str(files_dir))})
            except ImportError:
                return self.send_text_error(HTTPStatus.NOT_IMPLEMENTED,
                                            "sélecteur de dossier indisponible (tkinter absent) : saisir le chemin")
        self.send_text_error(HTTPStatus.NOT_FOUND, "adresse inconnue")

    def change_table(self, change):
        """Applique la modification d'une ligne de table (page Tables de la base), d'après les valeurs envoyées en
        JSON."""
        try:
            change(json.loads(self.read_body() or b"{}"))
        except KeyError:
            return self.send_text_error(HTTPStatus.NOT_FOUND, "table inconnue")
        except ValueError as e:   # couvre aussi un JSON invalide
            return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        self.send_json({"ok": True})

    def remove_file(self, file_id):
        """Retire un fichier et ce qui en vient. Renvoie ce qui a été retiré."""
        # Les exigences importées de ce fichier, ou le contenu de ce plan, quittent la base avec lui (la cascade SQL les
        # retirerait aussi ; ils sont retirés d'abord pour dire ce qui a été retiré).
        # requirements.remove_file et plans.remove renvoient ce qu'ils retirent, pour le message de la page.
        removed = {"deleted": file_id, "requirements_removed": requirements.remove_file(file_id),
                   "plan_removed": plans.remove(file_id)}
        database.remove(file_id)
        return removed

    def do_PUT(self):
        """Enregistre la configuration ou une ligne d'une table."""
        if m := re.fullmatch(r"/api/tables/(\w+)/(-?\d+)", urlparse(self.path).path):
            # PUT /api/tables/<table>/<id> : une ligne modifiée dans la page Tables de la base.
            return self.change_table(lambda values: sql.update_row(m[1], int(m[2]), values))
        if urlparse(self.path).path != "/api/config":
            return self.send_text_error(HTTPStatus.NOT_FOUND, "adresse inconnue")
        try:
            # Le contenu des STDHL dépend de ces réglages : quand ils changent, les fichiers enregistrés sont relus.
            # Les valeurs avant la mise à jour, pour savoir ensuite ce qui a changé.
            folders = {name: configuration.get(name) for name in FOLDER_VARIABLES}
            settings = ("requirement_regex", "fiche_regex", "stdhl_name_regex", "headers_stdhl", "release",
                        "stdhl2_test_regex", "stdhl2_requirements_label", "stdhl2_release_label",
                        "stdhl2_release_regex", "stdhl2_test_case_regex")
            tags = ("requirement_regex", "test_tag", "test_case_tag", "requirement_tag")
            reading = ("requirement_regex", "code_separator_regex", "headers_plan")   # comment les plans sont lus
            before = {name: configuration.get(name) for name in settings + tags + reading}
            result = configuration.update(json.loads(self.read_body() or b"{}"))
            # Seuls les documents dont les réglages de lecture ont changé sont relus.
            changed = {name for name, value in before.items() if configuration.get(name) != value}
            if changed & set(settings):
                stdhl.read_missing(database, configuration, again=True)
            if changed & set(tags):
                fiches.read_missing(database, configuration, again=True)
            if changed & set(reading):
                plans.read_missing(database, configuration, again=True)
            # Un sous-dossier renommé garde ses fichiers.
            for name, old in folders.items():
                new = configuration.get(name)
                if new != old and (files_dir / old).is_dir() and not (files_dir / new).exists():
                    (files_dir / old).rename(files_dir / new)
            self.send_json(result)
        except ValueError as e:   # couvre aussi un JSON invalide
            self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))

    def do_DELETE(self):
        """Retire une ligne d'une table ou un fichier."""
        if m := re.fullmatch(r"/api/tables/(\w+)/(-?\d+)", urlparse(self.path).path):
            # Une ligne de fichier est retirée comme depuis la page Gestion des fichiers : sa copie et son contenu
            # partent avec elle.
            if m[1] == "files":   # avec sa copie enregistrée
                return self.send_json(self.remove_file(int(m[2]))) if database.get(int(m[2])) else \
                    self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
            return self.change_table(lambda values: sql.delete_row(m[1], int(m[2])))
        file_id = self.file_id()
        if file_id is None or database.get(file_id) is None:
            return self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
        self.send_json(self.remove_file(file_id))


class Server(ThreadingHTTPServer):
    """Un fil par requête : une connexion de navigateur laissée ouverte ne doit pas bloquer les autres ; la base a
    son verrou."""
    request_queue_size = 64
    # Sous Windows, la réutilisation d'adresse laisserait deux serveurs partager un port, et écrire tour à tour dans la
    # même base.
    allow_reuse_address = sys.platform != "win32"


def main():
    """Lit les options, ouvre le dossier des fichiers et lance le serveur."""
    parser = argparse.ArgumentParser(description=HELP, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--base", type=Path, default=BASE_DIR,
                        help=f"Dossier de config.json, nexus-v-schema.sql et nexus-v.db, pour les essais (défaut : {BASE_DIR})")
    parser.add_argument("--port", type=int, help="Port d'écoute (défaut : variable server_port de la configuration)")
    parser.add_argument("--sans-navigateur", dest="no_browser", action="store_true",
                        help="Ne pas ouvrir le navigateur au démarrage (défaut : variable open_browser)")
    args = parser.parse_args()

    try:
        open_storage(args.base)
    except ValueError as e:
        sys.exit(f"Base inutilisable : {e}")

    # Le port donné en ligne de commande l'emporte sur la configuration.
    port = args.port or configuration.get("server_port")
    # Usage local seulement : écoute sur l'interface de bouclage, pas sur le réseau.
    try:
        server = Server(("127.0.0.1", port), Handler)
    except OSError:
        sys.exit(f"Le port {port} est déjà utilisé : un serveur NEXUS-V tourne sans doute déjà. Fermer sa fenêtre, "
                 "ou choisir un autre port avec --port.")
    url = f"http://localhost:{port}/"
    print(f"NEXUS-V — {configuration.get('slogan')}\n"
          f"Interface : {url}\nBase : {base_dir}\nFichiers importés : {files_dir}\nCtrl+C pour arrêter.", flush=True)
    if configuration.get("open_browser") and not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nArrêt du serveur.")


if __name__ == "__main__":
    main()
