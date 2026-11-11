"""NEXUS-V — local server of the page.

Serves the page and its JSON API. Started by nexus-v.py, at the root of NEXUS-V; this module is in the sub-folder
programme. The files of NEXUS-V itself are in the folder fichiers-references/base, at the root; its name is fixed:
  config.json         the configuration (see configuration.py);
  nexus-v-schema.sql  the definition of the tables and the version of the base;
  nexus-v.db          the SQLite base: files loaded, requirements, content of the documents and referentials.
Each imported file is kept in the folder of imported files, general.storage_dir of config.json (by default « . »:
the folder that holds base, that is fichiers-references), in one sub-folder per type, named « <n°>_<name of the file> ». That folder can be changed in the page.
The help of the command line (HELP, in French) gives the usage.
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
    # Usually the server was started with a Python that is not the project's one (.venv).
    sys.exit(f"Module Python manquant : {missing.name}\n"
             f"Python utilisé : {sys.executable}\n"
             f"Lancer NEXUS-V par nexus-v.py, qui utilise le Python du projet (.venv) s'il le faut,\n"
             f"ou installer les modules nécessaires : py -m pip install -r programme/requirements.txt")

# Help of the command line, shown to the user (in French, like the interface).
HELP = """NEXUS-V — Le fil d’Ariane de la traçabilité, de la spécification à l'essai.

Serveur local de l'interface. Se lance par nexus-v.py, à la racine de NEXUS-V.

Usage (depuis la racine de NEXUS-V) :
    py nexus-v.py            puis ouvrir http://localhost:8000
    py nexus-v.py --port 8080 --sans-navigateur
    py nexus-v.py --base ../essai/base --port 8001     (base d'essai séparée)
"""

ROOT = Path(__file__).resolve().parent   # programme/: the code
APP_DIR = ROOT.parent                    # folder of nexus-v.py, the file the user starts
INTERFACE_DIR = ROOT / "interface"
INDEX_PAGE = "Verification_coherence.html"
BASE_DIR = APP_DIR / BASE_FOLDER   # config.json, nexus-v-schema.sql, nexus-v.db: fixed, not chosen by the user
PIECE = 65536          # size of the pieces of an answer, in bytes

base_dir = None        # folder of the files of NEXUS-V (--base for tests)
files_dir = None       # storage directory
database = None        # FileDatabase of the storage directory
configuration = None   # Configuration of the storage directory
requirements = None    # RequirementsDatabase of the storage directory
plans = None           # PlansDatabase of the storage directory
stdhl = None           # StdhlDatabase of the storage directory
fiches = None          # FichesDatabase of the storage directory
traceability = None    # TraceabilityDatabase of the storage directory
verification = None    # Verification of the storage directory
file_check = None      # FileCheck: files checked before their import
sql = None             # NexusDatabase of the storage directory (file nexus-v.db)


# --------------------------------------------------------------------------- #
# Storage folder
# --------------------------------------------------------------------------- #

def open_storage(base):
    """Use `base`, the folder of the files of NEXUS-V (config.json, nexus-v-schema.sql, nexus-v.db); the imported files
    are in the folder general.storage_dir of the configuration, one sub-folder for each type."""
    # The objects below are module globals shared by every request thread; they are all rebuilt here, also when the
    # storage folder changes.
    global base_dir, files_dir, database, configuration, requirements, plans, stdhl, fiches, traceability, verification, sql
    global file_check
    base_dir = Path(base).resolve()
    # A previous base (storage folder changed) is closed before the new one is opened.
    if sql is not None:
        sql.close()
    sql = NexusDatabase(base_dir / DATABASE_NAME)
    # Order matters: referentials, then files, then what refers to them.
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
    # Documents whose content is missing from the base (new base, migration, tables emptied) are read again from their
    # stored copy: the base can always be rebuilt from the files.
    for base in (plans, stdhl, fiches, traceability):
        for name in base.read_missing(database, configuration):
            print(f"Base : contenu de « {name} » lu et enregistré")
    # Requirements and plans whose file is no longer in the base (deleted before the link was kept) are dropped.
    known = {r["id"] for r in database.list()}
    for tome, n in requirements.prune_files(lambda f: f in known).items():
        print(f"Base des exigences : {n} {'exigences' if n > 1 else 'exigence'} du tome « {tome} » sans fichier d'origine "
              f"{'supprimées' if n > 1 else 'supprimée'}")
    for name in plans.prune_files(lambda f: f in known):
        print(f"Base des plans : contenu de « {name} » sans fichier d'origine supprimé")


def storage_info():
    """Paths shown by the page (storage folder, database, configuration)."""
    return {"files_dir": str(files_dir), "database": str(database.path), "config": str(configuration.path),
            "requirements": str(requirements.path), "plans": str(plans.path), "sql": str(sql.path),
            "base": str(base_dir)}


def check_storage(path):
    """The folder as a resolved path, created if needed and writable. Raises ValueError otherwise."""
    path = (path or "").strip()
    if not path or not Path(path).is_absolute():
        raise ValueError("indiquer un chemin complet, par exemple D:\\NEXUS-V\\donnees")
    folder = Path(path).resolve()
    if folder.exists() and not folder.is_dir():
        raise ValueError("ce chemin désigne un fichier, pas un dossier")
    try:
        folder.mkdir(parents=True, exist_ok=True)
        # Writing and removing a small file is the only reliable way to know that the folder accepts writing.
        probe = folder / ".nexus-v-test"
        probe.write_bytes(b"")
        probe.unlink()
    except OSError as e:
        raise ValueError(f"dossier inaccessible en écriture : {e.strerror or e}") from None
    return folder


def change_storage(path, copy_files):
    """Use another folder for the imported files, optionally copying their sub-folders into it; it is kept in
    config.json (general.storage_dir)."""
    target = check_storage(path)
    if target == files_dir:
        raise ValueError("c'est déjà le dossier des fichiers importés")
    folders = [configuration.get(name) for name in FOLDER_VARIABLES]
    if copy_files:
        # The copy never overwrites: a folder that already holds NEXUS-V sub-folders is refused.
        if any((target / name).exists() for name in folders):
            raise ValueError("ce dossier contient déjà des fichiers NEXUS-V : copie refusée pour ne rien écraser")
        for name in folders:
            if (files_dir / name).is_dir():
                shutil.copytree(files_dir / name, target / name)
    # The new folder is saved in config.json, then every object is rebuilt on it.
    configuration.update({"storage_dir": str(target)})
    open_storage(base_dir)


def choose_directory(initial):
    """Open the native folder dialog on this PC. Returns the chosen folder, or "" if cancelled."""
    import tkinter
    from tkinter import filedialog
    # An invisible Tk window, only needed as the parent of the native dialog.
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)   # otherwise the dialog may open behind the browser
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
    """Static files from interface/ plus a small JSON API under /api/."""

    def __init__(self, *args, **kwargs):
        """Static files are served from the interface folder."""
        super().__init__(*args, directory=str(INTERFACE_DIR), **kwargs)

    def end_headers(self):
        """Add the cache header to every answer."""
        # The browser must check every time that it has the current version of the page (and of the
        # API answers): otherwise a page opened before an update keeps running the old code.
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

    def copyfile(self, source, outputfile):
        """Static files by pieces too, for the same reason as in send_body."""
        shutil.copyfileobj(source, outputfile, PIECE)

    def send_body(self, body, content_type, status=HTTPStatus.OK, file_name=None):
        """Send an answer with its type and length; with a file name, the browser saves it."""
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        if file_name:
            self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        # By pieces: a single large write sometimes stalled for 19 s, then was cut off at 130560 bytes.
        for start in range(0, len(body), PIECE):
            self.wfile.write(body[start:start + PIECE])

    def send_json(self, data, status=HTTPStatus.OK):
        """Send a value as JSON."""
        self.send_body(json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def send_text_error(self, status, message):
        """Send an error message, shown as is by the page."""
        self.send_body(message.encode("utf-8"), "text/plain; charset=utf-8", status)

    def read_body(self):
        """Bytes sent with the request."""
        return self.rfile.read(int(self.headers.get("Content-Length") or 0))

    def file_id(self, suffix=""):
        """File number of a /api/files/<n°><suffix> address, or None."""
        m = re.fullmatch(rf"/api/files/(\d+){suffix}", urlparse(self.path).path)
        return int(m[1]) if m else None

    def do_GET(self):
        """Read routes of the API; any other address is a static file of the interface."""
        path = urlparse(self.path).path
        # The root address serves the page; the other addresses are routes of the API, or static files of the interface.
        if path == "/":
            self.path = "/" + INDEX_PAGE
        elif path == "/api/files":
            return self.send_json(database.list())
        elif path == "/api/config":
            return self.send_json(configuration.to_dict())
        elif path == "/api/platform-color":   # a colour drawn for a new platform, away from those listed in taken
            # taken: the colours already used by the draft of the configuration page, so that the new colour stands
            # apart.
            taken = parse_qs(urlparse(self.path).query).get("taken", [""])[0].split(",")
            return self.send_json({"color": random_platform_color(taken)})
        elif path == "/api/requirements":   # with the share of each tome (cumulative bar of the page)
            return self.send_json({**requirements.to_dict(),
                                   "distribution": requirements.distribution(configuration.get("colors").get("tomes"))})
        elif path == "/api/plans":
            # Rows checked against the current configuration (platforms, IADT, MOC, delegation).
            return self.send_json(plans.to_dict(configuration))
        elif path == "/api/stdhl":
            return self.send_json(stdhl.to_dict())
        elif path == "/api/fiches":
            return self.send_json(fiches.to_dict())
        elif path == "/api/traceability":
            return self.send_json(traceability.to_dict())
        elif path in ("/api/verification", "/api/verification/report"):
            selected = []   # one release for the base: everything is checked
            if path == "/api/verification":
                return self.send_json(verification.run(selected))
            try:
                name, content = verification.report(selected)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_body(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  file_name=name)
        elif path == "/api/backup":
            # Whole database as SQL, saved by the browser.
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
            selected = []   # one release for the base
            return self.send_json(plans.check_tome(query.get("tome", ""), requirements, configuration, selected))
        elif path == "/api/info":
            return self.send_json(storage_info())
        # Raw bytes of a stored file (the page reads a requirements file again from it).
        elif (file_id := self.file_id("/content")) is not None:
            content = database.content(file_id)
            if content is None:
                return self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
            return self.send_body(content, "application/octet-stream")
        super().do_GET()

    def do_POST(self):
        """Routes that load, import or change data."""
        url = urlparse(self.path)
        if url.path == "/api/storage":
            try:
                data = json.loads(self.read_body() or b"{}")
                # {path, copy}: the new storage folder and whether the files are copied into it.
                change_storage(data.get("path"), bool(data.get("copy")))
            except ValueError as e:   # also covers invalid JSON
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(storage_info())
        if m := re.fullmatch(r"/api/tables/(\w+)", url.path):
            return self.change_table(lambda values: sql.insert_row(m[1], values))
        if url.path == "/api/files/delete":
            # Several files at once: {ids: [file numbers]}; numbers already gone are left out.
            try:
                ids = [int(i) for i in json.loads(self.read_body() or b"{}").get("ids", [])]
            except (ValueError, TypeError) as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            # One answer per file removed, in the order asked.
            return self.send_json([self.remove_file(i) for i in ids if database.get(i)])
        if url.path == "/api/plans/edit":
            # Manual change of the row of a requirement in a plan: {file_id, action: add|update|delete|restore,
            # requirement, values}.
            try:
                data = json.loads(self.read_body() or b"{}")
                row = plans.edit(int(data.get("file_id") or 0), data.get("action"), data.get("requirement"),
                                 data.get("values"), configuration)
                # A plan row changed: a requirement may have become N/A (or stopped being), and the STDHL leave out the
                # N/A ones.
                stdhl.read_missing(database, configuration, again=True)   # requirements N/A in the plans left out
            except ValueError as e:   # also covers invalid JSON
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json({"row": row})
        if url.path == "/api/requirements/edit":
            # Manual change of one requirement: {tome, action: add|update|delete, id, requirement}.
            try:
                data = json.loads(self.read_body() or b"{}")
                req = requirements.edit(str(data.get("tome") or ""), data.get("action"), data.get("id"),
                                        data.get("requirement"), configuration)
                if data.get("action") in ("add", "update"):   # a number cited by a document may now be known
                    plans.read_missing(database, configuration, again=True)
                    stdhl.read_missing(database, configuration, again=True)
                    fiches.read_missing(database, configuration, again=True)
                    traceability.read_missing(database, configuration, again=True)
            except ValueError as e:   # also covers invalid JSON
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json({"requirement": req})
        # Nothing can be loaded before the release of the base is set: every requirement and file belongs to it.
        if url.path.startswith("/api/") and url.path.endswith(("/analyse", "/import", "/load")) \
                and not configuration.get("release"):
            return self.send_text_error(HTTPStatus.BAD_REQUEST, "définir d'abord la release dans la configuration "
                                                                "(Paramètres généraux)")
        if url.path in ("/api/requirements/analyse", "/api/requirements/import"):
            # Requirements file sent as is (body); its reading and the import are done by RequirementsDatabase.
            query = {k: v[0] for k, v in parse_qs(url.query).items()}
            # Only the requirements of the release of the base are imported.
            selected = [configuration.get("release")]
            try:
                content = self.read_body()
                # The choices of the user, in JSON, come before the file.
                # The body is the JSON of the choices (choices_size bytes: groups forced into the release, rows,
                # versions kept, STDHL values), followed by the bytes of the Excel file.
                size = int(query.get("choices_size") or 0)
                choices = json.loads(content[:size]) if size else None
                content = content[size:]
                if url.path.endswith("/analyse"):
                    # With the name of the file: the tome it names, proposed in the import window.
                    return self.send_json({**requirements.analyse(content, query.get("sheet") or None,
                                                                  (query.get("tome") or "").strip() or None,
                                                                  query.get("mode", "replace"),
                                                                  configuration, selected, choices),
                                           "tome_name": requirements.tome_from_name(query.get("name"), configuration)})
                # Import: the file is stored and its requirements written, the conflicts being settled by the choices.
                if not query.get("name") or not (query.get("tome") or "").strip():
                    raise ValueError("paramètres name / tome manquants")
                summary = requirements.import_file(database, content, query["name"], query["tome"].strip(),
                                                   query.get("mode", "replace"), query.get("sheet") or None,
                                                   configuration, selected, choices)
                # Plans, STDHL and fiches keep only requirements of the base: read again.
                plans.read_missing(database, configuration, again=True)
                stdhl.read_missing(database, configuration, again=True)
                fiches.read_missing(database, configuration, again=True)
                traceability.read_missing(database, configuration, again=True)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(summary)
        if m := re.fullmatch(r"/api/(plans|stdhl|fiches|traceability)/(analyse|load)", url.path):
            # Plan, STDHL, fiches or traceability file sent as is (body): the class of the subject reads and stores it.
            # Same two steps for every document: /analyse reads the file and returns what the user must settle, nothing
            # stored; /load stores the file and its content with the user's decisions.
            base = {"plans": plans, "stdhl": stdhl, "fiches": fiches, "traceability": traceability}[m[1]]
            query = {k: v[0] for k, v in parse_qs(url.query).items()}
            try:
                content = self.read_body()
                if not query.get("name"):
                    raise ValueError("paramètre name manquant")
                if m[2] == "analyse":
                    if base is plans:   # with the tome and releases: unknown numbers and requirements in another plan
                        return self.send_json(plans.analyse(content, configuration, (query.get("tome") or "").strip(),
                                                            [configuration.get("release")], query["name"]))
                    if base is traceability:   # with the tome and release: the numbers to ignore
                        return self.send_json(base.analyse(content, configuration, (query.get("tome") or "").strip(),
                                                           [configuration.get("release")]))
                    if base is fiches:   # with the tome (none: every tome): the numbers to ignore
                        return self.send_json(fiches.analyse(content, configuration, [configuration.get("release")],
                                                             (query.get("tome") or "").strip()))
                    args = (content, query["name"], configuration) if base is stdhl else (content, configuration)
                    return self.send_json(base.analyse(*args))
                # Load: the tome chosen (none for a STDHL or fiches file that applies to every tome), the decisions on
                # the rows to settle, and for an Excel fiches file the platform it comes from.
                extra = {}
                decisions = json.loads(query["decisions"]) if query.get("decisions") else None
                if base in (plans, traceability, fiches):   # decisions of the user on the rows to settle, in JSON
                    extra = {**extra, "decisions": decisions}
                if base is traceability:   # the platform the fiches file comes from
                    extra["platform"] = (query.get("platform") or "").strip()
                result = base.load_file(database, requirements, content, query["name"], (query.get("tome") or "").strip(),
                                        [configuration.get("release")], configuration, **extra)
                if base is plans:   # the STDHL leave out the requirements N/A in the plans: read again
                    stdhl.read_missing(database, configuration, again=True)
            except ValueError as e:
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
            return self.send_json(result)
        if m := re.fullmatch(r"/api/check/(requirements|plans)", url.path):
            # Files of one stage checked before their import, nothing stored: the body is the JSON list of the files
            # [{name, size}] (manifest_size bytes), then the bytes of each file in that order. The answer carries the
            # HTML report written by FileCheck.
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
                    files.append((str(f["name"]), body[at:at + int(f["size"])]))
                    at += int(f["size"])
                if not files:
                    raise ValueError("aucun fichier à vérifier")
                if m[1] == "requirements":   # checked for both imports: replacement and complement
                    return self.send_json(file_check.check_requirements(files))
                return self.send_json(file_check.check_plans(files))
            except (ValueError, KeyError, TypeError) as e:   # also covers invalid JSON
                return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        if url.path == "/api/storage/choose":
            try:
                return self.send_json({"path": choose_directory(str(files_dir))})
            except ImportError:
                return self.send_text_error(HTTPStatus.NOT_IMPLEMENTED,
                                            "sélecteur de dossier indisponible (tkinter absent) : saisir le chemin")
        self.send_text_error(HTTPStatus.NOT_FOUND, "adresse inconnue")

    def change_table(self, change):
        """Apply a change of a table row (page Tables de la base), from the values sent as JSON."""
        try:
            change(json.loads(self.read_body() or b"{}"))
        except KeyError:
            return self.send_text_error(HTTPStatus.NOT_FOUND, "table inconnue")
        except ValueError as e:   # also covers invalid JSON
            return self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))
        self.send_json({"ok": True})

    def remove_file(self, file_id):
        """Remove a file and what came from it. Returns what was removed."""
        # The requirements imported from this file, or the content of this plan, leave the base with it
        # (the SQL cascade would remove them too; they are removed first to tell what was removed).
        # requirements.remove_file and plans.remove return what they remove, for the message of the page.
        removed = {"deleted": file_id, "requirements_removed": requirements.remove_file(file_id),
                   "plan_removed": plans.remove(file_id)}
        database.remove(file_id)
        return removed

    def do_PUT(self):
        """Save the configuration or a row of a table."""
        if m := re.fullmatch(r"/api/tables/(\w+)/(-?\d+)", urlparse(self.path).path):
            # PUT /api/tables/<table>/<id>: a row changed in the page Tables de la base.
            return self.change_table(lambda values: sql.update_row(m[1], int(m[2]), values))
        if urlparse(self.path).path != "/api/config":
            return self.send_text_error(HTTPStatus.NOT_FOUND, "adresse inconnue")
        try:
            # The STDHL content depends on these settings: when they change, the stored files are read again.
            # The values before the update, to know afterward what changed.
            folders = {name: configuration.get(name) for name in FOLDER_VARIABLES}
            settings = ("requirement_regex", "fiche_regex", "stdhl_name_regex", "headers_stdhl")
            tags = ("requirement_regex", "test_tag", "test_case_tag", "requirement_tag")
            reading = ("requirement_regex", "code_separator_regex", "headers_plan")   # how plans are read
            before = {name: configuration.get(name) for name in settings + tags + reading}
            result = configuration.update(json.loads(self.read_body() or b"{}"))
            # Only the documents whose reading settings changed are read again.
            changed = {name for name, value in before.items() if configuration.get(name) != value}
            if changed & set(settings):
                stdhl.read_missing(database, configuration, again=True)
            if changed & set(tags):
                fiches.read_missing(database, configuration, again=True)
            if changed & set(reading):
                plans.read_missing(database, configuration, again=True)
            # A renamed sub-folder keeps its files.
            for name, old in folders.items():
                new = configuration.get(name)
                if new != old and (files_dir / old).is_dir() and not (files_dir / new).exists():
                    (files_dir / old).rename(files_dir / new)
            self.send_json(result)
        except ValueError as e:   # also covers invalid JSON
            self.send_text_error(HTTPStatus.BAD_REQUEST, str(e))

    def do_DELETE(self):
        """Remove a row of a table or a file."""
        if m := re.fullmatch(r"/api/tables/(\w+)/(-?\d+)", urlparse(self.path).path):
            # A file row is removed like from the page Gestion des fichiers: its copy and its content go with it.
            if m[1] == "files":   # with its stored copy
                return self.send_json(self.remove_file(int(m[2]))) if database.get(int(m[2])) else \
                    self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
            return self.change_table(lambda values: sql.delete_row(m[1], int(m[2])))
        file_id = self.file_id()
        if file_id is None or database.get(file_id) is None:
            return self.send_text_error(HTTPStatus.NOT_FOUND, "fichier introuvable")
        self.send_json(self.remove_file(file_id))


class Server(ThreadingHTTPServer):
    """One thread per request: a browser connection left open must not block the others; the database has its lock."""
    request_queue_size = 64
    # Under Windows, address reuse would let two servers share a port, and write the same database in turn.
    allow_reuse_address = sys.platform != "win32"


def main():
    """Read the options, open the storage folder and run the server."""
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

    # The port given on the command line wins over the configuration.
    port = args.port or configuration.get("server_port")
    # Local use only: listen on the loopback interface, not on the network.
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
