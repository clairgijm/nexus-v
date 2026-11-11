"""NEXUS-V — starting the application.

To be started with any Python 3 (double-click on the file, or "py nexus-v.py", or "python nexus-v.py"): when that
Python lacks the modules needed (requirements.txt) and the project environment (.venv) exists, the server is started
again with the Python of that environment. The options are passed on to the server, for instance:
py nexus-v.py --port 8001 --sans-navigateur
"""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROGRAM_DIR = ROOT / "programme"   # the code; the root only holds this file and the data folders
VENV_PYTHON = ROOT / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def has_modules():
    """True when the running Python has the modules NEXUS-V needs."""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        return False
    return True


def wait_if_double_clicked():
    """Keep the window open on an error, so that the message can be read when the file was double-clicked."""
    # Only in a console window: started from a terminal or a service, nothing waits.
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Appuyer sur Entrée pour fermer.")
        except EOFError:
            pass


def main():
    """Run the server with this Python, or with the one of .venv when this one lacks the modules."""
    # Relative paths (base, fichiers-references) are taken from the folder of NEXUS-V, wherever it was started from.
    os.chdir(ROOT)
    if not has_modules():
        # Started again with the Python of .venv, with the same options; its exit code becomes ours. The check on the
        # path avoids starting it again forever if .venv itself lacks the modules.
        if VENV_PYTHON.is_file() and Path(sys.executable).resolve() != VENV_PYTHON.resolve():
            code = subprocess.call([str(VENV_PYTHON), str(ROOT / "nexus-v.py"), *sys.argv[1:]])
            sys.exit(code)
        print("Module openpyxl introuvable pour ce Python :", sys.executable)
        print("L'installer : py -m pip install -r programme/requirements.txt")
        print(f"(ou créer l'environnement du projet : py -m venv .venv puis {VENV_PYTHON} -m pip install -r "
              "programme/requirements.txt)")
        wait_if_double_clicked()
        sys.exit(1)
    # The code is in programme/: it is put first on the import path, then the server runs in this process.
    sys.path.insert(0, str(PROGRAM_DIR))
    try:
        import serveur
        serveur.main()
    except SystemExit as e:
        # A normal stop goes through; an error message of the server is shown and the window waits.
        if e.code in (None, 0):
            raise
        print(e.code if isinstance(e.code, str) else f"Arrêt (code {e.code}).")
        wait_if_double_clicked()
        sys.exit(1)


if __name__ == "__main__":
    main()
