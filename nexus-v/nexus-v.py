"""NEXUS-V — lancement de l'application.

À lancer avec n'importe quel Python 3 (double-clic sur le fichier, ou « py nexus-v.py », ou « python nexus-v.py ») :
quand ce Python n'a pas les modules nécessaires (requirements.txt) et que l'environnement du projet (.venv) existe,
le serveur est relancé avec le Python de cet environnement. Les options sont transmises au serveur, par exemple :
py nexus-v.py --port 8001 --sans-navigateur
"""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROGRAM_DIR = ROOT / "programme"   # le code ; la racine ne contient que ce fichier et les dossiers de données
VENV_PYTHON = ROOT / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def has_modules():
    """True quand le Python en cours a les modules dont NEXUS-V a besoin."""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        return False
    return True


def wait_if_double_clicked():
    """Garde la fenêtre ouverte sur une erreur, pour que le message puisse être lu quand le fichier a été lancé par
    un double-clic."""
    # Seulement dans une fenêtre de console : lancé depuis un terminal ou un service, rien n'attend.
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Appuyer sur Entrée pour fermer.")
        except EOFError:
            pass


def main():
    """Lance le serveur avec ce Python, ou avec celui de .venv quand celui-ci n'a pas les modules."""
    # Les chemins relatifs (base, fichiers-references) partent du dossier de NEXUS-V, d'où qu'il ait été lancé.
    os.chdir(ROOT)
    if not has_modules():
        # Relancé avec le Python de .venv, avec les mêmes options ; son code de sortie devient le nôtre. La vérification
        # du chemin évite de le relancer sans fin si .venv lui-même n'a pas les modules.
        if VENV_PYTHON.is_file() and Path(sys.executable).resolve() != VENV_PYTHON.resolve():
            code = subprocess.call([str(VENV_PYTHON), str(ROOT / "nexus-v.py"), *sys.argv[1:]])
            sys.exit(code)
        print("Module openpyxl introuvable pour ce Python :", sys.executable)
        print("L'installer : py -m pip install -r programme/requirements.txt")
        print(f"(ou créer l'environnement du projet : py -m venv .venv puis {VENV_PYTHON} -m pip install -r "
              "programme/requirements.txt)")
        wait_if_double_clicked()
        sys.exit(1)
    # Le code est dans programme/ : il est mis en tête du chemin d'import, puis le serveur tourne dans ce processus.
    sys.path.insert(0, str(PROGRAM_DIR))
    try:
        import serveur
        serveur.main()
    except SystemExit as e:
        # Un arrêt normal passe ; un message d'erreur du serveur est montré et la fenêtre attend.
        if e.code in (None, 0):
            raise
        print(e.code if isinstance(e.code, str) else f"Arrêt (code {e.code}).")
        wait_if_double_clicked()
        sys.exit(1)


if __name__ == "__main__":
    main()
