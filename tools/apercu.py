"""Banc d'essai du volet, hors Outlook.

Sert une copie de src\\taskpane.html ou le office.js du CDN Microsoft est
remplace par tools\\apercu\\stub-office.js. Le CSS et le JS servis sont les
vrais fichiers du complement : ce que l'on valide ici est bien le code livre.

    py -3.12 tools\\apercu.py [--port 3100]

Puis ouvrir http://127.0.0.1:3100/ dans un navigateur.
"""

import argparse
import re
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
APERCU = RACINE / "tools" / "apercu"

SCRIPT_CDN = re.compile(
    r'<script src="https://appsforoffice\.microsoft\.com[^"]*"></script>')


def fabriquer_index() -> Path:
    page = (RACINE / "src" / "taskpane.html").read_text(encoding="utf-8")
    if not SCRIPT_CDN.search(page):
        raise SystemExit("La balise office.js attendue est introuvable dans taskpane.html")

    page = SCRIPT_CDN.sub('<script src="stub-office.js"></script>', page)
    page = page.replace('href="taskpane.css', 'href="../../src/taskpane.css')
    page = page.replace('src="taskpane.js', 'src="../../src/taskpane.js')

    cible = APERCU / "index.html"
    cible.write_text(page, encoding="utf-8")
    return cible


class Gestionnaire(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, format, *args):
        pass


def main() -> None:
    a = argparse.ArgumentParser()
    a.add_argument("--port", type=int, default=3100)
    args = a.parse_args()

    fabriquer_index()
    # On sert la racine du projet pour que l'index puisse referencer les vrais
    # src\taskpane.css et src\taskpane.js par chemin relatif.
    serveur = ThreadingHTTPServer(
        ("127.0.0.1", args.port),
        partial(Gestionnaire, directory=str(RACINE)))
    print(f"Banc d'essai : http://127.0.0.1:{args.port}/tools/apercu/index.html")
    serveur.serve_forever()


if __name__ == "__main__":
    main()
