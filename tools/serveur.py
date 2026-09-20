"""Sert le complement en HTTPS sur la boucle locale.

Outlook exige une SourceLocation en HTTPS : meme pour un complement qui ne sort
jamais du poste, il faut un certificat. Celui-ci est fabrique par
tools\\installer.ps1 et depose dans tools\\cert\\.

    py -3.12 tools\\serveur.py [--port 3000]

Rien n'est expose hors du poste : on n'ecoute que sur 127.0.0.1.
"""

import argparse
import datetime as dt
import ssl
import sys
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
CERT = RACINE / "tools" / "cert" / "localhost.pem"
CLE = RACINE / "tools" / "cert" / "localhost.key"
JOURNAL = RACINE / "tools" / "serveur.log"


class Gestionnaire(SimpleHTTPRequestHandler):
    """Sert les fichiers du projet en interdisant toute mise en cache.

    Le WebView2 d'Outlook garde volontiers une vieille copie d'un .js ; en local
    on prefere payer une requete de plus que deboguer une version fantome.
    """

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    def log_message(self, format, *args):
        ligne = "%s  %s\n" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                              format % args)
        try:
            with open(JOURNAL, "a", encoding="utf-8") as f:
                f.write(ligne)
        except OSError:
            pass


def main() -> int:
    a = argparse.ArgumentParser()
    a.add_argument("--port", type=int, default=3000)
    args = a.parse_args()

    if not CERT.exists() or not CLE.exists():
        print(f"Certificat introuvable : {CERT}\n"
              f"Lancer d'abord tools\\installer.ps1", file=sys.stderr)
        return 1

    contexte = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    contexte.load_cert_chain(certfile=str(CERT), keyfile=str(CLE))

    serveur = ThreadingHTTPServer(("127.0.0.1", args.port),
                                  partial(Gestionnaire, directory=str(RACINE)))
    serveur.socket = contexte.wrap_socket(serveur.socket, server_side=True)

    print(f"Complement servi sur https://localhost:{args.port}/src/taskpane.html")
    try:
        serveur.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
