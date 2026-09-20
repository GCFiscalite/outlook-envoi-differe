"""Fabrique manifest.xml a partir du gabarit, pour une URL de base donnee.

    py -3.12 tools\\generer_manifeste.py --base https://127.0.0.1:3000
    py -3.12 tools\\generer_manifeste.py --base https://gcfiscalite.github.io/outlook-envoi-differe

La version du manifeste est incrementee a chaque generation si elle n'est pas
imposee : Outlook ne recharge un complement sideloade que lorsque la version
change, et son cache de manifeste peut tenir jusqu'a 24 heures.
"""

import argparse
import re
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
GABARIT = RACINE / "manifest.template.xml"
SORTIE = RACINE / "manifest.xml"


def version_courante() -> str:
    if not SORTIE.exists():
        return "1.0.0.0"
    trouve = re.search(r"<Version>([\d.]+)</Version>", SORTIE.read_text(encoding="utf-8"))
    return trouve.group(1) if trouve else "1.0.0.0"


def incrementer(version: str) -> str:
    morceaux = [int(x) for x in version.split(".")]
    while len(morceaux) < 4:
        morceaux.append(0)
    morceaux[3] += 1
    return ".".join(str(x) for x in morceaux)


def main() -> None:
    a = argparse.ArgumentParser()
    a.add_argument("--base", required=True,
                   help="URL de base, sans barre oblique finale")
    a.add_argument("--version", help="version imposee, sinon la derniere + 1")
    args = a.parse_args()

    base = args.base.rstrip("/")
    if not base.startswith("https://"):
        raise SystemExit("L'URL de base doit etre en https, Outlook refuse le reste.")

    version = args.version or incrementer(version_courante())

    # AppDomains attend l'origine seule, sans chemin.
    origine = "/".join(base.split("/")[:3])

    texte = GABARIT.read_text(encoding="utf-8")
    texte = texte.replace("<AppDomain>__BASE__</AppDomain>",
                          f"<AppDomain>{origine}</AppDomain>")
    texte = texte.replace("__BASE__", base).replace("__VERSION__", version)

    SORTIE.write_text(texte, encoding="utf-8")
    print(f"manifest.xml genere\n  version : {version}\n  base    : {base}")


if __name__ == "__main__":
    main()
