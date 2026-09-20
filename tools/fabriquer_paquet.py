"""Fabrique le paquet .zip du complement (manifeste unifie Microsoft 365).

    py -3.12 tools\\fabriquer_paquet.py --base https://localhost:3000
    py -3.12 tools\\fabriquer_paquet.py --base https://... --schema devPreview

Produit paquet\\manifest.json puis paquet\\envoi-programme.zip, qui contient
le manifeste et les deux icones exigees, a la racine de l'archive.

Le paquet se televerse dans le centre d'administration Microsoft 365, sous
Parametres > Applications integrees > Charger les applications personnalisees,
en choisissant le type « Complement Office ».
"""

import argparse
import json
import re
import zipfile
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
GABARIT = RACINE / "manifest.template.json"
PAQUET = RACINE / "paquet"
ARCHIVE = PAQUET / "envoi-programme.zip"

# Longueurs maximales imposees par le schema : les depasser fait echouer la
# validation avec un message generique, autant les attraper ici.
LIMITES = {
    "name.short": 30,
    "name.full": 100,
    "description.short": 80,
    "description.full": 4000,
}


def version_courante() -> str:
    cible = PAQUET / "manifest.json"
    if not cible.exists():
        return "1.0.0"
    try:
        return json.loads(cible.read_text(encoding="utf-8")).get("version", "1.0.0")
    except (json.JSONDecodeError, OSError):
        return "1.0.0"


def incrementer(version: str) -> str:
    morceaux = [int(x) for x in version.split(".")]
    while len(morceaux) < 3:
        morceaux.append(0)
    morceaux[2] += 1
    return ".".join(str(x) for x in morceaux[:3])


def controler(manifeste: dict) -> list[str]:
    soucis = []

    for chemin, maxi in LIMITES.items():
        a, b = chemin.split(".")
        valeur = manifeste.get(a, {}).get(b, "")
        if len(valeur) > maxi:
            soucis.append(f"{chemin} fait {len(valeur)} caracteres, maximum {maxi}")

    if not re.fullmatch(r"\d{1,5}\.\d{1,5}\.\d{1,5}", manifeste.get("version", "")):
        soucis.append(f"version {manifeste.get('version')!r} : attendu n.n.n")

    urls = re.findall(r'"(https?://[^"]+)"', json.dumps(manifeste))
    for u in urls:
        if u.startswith("http://"):
            soucis.append(f"{u} n'est pas en https")

    ext = manifeste.get("extensions", [{}])[0]
    actions = {a["id"] for r in ext.get("runtimes", []) for a in r.get("actions", [])}
    for ruban in ext.get("ribbons", []):
        for onglet in ruban.get("tabs", []):
            for groupe in onglet.get("groups", []):
                for ctrl in groupe.get("controls", []):
                    if ctrl.get("actionId") not in actions:
                        soucis.append(
                            f"le controle {ctrl.get('id')} pointe vers actionId "
                            f"{ctrl.get('actionId')!r}, qui n'existe dans aucun runtime")

    return soucis


def main() -> None:
    a = argparse.ArgumentParser()
    a.add_argument("--base", required=True, help="URL de base, sans barre oblique finale")
    a.add_argument("--version", help="version imposee, sinon la derniere + 1")
    a.add_argument("--schema", default="1.17",
                   help="version du schema : 1.17 (defaut) ou devPreview")
    args = a.parse_args()

    base = args.base.rstrip("/")
    if not base.startswith("https://"):
        raise SystemExit("L'URL de base doit etre en https.")

    version = args.version or incrementer(version_courante())

    texte = GABARIT.read_text(encoding="utf-8")
    texte = texte.replace("__BASE__", base)
    texte = texte.replace("__VERSION__", version)
    texte = texte.replace("__SCHEMA__", args.schema)

    try:
        manifeste = json.loads(texte)
    except json.JSONDecodeError as e:
        raise SystemExit(f"Le gabarit ne produit pas du JSON valide : {e}")

    soucis = controler(manifeste)
    if soucis:
        print("PROBLEMES :")
        for s in soucis:
            print(f"  - {s}")
        raise SystemExit(1)

    PAQUET.mkdir(parents=True, exist_ok=True)
    (PAQUET / "manifest.json").write_text(
        json.dumps(manifeste, ensure_ascii=False, indent=2), encoding="utf-8")

    icones = {
        "color.png": RACINE / "assets" / "color.png",
        "outline.png": RACINE / "assets" / "outline.png",
    }
    for nom, chemin in icones.items():
        if not chemin.exists():
            raise SystemExit(f"{nom} manquante : lancer tools\\generer_icones.py")

    with zipfile.ZipFile(ARCHIVE, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("manifest.json", json.dumps(manifeste, ensure_ascii=False, indent=2))
        for nom, chemin in icones.items():
            z.write(chemin, nom)

    print(f"Paquet fabrique : {ARCHIVE}")
    print(f"  schema  : {args.schema}")
    print(f"  version : {version}")
    print(f"  base    : {base}")
    print(f"  contenu : {', '.join(zipfile.ZipFile(ARCHIVE).namelist())}")


if __name__ == "__main__":
    main()
