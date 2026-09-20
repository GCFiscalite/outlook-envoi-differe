"""Controle manifest.xml avant tout sideload.

Outlook rejette un manifeste mal forme sans toujours dire pourquoi : le volet
n'apparait simplement jamais. On verifie donc ici ce qui se verifie sans Outlook :

  - XML bien forme, GUID et version au bon format ;
  - elements obligatoires presents et dans l'ordre impose par le schema ;
  - toutes les URL du manifeste repondent bien 200 sur l'hote vise.

    py -3.12 tools\\verifier.py
"""

import re
import ssl
import sys
import urllib.error
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
MANIFESTE = RACINE / "manifest.xml"
CERT_LOCAL = RACINE / "tools" / "cert" / "localhost.pem"

NS = "{http://schemas.microsoft.com/office/appforoffice/1.1}"

# L'ordre est impose par le schema OfficeApp : un element au mauvais rang
# invalide tout le manifeste.
ORDRE = [
    "Id", "Version", "ProviderName", "DefaultLocale", "DisplayName",
    "Description", "IconUrl", "HighResolutionIconUrl", "SupportUrl",
    "AppDomains", "Hosts", "Requirements", "FormSettings", "Permissions",
    "Rule", "DisableEntityHighlighting", "VersionOverrides",
]

erreurs: list[str] = []
avertissements: list[str] = []


def echec(m: str) -> None:
    erreurs.append(m)


def nom_local(tag: str) -> str:
    # VersionOverrides porte un espace de noms different du reste du manifeste :
    # on compare donc toujours sur le nom local.
    return tag.split("}")[-1]


def verifier_structure(racine: ET.Element) -> None:
    presents = [nom_local(e.tag) for e in racine]

    for attendu in ORDRE:
        if attendu not in presents:
            echec(f"element obligatoire absent : <{attendu}>")

    rangs = [ORDRE.index(p) for p in presents if p in ORDRE]
    if rangs != sorted(rangs):
        echec("les elements ne sont pas dans l'ordre impose par le schema : "
              + ", ".join(presents))

    ident = racine.findtext(f"{NS}Id", "")
    try:
        uuid.UUID(ident)
    except ValueError:
        echec(f"Id n'est pas un GUID valide : {ident!r}")

    version = racine.findtext(f"{NS}Version", "")
    if not re.fullmatch(r"\d+(\.\d+){1,3}", version):
        echec(f"Version mal formee : {version!r}")

    permissions = racine.findtext(f"{NS}Permissions", "")
    if permissions != "ReadWriteItem":
        echec(f"Permissions devrait valoir ReadWriteItem, pas {permissions!r} "
              f"(ecrire delayDeliveryTime l'exige)")


def collecter_urls(texte: str) -> list[str]:
    return sorted(set(re.findall(r'DefaultValue="(https://[^"]+)"', texte)))


def verifier_urls(urls: list[str]) -> None:
    contexte = ssl.create_default_context()
    if CERT_LOCAL.exists():
        try:
            contexte.load_verify_locations(cafile=str(CERT_LOCAL))
        except ssl.SSLError:
            pass

    for url in urls:
        if "gcfiscalite.com" in url:      # SupportUrl, hors du complement
            continue
        try:
            with urllib.request.urlopen(url, timeout=8, context=contexte) as r:
                if r.status != 200:
                    echec(f"{url} repond {r.status}")
                else:
                    print(f"  200  {url}")
        except urllib.error.HTTPError as e:
            echec(f"{url} repond {e.code}")
        except Exception as e:                       # noqa: BLE001
            echec(f"{url} injoignable : {type(e).__name__} {e}")


def main() -> int:
    if not MANIFESTE.exists():
        print("manifest.xml absent : lancer d'abord tools\\generer_manifeste.py",
              file=sys.stderr)
        return 1

    texte = MANIFESTE.read_text(encoding="utf-8")

    try:
        racine = ET.fromstring(texte)
    except ET.ParseError as e:
        print(f"XML invalide : {e}", file=sys.stderr)
        return 1

    verifier_structure(racine)

    urls = collecter_urls(texte)
    print(f"{len(urls)} URL(s) declarees dans le manifeste :")
    verifier_urls(urls)

    for a in avertissements:
        print(f"  avertissement : {a}")

    if erreurs:
        print("\nPROBLEMES :")
        for e in erreurs:
            print(f"  - {e}")
        return 1

    print("\nManifeste conforme.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
