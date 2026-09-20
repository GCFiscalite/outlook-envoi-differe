"""Fabrique un certificat auto-signe pour localhost, valable 3 ans.

Produit dans tools\\cert\\ :
    localhost.key  cle privee (ne quitte jamais le poste, hors du depot)
    localhost.pem  certificat servi par tools\\serveur.py
    localhost.cer  meme certificat, a importer dans le magasin Windows

Le certificat couvre localhost et 127.0.0.1, rien d'autre : il ne peut pas
servir a usurper un site public.
"""

import datetime as dt
import ipaddress
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

DOSSIER = Path(__file__).resolve().parent / "cert"
ANNEES = 3


def main() -> None:
    DOSSIER.mkdir(parents=True, exist_ok=True)

    cle = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    sujet = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "GC Fiscalite - complement Outlook"),
    ])

    maintenant = dt.datetime.now(dt.timezone.utc)
    certificat = (
        x509.CertificateBuilder()
        .subject_name(sujet)
        .issuer_name(sujet)
        .public_key(cle.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(maintenant - dt.timedelta(days=1))
        .not_valid_after(maintenant + dt.timedelta(days=365 * ANNEES))
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
            ]),
            critical=False,
        )
        # ca=False volontairement : ce certificat vaut pour lui-meme et ne peut pas
        # servir d'autorite pour signer quoi que ce soit d'autre, meme place dans
        # le magasin racine.
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.ExtendedKeyUsage([x509.oid.ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False,
        )
        .sign(cle, hashes.SHA256())
    )

    (DOSSIER / "localhost.key").write_bytes(cle.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    ))

    pem = certificat.public_bytes(serialization.Encoding.PEM)
    (DOSSIER / "localhost.pem").write_bytes(pem)
    (DOSSIER / "localhost.cer").write_bytes(pem)

    empreinte = certificat.fingerprint(hashes.SHA256()).hex().upper()
    print(f"Certificat cree dans {DOSSIER}")
    print(f"  expire le : {certificat.not_valid_after_utc:%Y-%m-%d}")
    print(f"  empreinte : {empreinte}")


if __name__ == "__main__":
    main()
