"""Genere les icones du complement aux couleurs de GC Fiscalite.

Motif : une horloge pleine taupe, aiguilles blanches, marqueur orange en haut.
Dessine 8x plus grand puis reduit, pour un anticrenelage propre a 16 px.
"""

from pathlib import Path

from PIL import Image, ImageDraw

TAUPE = (85, 77, 74, 255)
ORANGE = (233, 141, 32, 255)
BLANC = (255, 255, 255, 255)

TAILLES = (16, 32, 64, 80, 128)
SUR = 8  # facteur de suréchantillonnage

DESTINATION = Path(__file__).resolve().parent.parent / "assets"


def dessiner(taille: int) -> Image.Image:
    c = taille * SUR
    img = Image.new("RGBA", (c, c), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    marge = c * 0.09
    d.ellipse([marge, marge, c - marge, c - marge], fill=TAUPE)

    centre = c / 2
    rayon = centre - marge
    epaisseur = max(1, int(c * 0.055))

    # Aiguille des minutes vers midi, aiguille des heures vers 4 h : l'angle
    # obtus reste lisible a 16 px, contrairement a un 10 h 10 qui se lit « coche ».
    d.line([centre, centre, centre, centre - rayon * 0.62],
           fill=BLANC, width=epaisseur)
    d.line([centre, centre, centre + rayon * 0.34, centre + rayon * 0.30],
           fill=BLANC, width=epaisseur)

    # Pivot central, pour que les deux aiguilles ne forment pas un seul trait coude.
    p = epaisseur * 0.9
    d.ellipse([centre - p, centre - p, centre + p, centre + p], fill=BLANC)

    # Pastille orange de la marque, en haut a droite, debordant du cadran.
    r = c * 0.20
    cx = c - r - c * 0.02
    cy = r + c * 0.02
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=ORANGE)

    return img.resize((taille, taille), Image.LANCZOS)


def dessiner_contour(taille: int) -> Image.Image:
    """Icone « outline » du paquet : monochrome blanche sur fond transparent.

    Le manifeste unifie l'exige a cote de l'icone couleur ; Office la recolore
    selon le theme, donc seule la silhouette compte.
    """
    c = taille * SUR
    img = Image.new("RGBA", (c, c), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    marge = c * 0.09
    trait = max(1, int(c * 0.08))
    d.ellipse([marge, marge, c - marge, c - marge], outline=BLANC, width=trait)

    centre = c / 2
    rayon = centre - marge
    d.line([centre, centre, centre, centre - rayon * 0.60], fill=BLANC, width=trait)
    d.line([centre, centre, centre + rayon * 0.34, centre + rayon * 0.30],
           fill=BLANC, width=trait)

    return img.resize((taille, taille), Image.LANCZOS)


def main() -> None:
    DESTINATION.mkdir(parents=True, exist_ok=True)
    for taille in TAILLES:
        chemin = DESTINATION / f"icone-{taille}.png"
        dessiner(taille).save(chemin, "PNG")
        print(f"  {chemin.name}")

    # Les deux icones du paquet .zip du manifeste unifie.
    dessiner(192).save(DESTINATION / "color.png", "PNG")
    print("  color.png (192)")
    dessiner_contour(32).save(DESTINATION / "outline.png", "PNG")
    print("  outline.png (32, monochrome)")


if __name__ == "__main__":
    main()
