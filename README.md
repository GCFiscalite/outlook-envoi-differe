# Envoi programmé — complément Outlook classique

Ajoute un bouton **Programmer l'envoi** dans la fenêtre de rédaction d'Outlook
classique. Contrairement à l'option native « Ne pas envoyer avant », le délai est
traité **par Exchange, pas par Outlook** : le message part à l'heure prévue même
si Outlook est fermé et le poste éteint.

## Pourquoi ce complément existe

| | Natif « Ne pas envoyer avant » | Ce complément |
|---|---|---|
| Où le message attend | Boîte d'envoi locale (.ost) | Sur le serveur Exchange |
| Outlook doit tourner | **Oui** | Non |
| Poste peut être éteint | Non | **Oui** |
| Modifiable après Envoyer | Oui, dans la Boîte d'envoi | **Non**, il file dans Éléments envoyés |

Le mécanisme est l'API Office `item.delayDeliveryTime.setAsync()`
([documentation Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/delay-delivery)),
disponible à partir du requirement set **Mailbox 1.13**, soit Outlook **2304
(build 16327.20248)** avec un compte Microsoft 365 / Exchange Online.

Effet de bord appréciable : le message ne transite jamais par la Boîte d'envoi,
donc il échappe au piège connu où la lecture COM de la Boîte d'envoi fige un
envoi différé.

## Installation

Un complément Outlook ne s'installe **ni par le registre, ni par un dossier local** :
il vit dans la boîte aux lettres Exchange. Deux points appris à la dure :

- la clé `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` fonctionne pour
  Word et Excel, **pas pour Outlook** ;
- le manifeste **XML** est refusé par le centre d'administration Microsoft 365
  (« Fichier de manifeste non valide », sans plus de détail, y compris pour un
  manifeste minimal conforme au schéma publié). Seul le **manifeste unifié JSON**,
  livré en .zip, passe la validation.

### Fabriquer le paquet

```powershell
py -3.12 tools\generer_icones.py
py -3.12 tools\fabriquer_paquet.py --base https://localhost:3000
```

Produit `paquet\envoi-programme.zip`, qui contient `manifest.json`, `color.png` et
`outline.png` à la racine de l'archive.

### Déployer

Centre d'administration Microsoft 365 → **Paramètres** → **Applications intégrées**
→ **Charger les applications personnalisées** → type **Application Teams** →
choisir le .zip → **Seulement moi** (ou les utilisateurs visés).

Le complément apparaît dans Outlook sur le web en quelques minutes. Dans Outlook
classique, la propagation peut prendre jusqu'à 24 heures : c'est le cache du
client, pas une erreur.

### Héberger le volet

`--base` désigne l'hôte du volet. Deux options :

| | Serveur local | URL publique |
|---|---|---|
| Commande | `tools\installer.ps1` | héberger `src/` et `assets/` |
| Ce qui sort du poste | rien | rien non plus, le code est statique |
| Fonctionne sur | le poste qui héberge | tous les postes, OWA, mobile |
| À installer par poste | certificat + tâche planifiée | rien |

Le mode local suffit pour un poste et se monte avec `tools\installer.ps1`, qui
génère un certificat pour `localhost`, l'ajoute au magasin racine de l'utilisateur
(Windows demande confirmation une fois) et crée une tâche planifiée qui démarre le
serveur à l'ouverture de session. Pour plusieurs utilisateurs, une URL publique
est le bon choix : le manifeste étant déployé par le tenant, `localhost` ne
répondrait sur aucun autre poste.

### Désinstallation

Le complément se retire depuis la même page du centre d'administration. Pour le
serveur local et son certificat :

```powershell
.\tools\installer.ps1 -Desinstaller
```

## Développement

| Commande | Rôle |
|---|---|
| `py -3.12 tools\apercu.py` | Banc d'essai hors Outlook, sur `http://127.0.0.1:3100/tools/apercu/index.html` |
| `py -3.12 tools\fabriquer_paquet.py --base <url>` | Fabrique le paquet .zip a deployer |
| `py -3.12 tools\generer_manifeste.py --base <url>` | Fabrique `manifest.xml` (ancien format, refuse par le portail) |
| `py -3.12 tools\verifier.py` | Contrôle le manifeste et l'accessibilité de ses URL |
| `py -3.12 tools\generer_icones.py` | Régénère les icônes |
| `py -3.12 tools\serveur.py` | Serveur HTTPS local |

Le banc d'essai sert les **vrais** `src/taskpane.css` et `src/taskpane.js`, avec
un faux `Office.js` à la place du CDN Microsoft. Scénarios pilotables par l'URL :

| Paramètre | Effet |
|---|---|
| `?no113=1` | Simule un Outlook sans Mailbox 1.13 |
| `?echecSet=1` | Échec de `setAsync` |
| `?echecSave=1` | Échec de `saveAsync` |
| `?echecGet=1` | Échec de `getAsync` |
| `?delai=<epoch ms>` | Ouvre le volet avec une programmation déjà posée |

## Limites connues

- **Compléments COM au clic Envoyer.** Microsoft ne garantit aucun ordre
  d'exécution entre un complément web et un complément COM. Un outil qui
  reconstruit le message avant de l'envoyer (chiffrement, DLP) peut écraser la
  propriété de livraison différée. À tester de bout en bout avec Secure
  Exchanges et Bells & Whistles avant de s'y fier pour un envoi chiffré.
- **Horizon de 30 jours.** Microsoft ne documente aucune limite maximale pour
  `delayDeliveryTime`. Les 30 jours sont une borne que le volet s'impose, pas une
  contrainte d'Exchange.
- **Pas d'annulation.** Une fois Envoyer cliqué dans Outlook classique, le
  message n'est plus visible que dans Éléments envoyés. Pour garder la
  possibilité de le modifier, il faut passer par OWA ou le nouvel Outlook, où le
  message attend dans Brouillons.
- **L'heure choisie est une heure de soumission**, pas une garantie de livraison
  à la seconde : antispam, connecteurs et serveur du destinataire s'ajoutent.
