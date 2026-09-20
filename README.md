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

### Mode Distant (recommandé)

Le volet est servi par une URL HTTPS publique. Rien à installer sur le poste à
part le manifeste, et c'est le seul mode qui passe à l'échelle pour l'équipe.

```powershell
.\tools\installer.ps1 -Mode Distant -UrlBase https://<hote>/outlook-envoi-differe
```

Le contenu de `src/` et `assets/` est déployé tel quel sur l'hôte. Le code est du
HTML/CSS/JS statique sans backend : **aucun contenu de message ne quitte
Outlook**, tout s'exécute dans le WebView du client.

### Mode Local

Le volet est servi par un petit serveur HTTPS qui ne répond que sur
`127.0.0.1:3000`. Aucune dépendance externe, mais une installation par poste.

```powershell
.\tools\installer.ps1
```

L'installeur génère un certificat auto-signé pour `localhost`, l'ajoute au
magasin racine de l'utilisateur — **Windows demande confirmation une fois** —
crée une tâche planifiée qui démarre le serveur à l'ouverture de session, puis
inscrit le manifeste.

### Dans les deux cas

Le manifeste est chargé par la clé de sideload d'Office, sans passer par le
centre d'administration :

```
HKCU\Software\Microsoft\Office\16.0\WEF\Developer
    GCFiscalite.EnvoiProgramme = <chemin>\manifest.xml
```

**Outlook doit être fermé et rouvert** pour relire un manifeste. Il met parfois
jusqu'à une heure, et dans le pire des cas 24 h, à prendre en compte un
changement de manifeste : c'est pour cette raison que `generer_manifeste.py`
incrémente le numéro de version à chaque génération.

### Désinstallation

```powershell
.\tools\installer.ps1 -Desinstaller
```

Retire la clé de registre, la tâche planifiée, le serveur et le certificat.

## Développement

| Commande | Rôle |
|---|---|
| `py -3.12 tools\apercu.py` | Banc d'essai hors Outlook, sur `http://127.0.0.1:3100/tools/apercu/index.html` |
| `py -3.12 tools\generer_manifeste.py --base <url>` | Fabrique `manifest.xml` |
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
