# Signatures O365

![version](https://img.shields.io/badge/version-1.9-blue) ![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6) ![tests](https://github.com/Henri2214/signatures-o365/actions/workflows/tests.yml/badge.svg)

Petit outil portable qui lit un dossier de signatures Outlook (`Nom.htm`, `Nom.txt`, `Nom_fichiers\`) et les crée directement dans la boîte Exchange Online, comme une signature créée à la main dans Outlook Web. Elles apparaissent ensuite dans OWA, le nouvel Outlook et Outlook Microsoft 365, sans les recréer à la main.

## Prérequis côté tenant et poste

Les signatures doivent être **itinérantes** (stockées dans la boîte) :

- Exchange Online : `Get-OrganizationConfig | fl PostponeRoamingSignaturesUntilLater` doit renvoyer `False`. Sinon : `Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $false`.
- Outlook classique : la valeur `DisableRoamingSignaturesTemporaryToggle` (sous `HKCU\Software\Microsoft\Office\16.0\Outlook\Setup`) ne doit pas valoir `1`. L'outil affiche une alerte si c'est le cas sur le poste. Le script `scripts\Import-OutlookSignatures.ps1 -DisableRoaming` la positionne : ne pas l'utiliser avec cet outil.
- Les stratégies d'accès conditionnel du tenant doivent autoriser la connexion d'Office sur le poste.

## Construire l'exécutable

Sur un poste Windows, lancer `build.bat`. Il installe Python 3.12 si besoin, crée un environnement propre, **lance les tests**, puis produit `dist\SignaturesO365.exe`, un fichier unique copiable sur une clé ou un SSD. L'exe crée à côté de lui :

- `data\` : configuration et cache de connexion **chiffré** (DPAPI, lisible uniquement par le même utilisateur Windows sur le même PC) ;
- `exports\` : sauvegardes de boîtes, simulations, aperçus et journal du jour (`journal_AAAAMMJJ.txt`).

Chaque version publiée (tag `v*`) est aussi construite automatiquement par GitHub Actions et jointe à la release avec son empreinte SHA-256.

Pour le lancer sans compiler : `pip install -r requirements.txt` puis `python sig_o365.py`.

## Utilisation (interface)

1. **Dossier** : *Mes signatures Outlook* prend directement le dossier d'Outlook classique du poste ; sinon *Parcourir…*. La liste affiche les images, la taille finale (images intégrées en base64) et les **alertes** en orange : image introuvable, format non affiché par OWA (EMF, WMF, SVG…), image lourde, lien vers un fichier local, nom invalide. *Aperçu* (ou double-clic) ouvre le HTML tel qu'il sera envoyé.
   Cocher les signatures à importer en cliquant sur leur ligne (ou avec la barre d'espace).
2. **Connexion** : mettre le domaine dans *Tenant* (ex. `exemple.fr`), puis *Se connecter (Office 365)*. La fenêtre de connexion Microsoft de Windows s'ouvre (MFA comprise). Aucune application Entra n'est nécessaire.
3. **Signatures par défaut** : choisir la signature des nouveaux messages et des réponses, ou *Reprendre ceux d'Outlook* pour recopier les choix d'Outlook classique.
4. **Lire les signatures du cloud** : liste les signatures déjà présentes, apprend leur format et repère les paramètres de signature par défaut. Copie brute dans `exports\`.
5. **Simulation** (cochée par défaut) : génère seulement les JSON dans `exports\`.
6. **Envoi réel** : décocher *Simulation*, cocher une signature, *ENVOYER*. Une **sauvegarde de la boîte** est faite avant chaque envoi (`exports\sauvegarde_*.json`), puis l'outil relit la boîte pour confirmer. *Restaurer une sauvegarde…* renvoie les signatures d'un de ces fichiers.

Une signature portant le même nom qu'une signature existante est remplacée.

## Ligne de commande (déploiement)

Avec des arguments, l'exe fonctionne sans fenêtre (script de logon, Intune, tâche planifiée). **Sans `--apply`, rien n'est envoyé.**

```
SignaturesO365.exe --list
SignaturesO365.exe --all --outlook-defaults
SignaturesO365.exe --all --outlook-defaults --apply
SignaturesO365.exe --folder "\\srv\Signatures" --name Societe --default-new Societe --apply
SignaturesO365.exe --read-cloud
SignaturesO365.exe --restore exports\sauvegarde_x.json --apply
SignaturesO365.exe --all --apply --no-interactive     (échoue si aucun jeton en cache)
```

Sans `--folder`, le dossier d'Outlook classique de l'utilisateur est utilisé. Codes de retour : `0` succès, `1` échec partiel, `2` paramètre ou dossier invalide, `3` connexion requise. Le journal est écrit dans `exports\`.

## Limites à connaître

- L'outil utilise le service interne `substrate.office.com/ows/beta/outlookcloudsettings`, celui qu'emploient Outlook et OWA, avec l'identifiant client d'Office. **Microsoft ne le documente pas** : il fonctionne aujourd'hui mais peut changer ou être bloqué par une stratégie du tenant. Toujours tester sur une boîte pilote.
- **Signatures par défaut** : le nom du paramètre côté boîte n'est pas documenté. L'outil le repère s'il existe déjà ; sur une boîte où aucun défaut n'a jamais été choisi, le choisir une fois dans OWA puis relancer *Lire les signatures du cloud*.
- Les images sont intégrées en base64 (`data:`), comme le fait Outlook. Certains clients de destinataires les bloquent : préférer des images légères, ou des images hébergées (`https://`).
- Écrire dans une autre boîte que la sienne (boîte partagée) n'est possible qu'avec un accès complet, et ce n'est pas garanti par ce service.
- Les seuils d'alerte (200 Ko par image, 1 Mo par signature) sont indicatifs : la limite réelle du service n'est pas publiée.
- Solution de repli documentée par Microsoft, si le service change : `Set-MailboxMessageConfiguration -SignatureHtml … -AutoAddSignature $true` (Exchange Online PowerShell), qui ne gère qu'une seule signature par boîte.

## Développement

```
pip install -r requirements-dev.txt
python -m pytest -q tests
```

- `sigcore.py` : lecture du dossier, nettoyage du HTML Word/VML, conversion texte, format cloud, registre Outlook. Sans dépendance externe.
- `sig_o365.py` : connexion (MSAL + WAM), appels au service, interface Tkinter et ligne de commande.
- `tests\` : tests pytest, sans réseau.
- `build.bat` : construction de l'exe portable avec PyInstaller.
- `scripts\Import-OutlookSignatures.ps1` : script PowerShell séparé qui copie des signatures dans le dossier **local** d'Outlook classique (sans cloud).

## Licence

MIT, voir `LICENSE`.
