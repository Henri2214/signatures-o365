# Changelog

## 1.9.1
- `build.bat` construit dans un dossier temporaire : en cas d'échec, l'exe précédent reste intact.
- `dist\` ne contient plus que le nouvel exe et son empreinte SHA-256 ; `dist\data` et `dist\exports` sont conservés.
- Suppression de tous les fichiers temporaires après la construction (environnement, `build\`, `.spec`, caches Python et pytest).
- Tests GitHub lancés une seule fois par modification.

## 1.9
- Nouveau module `sigcore.py` (sans dépendance) et tests pytest, lancés par `build.bat` et GitHub Actions.
- Nettoyage du HTML Word/VML avant import (blocs VML en double, îlots XML, balises `o:`/`v:`/`w:`, liens locaux).
- Alertes dès l'analyse : images introuvables, formats non affichés par OWA (EMF, WMF, SVG…), images lourdes, liens vers des fichiers locaux, noms invalides, doublons `.htm`/`.html`.
- Version texte : les cellules de tableau sont séparées.
- Bouton *Mes signatures Outlook* (dossier localisé lu dans le registre) et *Aperçu* du HTML envoyé.
- Signatures par défaut (nouveau message, réponse/transfert), reprise possible depuis Outlook classique.
- Sauvegarde automatique de la boîte avant envoi et bouton *Restaurer une sauvegarde*.
- Lecture du format de la boîte aussi en simulation, quand on est connecté.
- Nouvel essai automatique sur erreurs 429/5xx et réseau.
- Cache de jeton chiffré par DPAPI ; l'ancien cache en clair est migré puis supprimé.
- Journal enregistré dans `exports\`, barre d'activité, une seule opération à la fois.
- Mode ligne de commande pour le déploiement (`--all`, `--apply`, `--outlook-defaults`, `--restore`…).
- Alerte si les signatures itinérantes sont désactivées sur le poste ; prérequis documentés.
- `requirements.txt` limité à l'exécution, `requirements-dev.txt` pour la construction et les tests.
- Release automatique de l'exe (avec SHA-256) sur les tags `v*`.
- Suppression du code mort (connexion OAuth par navigateur, aide sur l'application Entra).

## 1.8
- `build.bat` utilise Python 3.12, recrée l'environnement à chaque construction et vérifie la présence de `pymsalruntime`.

## 1.7
- Connexion unique « Office 365 » par la fenêtre de connexion Windows (WAM), sans application Entra.
- Le format des signatures est copié sur celles créées par Outlook.
- Relecture automatique de la boîte après envoi pour confirmer la présence des signatures.
- Bouton d'envoi explicite selon le mode (simulation ou envoi réel).

## 1.6
- Le diagnostic HTTP 401/403 est affiché lors de la lecture du cloud.

## 1.5
- Essai de connexion OAuth2 avec application Entra (abandonné : le service refuse les applications tierces).
- Tenant mémorisé dans `data\config.json`.

## 1.4
- Connexion via le gestionnaire de comptes Windows (WAM).
- Messages d'erreur AADSTS expliqués en français.

## 1.3
- Suppression de la dépendance à `html.parser` (absente de l'exe PyInstaller).

## 1.2
- Correction de l'import de `html.parser`.

## 1.1
- Cases à cocher pour choisir les signatures à importer.

## 1.0
- Première version : analyse d'un dossier de signatures, images intégrées en base64, envoi dans la boîte Exchange Online, mode simulation.
