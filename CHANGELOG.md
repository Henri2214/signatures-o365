# Changelog

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
