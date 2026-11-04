# Signatures O365

![version](https://img.shields.io/badge/version-1.8-blue) ![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6)

Petit outil portable qui lit un dossier de signatures Outlook (`Nom.htm`, `Nom.txt`, `Nom_fichiers\`) et les crée directement dans la boîte Exchange Online, comme une signature créée à la main dans Outlook Web. Elles apparaissent ensuite dans OWA, le nouvel Outlook et Outlook Microsoft 365, sans copie locale.

## Construire l'exécutable

Sur un poste Windows avec Python 3.10 ou plus récent, lancer `build.bat`. Le résultat est `dist\SignaturesO365.exe`, un fichier unique, copiable sur une clé ou un SSD. L'exe crée à côté de lui deux dossiers : `data\` (cache de connexion) et `exports\` (simulations, lectures cloud).

Pour le lancer sans compiler : `pip install msal requests` puis `python sig_o365.py`.

## Utilisation

1. **Dossier** : choisir le dossier de signatures puis cliquer sur *Analyser*. La liste affiche le nombre d'images et la taille finale (images intégrées en base64).
   Cocher les signatures à importer en cliquant sur leur ligne (ou avec la barre d'espace). Les boutons *Tout cocher*, *Tout décocher* et *Inverser* aident pour les gros dossiers.
2. **Connexion** : mettre le domaine dans *Tenant* (ex. `exemple.fr`), puis cliquer sur *Se connecter (Office 365)*. La fenêtre de connexion Microsoft de Windows s'ouvre (MFA comprise), comme pour Outlook. Aucune application Entra n'est nécessaire.
3. **Test** : cliquer sur *Lire les signatures du cloud*. On obtient la liste des signatures déjà présentes et une copie brute dans `exports\`.
4. **Simulation** : la case est cochée par défaut. *Envoyer* génère seulement les JSON dans `exports\`.
5. **Envoi réel** (le bouton affiche alors *ENVOYER dans la boîte mail*) : décocher *Simulation*, cocher une seule signature, puis *Envoyer*. Vérifier dans OWA. Si c'est bon, cocher les autres et les envoyer. Seules les signatures cochées sont importées. Après l'envoi, l'outil relit la boîte et indique quelles signatures y sont bien présentes.

Une signature portant le même nom qu'une signature existante est remplacée.

## Limites à connaître

- L'outil utilise le service interne `substrate.office.com/ows/beta/outlookcloudsettings`, celui qu'emploient Outlook et OWA. **Microsoft ne le documente pas** : il fonctionne aujourd'hui mais peut changer.
- La signature **par défaut** (nouveau message / réponse) n'est pas encore définie par l'outil. Il suffit de la choisir une fois dans OWA.
- Écrire dans une autre boîte que la sienne (boîte partagée) n'est possible qu'avec un accès complet, et ce n'est pas garanti par ce service.

## Contenu du dépôt

- `sig_o365.py` : l'application (Tkinter, MSAL avec connexion Windows WAM, requests).
- `build.bat` : construction de l'exe portable avec PyInstaller (installe Python 3.12 si besoin).
- `scripts/Import-OutlookSignatures.ps1` : script PowerShell séparé qui copie des signatures dans le dossier **local** d'Outlook classique et définit les signatures par défaut (sans cloud).

## Licence

MIT, voir `LICENSE`.
