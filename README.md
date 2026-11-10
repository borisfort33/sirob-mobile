# SIRØB Mobile

La version iPhone du SIRØB Project Browser, réduite à l'essentiel : **chercher un morceau dans
tout le catalogue et envoyer son master, son bounce ou ses stems** (lien OneDrive en lecture seule)
par WhatsApp, SMS, AirDrop… ou par mail, en 3 touches.

C'est une « web app » : elle s'installe sur l'écran d'accueil depuis Safari, sans App Store, et
s'ouvre en plein écran comme une vraie app. Elle est hébergée gratuitement sur GitHub Pages et
parle directement à ton OneDrive : aucun serveur, aucun abonnement.

## Comment ça marche

1. L'app du Mac (v17 et plus) écrit à chaque scan un petit fichier
   `OneDrive/SIROB Browser/catalogue.json` : la liste de tes morceaux avec leurs bounces,
   masters et stems, trouvés avec exactement les mêmes règles que sur le Mac.
2. L'iPhone lit ce fichier via ton compte Microsoft : la recherche est instantanée, même sur 600
   projets. Le catalogue reste en mémoire sur le téléphone : il s'affiche même avec un mauvais
   réseau.
3. « Envoyer » crée le lien OneDrive « toute personne disposant du lien, lecture seule » (le même
   que le bouton 🔗 du Mac), puis te propose :
   - **Partager…** : la feuille de partage de l'iPhone (WhatsApp, Messages, AirDrop…) ;
   - **Envoyer par mail** : destinataire (avec tes contacts déjà utilisés sur le Mac), objet et
     message pré-remplis comme sur le Mac, signature comprise. L'app Mail s'ouvre avec tout
     prêt, il reste à appuyer sur Envoyer ;
   - **Copier le lien**.
4. Un bounce exporté après le dernier scan du Mac ? « Chercher dans tout OneDrive » en bas des
   résultats fouille directement OneDrive.

## Installation (une seule fois, environ 15 minutes)

### Étape 1 : mettre à jour l'app du Mac

1. Quitte l'app du Mac (⌘Q).
2. Remplace le dossier `flbrowser` par celui de `fl-project-browser-v17.zip`.
3. Relance avec `lancer.command`. Après le scan, le fichier
   `OneDrive-Personnel/SIROB Browser/catalogue.json` apparaît (laisse OneDrive le synchroniser).

### Étape 2 : mettre l'app en ligne sur GitHub Pages

1. Va sur https://github.com/new (connecté avec ton compte `borisfort33`).
2. **Repository name** : `sirob-mobile`. Laisse **Public** (seul le code est public : ni tes
   fichiers, ni ta connexion n'y sont). Clique **Create repository**.
3. Sur la page du dépôt vide, clique le lien **uploading an existing file**.
4. Dézippe `sirob-mobile.zip` sur le Mac, ouvre le dossier `sirob-mobile`, sélectionne **tout son
   contenu** (⌘A : `index.html`, `app.js`, le dossier `icons`…) et glisse-le dans la page.
   Clique **Commit changes**.
5. Dans le dépôt : **Settings > Pages**. Sous « Build and deployment », Source :
   **Deploy from a branch**, Branch : **main**, dossier **/ (root)**, puis **Save**.
6. Une à deux minutes plus tard, l'app est en ligne sur
   **https://borisfort33.github.io/sirob-mobile/**

### Étape 3 : autoriser cette adresse dans ton app Azure

C'est la même app Azure que pour les liens du Mac, on lui ajoute juste l'adresse de l'iPhone.

1. Ouvre https://entra.microsoft.com > **Applications > Inscriptions d'applications** >
   onglet **Toutes les applications** > **SIRØB Project Browser**.
2. Menu de gauche **Gérer > Authentification** > **+ Ajouter une plateforme** >
   **Application monopage** (« Single-page application »).
3. URI de redirection : `https://borisfort33.github.io/sirob-mobile/` (avec le `/` final),
   puis **Configurer**. Ne coche rien d'autre.

### Étape 4 : installer sur l'iPhone

1. Ouvre **Safari** (pas Chrome) sur https://borisfort33.github.io/sirob-mobile/
2. Bouton **Partager** (carré avec la flèche) > **Sur l'écran d'accueil** > **Ajouter**.
3. Ouvre **SIRØB** depuis l'écran d'accueil > **Se connecter à OneDrive** > ton compte
   Microsoft perso > **Accepter**.

### Pour les mails depuis ton adresse pro

L'iPhone envoie avec l'app Mail. Pour que le mail parte de ton adresse Hostinger, ajoute-la une
fois : **Réglages > Apps > Mail > Comptes > Ajouter un compte > Autre > Ajouter un compte Mail**
(serveur entrant `imap.hostinger.com`, sortant `smtp.hostinger.com`), puis
**Réglages > Apps > Mail > Compte par défaut** : ton adresse pro.

## Bon à savoir

- **Reconnexion** : Microsoft limite la connexion des web apps à 24 h. Si tu ne l'as pas ouverte
  depuis plus d'un jour, l'app affiche « Se connecter à OneDrive » : une touche, et Microsoft te
  reconnaît en général sans redemander le mot de passe.
- **Catalogue à jour** : il est réécrit par le Mac à chaque scan ou nouveau bounce détecté. La date
  du dernier envoi par le Mac est affichée sous la barre de recherche.
- **Écouter avant d'envoyer** : ▶ lit le fichier en streaming depuis OneDrive.
- **Mise à jour de l'app** : dans le dépôt GitHub, Add file > Upload files, glisse les fichiers modifiés, Commit changes.
  L'iPhone prend la nouvelle version à la prochaine ouverture (ferme l'app et rouvre-la).
- **Emplacement du catalogue** : modifiable dans ⚙︎ si tu déplaces le dossier.

## SIRØB Planning

Une deuxième app vit dans le dossier [`planning`](planning/README.md) : ta journée au même endroit
(rendez-vous Google Agenda, tes tâches, deadlines et morceaux en attente), à l'adresse
**https://borisfort33.github.io/sirob-mobile/planning/**. Installation : voir
[planning/README.md](planning/README.md).

## 💡 Idées de prod (1.6)

Onglet « Idées de prod » (façon Notes) : liste, ✎ nouvelle idée, « De la part de », 📌, ✓ faite, 🗑.
Les liens YouTube, SoundCloud, Spotify, Deezer, Apple Music se jouent dans l'app ; un lien Shazam
est retrouvé sur Apple Music (extrait 30 s). Synchro avec l'app du Mac par
`OneDrive/SIROB Browser/idees.json` (fusion idée par idée, eTag / If-Match).
