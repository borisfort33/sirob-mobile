# SIRØB Planning

Ta journée au même endroit, jour par jour :

- **Rendez-vous** de ton Google Agenda (lecture seule) ;
- **Tes tâches** : tu les ajoutes, coches, déplaces à demain ou à un autre jour, avec une heure si
  tu veux. Une tâche pas faite reste en haut d'« Aujourd'hui » jusqu'à ce que tu la coches ;
- **Deadlines** des morceaux (celles du jour et celles en retard) ;
- **Morceaux en attente** : tous ceux qui ont un statut Idée, Active ou Mix (ou une deadline) dans
  l'app du Mac. Touche un morceau pour le planifier aujourd'hui, demain ou un autre jour.

Glisse à gauche ou à droite pour changer de jour, ou touche un jour dans la barre.

Comme SIRØB Mobile, c'est une web app hébergée sur GitHub Pages, sans serveur ni abonnement :

- les tâches sont enregistrées dans `OneDrive/SIROB Browser/planning.json`. L'iPhone écrit
  directement dans OneDrive : **le Mac n'a pas besoin d'être allumé**. Hors réseau, les tâches
  restent sur le téléphone et partent au retour du réseau ;
- les morceaux viennent du même `catalogue.json` que SIRØB Mobile (mis à jour par le Mac à
  chaque scan ; la dernière version reste affichée quand le Mac est éteint) ;
- l'agenda est lu directement chez Google.

Adresse : **https://borisfort33.github.io/sirob-mobile/planning/**

## Installation (une seule fois)

### Étape 1 : autoriser l'adresse dans ton app Azure (2 minutes)

Même chose que pour SIRØB Mobile, avec la nouvelle adresse.

1. Ouvre https://entra.microsoft.com > **Applications > Inscriptions d'applications** >
   onglet **Toutes les applications** > **SIRØB Project Browser**.
2. **Gérer > Authentification**. Dans la section **Application monopage** (là où il y a déjà
   l'adresse de SIRØB Mobile), clique **Ajouter un URI**.
3. Colle `https://borisfort33.github.io/sirob-mobile/planning/` (avec le `/` final), puis
   **Enregistrer**.

### Étape 2 : installer sur l'iPhone

1. Ouvre **Safari** sur https://borisfort33.github.io/sirob-mobile/planning/
2. **Partager > Sur l'écran d'accueil > Ajouter**.
3. Ouvre **Planning** > **Se connecter à OneDrive**.

L'app marche déjà : tâches et morceaux. Il reste l'agenda.

### Étape 3 : relier Google Agenda (environ 10 minutes)

Google demande un « identifiant client » pour qu'une app puisse lire ton agenda. C'est gratuit
et il reste privé : seul ton compte pourra s'en servir.

1. Va sur https://console.cloud.google.com (connecté avec `imsirob@gmail.com`), accepte les
   conditions si on te le demande.
2. En haut, menu des projets > **Nouveau projet** > nom `SIROB Planning` > **Créer**, puis
   sélectionne ce projet.
3. Active l'API : https://console.cloud.google.com/apis/library/calendar-json.googleapis.com >
   **Activer**.
4. Écran de consentement : menu **API et services > Écran de consentement OAuth** (ou
   « Google Auth Platform ») > **Commencer** :
   - nom de l'application `SIROB Planning`, e-mail d'assistance : ton adresse ;
   - audience : **Externe** ;
   - coordonnées : ton adresse ; accepte et **Créer**.
   - Dans **Audience** > **Utilisateurs test** > **Add users** : `imsirob@gmail.com` > Enregistrer.
     (Laisse l'app « en test » : pas besoin de validation par Google.)
5. Identifiant : **Clients** (ou API et services > Identifiants) > **Créer un client** :
   - type : **Application Web** ;
   - **Origines JavaScript autorisées** : `https://borisfort33.github.io`
   - **URI de redirection autorisés** : `https://borisfort33.github.io/sirob-mobile/planning/`
   - **Créer**, puis copie l'**ID client** (il finit par `.apps.googleusercontent.com`).
6. Dans le dépôt GitHub, ouvre `planning/config.js` > crayon ✏️ (Modifier) et colle l'ID entre
   les guillemets de `googleClientId: ""`. **Commit changes**.
7. Dans l'app (ferme-la et rouvre-la) : **Connecter Google Agenda** > ton compte > Google
   affiche « Google n'a pas validé cette application » : **Continuer** (c'est ton app) > coche
   l'accès à l'agenda > **Continuer**.

Dans ⚙︎, tu choisis quels agendas afficher (perso, jours fériés…).

### Étape 4 : les mails (résumé du matin)

Le bouton ✉︎ en haut de l'app affiche le **résumé du matin** écrit par Claude et tes **mails non
lus importants** depuis hier (touche un mail pour l'ouvrir dans Gmail). Quand le résumé du jour
est arrivé, une carte ☀️ apparaît en haut d'« Aujourd'hui ».

1. Active l'API Gmail dans le même projet Google Cloud :
   https://console.cloud.google.com/apis/library/gmail.googleapis.com > **Activer**.
2. Dans l'app : ✉︎ > **Autoriser Gmail** > coche l'accès en lecture à Gmail > **Continuer**.
   (Lecture seule : l'app ne peut ni envoyer, ni supprimer.)

Le résumé est un mail que Claude t'envoie chaque matin avec l'objet
« Résumé mails SIRØB · <date> » ; l'app le retrouve et l'affiche joliment.

## Bon à savoir

- **Google** : sa connexion dure 1 h. Ensuite l'app repasse chez Google toute seule, en une
  seconde, quand tu l'ouvres. Si Google veut une confirmation, un bouton « Connecter Google
  Agenda » s'affiche.
- **Microsoft** : comme SIRØB Mobile, reconnexion d'une touche si l'app n'a pas été ouverte
  depuis plus d'un jour.
- **Morceaux en attente** : renseigne le statut (Idée, Active, Mix, Done) et la deadline sur le
  Mac. Passer un morceau en Done le retire de la liste.
- **Mise à jour de l'app** : comme SIRØB Mobile (Add file > Upload files dans le dossier
  `planning`).
