# script-alternance

Script Python qui envoie des e-mails de candidature en alternance, un par un, avec un message adapté au domaine de chaque entreprise. Il limite le rythme d'envoi, évite les doublons et garde un journal de tout ce qui a été envoyé.

Je l'ai écrit pour ma recherche d'alternance en cybersécurité : démarcher directement les entreprises qui m'intéressent plutôt que d'attendre une offre sur un site d'emploi, sans passer des heures à copier-coller les mêmes mails.

## Fonctionnalités

- **Envoi par SMTP** (Gmail recommandé, Outlook / Microsoft 365 possible) avec chiffrement STARTTLS.
- **Modèles par domaine** : `cyber`, `ia`, `web`, `salarie`, chacun dans `templates/`.
- **Personnalisation automatique** : prénom, entreprise, liens vers le CV, GitHub et LinkedIn.
- **Anti-doublons** : une adresse déjà contactée avec succès n'est jamais relancée.
- **Limitation du rythme** : pause entre deux mails et plafond d'envois par jour.
- **Mode simulation** (`--dry-run`) : affiche le premier mail sans rien envoyer.
- **Journal** des envois et des erreurs dans `envoyes.txt`.

## Prérequis

- Python 3.9 ou plus récent
- Aucune dépendance externe : le script n'utilise que la bibliothèque standard

## Installation

```bash
git clone https://github.com/ilansah/script-alternance.git
cd script-alternance
cp .env.example .env
```

Ouvre ensuite `.env` et remplis tes informations (voir ci-dessous).

## Configuration (`.env`)

| Variable | Rôle |
|---|---|
| `OUTLOOK_EMAIL` | Adresse d'envoi (nom historique, fonctionne aussi avec Gmail) |
| `OUTLOOK_PASSWORD` | Mot de passe d'application (voir plus bas), jamais ton mot de passe habituel |
| `SMTP_HOST` / `SMTP_PORT` | Serveur SMTP, par défaut `smtp.gmail.com` / `587` |
| `REPLY_TO` | Adresse où arriveront les réponses (facultatif) |
| `NOM_COMPLET`, `TELEPHONE` | Informations de signature |
| `CV_URL`, `GITHUB`, `LINKEDIN`, `SCRIPT_URL` | Liens insérés dans les mails |
| `PAUSE_SECONDES` | Pause entre deux envois (90 par défaut) |
| `MAX_PAR_JOUR` | Nombre maximum d'envois par exécution (25 par défaut) |

`CV_URL`, `OUTLOOK_EMAIL` et `OUTLOOK_PASSWORD` sont obligatoires pour un envoi réel.

> Le fichier `.env` contient des secrets. Il est ignoré par Git (`.gitignore`) : ne le publie jamais.

### Utiliser Gmail (recommandé)

1. Active la validation en deux étapes sur ton compte Google.
2. Crée un **mot de passe d'application** dans les paramètres de sécurité du compte.
3. Mets-le dans `OUTLOOK_PASSWORD`, avec `SMTP_HOST=smtp.gmail.com` et `SMTP_PORT=587`.

Un compte Outlook ou Microsoft 365 d'école a souvent l'authentification SMTP désactivée par l'administrateur. Dans ce cas, le script affiche une erreur d'authentification et il vaut mieux passer par un Gmail personnel, avec `REPLY_TO` sur l'adresse de l'école si besoin.

## Préparer la liste de contacts

Crée un fichier `contacts.csv` à la racine, séparé par des **points-virgules** :

```csv
prenom;nom;email;entreprise;domaine;source_url
;;recrutement@exemple.fr;Exemple SAS;cyber;https://exemple.fr/recrutement/
Camille;Martin;c.martin@exemple.fr;Autre SA;web;https://exemple.fr/equipe/
```

- `domaine` doit valoir `cyber`, `ia`, `web` ou `salarie`. Les autres lignes sont ignorées.
- `prenom` peut rester vide : le mail commence alors par « Bonjour, ».
- `source_url` indique d'où vient l'adresse, pour garder une trace.

Ce fichier contient des données personnelles et n'est **pas** destiné à être publié dans le dépôt.

## Modèles de mail

Chaque fichier de `templates/` commence par une ligne `Sujet: ...`, suivie du corps du message. Variables disponibles :

`{salutation}`, `{prenom}`, `{entreprise}`, `{nom_complet}`, `{tel}`, `{cv_url}`, `{linkedin}`, `{github}`, `{script_url}`

Relis et adapte les modèles à ta situation (ville, date de début, ton) avant le premier envoi.

## Utilisation

Commence toujours par une simulation :

```bash
python envoyer.py --dry-run
```

Elle affiche le nombre de contacts retenus et l'aperçu du premier mail, sans rien envoyer. Quand le résultat te convient :

```bash
python envoyer.py
```

Le script se connecte, envoie les mails un par un en respectant la pause, puis écrit chaque résultat dans `envoyes.txt` au format `date;adresse;entreprise;statut;message`. Au prochain lancement, les adresses déjà au statut `ok` sont sautées.

## Fonctionnement

```
contacts.csv ──► filtre (domaine valide, adresse non contactée)
                   │
                   ▼
           limite MAX_PAR_JOUR
                   │
                   ▼
   modèle du domaine + personnalisation
                   │
                   ▼
        envoi SMTP (STARTTLS) ──► pause ──► suivant
                   │
                   ▼
              envoyes.txt
```

## Bonnes pratiques

- N'écris qu'à des **adresses professionnelles publiques** (page recrutement ou contact d'une entreprise), jamais à des adresses privées.
- Garde un volume raisonnable : plafond quotidien et pause longue limitent le risque d'être classé comme spam et de voir ton compte bloqué.
- Personnalise ton message. Un envoi ciblé et bien rédigé donne de meilleurs résultats qu'un envoi en masse.
- Réponds vite et poliment si quelqu'un te demande de ne plus le contacter, et retire son adresse de la liste.
- Ne publie jamais `.env`, `contacts.csv` ni `envoyes.txt`.

## Structure du dépôt

```
script-alternance/
├── envoyer.py        # Script principal
├── templates/        # Modèles de mail : cyber, ia, web, salarie
├── .env.example      # Exemple de configuration (sans secrets)
├── contacts.csv      # Liste des contacts (à créer, non publiée)
└── envoyes.txt       # Journal des envois (généré)
```

## Limites connues

- Pas d'envoi de pièce jointe : le CV est partagé par lien (`CV_URL`).
- Pas de relance automatique.
- Les mails sont envoyés en texte brut.

## Licence

À préciser (par exemple MIT).
