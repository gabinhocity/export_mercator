Mercator Backup
https://github.com/sourcentis/mercator

Script PowerShell permettant la sauvegarde complète d'une instance Mercator via l'API REST avec génération d'un référentiel de crise autonome consultable hors ligne.

Fonctionnalités
Export des données Mercator

Le script exporte automatiquement l'ensemble des objets de cartographie disponibles via l'API Mercator :

Applications
Processus
Flux
Serveurs logiques
Serveurs physiques
Bases de données
Réseaux
VLANs
Baies
Sites
Contrôles de sécurité
Sauvegardes
Documents
Et toutes les autres ressources configurées dans Mercator.

Chaque ressource est exportée dans trois formats :

JSON
CSV
HTML autonome

Export des rapports Word Mercator

Le script télécharge automatiquement les rapports Word générés par Mercator.

Mercator\
│
├── 2026-09-11_12-00-00
│   │
│   ├── JSON
│   ├── CSV
│   ├── HTML
│   ├── WORD
│   │
│   ├── index.html
│   ├── synthese.csv
│   ├── controle.json
│   └── BackupMercator.log
│
└── Mercator_2026-09-11_12-00-00.zip


Pré-requis
Version PowerShell

Compatible :

Windows PowerShell 5.1
PowerShell 7+


Mercator doit disposer :

d'une API activée
de Laravel Passport configuré
d'un compte API actif

Documentation officielle API Mercator : Documentation API Mercator.

Création du compte API

Créer un utilisateur dédié :

backup-mercator

Attribuer les droits de lecture sur toutes les ressources nécessaires.

Création du fichier d'identifiants sécurisé

Le script utilise un stockage sécurisé DPAPI Windows.

Créer une seule fois le fichier :

$Path = "C:\ROBOCOPY\Mercator\mercator_cred.xml"

Get-Credential | Export-Clixml -Path $Path


Une fenêtre d'authentification apparaît :


Utilisateur : backup-mercator
Mot de passe : ********


Le fichier créé :
C:\ROBOCOPY\Mercator\mercator_cred.xml


est chiffré par Windows et ne peut être relu que par :

l'utilisateur qui l'a créé ;
sur la machine qui l'a créé.
