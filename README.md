
# Powershell sauvegarde Mercator par API REST 

Script PowerShell permettant la sauvegarde complète d'une instance Mercator via l'API REST avec génération d'un référentiel de crise autonome consultable hors ligne.
https://github.com/sourcentis/mercator

Le script exporte automatiquement l'ensemble des objets de cartographie disponibles via l'API Mercator :

* Applications 
* Processus 
* Flux 
* Serveurs logiques 
* Serveurs physiques 
* Bases de données 
* Réseaux 
* VLANs 
* Baies 
* Sites 
* Contrôles de sécurité 
* Sauvegardes Documents
* toutes les autres ressources configurées dans Mercator.

Chaque ressource est exportée dans trois formats :

* JSON 
* CSV 
* HTML autonome 
* Exports des rapports Word/Excel Mercator 

Le script télécharge automatiquement les rapports Word générés par Mercator.

Mercator
│ ├── 2026-09-11_12-00-00 │ │ │ ├── JSON │ ├── CSV │ ├── HTML │ ├── WORD │ │ │ ├── index.html │ ├── synthese.csv │ ├── controle.json │ └── BackupMercator.log │ └── Mercator_2026-09-11_12-00-00.zip

# Prérequis

* Version PowerShell Compatible : Windows PowerShell 5.1 PowerShell 7+

* API activée de Laravel Passport configuré d'un compte API actif



## Création du compte API

Créer un utilisateur dédié pour la sauvegarde API dans mercator et lui attribuer les droits de lecture sur toutes les ressources nécessaires
#### backup-mercator


#### Création du fichier d'identifiants sécurisé : 

Le script utilise un stockage sécurisé DPAPI Windows.

Créer une seule fois le fichier dans powershell :
```http
  $Path = "C:\Mercator\mercator_cred.xml"
  Get-Credential | Export-Clixml -Path $Path
```
Une fenêtre d'authentification apparaît :
```http
  Utilisateur : backup-mercator Mot de passe : ********
```

Le fichier créé : C:\Emplacement_Mercator\mercator_cred.xml est chiffré par Windows et ne peut être relu que par l'utilisateur qui l'a créé, sur la machine qui l'a créé. 





## Auteur

- [@Gabinhocity](https://github.com/gabinhocity)

