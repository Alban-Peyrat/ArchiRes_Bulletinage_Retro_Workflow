# Workflow du bulletinage rétrospectif dans ArchiRès

[![Active Development](https://img.shields.io/badge/Maintenance%20Level-Actively%20Developed-brightgreen.svg)](https://gist.github.com/cheerfulstoic/d107229326a01ff0f333a1d3476e068d)

_Remplace le repo déprécié_ ArchiRes_Prepare_Files_Subscription_Import _qui contenait la version Python de l'utilitaire et des traces de version sur OpenRefine_

Workflow permettant de préparer les fichiers à transmettre au prestataire pour le bulletinage rétrospectif des périodiques.

## Traitement d'un titre

Soit :

* Alice : reponsable du chantier de bulletinage rétrospectif
* Bob : administration du SID

Circuit :

1. Un référent de la commission périodique prépare le fichier Google Sheets pour un titre avec une feuille modèle
1. Chaque école qui possède le titre duplique la feuille modèle et la remplie pour son école
1. Une fois le fichier remplie, Alice télécharge le fichier et le dépose dans le SharePoint dédié
1. [Power Automate détecte le dépôt, exécute le script de modification](./PowerAutomate/flux_principal.md) et notifie par mail Alice que le fichier a été traité
    * Le script de modification duplique le fichier originellement déposé
    * En cas d'échec de la modification, Alice & Bob reçoivent un mail 
1. Alice analyse le résultat du script disponible dans SharePoint :
    * Les `INFO` sont ignorables
    * Les `WARNING` doivent être lus mais ne nécessitent aucune action sauf si le traitement effectué par le script est erroné
    * Les `ERROR` doivent être traitées dans le ficheir original Google Sheets par Alice ou l'équipe de la bibliothèque concernée
    * _Retourner à l'étape de dépôt du fichier dans SharePoint jusqu'à ce que le script de renvoie plus d'`ERROR`_
1. Alice notifie Bob que le fichier est prêt
1. Bob télécharge le résultat du script
1. Bob exécute les opérations OpenRefine
    * Si des erreurs apparaissent, notifier Alice et repartir de l'étape du dépot dans SharePoint
1. Bob exporte [le fichier dans le format demandé](#spécifications-du-fichier-attendu-par-le-prestataire) et le transmet au prestataire

## Spécifications du fichier attendu par le prestataire

Fichier :

* Format : CSV
* Encodage : UTF-8
* Séparateur : `;`
* **Sans encapsulation des chaînes de caractères**

Colonnes (obligatoirement remplies sauf si mention du contraire) :

* _branchcode_ : contient le code d'une bibliothèque Koha
* _biblionumber_ : contient le biblionumber de la notice du périodique, donc le même au sein d'un même fichier
* *no_abonnement_koha* : facultatif, contient le numéro de l'abonnement de l'école. Si vide, un nouvel abonnement sera créé
* _numero_ : forme finale du numéro (ex : `2014 n°20`), doit être unique au couple biblionumber + branchcode
* *date_parution* : date de parution du fascicule, au format YYYY-MM-DD
* *date_reception* : date de réception du fascicule, au format YYYY-MM-DD
* *statut_arrive_manquant* : si le fascicule est arrivé ou non, prend les values : _manquant_, _arrivé_ ou _arrivée_ (accent & casse ignorés)
* *Localisation* (avec la majuscule) : facultatif, **libellé** de la localisation du fascicule dans Koha
* *Cote* (avec la majuscule) : facultatif, cote du fascicule

Notes sur le script du prestataire :

* S'il existe déjà dans la base un fascicule de même biblionumber, même identifiant d'abonnement, même numéro, même statut, le fascicule n'est pas créé
* (?) Une seule localisation et cote par couple branchcode + biblionumber