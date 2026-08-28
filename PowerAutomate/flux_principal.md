# Flux principal du workflow du bulletinage rétrospectif dans ArchiRès

Ce flux est utilisé par les responsables du chantier côté métiers pour effectuer des vérifications & corrections en toute autonomie.
Une fois que les erreurs sont corrigées, l'administration du SID est notifiée pour:

* Extraire le fichier pré-généré
* Effectuer les vérifications finales via OpenRefine
* Générer le fichier final
* Transmettre au prestataire le fichier final

## Script de transformation du fichier Excel

### Gestion du script

Le script ([*bulletinage_retro_script.js*](./bulletinage_retro_script.js)) est un Office Script (donc Typescript), maintenu dans un fichier Javascript qui est transformé en Office Scripts via [un flux Power Automate de "conversion"](./js_to_osts.md).

**Ne pas toucher à :**

``` Typescript 
main(workbook: ExcelScript.Workbook)
```

Le reste peut être géré sans problème en Javascript pour éviter qu'IntelliSense signale des dizaines d'erreurs.

Pour mettre à jour la liste des localisations dans Koha, utiliser la requête SQL ci-dessous et remplacer la déclaration de la variable `kohaLocations` dans le script :

``` SQL
SELECT CONCAT(
    "const kohaLocations = [",
    GROUP_CONCAT(CONCAT('"', lib, '"') SEPARATOR ','),
    "];"
) AS typescript_array
FROM authorised_values
WHERE category = "LOC"
```

### Traitements effectués par le script

1. Fusion des feuilles de chaque école en une feuille unique
1. Signale une erreur si le nombre de feuille traité n'est pas pertinent
1. Si un biblionumber représente au moins 80% des lignes, corrige les lignes qui ne le possèdent pas
1. Si le branchcode d'une ligne est différent du nom de la feuille de son école, le remplace par le nom de sa feuille
1. Normalise le statut (suppression des diacritiques, suppression des espaces en début et fin, passage en minuscule)
1. Signale une erreur si le statut d'une ligne n'est pas valide
1. Supprime la localisation si une ligne est signalée comme manquante
1. Si une localisation n'est pas valide, essaye de la corriger. S'il n'y arrive pas, signale une erreur
1. Formatte correctement les dates et signale une erreur si certaines ne peuvent être interprétées :
    * Si la valeur originale d'une date est un nombre entre 1800 & 2050, est considéré comme une année et se transforme en le premier janvier de celle-ci. Ceci implique que toute date comprise entre le 04 décembre 1904 & le 11 août 1905 sera mal transformé
1. Signale une erreur si certaines informations obligatoires sont absentes

### Gestion du fichier de localisations

* Exécuter dans Koha un rapport SQL avec la requête : `SELECT lib FROM authorised_values where category="LOC"`
* Exporter au format CSV le fichier sous le nom *koha_locations.csv*
* Le déposer dans dans le même dossier que le Office Script

## Étapes du flux dans Power Automate

* 1) Déclencheur : _SharePoint : When a file is created or modified (properties only)_ :
  * _Site Address_ : SharePoint du SID ArchiRès
  * _Library Name_ : _Documents_
  * _Advanced parameters → Folder_ : le dossier dédié aux fichiers originaux
* 2) _Initialize variable_ :
  * _Name_ : _MailOutput_
  * _Type_ : _String_
  * _Value_ : laisser vide
* 2) _Condition_ :
  * Opérateur : _Or_
  * _Nom de fichier avec l'extension_ → _ends with_ → _.ods_
  * _Nom de fichier avec l'extension_ → _ends with_ → _.xlsx_

Dans _True_ :

* 3) _SharePoint : Get file content using path_ :
  * _Note: pour je ne sais quelle raison, la récupération via ID renvoyait une erreur 404 donc bon_
  * _Site Address_ : SharePoint du SID ArchiRès
  * _File Path_ : `concat('/Documents partages/Bulletinage_Retro/Originaux/',triggerBody()?['{FilenameWithExtension}'])` 
    * `triggerBody()?['{FilenameWithExtension}']` = _Nom de fichier avec l'extension_
* 4) _SharePoint : Create File_ :
  * _Site Address_ : SharePoint du SID ArchiRès
  * _Folder Path_ : le dossier dédié aux fichiers de sortie
  * _File Name_ : _Nom de fichier avec l'extension_
  * _File Content_ : _File content_ de l'étape précédente
* 5) _Excel Online (Business) : Run script from SharePoint library_ :
  * _Workbook Location_ : SharePoint du SID ArchiRès
  * _Workbook Library_ : _Documents_
  * _Workbook_ : `concat('/Bulletinage_Retro/Resultat/',triggerBody()?['{FilenameWithExtension}'])`
    * `triggerBody()?['{FilenameWithExtension}']` = _Nom de fichier avec l'extension_ de la première étape
  * _Script Location_ : _OneDrive for Business_
  * _Script Library_ : _OneDrive_
  * _Script_ : sélectionner le fichier _.osts_ du script
* X) _Set variable_ :
  * _Name_ : _MailOutput_
  * _Value_ : `concat('Fichier ', triggerBody()?['{FilenameWithExtension}'],' correctement traité et disponible dans Bulletinage_Retro/Resultat.')`
    * `triggerBody()?['{FilenameWithExtension}']` = _Nom de fichier avec l'extension_ de la première étape
  * _Settings_ → _Run after_ : _Run script from SharePoint library_ seulement en cas de succès
* X) _Set variable_ :
  * _Name_ : _MailOutput_
  * _Value_ : `concat('Le fichier ', triggerBody()?['{FilenameWithExtension}'],' n a pas pu être traité.')`
    * `triggerBody()?['{FilenameWithExtension}']` = _Nom de fichier avec l'extension_ de la première étape
  * _Settings_ → _Run after_ : _Run script from SharePoint library_ sauf en cas de succès
* X) _Office 365 Outlook : Send an email (V2)_ :
  * _To_ : destinaires
  * _Subject_ : `[SID ArchiRès PA] Bulletinage rétrospectif : traitement du fichier` + _Nom de fichier avec l'extension_ de la première étape
  * _Body_ : variable _Mailoutput_
  * _Settings_ : dans tous les cas après les 2 set variable _mailOutput_

## Ressources

* [Office Scripts API reference / Microsoft](https://learn.microsoft.com/en-us/javascript/api/office-scripts/overview?view=office-scripts)
* [Office Script pour Excel : fusionner les feuilles / Bastien Perez](https://github.com/bastienperez/office-scripts-excel)
* [Automate Saving Excel Attachments to the SharePoint folder or the Blob storage / Ravi Kumar](https://medium.com/@ravikumar10593/automate-saving-excel-attachments-to-the-sharepoint-folder-or-the-blob-storage-22228b55322b)
* [Convert JS files to OSTS files with Power Automate / AymKdn](https://stackoverflow.com/questions/61477389/editing-office-scripts-in-visual-studio-code#answer-79816457)