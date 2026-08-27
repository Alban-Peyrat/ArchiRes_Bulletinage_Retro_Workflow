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
1. Formatte correctement les dates et signale une erreur si certaines ne peuvent être interprétées
1. Signale une erreur si certaines informations obligatoires sont absentes

### Gestion du fichier de localisations

* Exécuter dans Koha un rapport SQL avec la requête : `SELECT lib FROM authorised_values where category="LOC"`
* Exporter au format CSV le fichier sous le nom *koha_locations.csv*
* Le déposer dans dans le même dossier que le Office Script

## Étapes du flux dans Power Automate

_Mettre au propre quand le flux théorique est correctement défini_

```
Déclencheru :
When a nw mail arrive V3 on specific folder
Apply to each : Attachments
Condition : AttachmentName contains /xlsx
Get attachment : Message Is = Message Id ; Attachnment Id = attachment id
Crteate file : Folder path = /ArchiRes_Auto/received ; File Name = Get Attachment Name ; File content = Get aatttahcment content bytes
Copy File : File = CreateFile Id ; Destination = `concat('/ArchiRes_Auto/edited/', outputs('Create_file')?['body/DisplayName'])`
Run script from sharepoint lirbary:
* OneDrive for Bsuiness
* OneDrive
* `concat('/ArchiRes_Auto/edited/', outputs('Copy_file')?['body/Name'])`
* OneDrive for Bsuiness
* OneDrive
* /ArchiRes_Auto/concatenate-worksheets-into-one.osts
```

## Ressources

* [Office Scripts API reference / Microsoft](https://learn.microsoft.com/en-us/javascript/api/office-scripts/overview?view=office-scripts)
* [Office Script pour Excel : fusionner les feuilles / Bastien Perez](https://github.com/bastienperez/office-scripts-excel)
* [Automate Saving Excel Attachments to the SharePoint folder or the Blob storage / Ravi Kumar](https://medium.com/@ravikumar10593/automate-saving-excel-attachments-to-the-sharepoint-folder-or-the-blob-storage-22228b55322b)
* [Convert JS files to OSTS files with Power Automate / AymKdn](https://stackoverflow.com/questions/61477389/editing-office-scripts-in-visual-studio-code#answer-79816457)