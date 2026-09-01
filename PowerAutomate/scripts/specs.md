# Données sur le script Power Automate

Nom des feuilles :

* `Fusionnee` : feuille avec les données fusionnées
* `Erreurs` : feuille avec les données d'erreurs
* `Synthese` : feuille avec le tableau crosié dynamique
* `PowerAutomateHidden` : feuille contenant en A1 le biblionumber principal

Colonnes après la 1re étape :

``` JS
const colIndex = {
  "branchcode":0,
  "biblionumber":1,
  "no_abonnement_koha":2,
  "numero":3,
  "date_parution":4,
  "date_reception":5,
  "statut_arrive_manquant":6,
  "Localisation":7,
  "Cote":8,
  "sheetName":9
}
```