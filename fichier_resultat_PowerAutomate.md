# Contenu du fichier produit par Power Automate

Le fichier est une copie du fichier original avec 3 nouvelles feuilles :

* _Synthese_ : contient une synthèse des erreurs par type d'erreur pour chaque école, à titre informatif
* _Erreurs_ : contient [la liste des erreurs](#liste-des-erreurs-possibles)
* _Fusionnee_ : contient toutes les données fusionnées. C'est la feuille qui est ensuite utilisée pour générer le fichier final

Les erreurs sont classées selon 3 niveaux de gravité :

* `[1] ERROR` : le problème doit être traité
* `[2] WARNING` : le problème doit être lu mais ne nécessite aucune action sauf si la résolution effectuée par le script est mauvaise
* `[3] INFO` : le problème peut être ignoré

## Liste des erreurs possibles

`[1] ERROR` :

* _Impossible d'interpréter_ `date_reception` ou `date_parution` : le script n'a pas réussi à produire une date au format YYYY-MM-DD pour ce fascicule
* _Informations vitales absentes_ : certaines informations obligatoires sont absentes pour ce fascicule
* _Localisation inexistante_ : le fasicule contient une localisation, cependant elle n'existe pas dans Koha et le script n'a pas trouvé celle qui correspondrait
* _Nombre d'écoles traitées_ : le nombre d'écoles traitées semble anormale, il est probable qu'une école ait été ignorée
* _Statut arrivé manquant erronné_ : le statut du fascicule n'a pas pu être interprété comme _arrive_ ou _manquant_

`[2] WARNING` :

* _Biblionumber différent du biblionumber principal_ : au moins 80% des fascicules contiennent le même biblionumber, cependant cette ligne contenait un biblionumber différent : il a été corrigé
* _Correction de la localisation_ : le fascicule contient une localisation mais elle n'était pas correctement écrite : elle a été corrigée
* _Branchcode différent du nom de la feuille_ : le branchcode du fascicule ne correspond pas au nom de la feuille : il a été remplacé par le nom de la feuille

`[3] INFO` :

* _Nombre d'écoles traitées_ : toute les écoles ont a priori été traitées
* _Manquant, suppression de sa localisation_ : fascicule manquant, la localisation a été supprimée