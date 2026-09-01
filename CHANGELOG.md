# Liste des changements

## [Unreleased]

### Correctifs

* Les valeurs sont désormais récupérées comme elles sont affichées plutôt que par leur valeur interne (notamment pour que Excel arrête de convertir tout et n'importe quoi en date)
* Si la valeur originale d'une date est un nombre entre 1800 & 2050, est désormais considéré comme une année et se transforme en le premier janvier de celle-ci. Ceci implique que toute date comprise entre le 04 décembre 1904 & le 11 août 1905 sera mal transformé.
* Les dates au format `JJ/MM/YYYY` sont correctement interprétées
* Avoir un statut vide ne provoque plus qu'une seule erreur
* La détection du type de donnée ne se base plus sur la position de la colonne dans la feuille originale mais sur le nom de la colonne
* Détecte si des colonnes obligatoires ou facultatives sont manquantes
* Optimisations techniques dont divisions du script en pluseiurs parties pour éviter un TimeOut sur les fichiers les plus lourds