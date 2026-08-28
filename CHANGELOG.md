# Liste des changements

## [Unreleased]

### Correctifs

* Si la valeur originale d'une date est un nombre entre 1800 & 2050, est désormais considéré comme une année et se transforme en le premier janvier de celle-ci. Ceci implique que toute date comprise entre le 04 décembre 1904 & le 11 août 1905 sera mal transformé.
* Les dates au format `JJ/MM/YYYY` sont correctement interprétées
* Avoir un statut vide ne provoque plus qu'une seule erreur
* Optimisations techniques