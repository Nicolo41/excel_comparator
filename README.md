# Ecarts vidanges
***
![logo_jr](https://github.com/Nicolo41/excel_comparator/assets/72193849/02109d10-a47e-44f4-a301-e4a39d5796ac)

Programme pour l'entreprise Jacques Remy & fils.

Programme ```excel_comparator.py``` permettant une comparaison rapide et efficace de grands tableaux Excel.
***
## Objectif
L'enjeu était de trouver un moyen de voir rapidement les écarts qui peuvent survenir entre deux tableurs Excel.

## Fonctionnalités
Traitement de fichiers Excel pour extraire et regrouper les données.
Comparaison de deux fichiers Excel pour identifier les différences entre eux.
Export des résultats dans un nouveau fichier Excel.
## Installation
Clonez le dépôt : ```git clone https://github.com/Nicolo41/excel_comparator.git```

Installez les dépendances : ```pip install -r requirements.txt```

Choix de ne pas utiliser la lib PIL, Pillow car non exempte de failles de sécurité -> gestion des images directement avec Tkinter -> redim manuel
| Python 3.11.x   | :white_check_mark: |
| ----------------| ------------------ |
| pandas==2.0.3   | :white_check_mark: |
| collections     | :white_check_mark: |
| tabulate==0.9.0 | :white_check_mark: |
| PIL             | :x:                |
## Utilisation
Exécutez le programme en utilisant Python : ```python excel_comparator.py```


## Environnement virtuel
Créer un exécutable à partir du script en utilisant Pyinstaller : ```pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py```

## Auteur
Developpé par : BROAGE Nicolas employé pour 2 mois.
https://www.linkedin.com/in/nicolas-broage/ 
