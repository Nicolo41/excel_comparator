# Ecarts vidanges
***
![logo_jr](https://github.com/Nicolo41/excel_comparator/assets/72193849/02109d10-a47e-44f4-a301-e4a39d5796ac)


Programme `excel_comparator.py` permettant une comparaison rapide et efficace de grands tableaux Excel.
***
## Objectif
L'enjeu était de trouver un moyen de voir rapidement les écarts qui peuvent survenir entre deux tableurs Excel et d'en avoir un executable pour des machines sans dépendances Python.

## Fonctionnalités
Traitement de fichiers Excel pour extraire et regrouper les données.
Comparaison de deux fichiers Excel pour identifier les différences entre eux.
Export des résultats dans un nouveau fichier Excel.
## Installation
Clonez le dépôt : `git clone https://github.com/Nicolo41/excel_comparator.git`

Installez les dépendances : `pip install -r requirements.txt`

| Python 3.11.x   | :white_check_mark: |
| ----------------| ------------------ |
| pandas==1.3.3   | :white_check_mark: |
| tk==0.1.0       | :white_check_mark: |
| tabulate==0.9.0 | :white_check_mark: |
| PIL             | :x:                |
| openpyxl==3.0.9 | :white_check_mark: |
| colorama==0.4.4 | :white_check_mark: |

Choix de ne pas utiliser la lib PIL, Pillow car il y a des alertes de failles de sécurité -> gestion des images directement avec Tkinter -> redimentionnement manuel (Gimp)

## Utilisation
Exécutez le programme en utilisant Python : `python excel_comparator.py`


## Environnement virtuel
Pour que le programme soit executable depuis n'importe quelle machine.

Création d'un environnement virtuel PyInstaller -> pas de dépendance à avoir pour executer le programme ni d'installation python.

Créer un exécutable à partir du script en utilisant Pyinstaller : `pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py`
## Processus
- L'utilisateur doit sélectionner un fichier Excel contenant les données encodées des livreurs. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des livraisons par client.

- L'utilisateur doit sélectionner un fichier Excel contenant les données de Descartes. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des vidanges par client, ce fichier va être réorganisé pour la comparaison.

- Lorsque l'utilisateur appuie sur le bouton 'Comparer les fichiers traités', le script va comparer les deux fichiers générés lors des deux précédentes étapes et créer un nouveau fichier Excel contenant les différences entre les deux. S'il n'y a pas de différence, un message l'indique et le fichier des écarts ne sera pas généré.

- Le code est fait pour que chaque fichier généré soit enregistré dans le dossier des téléchargements à la date de sa création.

- Une gestion des erreurs est fonctionnelle, des messages d'erreurs apparaîtrons s'il y en a. De plus, un système de logs a également été mis en place pour debug et info.


## Auteur
*Developpé par :* ***BROAGE Nicolas*** */ 07-2023.*


*https://www.linkedin.com/in/nicolas-broage/*
