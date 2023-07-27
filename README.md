# Ecarts vidanges
***
![logo_jr](https://github.com/Nicolo41/excel_comparator/assets/72193849/02109d10-a47e-44f4-a301-e4a39d5796ac)

Programme pour l'entreprise ***SA Jacques Remy & Fils***.

Programme `excel_comparator.py` permettant une comparaison rapide et efficace de grands tableaux Excel.
***
## Objectif
L'enjeu était de trouver un moyen de voir rapidement les écarts qui peuvent survenir entre deux tableurs Excel et d'avoir un executable pour des machines sans dépendances Python;

## Fonctionnalités
Traitement de fichiers Excel pour extraire et regrouper les données.
Comparaison de deux fichiers Excel pour identifier les différences entre eux.
Export des résultats dans un nouveau fichier Excel.
## Processus
- L'utilisateur doit sélectionner un fichier Excel contenant les données des livraisons. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des livraisons par client.

- L'utilisateur doit sélectionner un fichier Excel contenant les données des vidanges. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des vidanges par client.

- Lorsque l'utilisateur appuie sur le bouton 'Comparer les fichiers générés', le script va comparer les deux fichiers générés et créer un nouveau fichier Excel contenant les différences entre les deux. 

- Le code est fait pour que chaque fichier généré soit enregistré dans le dossier des téléchargements à la date de sa création.
### Détail processus
#### Fonction 1 (traiter_chauffeur) :
- Elle extrait les colonnes des vidanges du DataFrame df_livraisons.

- Elle crée un dictionnaire appelé clients_vidanges pour stocker les informations de chaque client, y compris les quantités pour chaque type de vidange.

- Pour chaque ligne du DataFrame df_livraisons, la fonction identifie le client, puis parcourt les colonnes des vidanges pour récupérer les quantités associées à chaque type de vidange pour ce client.

- Elle stocke les informations dans le dictionnaire clients_vidanges en utilisant le client comme clé principale et en ajoutant les quantités de vidanges associées à chaque type de vidange comme valeurs dans des sous-dictionnaires.

- Si le même type de vidange pour le même client apparaît dans plusieurs lignes du fichier de livraisons, la fonction ajoute les quantités pour ce type de vidange au sous-dictionnaire correspondant du dictionnaire clients_vidanges.

La fonction regroupe les données de vidanges pour chaque client et chaque type de vidange séparément dans le dictionnaire clients_vidanges. Ainsi, à la fin du traitement, clients_vidanges contiendra toutes les informations nécessaires pour chaque client, y compris les quantités pour chaque type de vidange.
#### Fonction 2 


## Installation
Clonez le dépôt : `git clone https://github.com/Nicolo41/excel_comparator.git`

Installez les dépendances : `pip install -r requirements.txt`

| Python 3.11.x   | :white_check_mark: |
| ----------------| ------------------ |
| pandas==2.0.3   | :white_check_mark: |
| collections     | :white_check_mark: |
| tabulate==0.9.0 | :white_check_mark: |
| PIL             | :x:                |

Choix de ne pas utiliser la lib PIL, Pillow car non exempte de failles de sécurité -> gestion des images directement avec Tkinter -> redimentionnement manuel (Gimp)

## Utilisation
Exécutez le programme en utilisant Python : `python excel_comparator.py`


## Environnement virtuel
Je voulais faire en sorte que le programme soit executable depuis n'importe quelle machine.

Création d'un environnement virtuel PyInstaller -> pas de dépendance à avoir pour executer le programme.

Créer un exécutable à partir du script en utilisant Pyinstaller : `pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py`

## Auteur
*Developpé par :* ***BROAGE Nicolas*** */ 07-2023.*


*https://www.linkedin.com/in/nicolas-broage/*
