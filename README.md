# Ecarts vidanges
***
![logo_jr](https://github.com/Nicolo41/excel_comparator/assets/72193849/4d269ab8-575b-4c21-9807-e826af440e03)


Programme `excel_comparator.py` permettant une comparaison rapide et efficace de grands tableaux Excel.
***
## Objectif
L'enjeu était de trouver un moyen de voir rapidement les écarts qui peuvent survenir entre deux tableurs Excel et d'en avoir un executable pour des machines sans dépendances Python.

## Fonctionnalités
Traitement de fichiers Excel pour extraire et regrouper les données.
Comparaison de deux fichiers Excel pour identifier les différences entre eux.
Export des résultats dans un nouveau fichier Excel.
## Téléchargement du programme
Selectionnez "<> Code" en haut à droite puis -> "Download ZIP".

Décompressez le dossier et placez le où vous le souhaitez, lancer le programme avec le fichier `excel_comparator.exe` dans le dossier `dist`.
(Vous pouvez créer un raccourci sur le bureau)

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

Créer un exécutable à partir du script en utilisant Pyinstaller : `pyinstaller --onefile --add-data "img/*.png:img/" excel_comparator.py`
## Processus
- L'utilisateur doit sélectionner un fichier Excel contenant les données encodées des livreurs. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des livraisons par client.

- L'utilisateur doit sélectionner un fichier Excel contenant les données de Descartes. Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des vidanges par client, ce fichier va être réorganisé pour la comparaison.

- Lorsque l'utilisateur appuie sur le bouton 'Comparer les fichiers traités', le script va comparer les deux fichiers générés lors des deux précédentes étapes et créer un nouveau fichier Excel contenant les différences entre les deux. S'il n'y a pas de différence, un message l'indique et le fichier des écarts ne sera pas généré.

- Le code est fait pour que chaque fichier généré soit enregistré dans le dossier des téléchargements à la date de sa création.

- Une gestion des erreurs est fonctionnelle, des messages d'erreurs apparaîtrons s'il y en a. De plus, un système de logs a également été mis en place pour debug et info.
## Générer ses fichiers pour la comparaison
Seuls les fichier en .xlsx sont prit en charge, la conversion CSV -> XLSX est simple à faire (ouvrir le fichier -> "enregistrer sous" -> dans la selection du format choisir "Classeur Excel (*.xlsx)")

### Obtenir le fichier de Odoo
Le fichier encodé par les chauffeurs (provenant de Odoo) doit avoir cette structure : 

['Client', 'Lignes de la commande/Article', 'Lignes de la commande/Quantité']

Pour générer le fichier d'Odoo :
* Se rendre dans le module "Cockpit des ventes"
* Dans favoris, appliquez le filtre "RETOURS VIDANGES DESCARTES"

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/c921d06a-7971-4ede-94a8-bb955a7d21b3)

* Appliquer un filtre pour selectionner les dates : ```Filtres -> Ajouter un filtre personnalisé``` et sélectionner comme suit :
![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/16d843e4-439d-472b-91d6-cd3d26701d6b)

Le mieux est de prendre sur un mois complet, évitez des plages trop grande. 

N.B : Certains écarts peuvent être présent car les dates d'encodage dans Odoo et dans Descartes ne sont pas les mêmes et se retrouver sur le mois d'après -> Nécessite une supervision.

* Une fois les filtres appliqués, il faut selectionner tous les SO en cliquant sur la case (entourée en rouge) puis sur "Tout sélectionner" (souligné en vert) :

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/a769c38c-3f77-47f5-b359-bad086e152cd)

* Un bouton ```Action``` est maintenant disponible, sélectionnez "Exporter" :

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/0c7924ed-b467-422a-b716-e90950fc92e6)

Ici vous pouvez choisir les champs à exporter dans un tableur Excel, il faut obligatoirement ces champs :

['Client', 'Lignes de la commande/Article', 'Lignes de la commande/Quantité']

Pour simplifier le processus, sélectionnez le modèle créé à cet effet (```Export Vidanges Descartes```) : 

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/2b22a2b8-5672-4123-b3ab-b4fad8dd7538)

* Cliquez sur "Exporter"

### Obtenir le fichier de Descartes
Le fichier Descartes doit avoir cette structure lors de son export : 

['Customer Name', 'Palette Euro NEW', 'Caisses vertes', 'VID-T', 'VID-S', 'Vidange F', 'FRIGO BOX', 'Palette Truval', 'Palette banane', 'Palette Plastique', 'Palette Pool']

* Sélectionnez "Rapports" à gauche dans Descartes puis ```Export_Livraisons``` :

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/75488bfe-ce9e-469c-b64f-478f2e4e416d)

* Sélectionnez vos dates et cliquez sur "Générer"

![image](https://github.com/Nicolo41/excel_comparator/assets/72193849/f47367fa-02e4-4738-b948-e4d01540acdd)

ATTENTION ! Ce fichier est automatiquement généré en ```.csv``` et nécessite donc une conversion en ```.xlsx```.

Pensez également à prendre des dates de comparaison cohérentes, par exemple, 2 fichiers du 01 au 30 juin, ne pas oublier que certains écarts apparents peuvent avoir été crédité à une date ultérieure ! L'ajout des dates aux écarts n'est pas sûr à 100%, le mieux est de vérifier.

Les 'PALETTE EU 11' et 'PALETTE EU 9' sont gérées, si de nouvelles vidanges s'ajoutent des erreurs peuvent être causées ! Dans ce cas il faudra modifier le code de l'application.

Si des modifications du code sont effectuées, il faudra télécharger à nouveau la solution pour pouvoir en bénéficier.
