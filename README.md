# ecarts_vidanges

Programme `excel.comparator.py` permettant une comparaison rapide et efficace de grands tableaux Excel.

L'enjeu était de trouver un moyen de voir rapidement les ecarts qui peuvent survenirs entre deux tableurs.

Le programme doit être executable depuis n'importe quelle machine -> Pyinstaller créé un environnenement virtuel et permet d'avoir le script, les images et les dépendances dans un .exe .

Libs : requirements.txt

pandas==1.3.3

tkinter

tabulate==0.8.9

Pour installer les libs : `pip install -r requirements.txt`

Requirements :

`pip install pandas`

`pip install tkinter`

`pip install tabulate` (optional)

`pip install collections`

`pip install subprocess`

`pip freeze > requirements.txt`

Environnement virtuel :

`pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py`
