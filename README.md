# ecarts_vidanges

Programme ```excel.comparator.py``` permettant une comparaison rapide et efficace de grands tableaux Excel.
Programme `excel.comparator.py` permettant une comparaison rapide et efficace de grands tableaux Excel.

L'enjeu était de trouver un moyen de voir rapidement les ecarts qui peuvent survenirs entre deux tableurs.

Le programme doit être executable depuis n'importe quelle machine -> Pyinstaller créé un environnenement virtuel et permet d'avoir le script, les images et les dépendances dans un .exe .


Libs :  requirements.txt
Libs : requirements.txt

pandas==1.3.3

tkinter

tabulate==0.8.9

Pillow==8.3.2


Requirements : 
Pour installer les libs : `pip install -r requirements.txt`

``` pip install pandas ```
Requirements :

``` pip install tkinter ```
`pip install pandas`

``` pip install tabulate ``` (optional)
`pip install tkinter`

``` pip install collections ```
`pip install tabulate` (optional)

``` pip install subprocess ```
`pip install collections`

``` pip install Pillow ```
`pip install subprocess`

`pip freeze > requirements.txt`

Environnement virtuel :

```pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py``` 
`pyinstaller --onefile --add-data "requirements.txt;." excel_comparator.py`
