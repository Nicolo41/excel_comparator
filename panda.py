import pandas as pd

# Importer les deux tableaux
table1 = pd.read_csv("Export_Livraisons - 20230713.csv")
table2 = pd.read_excel("Export vidanges Descartes 01-06 au 13-07.xlsx")

# Comparer les deux tableaux
diff = table1.compare(table2, indicator='difference')

# Imprimer les différences
print(diff)
