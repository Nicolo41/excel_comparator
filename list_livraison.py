import pandas as pd
from collections import defaultdict
from tabulate import tabulate

# Chemin du fichier Excel
fichier_livraisons = 'Export_Livraisons - 20230713.xlsx'

# Charger le fichier des livraisons en DataFrame pandas
df_livraisons = pd.read_excel(fichier_livraisons)

# Extraire les colonnes des vidanges (de D à N)
colonnes_vidanges = df_livraisons.columns[3:14]  # Adapté pour les colonnes D à N (indices 3 à 14 exclus)

# Créer un dictionnaire pour stocker les clients, leurs vidanges et les quantités correspondantes
clients_vidanges = defaultdict(dict)

# Parcourir chaque ligne du DataFrame des livraisons
for index, row in df_livraisons.iterrows():
    client = row['Customer Name']
    for vidange in colonnes_vidanges:
        quantite = row[vidange]
        if pd.notna(quantite):
            if vidange not in clients_vidanges[client]:
                clients_vidanges[client][vidange] = quantite
            else:
                clients_vidanges[client][vidange] += quantite

# Créer une liste de listes pour le tableau
table_data = []
for client, vidanges in clients_vidanges.items():
    row = [client] + [vidanges.get(vidange, 0) for vidange in colonnes_vidanges]
    table_data.append(row)

# Afficher le tableau
headers = ['Client'] + list(colonnes_vidanges)
print(tabulate(table_data, headers=headers, tablefmt='grid'))

# Convertir la liste de listes en DataFrame pandas
df_output = pd.DataFrame(table_data, columns=headers)

# Enregistrer le DataFrame dans un fichier Excel
nom_fichier_excel = 'output_livraisons.xlsx'
df_output.to_excel(nom_fichier_excel, index=False)

print("Le fichier Excel a été enregistré avec succès sous le nom:", nom_fichier_excel)
