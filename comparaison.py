import pandas as pd

# Chemin des fichiers Excel à comparer
fichier_vidanges = 'Export_vidanges_tableau.xlsx'
fichier_livraisons = 'output_livraisons.xlsx'

# Charger les fichiers Excel en DataFrames
df_vidanges = pd.read_excel(fichier_vidanges)
df_livraisons = pd.read_excel(fichier_livraisons)

# Extraire les colonnes des articles du DataFrame df_vidanges
colonnes_articles = list(df_vidanges.columns)[1:]

# Grouper par 'Client' et agréger les valeurs en les sommant
df_grouped = df_vidanges.groupby('Client', as_index=False).sum()

# Filtrer les lignes avec des écarts non nuls dans au moins une colonne
df_diff = df_grouped[df_grouped[colonnes_articles].ne(0).any(axis=1)]

# Afficher les différences
print("Différences entre les deux fichiers :")
print(df_diff)

# Exporter les différences en fichier Excel
fichier_sortie = 'differences.xlsx'
df_diff.to_excel(fichier_sortie, index=False)

print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès.")
