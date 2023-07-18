import pandas as pd

# Chemins des fichiers
fichier_livraisons = 'Export_Livraisons - 20230713.xlsx'
fichier_vidanges = 'Export vidanges Descartes 01-06 au 13-07.xlsx'

# Charger le fichier des vidanges en DataFrame pandas
df_vidanges = pd.read_excel(fichier_vidanges, usecols='C:G')
df_vidanges = df_vidanges.rename(columns={'Lignes de la commande/Quantité facturée': 'Quantite facturee'})

# Charger le fichier des livraisons en DataFrame pandas, en ne sélectionnant que les colonnes C à N
df_livraisons = pd.read_excel(fichier_livraisons, usecols='C:N')

# Pivoter le DataFrame des livraisons pour avoir les vidanges en lignes
df_livraisons_pivot = df_livraisons.melt(var_name='Vidange', value_name='Quantite')

# Supprimer les lignes vides du DataFrame pivoté
df_livraisons_pivot = df_livraisons_pivot.dropna()

# Supprimer les colonnes inutiles du DataFrame des vidanges
df_vidanges = df_vidanges[['Client', 'Quantite facturee']]

# Fusionner les DataFrames des livraisons pivotées et des vidanges
merged_df = pd.merge(df_livraisons_pivot, df_vidanges, left_on='Vidange', right_on='Client', how='left')

# Calculer les écarts
merged_df['Écart'] = merged_df['Quantite'] - merged_df['Quantite facturee']

# Afficher les écarts
print(merged_df[['Vidange', 'Quantite', 'Quantite facturee', 'Écart']])
print('-----------------')
print(df_livraisons_pivot['Vidange'].value_counts())
print(df_vidanges['Client'].value_counts())
print('-----------------')
print('Lignes erreurs :')
# Afficher les lignes d'erreur
lignes_erreur = merged_df[merged_df['Écart'].notna()]
print(lignes_erreur)

