import pandas as pd

# Ouvrir le fichier CSV en utilisant open pour gérer les erreurs de tokenisation
with open("C:/Users/nicol/Downloads/vidanges/Export_Livraisons - 20230713.csv", errors='ignore') as file:
    # Lire le contenu du fichier CSV
    lines = file.readlines()

# Supprimer les caractères de nouvelle ligne
lines = [line.strip() for line in lines]

# Identifier les lignes correctement formatées (avec un seul champ)
formatted_lines = []
for line in lines:
    if len(line.split(',')) == 1:
        formatted_lines.append(line)

# Créer le DataFrame à partir des lignes correctement formatées
livraisons_df = pd.read_csv("\n".join(formatted_lines), delimiter=',')

# Charger le fichier Excel
vidanges_df = pd.read_excel("Export vidanges Descartes 01-06 au 13-07.xlsx")

# Identifier les références de commande des livraisons et vidanges respectives
ref_livraisons = set(livraisons_df["Num unique Liv"].dropna())
ref_vidanges = set(vidanges_df["Référence commande"].dropna())

# Trouver les références de commande présentes dans les livraisons mais absentes des vidanges
references_manquantes = ref_livraisons - ref_vidanges

# Afficher les références de commande manquantes dans les vidanges
print("Références de commande manquantes dans les vidanges :")
for reference in references_manquantes:
    print(reference)

# Rechercher les lignes de livraisons correspondant aux références manquantes
livraisons_manquantes = livraisons_df[livraisons_df["Num unique Liv"].isin(references_manquantes)]

# Afficher les lignes de livraisons correspondant aux références manquantes
print("\nLignes de livraisons correspondant aux références manquantes dans les vidanges :")
print(livraisons_manquantes)
