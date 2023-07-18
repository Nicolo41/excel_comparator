import pandas as pd
from tabulate import tabulate

# Chemin du fichier Excel
fichier_vidanges = 'Export vidanges Descartes 01-06 au 13-07.xlsx'

# Charger le fichier des vidanges en DataFrame pandas
df_vidanges = pd.read_excel(fichier_vidanges)

# Extraire les colonnes des clients, des quantités facturées et des articles
colonnes_clients = ['Client']
colonnes_quantites = ['Lignes de la commande/Quantité facturée']
colonnes_articles = ['Lignes de la commande/Article']

# Créer un dictionnaire pour stocker les articles par client
articles_par_client = {}

# Parcourir chaque ligne du DataFrame
for _, row in df_vidanges.iterrows():
    client = row[colonnes_clients[0]]
    quantite = row[colonnes_quantites[0]]
    article = row[colonnes_articles[0]]

    # Vérifier si le client existe déjà dans le dictionnaire
    if client in articles_par_client:
        # Vérifier si l'article existe déjà pour ce client
        if article in articles_par_client[client]:
            articles_par_client[client][article] += quantite
        else:
            articles_par_client[client][article] = quantite
    else:
        articles_par_client[client] = {article: quantite}

# Créer une liste de listes pour le tableau
table_data = []
for client, articles_quantites in articles_par_client.items():
    for article, quantite in articles_quantites.items():
        row_data = [client, article, quantite]
        table_data.append(row_data)

# Afficher le tableau
headers = ['Client', 'Article', 'Quantité facturée']
print(tabulate(table_data, headers=headers, tablefmt='grid'))

# Convertir la liste de listes en DataFrame pandas
df_output = pd.DataFrame(table_data, columns=headers)

# Enregistrer le DataFrame dans un fichier Excel
nom_fichier_excel = 'output_vidanges.xlsx'
df_output.to_excel(nom_fichier_excel, index=False)

print("Le fichier Excel a été enregistré avec succès sous le nom:", nom_fichier_excel)
