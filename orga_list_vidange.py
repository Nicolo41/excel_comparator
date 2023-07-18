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

# Créer un dictionnaire pour stocker les articles et leurs quantités par client
articles_par_client = {}

# Variables temporaires pour stocker le client actuel et les articles associés
client_actuel = None
articles_actuels = {}

# Parcourir chaque ligne du DataFrame
for _, row in df_vidanges.iterrows():
    client = row[colonnes_clients[0]]
    quantite = row[colonnes_quantites[0]]
    article = row[colonnes_articles[0]]

    # Si on rencontre un nouveau client, on ajoute les articles associés au client actuel au dictionnaire
    if pd.notna(client):
        if client_actuel is not None:
            articles_par_client[client_actuel] = articles_actuels

        # Réinitialiser les articles actuels pour le nouveau client
        client_actuel = client
        articles_actuels = {}

    # Ajouter l'article et sa quantité au dictionnaire des articles actuels
    if article not in articles_actuels:
        articles_actuels[article] = quantite
    else:
        articles_actuels[article] += quantite

# Ajouter les articles et quantités du dernier client rencontré
if client_actuel is not None:
    articles_par_client[client_actuel] = articles_actuels

# Créer une liste de toutes les colonnes possibles en combinant les articles de tous les clients
toutes_colonnes = sorted(set(article for articles_quantites in articles_par_client.values() for article in articles_quantites))

# Créer une liste de listes pour le tableau final
table_data = []
for client, articles_quantites in articles_par_client.items():
    row_data = [client] + [articles_quantites.get(article, 0) for article in toutes_colonnes]
    table_data.append(row_data)

# Afficher le tableau
headers = ['Client'] + toutes_colonnes
print(tabulate(table_data, headers=headers, tablefmt='grid'))

# Exporter le tableau en fichier Excel
fichier_sortie = 'Export_vidanges_tableau.xlsx'
df_export = pd.DataFrame(table_data, columns=headers)
df_export.to_excel(fichier_sortie, index=False)

print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès.")
