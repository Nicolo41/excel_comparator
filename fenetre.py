import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tabulate import tabulate
from collections import defaultdict

# Créer la fenêtre principale
root = tk.Tk()
root.title('Traitement des fichiers Excel')


# Fonction pour traiter le fichier "Export_Livraisons - 20230713.xlsx"
def traiter_livraisons():
    fichier_livraisons = filedialog.askopenfilename(filetypes=[('Export_Livraisons - 20230713.xlsx', '*.xlsx')])
    df_livraisons = pd.read_excel(fichier_livraisons)

    # Extraire les colonnes des vidanges (de D à N)
    colonnes_vidanges = df_livraisons.columns[3:14]  # Adapté pour les colonnes D à N (indices 3 à 14 exclus)
    # Créer un dictionnaire pour stocker les clients, leurs vidanges et les quantités correspondantes
    clients_vidanges = defaultdict(dict)

    # print(tabulate(table_data, headers=headers, tablefmt='grid'))
    print('--------------------')
    print(f"Lecture du fichier '{fichier_livraisons}' en cours...")
    # Afficher le résultat dans la fenêtre
    result_label.config(text=f"Lecture du fichier '{fichier_livraisons}' en cours...")
    
    
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


    # Convertir la liste de listes en DataFrame pandas
    df_output = pd.DataFrame(table_data, columns=headers)

    # Enregistrer le DataFrame dans un fichier Excel
    nom_fichier_excel = 'output_livraisons.xlsx'
    df_output.to_excel(nom_fichier_excel, index=False)

    print(f"Le fichier Excel concernant '{fichier_livraisons}' a été enregistré avec succès sous le nom:", nom_fichier_excel)
    # Afficher le résultat dans la fenêtre
    result_label.config(text=f"Le fichier Excel a été enregistré avec succès sous le nom: {nom_fichier_excel}\n\nVous pouvez maintenant traiter le fichier des vidanges.")


# Fonction pour traiter le fichier "Export vidanges Descartes 01-06 au 13-07.xlsx"
def traiter_vidanges():
    fichier_vidanges = filedialog.askopenfilename(filetypes=[('Fichiers Excel', '*.xlsx')])
    df_vidanges = pd.read_excel(fichier_vidanges)

    # Votre code de traitement pour le fichier des vidanges ici...
    # Extraire les colonnes des clients, des quantités facturées et des articles
    colonnes_clients = ['Client']
    colonnes_quantites = ['Lignes de la commande/Quantité facturée']
    colonnes_articles = ['Lignes de la commande/Article']

    # Créer un dictionnaire pour stocker les articles et leurs quantités par client
    articles_par_client = {}

    # Variables temporaires pour stocker le client actuel et les articles associés
    client_actuel = None
    articles_actuels = {}
    
    
    print('--------------------')
    print(f"Lecture du fichier '{fichier_vidanges}' en cours...")
    result_label.config(text=f"Lecture du fichier '{fichier_vidanges}' en cours...")

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
    # print(tabulate(table_data, headers=headers, tablefmt='grid'))

    # Exporter le tableau en fichier Excel
    fichier_sortie = 'Export_vidanges_tableau.xlsx'
    df_export = pd.DataFrame(table_data, columns=headers)
    df_export.to_excel(fichier_sortie, index=False)

    print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès.")
    result_label.config(text=f"Le fichier Excel a été enregistré avec succès sous le nom: {fichier_sortie}\n\nVous pouvez maintenant comparer les deux fichiers générés.")

# Fonction pour comparer les deux fichiers générés
def comparer_fichiers():
    # Chemin des fichiers Excel à comparer
    fichier_vidanges = 'Export_vidanges_tableau.xlsx'
    fichier_livraisons = 'output_livraisons.xlsx'

    if os.path.exists(fichier_vidanges) and os.path.exists(fichier_livraisons):
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
        print('--------------------')
        print("Traitement des fichiers en cours...")
        # print(df_diff)
        result_label.config(text=f"Comparaison des fichiers en cours...")

        # Exporter les différences en fichier Excel
        fichier_sortie = 'differences.xlsx'
        df_diff.to_excel(fichier_sortie, index=False)

        print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès.")
        result_label.config(text=f"La comparaison est terminée !\nLe fichier Excel a été enregistré avec succès sous le nom: {fichier_sortie}\n\nVous pouvez maintenant ouvrir le fichier généré.")
    else: 
        print(f"Les fichiers '{fichier_vidanges}' et '{fichier_livraisons}' n'existent pas.")
        result_label.config(text=f"Les fichiers '{fichier_vidanges}' et '{fichier_livraisons}' n'existent pas.")
    

# Fonction pour ouvrir le dernier fichier généré
def ouvrir_dernier_fichier():
    fichier_genere = 'differences.xlsx'  
    print('--------------------')
    if os.path.exists(fichier_genere):
        os.startfile(fichier_genere)
        print(f"Le dernier fichier généré '{fichier_genere}' a été ouvert.")
        result_label.config(text=f"Le dernier fichier généré a été ouvert : {fichier_genere}")
    else:
        print(f"Le fichier '{fichier_genere}' n'existe pas.")
        result_label.config(text=f"Le fichier '{fichier_genere}' n'existe pas.")


# Ajouter un widget Label pour afficher les résultats
result_label = tk.Label(root, text="")
result_label.pack(padx=20, pady=10)

result_label.config(text="Interface pour traiter les fichiers Excel.\nSeuls les fichiers Excel en .xlsx sont acceptés.\n\n Cliquez sur les boutons dans l'ordre\n\n Attendez l'autorisation avant de cliquer sur le bouton suivant.")


# Créer des boutons pour les différentes opérations
btn_traiter_livraisons = tk.Button(root, text='1. Traiter le fichier des livraisons', command=traiter_livraisons)
btn_traiter_livraisons.pack(padx=20, pady=10)

btn_traiter_vidanges = tk.Button(root, text='2. Traiter le fichier des vidanges', command=traiter_vidanges)
btn_traiter_vidanges.pack(padx=20, pady=10)

btn_comparer_fichiers = tk.Button(root, text='3. Comparer les fichiers générés', command=comparer_fichiers)
btn_comparer_fichiers.pack(padx=20, pady=10)

# Créer un bouton pour ouvrir le dernier fichier généré
btn_ouvrir_fichier = tk.Button(root, text='4. Ouvrir le dernier fichier généré', command=ouvrir_dernier_fichier)
btn_ouvrir_fichier.pack(padx=20, pady=10)

# Lancer l'interface graphique
root.mainloop()
