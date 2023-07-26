import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
import subprocess
from pathlib import Path
from tkinter import ttk
import tkinter.messagebox as messagebox
from tkinter import Button
from tkinter import PhotoImage
from datetime import date
import webbrowser


# Créer la fenêtre principale
print('Lancement de l\'application...')
print ('Lançement de l\'interface graphique...')
root = tk.Tk()
root.title('Traitement des fichiers Excel')
print('Ouverture fenêtre principale...')

#icone fenêtres
icone = tk.PhotoImage(file='img/logo_jr2.png')
# Définir l'icône pour la fenêtre principale
root.tk.call('wm', 'iconphoto', root._w, icone)

# Définir la taille de la fenêtre
root.geometry("850x680")

                                                    ### FONCTIONS ###

# Fonction pour changer l'icône des nouvelles fenêtres
def changer_icone_fenetre(fenetre):
    fenetre.iconphoto(True, icone)


# Fonction pour traiter le fichier des livraisons
def traiter_livraisons():
    fichier_livraisons = filedialog.askopenfilename(filetypes=[('Fichiers Excel', '*.xlsx')])

    # Valider le fichier avant de continuer                                                                              MODIF A FAIRE POUR NOUV EXCEL
    if not valider_fichier_livraison(fichier_livraisons):
        messagebox.showerror("Erreur #100", "Le fichier sélectionné n'est pas au bon format ou ne contient pas les données attendues.")
        return
    
    df_livraisons = pd.read_excel(fichier_livraisons)
    
    # Extraire les colonnes des vidanges (de D à N)
    colonnes_vidanges = df_livraisons.columns[3:14]  # Adapté pour les colonnes D à N (indices 3 à 14 exclus)            MODIF A FAIRE POUR NOUV EXCEL
    # Créer un dictionnaire pour stocker les clients, leurs vidanges et les quantités correspondantes
    clients_vidanges = defaultdict(dict)

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

    # Fonction pour mettre à jour la barre de progression
    def update_progress(progress):
        progress_bar.step(progress)
        root.update_idletasks()

    nb_etapes = 5

    # Mettre à jour la barre de progression progressivement
    for i in range(1, nb_etapes + 1):
        root.after(i * 100, update_progress, 49 // nb_etapes)
    

    # Créer une liste de listes pour le tableau
    table_data = []
    for client, vidanges in clients_vidanges.items():
        row = [client] + [vidanges.get(vidange, 0) for vidange in colonnes_vidanges]
        table_data.append(row)

    # Afficher le tableau
    headers = ['Client'] + list(colonnes_vidanges)

    # Convertir la liste de listes en DataFrame pandas
    df_output = pd.DataFrame(table_data, columns=headers)

    # Enregistrer le DataFrame dans un fichier Excel dans le dossier des téléchargements
    nom_fichier_excel = os.path.join(os.path.expanduser('~'), 'Downloads', f'output_livraisons_{date.today()}.xlsx')
    df_output.to_excel(nom_fichier_excel, index=False)
          
    print(f"Le fichier Excel concernant '{fichier_livraisons}' a été enregistré avec succès dans le dossier téléchargements sous le nom:", nom_fichier_excel)
    # Afficher le résultat dans la fenêtre
    result_label.config(text=f"Le fichier Excel a été enregistré avec succès\n")
    messagebox.showinfo("Prêt", "Le fichier a bien été enregistré !\n\nVous pouvez maintenant traiter le fichier des vidanges.")
    up_label.config(text="Vous pouvez maintenant traiter le fichier des vidanges !")
    
    
# Fonction pour traiter le fichier des vidanges
def traiter_vidanges():
    fichier_vidanges = filedialog.askopenfilename(filetypes=[('Fichiers Excel', '*.xlsx')])
    
     # Valider le fichier avant de continuer
    if not valider_fichier_vidanges(fichier_vidanges):
        messagebox.showerror("Erreur #100", "Le fichier sélectionné n'est pas au bon format ou ne contient pas les données attendues.")
        return
    
    df_vidanges = pd.read_excel(fichier_vidanges)

    # Extraire les colonnes des clients, des quantités facturées et des articles                                                              MODIF A FAIRE POUR NOUV EXCEL
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
    fichier_sortie = os.path.join(os.path.expanduser('~'), 'Downloads', f'Export_vidanges_tableau_{date.today()}.xlsx')
    df_export = pd.DataFrame(table_data, columns=headers)
    df_export.to_excel(fichier_sortie, index=False)
    
    def update_progress(progress):
        progress_bar.step(progress)
        root.update_idletasks()
    
    nb_etapes = 5

    # Mettre à jour la barre de progression progressivement
    for i in range(1, nb_etapes + 1):
        root.after(i * 100, update_progress, 50 // nb_etapes)
        
    print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès dans le dossier téléchargements.")
    result_label.config(text=f"Le fichier Excel a été enregistré avec succès \n")
    messagebox.showinfo("Prêt", "Le fichier a bien été enregistré !\n\nVous pouvez maintenant comparer les deux fichiers générés.")
    up_label.config(text="Vous pouvez maintenant comparer les deux fichiers générés.", foreground="blue")
    
    
# Fonctions de validation pour les fichiers

def valider_fichier_vidanges(fichier):
    # Vérifier que le fichier est au format xlsx
    print('----------')
    print(f"Vérification du format du fichier en cours...")
    if not fichier.lower().endswith('.xlsx'):
        messagebox.showerror("Erreur #101", "Le fichier sélectionné n'est pas au bon format.")
        return False

    # Vérifier que le fichier contient les colonnes attendues                                                                   MODIF A FAIRE POUR NOUV EXCEL
    print('Vérification du contenu du fichier en cours...')
    colonnes_attendues = ['Client', 'Lignes de la commande/Article', 'Lignes de la commande/Quantité facturée' ]
    df = pd.read_excel(fichier)
    colonnes_fichier = df.columns.tolist()

    if not all(colonne in colonnes_fichier for colonne in colonnes_attendues):
        messagebox.showerror("Erreur #102", "Le fichier sélectionné ne contient pas les données attendues.")
        return False

    return True


def valider_fichier_livraison(fichier):
    # Vérifier que le fichier est au format xlsx
    print('----------')
    print(f"Vérification du format du fichier en cours...")
    if not fichier.lower().endswith('.xlsx'):
        messagebox.showerror("Erreur #101", "Le fichier sélectionné n'est pas au bon format.")
        return False

    # Vérifier que le fichier contient les colonnes attendues                                                                   MODIF A FAIRE POUR NOUV EXCEL
    print('Vérification du contenu du fichier en cours...')
    colonnes_attendues = ['Customer Name', 'Palette Euro NEW', 'Caisses vertes', 'VID-T', 'VID-S', 'Vidanges champignons', 'Vidange F', 'FRIGO BOX', 'Palette Truval', 'Palette banane', 'Palette Plastique', 'Palette Pool' ]
    df = pd.read_excel(fichier)
    colonnes_fichier = df.columns.tolist()

    if not all(colonne in colonnes_fichier for colonne in colonnes_attendues):
        messagebox.showerror("Erreur #102", "Le fichier sélectionné ne contient pas les données attendues.")
        return False

    return True


# Fonction pour comparer les deux fichiers générés
def comparer_fichiers():
    # Chemin des fichiers Excel à comparer
    fichier_vidanges = os.path.join(os.path.expanduser('~'), 'Downloads', f'Export_vidanges_tableau_{date.today()}.xlsx')
    fichier_livraisons = os.path.join(os.path.expanduser('~'), 'Downloads', f'output_livraisons_{date.today()}.xlsx')

    if os.path.exists(fichier_vidanges) and os.path.exists(fichier_livraisons):
        # Charger les fichiers Excel en DataFrames
        df_vidanges = pd.read_excel(fichier_vidanges)
        df_livraisons = pd.read_excel(fichier_livraisons)
        
        # Vérifier que les fichiers contiennent des données
        if df_vidanges.empty or df_livraisons.empty:
            messagebox.showerror("Erreur #103", "Les fichiers ne contiennent pas de données.")
            print("Les fichiers ne contiennent pas de données.")
            return

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

        # Exporter les différences en fichier Excel dans le dossier des téléchargements
        fichier_sortie = os.path.join(os.path.expanduser('~'), 'Downloads', f'differences_{date.today()}.xlsx')
        df_diff.to_excel(fichier_sortie, index=False)

        # Mettre à jour la barre de progression pendant le traitement
        def update_progress():
            if progress < 100:
                progress_bar.step(1)
                root.after(100, update_progress)

        progress = 0
        progress_bar.step(0)  # Début du traitement, barre de progression à 0%
        root.after(100, update_progress)  # Démarre la mise à jour de la barre de progression

        progress = 100
     
        print('La comparaison est terminée !')
        print(f"Le fichier Excel '{fichier_sortie}' a été créé avec succès.")
        messagebox.showinfo("Validation", "Les fichiers ont bien été comparés !\n\nVous pouvez maintenant ouvrir le fichier généré.\n\nTous les fichiers générés sont disponibles dans le dossier des téléchargements.")
        result_label.config(text=f"La comparaison est terminée !\nLe fichier Excel a été enregistré avec succès\n")
        up_label.config(text="Vous pouvez maintenant ouvrir le fichier généré !")
                    
    else: 
        print(f"Les fichiers '{fichier_vidanges}' et/ou '{fichier_livraisons}' n'existent pas.")
        warn_label.config(text=f"!! ATTENTION !! Les fichiers '{fichier_vidanges}' \net/ou '{fichier_livraisons}' n'existent pas.", foreground="red")
        messagebox.showerror("Erreur #200", "Il n'y a pas de fichier à comparer et/ou il en manque un !\nVeuillez traiter les fichiers avant de les comparer")
    

# Fonction pour ouvrir le dernier fichier généré
def ouvrir_dernier_fichier():
    progress_bar.step(100)
    dossier_telechargements = Path.home() / 'Downloads'

    fichier_genere = f'differences_{date.today()}.xlsx'

    chemin_fichier_genere = dossier_telechargements / fichier_genere

    if chemin_fichier_genere.exists():
        progress_bar.step(100)
        os.startfile(str(chemin_fichier_genere))
        print(f"Le dernier fichier généré '{fichier_genere}' a été ouvert.")
        result_label.config(text=f"Le dernier fichier généré a été ouvert : {chemin_fichier_genere}")
    else:
        print(f"Le fichier '{fichier_genere}' n'existe pas.")
        warn_label.config(text=f"Le fichier '{fichier_genere}' n'existe pas.")
        messagebox.showerror("Erreur #300", "Le fichier des comparaisons n'existe pas.\nVeuillez traiter les fichiers avant de les comparer")


# Fonction pour ouvrir le dossier des téléchargements
def ouvrir_dossier_telechargements():
    progress_bar.step(100)
    dossier_telechargements = os.path.join(os.path.expanduser('~'), 'Downloads')
    subprocess.Popen(f'explorer "{dossier_telechargements}"')
    print(f"Le dossier des téléchargements a été ouvert : {dossier_telechargements}")
    result_label.config(text=f"Le dossier des téléchargements a été ouvert : {dossier_telechargements}")
    
    
# Fonction pour quitter la fenêtre
def quitter_fenetre():
    root.quit()
    print('Fermeture de l\'application...')
    quit_label = tk.Label(root, text="Fermeture de l'application...")
    quit_label.pack()
    
    
def afficher_aide():
    message_aide = """
    RUBRIQUE D'AIDE : \n\n
    Tous les fichiers générés sont disponibles dans le dossier des téléchargements. \n\nLeur nom contient la date du jour. \n\n
    IMPORTANT : \n
    - Seuls les fichiers Excel en .xlsx sont acceptés. 
    - Bien attendre la fenêtre pop-up avant de cliquer sur le bouton suivant. 
    - Il est conseillé de suivre l'ordre des boutons pour éviter les erreurs. \n\n
    En cas d'erreur : \n
    - Lire le message d'erreur, une rubrique traitant les erreurs est disponible dans l'onglet 'Aide'.
    - Vérifiez le format des fichiers Excel (.xlsx).
    - Vérifiez que les fichiers Excel sont correctement structuré et ne contiennent pas d'erreurs.\n
    Si toutes ces vérifications sont correctes, veuillez réessayer les opérations dans l'ordre.
    
    
    
    Développé par : BROAGE Nicolas
    """
    messagebox.showinfo("Aide", message_aide)


def afficher_instructions_boutons():
    message = f"""
    Instructions d'utilisation :
    - Bouton 'Traiter le fichier des livraisons': Cliquez sur ce bouton pour traiter le fichier des livraisons. Une fois le traitement terminé, le fichier généré sera enregistré dans le dossier des téléchargements.

    - Bouton 'Traiter le fichier des vidanges': Cliquez sur ce bouton pour traiter le fichier des vidanges. Une fois le traitement terminé, le fichier généré sera enregistré dans le dossier des téléchargements.

    - Bouton 'Comparer les fichiers générés': Cliquez sur ce bouton pour comparer les deux fichiers générés. Une fois la comparaison terminée, le résultat sera enregistré dans un nouveau fichier Excel dans le dossier des téléchargements.

    - Bouton 'Ouvrir le dernier fichier généré': Cliquez sur ce bouton pour ouvrir le dernier fichier Excel généré. Assurez-vous que le fichier existe dans le dossier des téléchargements.

    - Bouton 'Ouvrir le dossier des téléchargements': Cliquez sur ce bouton pour ouvrir le dossier des téléchargements où tous les fichiers générés sont enregistrés.

    - Bouton 'Quitter l'application': Cliquez sur ce bouton pour fermer l'application.

    Note : Suivez l'ordre des boutons pour éviter les erreurs lors du traitement des fichiers.
    """
    messagebox.showinfo("Instructions d'utilisation", message)
    
    
def err_100():
    message_100 = """
    Erreur 100 : Le fichier sélectionné n'est pas au bon format ou ne contient pas les données attendues.
    Pour régler cette erreur :
    - Assurez-vous que le fichier est au format Excel (.xlsx).
    - Vérifiez que le fichier contient les colonnes attendues : 'Client', 'Lignes de la commande/Article', 'Lignes de la commande/Quantité facturée'.
    """ 
    messagebox.showinfo("Erreur 100", message_100)
        
def err_101():
    message_101 = """
    Erreur 101 : Le fichier sélectionné n'est pas au bon format.
    Pour régler cette erreur :
    - Assurez-vous que le fichier est au format Excel (.xlsx).
    """
    messagebox.showinfo("Erreur 101", message_101)
    
def err_102():
    message_102 = """
    Erreur 102 : Le fichier sélectionné ne contient pas les données attendues.
    Pour régler cette erreur :
    - Vérifiez que le fichier contient les colonnes attendues : 'Client', 'Lignes de la commande/Article', 'Lignes de la commande/Quantité facturée'.
    """
    messagebox.showinfo("Erreur 102", message_102)
        
def err_103():
    message_103 = """
    Erreur 103 : Les fichiers ne contiennent pas de données.
    Ce message indique que le fichier est vide. Assurez-vous de choisir le bon fichier.
    """
    messagebox.showinfo("Erreur 103", message_103)
        
def err_200():
    message_200 = """
    Erreur 200 : Il n'y a pas de fichier à comparer et/ou il en manque un !
    Ce message indique qu'il manque un fichier à comparer. Assurez-vous de traiter les fichiers avant de les comparer.
    """
    messagebox.showinfo("Erreur 200", message_200)
    
def err_300():
    message_300 = """
    Erreur 300 : Le fichier des comparaisons n'existe pas.
    Ce message indique que la comparaison n'a pas été effectuée.
    """
    messagebox.showinfo("Erreur 300", message_300)


def type_erreur():
    types_errors_window = tk.Toplevel(root)
    types_errors_window.title("Types d'erreurs")
    types_errors_window.geometry("700x400")  # Définir la taille de la fenêtre

    # Ajouter les boutons d'erreurs à la fenêtre pop-up
    tk.Button(types_errors_window, text="Erreur 100 : Le fichier sélectionné n'est pas au bon format ou ne contient pas les données attendues", command=err_100, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    tk.Button(types_errors_window, text="Erreur 101 : Le fichier sélectionné n'est pas au bon format", command=err_101, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    tk.Button(types_errors_window, text="Erreur 102 : Le fichier sélectionné ne contient pas les données attendues", command=err_102, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    tk.Button(types_errors_window, text="Erreur 103 : Les fichiers ne contiennent pas de données", command=err_103, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    tk.Button(types_errors_window, text="Erreur 200 : Il n'y a pas de fichier à comparer et/ou il en manque un !\nVeuillez traiter les fichiers avant de les comparer", command=err_200, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    tk.Button(types_errors_window, text="Erreur 300 : Le fichier des comparaisons n'existe pas.\nVeuillez traiter les fichiers avant de les comparer", command=err_300, image=ico_error, compound='left', width=410,wraplength=400).pack(padx=20, pady=10)
    
    # Ajouter le bouton "OK" pour fermer la fenêtre
    btn_ok = tk.Button(types_errors_window, text="OK", command=types_errors_window.destroy, image=ico_ok, compound='left')
    btn_ok.pack(pady=10)


def ouvrir_github():
    # Ouvrir le lien GitHub dans le navigateur web par défaut
    webbrowser.open("https://github.com/Nicolo41/excel_comparator")
    
def afficher_fct() :
    message = """
    L'application suivante sert à traiter des fichiers Excel. \nCette interface utilisateur ainsi que les fonctions ont été développées en Python. \n\n
    L'application permet de comparer deux fichiers Excel et de voir s'il y a des différences entre les deux. \n\n
    
    PROCESUS : \n
    - L'utilisateur doit sélectionner un fichier Excel contenant les données des livraisons. \n Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des livraisons par client. \n\n
    - L'utilisateur doit sélectionner un fichier Excel contenant les données des vidanges. \n Le script va lire le fichier et créer un nouveau fichier Excel contenant les données des vidanges par client. \n\n
    - Lorsque l'utilisateur appuie sur le bouton 'Comparer les fichiers générés', le script va comparer les deux fichiers générés et créer un nouveau fichier Excel contenant les différences entre les deux. \n\n
    
    Le code est fait pour que chaque fichier généré soit enregistré dans le dossier des téléchargements à la date de la création. \n\n
    Lien vers le code source :  \n\n
    """
    # messagebox.showinfo("Fonctionnement de l'application", message)
    types_errors_window = tk.Toplevel(root)
    types_errors_window.title("Fonctionnement de l'application")
    
    # Afficher le message dans la fenêtre pop-up
    tk.Label(types_errors_window, text=message, wraplength=600).pack(padx=20, pady=10)

    # Ajouter un bouton pour ouvrir le lien GitHub
    btn_github = tk.Button(types_errors_window, text="GitHub", command=ouvrir_github, image=ico_git, compound='left')
    btn_github.pack(pady=20, padx=20)
    # Ajouter le bouton "OK" pour fermer la fenêtre
    btn_ok = tk.Button(types_errors_window, text="OK", command=types_errors_window.destroy, image=ico_ok, compound='left')
    btn_ok.pack(pady=10)
  
                                              ### ICONES ET IMAGES ###  

# Charger les icônes au format .png avec tkinter
ico_excel = PhotoImage(file='img/excel2.png')
ico_compare = PhotoImage(file='img/compare2.png')
ico_exit = PhotoImage(file='img/exit2.png')
ico_fichier = PhotoImage(file='img/fichier2.png')
ico_folder = PhotoImage(file='img/folder2.png')
ico_error = PhotoImage(file='img/error2.png')
ico_git = PhotoImage(file='img/git2.png')
ico_ok = PhotoImage(file='img/ok.png')
#background
bg = PhotoImage(file='img/logo_jr.png')


                                            ### WIDGETS, BOUTONS ET MENUS ###
                                            

# Créer un widget Label pour afficher l'image en arrière-plan
background_label = tk.Label(root, image=bg)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Ajouter d'autres widgets sur la fenêtre
label1 = tk.Label(root, text="Bienvenue !", font=('Arial', 20))
label1.pack(pady=20)

# Créer un menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# Créer un sous-menu "Aide"
menu_aide = tk.Menu(menu_bar, tearoff=False)
menu_bar.add_cascade(label='Aide', menu=menu_aide)

# Ajouter un élément dans le sous-menu "Aide" pour afficher l'aide
menu_aide.add_command(label='Général', command=afficher_aide)
menu_aide.add_command(label='Fonctionnement de l\'application', command=afficher_fct)
menu_aide.add_command(label='Afficher les instructions d\'utilisation des boutons', command=afficher_instructions_boutons)
menu_aide.add_command(label='Types d\'erreurs', command=type_erreur)


# Ajouter un widget Label pour afficher les résultats
result_label = tk.Label(root, text="")
result_label.pack(padx=20, pady=10)

# Ajouter un widget Label pour afficher les instructions
up_label = tk.Label(root, text="", foreground="blue")
up_label.pack(padx=20, pady=10)

# Ajouter un widget Label pour afficher les erreurs
warn_label = tk.Label(root, text="", foreground="red")
warn_label.pack(padx=20, pady=10)

result_label.config(text="Rendez-vous dans la rubrique 'Aide' en haut à gauche pour plus d'informations\nUne pop-up vous indiquera quand vous pourrez cliquer sur le bouton suivant.", font=('Arial', 10))

# Créer des boutons pour les différentes opérations
btn_traiter_livraisons = tk.Button(root, text='1. Traiter le fichier des livraisons', command=traiter_livraisons, image = ico_excel, compound='left', font=('Arial', 10))
btn_traiter_livraisons.pack(padx=20, pady=10)

btn_traiter_vidanges = tk.Button(root, text='2. Traiter le fichier des vidanges', command=traiter_vidanges, image = ico_excel, compound='left', font=('Arial', 10))
btn_traiter_vidanges.pack(padx=20, pady=10)

btn_comparer_fichiers = tk.Button(root, text='3. Comparer les fichiers générés', command=comparer_fichiers, image = ico_compare, compound='left', font=('Arial', 10))
btn_comparer_fichiers.pack(padx=20, pady=10)

# Créer un bouton pour ouvrir le dernier fichier généré
btn_ouvrir_fichier = tk.Button(root, text='4. Ouvrir le dernier fichier généré', command=ouvrir_dernier_fichier, image = ico_fichier, compound='left', font=('Arial', 10))
btn_ouvrir_fichier.pack(padx=20, pady=10)

# Créer un bouton pour ouvrir le dossier des téléchargements
btn_ouvrir_dossier = tk.Button(root, text='Ouvrir le dossier des téléchargements', command=ouvrir_dossier_telechargements, image = ico_folder, compound='left', font=('Arial', 10))
btn_ouvrir_dossier.pack(padx=20, pady=10)

# Ajouter une barre de progression
progress_bar = ttk.Progressbar(root, mode='determinate', maximum=100)
progress_bar.pack(fill='x', padx=20, pady=10)

# Créer un bouton pour quitter la fenêtre
btn_quitter = Button(root, text='Quitter l\'application', foreground="red", command=root.quit, image = ico_exit, compound='left', font=('Arial', 10))
btn_quitter.pack(padx=20, pady=10)


# Changer l'icône des nouvelles fenêtres
changer_icone_fenetre(root)

# Lancer l'interface graphique
root.mainloop()