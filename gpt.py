import csv
import openpyxl

# Chargement des données du fichier CSV
livraisons_data = []
with open('Export_Livraisons - 20230713.csv', 'r', encoding='utf-8') as csv_file:
    csv_reader = csv.DictReader(csv_file)
    for row in csv_reader:
        livraisons_data.append(row)

# Chargement des données du fichier Excel
vidanges_data = []
workbook = openpyxl.load_workbook('Export vidanges Descartes 01-06 au 13-07.xlsx')
worksheet = workbook.active
for row in worksheet.iter_rows(min_row=2, values_only=True):
    vidange = {
        'reference_commande': row[0],
        'date_commande': row[1],
        'client': row[2],
        'vendeur': row[3],
        'montant_facture': row[4]
        # Ajoutez ici les autres colonnes nécessaires
    }
    vidanges_data.append(vidange)

# Comparaison des données
ecarts = []
for vidange in vidanges_data:
    vidange_key = vidange['reference_commande']
    found = False
    for livraison in livraisons_data:
        if vidange_key == livraison.get('Remarques', ''):
            found = True
            break
    if not found:
        ecarts.append(vidange)

# Génération du rapport des écarts
if ecarts:
    print("Ecarts détectés :")
    for ecart in ecarts:
        if ecart['client'] is not None:  # Vérifie si le champ 'Client' n'est pas None
            print("Client :", ecart['client'])
            print("Montant facturé :", ecart['montant_facture'])
            print("Date de la commande :", ecart['date_commande'])
            print("-----")
  
        
else:
    print("Aucun écart détecté.")
