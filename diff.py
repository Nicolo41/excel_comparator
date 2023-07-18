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
    vidanges_data.append({
        'Référence commande': row[0],
        'client': row[1],
        'montant': row[2]
        # Ajoutez ici les autres colonnes nécessaires
    })

# Comparaison des données
ecarts = []
for livraison in livraisons_data:
    livraison_key = livraison['date'] + '_' + livraison['client']  # Utilisation de la clé date_client
    found = False
    for vidange in vidanges_data:
        vidange_key = vidange['date'] + '_' + vidange['client']
        if livraison_key == vidange_key:
            found = True
            break
    if not found:
        ecarts.append(livraison)

# Génération du rapport des écarts
if ecarts:
    print("Ecarts détectés :")
    for ecart in ecarts:
        print("Client :", ecart['client'])
        print("Montant :", ecart['montant'])
        print("Date :", ecart['date'])
        print("-----")
else:
    print("Aucun écart détecté.")

