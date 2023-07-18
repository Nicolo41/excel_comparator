import pandas as pd
import openai

# Configuration de l'API d'OpenAI
openai.api_key = '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'

# Chemins des fichiers Excel
fichier_livraisons = 'Export_Livraisons - 20230713.xlsx'
fichier_vidanges = 'Export vidanges Descartes 01-06 au 13-07.xlsx'

# Fonction pour comparer les fichiers Excel et identifier les écarts
def comparer_fichiers_excel():
    # Charger le fichier des vidanges en DataFrame pandas
    df_vidanges = pd.read_excel(fichier_vidanges, usecols='C:G')
    df_vidanges = df_vidanges.rename(columns={'Lignes de la commande/Quantité facturée': 'Quantite facturee'})

    # Charger le fichier des livraisons en DataFrame pandas
    df_livraisons = pd.read_excel(fichier_livraisons, usecols='C:N')

    # Pivoter le DataFrame des livraisons pour avoir les vidanges en lignes
    df_livraisons_pivot = df_livraisons.melt(var_name='Vidange', value_name='Quantite')
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

    # Lignes d'erreur
    lignes_erreur = merged_df[merged_df['Écart'].notna()]

    # Appels à l'API d'OpenAI
    for index, row in lignes_erreur.iterrows():
        vidange = row['Vidange']
        quantite = row['Quantite']
        quantite_facturee = row['Quantite facturee']

        # Appel à l'API pour effectuer une tâche spécifique avec les données de la ligne d'erreur
        # Remplacez `...` par le code approprié pour appeler l'API d'OpenAI et effectuer votre tâche souhaitée

        # Exemple : Utilisation de l'API pour générer une suggestion de résolution de l'écart
        prompt = f"Pour résoudre l'écart pour la vidange {vidange} avec une quantité de {quantite} et une quantité facturée de {quantite_facturee}, vous pouvez..."
        response = openai.Completion.create(
            engine='text-davinci-002',
            prompt=prompt,
            max_tokens=100,
            n=1,
            stop=None,
            temperature=0.7
        )

        suggestion = response.choices[0].text.strip()

        # Affichage de la suggestion générée par l'IA
        print(f"Suggestion pour résoudre l'écart pour la vidange {vidange}:")
        print(suggestion)
        print('--------ia---------')

# Appeler la fonction pour comparer les fichiers Excel
comparer_fichiers_excel()
