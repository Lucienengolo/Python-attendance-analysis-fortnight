import pandas as pd  # Importation de la bibliothèque pandas pour la manipulation des données

# Chemins vers les fichiers Excel
file_path1 = 'JUILLET.xlsx'
file_path2 = 'JUIN.xlsx'
file_path3 = 'MAY.xlsx'

# Fonction pour charger les feuilles d'un fichier Excel en DataFrames
def load_excel_to_dataframes(file_path):
    # Chargement du fichier Excel
    excel_file = pd.ExcelFile(file_path)
    # Dictionnaire pour stocker les DataFrames, avec les noms des feuilles comme clés
    dataframes = {}
    # Boucle sur chaque feuille dans le fichier Excel
    for i, sheet_name in enumerate(excel_file.sheet_names, start=1):
        # Lecture de la feuille et stockage dans un DataFrame
        dataframes[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
    # Retourne le dictionnaire de DataFrames
    return dataframes

# Charger les DataFrames de chaque fichier Excel
dataframes1 = load_excel_to_dataframes(file_path1)  # Chargement des DataFrames du premier fichier
dataframes2 = load_excel_to_dataframes(file_path2)  # Chargement des DataFrames du deuxième fichier
dataframes3 = load_excel_to_dataframes(file_path3)  # Chargement des DataFrames du troisième fichier

# Dictionnaire pour stocker les DataFrames fusionnés
merged_dataframes = {}

# Boucle sur chaque feuille présente dans le premier fichier Excel
for sheet_name in dataframes1.keys():
    # Vérifie si la feuille existe également dans les deux autres fichiers
    if sheet_name in dataframes2 and sheet_name in dataframes3:
        # Fusionne les DataFrames correspondant à cette feuille des trois fichiers
        merged_dataframes[sheet_name] = pd.concat(
            [dataframes1[sheet_name], dataframes2[sheet_name], dataframes3[sheet_name]],
            ignore_index=True  # Réindexe les lignes pour avoir une indexation continue
        )


# Exportation des DataFrames fusionnés dans des fichiers Excel distincts
for sheet_name, df in merged_dataframes.items():
    # Définir le nom de fichier pour chaque DataFrame fusionné
    output_file_name = f'fichier_fusionne_{sheet_name}.xlsx'
    # Exporter le DataFrame dans un fichier Excel
    df.to_excel(output_file_name, index=False)
    # Afficher un message confirmant l'exportation
    print(f'DataFrame fusionné pour la feuille {sheet_name} exporté dans {output_file_name}')


