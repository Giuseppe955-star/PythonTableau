import pandas as pd

# Caricare il file Excel aggiornato
file_path = 'C:/Users/giuseppe.diperna/PycharmProjects/pythonProjectTableau/Tableau Business Case Data Analyst July 2024 (2).xlsx'
xls = pd.ExcelFile(file_path)

# Caricare i fogli rilevanti
dati_progetti = pd.read_excel(xls, sheet_name='Project Data')
metadata = pd.read_excel(xls, sheet_name='Metadata')

# Caricare il foglio 'Taxonomy' considerando le specifiche delle righe
taxonomy_codes = pd.read_excel(xls, sheet_name='Taxonomy', nrows=1)
taxonomy_interventions = pd.read_excel(xls, sheet_name='Taxonomy', skiprows=1)

# Verificare le intestazioni e i dati
print("Intestazioni:")
print(taxonomy_codes.head())

print("\nInterventi:")
print(taxonomy_interventions.head())

# Assicurarsi che il numero di colonne corrisponda e rinominare le colonne
expected_columns = ['First Order Code', 'Second Order Code', 'Third Order Code', 'Unnamed: 1', 'Unnamed: 2']
taxonomy_interventions.columns = expected_columns

# Combinare le colonne 'Unnamed: 1' e 'Unnamed: 2' per creare la colonna 'Interventions'
taxonomy_interventions['Interventions'] = taxonomy_interventions['Unnamed: 1'].combine_first(taxonomy_interventions['Unnamed: 2'])
taxonomy_interventions.drop(['Unnamed: 1', 'Unnamed: 2'], axis=1, inplace=True)

# Riempire i valori mancanti nei codici gerarchici
taxonomy_interventions['First Order Code'].ffill(inplace=True)
taxonomy_interventions['Second Order Code'].ffill(inplace=True)
taxonomy_interventions['Third Order Code'].ffill(inplace=True)

# Visualizzare la struttura del dataframe dopo il riempimento
print(taxonomy_interventions.head())
