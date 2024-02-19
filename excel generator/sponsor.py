import pandas as pd
from random import choice

data_status = ["CP métier", "CP organisation"," CP IT"]

data = [
    ["ID", "Type_Sponsors"]
]

for i, status in enumerate(data_status,1):
    data.append([
        str(i).zfill(2),
        status
    ])

# Créer un DataFrame à partir des données
df = pd.DataFrame(data[1:], columns=data[0])

# Générer le fichier Excel
file_name = "sponsor.xlsx"
df.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' avec les lignes de données pour chaque statut a été généré avec succès.")
