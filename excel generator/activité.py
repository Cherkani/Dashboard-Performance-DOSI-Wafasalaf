import pandas as pd
data_activite = ["Recette", "Réalisation"]
data = [
    ["ID", "Type_Activité"]
]
for i, activite in enumerate(data_activite, 1):
    data.append([
        str(i).zfill(1),
        activite
    ])
# Créer un DataFrame à partir des données
df = pd.DataFrame(data[1:], columns=data[0])
# Générer le fichier Excel
file_name = "activité.xlsx"
df.to_excel(file_name, index=False)
print(f"Le fichier Excel '{file_name}' avec les lignes de données pour chaque statut a été généré avec succès.")
