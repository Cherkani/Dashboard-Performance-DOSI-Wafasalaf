import pandas as pd
from datetime import datetime

# projet
projet = pd.read_excel("projet.xlsx")
projet_colonne_choisi = projet[["ID_Projet", "Date début", "Date fin"]].sample(80, replace=True)
filtered_projet = projet.loc[projet_colonne_choisi.index]

# Convertir pour permettre comparaison pour status
filtered_projet.loc[:, "Date début"] = pd.to_datetime(filtered_projet["Date début"], dayfirst=True)
filtered_projet.loc[:, "Date fin"] = pd.to_datetime(filtered_projet["Date fin"], dayfirst=True)

# Responsables
responsable = pd.read_excel("responsables.xlsx")
filtered_responsable = responsable[(responsable["Fonction"] != "client") & (responsable["Fonction"] != "Support Technique")]
sampled_responsable = filtered_responsable["ID_Responsable"].sample(80, replace=True)

# Activité
activité = pd.read_excel("activité.xlsx")
sampled_activité = activité["ID_Activité"].sample(80, replace=True)

# Action
action = pd.read_excel("Action.xlsx")
sampled_action = action["ID_Action"].sample(80, replace=True)

# Charge
charge = pd.read_excel("charge.xlsx")
sampled_charge = charge["ID_Charge"].sample(80, replace=True)

# Status
current_date = datetime.today()

# Function to determine the ID based on the date comparison
def get_id(row):
    start_date = row["Date début"]
    end_date = row["Date fin"]
    if start_date <= current_date <= end_date:
        return 2
    elif current_date < start_date:
        return 1
    else:
        return 3

filtered_projet.loc[:, "ID_Status"] = filtered_projet.apply(get_id, axis=1)
filtered_projet = filtered_projet[["ID_Projet", "Date début", "Date fin","ID_Status"]]

# Reset les index
filtered_projet.reset_index(drop=True, inplace=True)
sampled_responsable.reset_index(drop=True, inplace=True)
sampled_activité.reset_index(drop=True, inplace=True)
sampled_action.reset_index(drop=True, inplace=True)
sampled_charge.reset_index(drop=True, inplace=True)

# Concatenate the DataFrames
merged_projet = pd.concat([filtered_projet, sampled_responsable, sampled_activité, sampled_action, sampled_charge], axis=1)
merged_projet = merged_projet.sort_values(by="ID_Projet")
# Save the merged DataFrame to 'fait_projetinformatique.xlsx'
new_file_name = "fait_projetinformatique.xlsx"
merged_projet.to_excel(new_file_name, index=False)

print(f"The Excel file '{new_file_name}' with the selected columns and the 'Nom du Responsable' column has been generated successfully.")

