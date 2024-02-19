import pandas as pd
from datetime import datetime, timedelta
import random

# incident
incident = pd.read_excel("incident.xlsx")
incident_clonne_choisi = incident[["Id_Incident"]].copy()

# responsable
responsable = pd.read_excel("responsables.xlsx")
filtered_responsable = responsable[(responsable["Fonction"] != "Support Technique")]
sampled_responsable = filtered_responsable["ID_Responsable"].sample(n=len(incident), replace=True)

# agence
agence = pd.read_excel("agence.xlsx")
sampled_agence = agence["ID_Agence"].sample(n=len(incident), replace=True)

# date
year = 2023  # Change the year if needed
date_debut_incident = [
    (datetime(year, 1, 1) + timedelta(days=random.randint(0, 364))).strftime('%d/%m/%Y')
    for _ in range(len(incident))
]
date_fin_incident = [
    (datetime.strptime(date, "%d/%m/%Y") + timedelta(days=random.randint(1, 365 - int(date[0:2])))).strftime('%d/%m/%Y')
    for date in date_debut_incident
]
date_df = pd.DataFrame({"Date début": date_debut_incident, "Date fin": date_fin_incident})

# reset l index
incident_clonne_choisi.reset_index(drop=True, inplace=True)
sampled_responsable.reset_index(drop=True, inplace=True)
sampled_agence.reset_index(drop=True, inplace=True)

# concatener
merged_incident = pd.concat([incident_clonne_choisi, sampled_responsable, sampled_agence, date_df], axis=1)

# status
def get_id(row):
    current_date = datetime.today()
    start_date = datetime.strptime(row["Date début"], "%d/%m/%Y")
    end_date = datetime.strptime(row["Date fin"], "%d/%m/%Y")
    if start_date <= current_date <= end_date:
        return 2
    elif current_date < end_date:
        return 1
    else:
        return 3

merged_incident.loc[:, "ID_Status"] = merged_incident.apply(get_id, axis=1)

new_file_name = "fait_incident.xlsx"
merged_incident.to_excel(new_file_name, index=False)

print(f"Le fichier Excel '{new_file_name}' avec les colonnes sélectionnées et les colonnes 'Nom du Responsable', 'ID_Agence', 'Date début', 'Date fin', et 'ID_Status' a été généré avec succès.")
