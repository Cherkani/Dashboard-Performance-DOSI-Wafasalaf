import pandas as pd
from datetime import datetime, timedelta
import random

# ticket
ticket = pd.read_excel("tickets.xlsx")
ticket_clonne_choisi = ticket[["ID_Status_Ticket"]].copy()

# responsable support technique
responsable = pd.read_excel("responsables.xlsx")
filtered_responsable = responsable[(responsable["Fonction"] == "Support Technique")]

# agence 
agence = pd.read_excel("agence.xlsx")

# Merge data 
merged_ticket = ticket_clonne_choisi.copy()
merged_ticket["ID_Responsable"] = filtered_responsable["ID_Responsable"].sample(n=len(ticket), replace=True).reset_index(drop=True)
merged_ticket["ID_Agence"] = agence["ID_Agence"].sample(n=len(ticket), replace=True).reset_index(drop=True)

# tamps
def random_date_2023():
    date_creation = datetime(2023, 1, 1) + timedelta(days=random.randint(1, 365))
    date_cloture = datetime(2023, 1, 1) + timedelta(days=random.randint(1, 365))
    while date_cloture <= date_creation:
        date_cloture = datetime(2023, 1, 1) + timedelta(days=random.randint(1, 365))
    return date_creation, date_cloture
date_creation_ticket = []
date_cloture_ticket = []
for _ in range(len(ticket)):
    date_creation, date_cloture = random_date_2023()
    date_creation_ticket.append(date_creation.strftime('%d/%m/%Y'))
    date_cloture_ticket.append(date_cloture.strftime('%d/%m/%Y'))

# Ajouter col au dataframe
merged_ticket["Date début"] = date_creation_ticket
merged_ticket["Date fin"] = date_cloture_ticket





new_file_name = "fait_supporttechnique.xlsx"
merged_ticket.to_excel(new_file_name, index=False)

print(f"Le fichier Excel '{new_file_name}' avec les colonnes sélectionnées et les colonnes 'Nom du Responsable', 'ID_Agence', 'Date début', 'Date fin', et 'ID_Status' a été généré avec succès.")
