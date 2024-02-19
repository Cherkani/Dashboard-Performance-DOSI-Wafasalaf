import pandas as pd
from random import randint, choice, choices
from datetime import datetime, timedelta
import random

num_rows = 150

# Liste des valeurs possibles pour certaines colonnes
projets = ["Projet A", "Projet B", "Projet C", "Projet D", "Projet E", "Projet F", "Projet H", "Projet I"]
phases = ["Réalisation", "Recette", 'Test', 'Maintenance']
actions = ["Action_1", "Action_2", "Action_3", "Action_4", "Action_5", "Action_6", "Action_7", "Action_8", "Action_9", "Action_10", "Action_11", "Action_12", "Action_13", "Action_14", "Action_15", "Action_16", "Action_17", "Action_22", "Action_23", "Action_19", "Action_25", "Action_24"]
responsables = ["DEV1", "DEV2", "DEV3", "DBA", "Sécurité", "CPI_1"]
responsables_weights = [1, 1, 1, 1, 1, 0.3]  # Adjust the weights as per your preference

# Générer les 100 lignes de données aléatoires similaires
data = [
    ["ID", "Date MAJ", "Projet", "Phase", "Responsable projet", "Date début projet", "Date fin projet", "Activité", "Action", "Responsable tache", "Date début", "Date fin", "Charge estimé", "Charge consommée", "Charge RAF"]
]

start_date = datetime(2021, 1, 1)
for i in range(1, num_rows):
    id = str(i)
    date_maj = (start_date + timedelta(days=random.randint(1, 365))).strftime('%d/%m/%Y')
    projet = choice(projets)
    phase = choice(phases)
    responsable_projet = "CPI_1"
    date_debut_projet = (start_date + timedelta(days=random.randint(1, 30))).strftime('%d/%m/%Y')
    date_fin_projet = (start_date + timedelta(days=random.randint(31, 90))).strftime('%d/%m/%Y')
    activite = "réalisation" if phase == "Réalisation" else "Recette"
    action = choice(actions)
    responsable_tache = choices(responsables, weights=responsables_weights)[0]

    # Generate random dates within the project start and end dates
    date_debut = (start_date + timedelta(days=random.randint(0, (datetime.strptime(date_fin_projet, '%d/%m/%Y') - start_date).days))).strftime("%d/%m/%Y")
    date_fin = (start_date + timedelta(days=random.randint((datetime.strptime(date_debut, '%d/%m/%Y') - start_date).days, (datetime.strptime(date_fin_projet, '%d/%m/%Y') - start_date).days))).strftime("%d/%m/%Y")

    charge_estimee = str(randint(1, 10))
    charge_consommee = str(randint(0, int(charge_estimee)))
    charge_raf = str(int(charge_estimee) - int(charge_consommee))

    data.append([
        id, date_maj, projet, phase, responsable_projet, date_debut_projet, date_fin_projet,
        activite, action, responsable_tache, date_debut, date_fin, charge_estimee, charge_consommee, charge_raf
    ])

    start_date += timedelta(days=randint(1, 7))  # Pour varier les dates
# Créer un DataFrame à partir des données
df = pd.DataFrame(data[1:], columns=data[0])

# Convertir les noms de projet en catégorie pour spécifier l'ordre de tri
df['Projet'] = pd.Categorical(df['Projet'], categories=["Projet " + c for c in "ABCDEFGHI"], ordered=True)

# Trier le DataFrame par le nom de projet et réinitialiser l'index
df_sorted = df.sort_values(by=['Projet']).reset_index(drop=True)

# Réinitialiser les ID de 1 à la fin
df_sorted['ID'] = range(1, len(df_sorted) + 1)

# Générer le fichier Excel
file_name = "zzz.xlsx"
df_sorted.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' avec 100 lignes de données similaires a été généré avec succès.")