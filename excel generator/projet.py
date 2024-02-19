import pandas as pd
from random import randint, choice
from datetime import datetime, timedelta
import random

projets = [f"Projet {i+1}" for i in range(25)]
phases_projet = ["cadrage", "conception", "réalisation", "recette"]
poles = ["Commerce", "conformité", "crédit", "recouvrement", "flux", "finance"]
sponsors = ["CP IT", "CP métier", " CP organisation"]

data = [
    ["ID_Projet", "Nom_Projet", "Date MAJ", "Phase", "Pole", "Sponsor", "Date début projet", "Date fin projet", "Date début",
     "Date fin"]
]




for i in range(len(projets)):
    id = i + 1
    projet = projets[i]
    phase = choice(phases_projet)
    pole = choice(poles)
    sponsor = choice(sponsors)

    date_debut_projet = (datetime(2023, 1, 1) + timedelta(days=random.randint(1, 30))).strftime('%d/%m/%Y')
    date_fin_projet = (datetime(2023, 1, 1) + timedelta(days=random.randint(31, 365))).strftime('%d/%m/%Y')

    date_maj = datetime.strptime(date_fin_projet, '%d/%m/%Y') + timedelta(days=random.randint(1, 365 - (datetime.strptime(date_fin_projet, '%d/%m/%Y') - datetime.strptime(date_debut_projet, '%d/%m/%Y')).days))
    date_maj = date_maj.strftime('%d/%m/%Y')

    date_debut = (datetime(2023, 1, 1) + timedelta(days=random.randint(0, (datetime.strptime(date_fin_projet, '%d/%m/%Y') - datetime(2023, 1, 1)).days))).strftime("%d/%m/%Y")
    date_fin = (datetime(2023, 1, 1) + timedelta(days=random.randint((datetime.strptime(date_debut, '%d/%m/%Y') - datetime(2023, 1, 1)).days, (datetime.strptime(date_fin_projet, '%d/%m/%Y') - datetime(2023, 1, 1)).days))).strftime("%d/%m/%Y")

    data.append([
        id, projet, date_maj, phase, pole, sponsor, date_debut_projet, date_fin_projet, date_debut, date_fin
    ])




df = pd.DataFrame(data[1:], columns=data[0])


df["Date fin projet"] = pd.to_datetime(df["Date fin projet"], format='%d/%m/%Y')

#cas exeptionnel
today = datetime.now()
df.loc[df["Date fin projet"] < today, "Phase"] = "mise en production"


file_name = "projet.xlsx"
df.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' avec {len(projets)} lignes de données similaires a été généré avec succès.")


df = pd.read_excel(file_name)

df["Date début projet"] = pd.to_datetime(df["Date début projet"], format='%d/%m/%Y')
df["Date fin projet"] = pd.to_datetime(df["Date fin projet"], format='%d/%m/%Y')

# Sélectionner les lignes dans l'intervalle spécifié
date_debut_projet_min = datetime(2023, 1, 1)
date_fin_projet_max = datetime(2023, 12, 30)
filtered_df = df.loc[
    (df["Date début projet"] >= date_debut_projet_min) & (df["Date fin projet"] <= date_fin_projet_max)
]

print(filtered_df)
