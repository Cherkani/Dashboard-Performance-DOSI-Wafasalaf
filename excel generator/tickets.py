import pandas as pd
import random

# Liste des valeurs possibles pour certaines colonnes
noms_statuts = ["initié", "en étude", "en attente", "annulé", "livré", "en cours", "bloqué"]
priorites = ["Basse", "Moyenne", "Haute", "Urgente", "Critique"]

# Générer les données pour la table de ticket
ticket_data = []
for i in range(1, 101):  # Générer 100 tickets
    id_ticket = i
    nom_statut = random.choice(noms_statuts)
    description = f"Description du ticket {i}"
    priorite = random.choice(priorites)
    
    # Définir le temps de résolution en fonction de la priorité
    if priorite == "Basse":
        temps_resolution = random.randint(24, 72)  # Temps de résolution pour les tickets de priorité basse
    elif priorite == "Moyenne":
        temps_resolution = random.randint(12, 48)  # Temps de résolution pour les tickets de priorité moyenne
    elif priorite == "Haute":
        temps_resolution = random.randint(6, 24)  # Temps de résolution pour les tickets de priorité haute
    elif priorite == "Urgente":
        temps_resolution = random.randint(2, 12)  # Temps de résolution pour les tickets de priorité urgente
    else:
        temps_resolution = random.randint(1, 6)  # Temps de résolution pour les tickets de priorité critique
    
    taux_satisfaction = random.randint(1, 100)  # Taux de satisfaction aléatoire en pourcentage

    ticket_data.append([id_ticket, nom_statut, description, priorite, temps_resolution, taux_satisfaction])

ticket_df = pd.DataFrame(ticket_data, columns=[
    'ID_Ticket', 'Nom_statut', 'Description', 'Priorité', 'Temps de résolution', 'Taux de satisfaction des utilisateurs'
])

# Enregistrer le DataFrame dans un fichier Excel
file_name = "tickets.xlsx"
ticket_df.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' contenant les informations des tickets a été généré avec succès.")
