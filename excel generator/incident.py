import pandas as pd
from random import randint, choice

num_rows = 150

# Liste des valeurs possibles pour certaines colonnes
types_incident = ["Problème de logiciel: Panne de système", "Problème matériel", "Problème de sécurité", "Problème d'accès aux applications", "Problème de sauvegarde", "Problème de connexion réseau", "Problème de performance"]
priorites = ["Basse", "Moyenne", "Haute", "Urgente", "Critique"]

# Fonction pour générer le temps de résolution en fonction de la priorité de l'incident
def generate_temps_resolution(priorite):
    if priorite == "Critique":
        return randint(1, 24)
    elif priorite == "Urgente":
        return randint(4, 72)
    elif priorite == "Haute":
        return randint(8, 120)
    elif priorite == "Moyenne":
        return randint(12, 168)
    elif priorite == "Basse":
        return randint(24, 240)
    else:
        return 0

# ...
# Générer les données pour la table "incident"
data = [
    ["Id incident", "type_incident", "Priorité de l'incident", "temps de resolution(en heures)", "impact de l'incident sur les opérations", "description", "service", "type incident service"]
]

for i in range(1, num_rows):
    id_incident = str(i)
    type_incident = choice(types_incident)
    priorite_incident = choice(priorites)
    temps_resolution = generate_temps_resolution(priorite_incident)
    impact_incident = choice(["Faible", "Moyen", "Élevé"])
    description = f"Description de l'incident {i}"
    service = choice(["Réseau et Telecom", "Sécurité Info", "Système d'exploitation", "Base de Données"])
    if service == "Réseau et Telecom":
        type_incident_service = choice(["Erreur de Configuration", "Temps de Reponse", "Connexion Internet"])
    elif service == "Sécurité Info":
        type_incident_service = choice(["Authentification", "Vulnerabilite-Faille", "Attaque"])
    elif service == "Système d'exploitation":
        type_incident_service = choice(["Mise à jours os", "Compatibilite Logiciel", "Configuration Système", "Malfonctionnement APP", "Problème Materiel"])
    elif service == "Base de Données":
        type_incident_service = choice(["Perte de Données", "Erreur de Requettes", "Panne Serveur"])
    else:
        type_incident_service = ""
    
    data.append([
        id_incident, type_incident, priorite_incident, temps_resolution, impact_incident, description, service, type_incident_service
    ])

# Créer un DataFrame à partir des données
df_incident = pd.DataFrame(data[1:], columns=data[0])

# Générer le fichier Excel
file_name = "incident.xlsx"
df_incident.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' avec la table de dimension 'incident' a été généré avec succès.")
