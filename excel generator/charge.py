import pandas as pd
import random

# Liste des valeurs possibles pour le type de charge
types_charge = ["Interne", "Externe"]

# Générer les données pour la table "charge"
charge_data = []
for i in range(1, 51):  # Générer 100 lignes de données
    for type_charge in types_charge:
        charge_estimee = random.randint(5, 20)
    
        # 90% de charges RAF positives
        if random.random() <= 0.9:
            charge_consommee = random.randint(1, charge_estimee - 1)
        else:
            # 10% de charges RAF négatives (charge consommée supérieure à la charge estimée)
            charge_consommee = random.randint(charge_estimee + 1, charge_estimee + 10)
        
        charge_raf = charge_estimee - charge_consommee

        charge_data.append([i, type_charge, charge_estimee, charge_consommee, charge_raf])

charge_df = pd.DataFrame(charge_data, columns=[
    'ID_Charge', 'Type_charge', 'Charge_estimee', 'Charge_consommee', 'Charge_RAF'
])

# Enregistrer le DataFrame dans un fichier Excel
file_name = "charge.xlsx"
charge_df.to_excel(file_name, index=False)

print(f"Le fichier Excel '{file_name}' contenant les informations de charge a été généré avec succès.")
