import pandas as pd
from datetime import datetime, timedelta

# projet
projet = pd.read_excel("projet.xlsx")

# Convert 'Date début' and 'Date fin' to datetime
projet.loc[:, "Date début"] = pd.to_datetime(projet["Date début"], dayfirst=True)
projet.loc[:, "Date fin"] = pd.to_datetime(projet["Date fin"], dayfirst=True)

# Responsables
responsable = pd.read_excel("responsables.xlsx")
filtered_responsable = responsable[(responsable["Fonction"] != "client") & (responsable["Fonction"] != "Support Technique")]
sampled_responsable = filtered_responsable["ID_Responsable"].sample(5 * len(projet), replace=True)

# Activité
activité = pd.read_excel("activité.xlsx")
sampled_activité = activité["ID_Activité"].sample(5 * len(projet), replace=True)

# Action
action = pd.read_excel("Action.xlsx")
sampled_action = action["ID_Action"].sample(5 * len(projet), replace=True)

# Charge
charge = pd.read_excel("charge.xlsx")
sampled_charge = charge["ID_Charge"].sample(5 * len(projet), replace=True)

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

projet.loc[:, "ID_Status"] = projet.apply(get_id, axis=1)
projet_selected = projet[["ID_Projet", "Date début", "Date fin", "ID_Status"]]


# Reset les index
projet.reset_index(drop=True, inplace=True)
sampled_responsable.reset_index(drop=True, inplace=True)
sampled_activité.reset_index(drop=True, inplace=True)
sampled_action.reset_index(drop=True, inplace=True)
sampled_charge.reset_index(drop=True, inplace=True)

# Create an empty list to store the final rows for each 'ID_Projet'
rows_list = []

# Group by 'ID_Projet' and extract each group for further processing
for _, group in projet.groupby("ID_Projet"):
    # Generate 5 rows for each 'ID_Projet' with different durations (5%, 10%, 5%, 20%, 60%)
    start_date = group.iloc[0]["Date début"]
    end_date = group.iloc[0]["Date fin"]
    total_duration = (end_date - start_date).total_seconds()
    percentages = [0.05, 0.1, 0.50, 0.15, 0.2]
    
    for percentage in percentages:
        duration = total_duration * percentage
        new_row = group.iloc[0].copy()
        new_row["Début"] = start_date
        new_row["Fin"] = start_date + timedelta(seconds=duration)
        start_date = new_row["Fin"]
        rows_list.append(new_row)

# Create DataFrame from the list of rows
merged_projet = pd.DataFrame(rows_list)

# Concatenate the DataFrames 'merged_projet', 'sampled_responsable', 'sampled_activité', 'sampled_action', 'sampled_charge'
final_merged = pd.concat([merged_projet.reset_index(drop=True), sampled_responsable.reset_index(drop=True), sampled_activité.reset_index(drop=True), sampled_action.reset_index(drop=True), sampled_charge.reset_index(drop=True)], axis=1)

# Save the modified DataFrame to 'fait_projetinformatique.xlsx'
new_file_name = "fait_projetinformatique.xlsx"
final_merged.to_excel(new_file_name, index=False)

print(f"The Excel file '{new_file_name}' with the selected columns, 'Nom du Responsable' column, and the added 'Début' and 'Fin' columns has been generated successfully.")
