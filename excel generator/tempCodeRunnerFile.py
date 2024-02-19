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

# Reset les index
projet.reset_index(drop=True, inplace=True)
sampled_responsable.reset_index(drop=True, inplace=True)
sampled_activité.reset_index(drop=True, inplace=True)
sampled_action.reset_index(drop=True, inplace=True)
sampled_charge.reset_index(drop=True, inplace=True)

# Repeat each 'ID_Projet' 5 times to get 5 rows for each 'ID_Projet'
projet = projet.loc[projet.index.repeat(5)].reset_index(drop=True)

# Concatenate the DataFrames
merged_projet = pd.concat([projet, sampled_responsable, sampled_activité, sampled_action, sampled_charge], axis=1)

# Sort the DataFrame by 'ID_Projet'
merged_projet = merged_projet.sort_values(by="ID_Projet")

# Function to generate successive dates with varying durations for each 'ID_Projet'
def generate_successive_dates(row):
    start_date = row["Date début"]
    end_date = row["Date fin"]
    total_duration = (end_date - start_date).total_seconds()
    
    # Define the division percentages for each successive date
    percentages = [0.05, 0.1, 0.05, 0.2, 0.6]
    dates = []
    for percentage in percentages:
        duration = total_duration * percentage
        new_date = start_date + timedelta(seconds=duration)
        dates.append(new_date)
        start_date = new_date
    
    return dates[:5]  # Keep only the first 5 dates

# Apply the function to create successive dates for each row
merged_projet["Successive_Date"] = merged_projet.apply(generate_successive_dates, axis=1)

# Explode the 'Successive_Date' list to multiple rows
merged_projet = merged_projet.explode("Successive_Date")

# Add two new columns 'Début' and 'Fin' to split the 'Successive_Date' into separate start and end dates
merged_projet["Début"] = merged_projet["Successive_Date"]
merged_projet["Fin"] = merged_projet.groupby("ID_Projet")["Successive_Date"].shift(-1)

# Drop the 'Successive_Date' column, we no longer need it
merged_projet.drop(columns="Successive_Date", inplace=True)

# Save the modified DataFrame to 'fait_projetinformatique.xlsx'
new_file_name = "fait_projetinformatique.xlsx"
merged_projet.to_excel(new_file_name, index=False)

print(f"The Excel file '{new_file_name}' with the selected columns, 'Nom du Responsable' column, and the added 'Début' and 'Fin' columns has been generated successfully.")
