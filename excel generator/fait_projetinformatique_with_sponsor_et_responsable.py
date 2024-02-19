import pandas as pd
import random
from datetime import datetime

# Read the existing 'fait_projetinformatique.xlsx' file
merged_projet = pd.read_excel("fait_projetinformatique_with_sponsor.xlsx")

# Function to generate random ID_Sponsor values
def get_random_sponsor_id():
    return random.choice([9, 10, 11,12,13])

# Add the ID_Sponsor column and fill it with random values
merged_projet["ID_Responsable_Projet"] = [get_random_sponsor_id() for _ in range(len(merged_projet))]

# Save the updated DataFrame to the Excel file
new_file_name = "zz.xlsx"
merged_projet.to_excel(new_file_name, index=False)

print(f"The Excel file '{new_file_name}' ")
