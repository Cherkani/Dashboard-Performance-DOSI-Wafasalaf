import pandas as pd
import random

# Data for the Excel file
data = [
    [f"DEV{i}", "dev", f"dev{i}@gmail.com"] for i in range(1, 4)
] + [
    [f"Sécurité{i}", "sécurité", f"securite{i}@gmail.com"] for i in range(1, 6)
] + [
    [f"CPI_{i}", "cpi", f"cpi{i}@gmail.com"] for i in range(1, 6)
] + [
    [f"Support Technique {i}", "Support Technique", f"SupportTechnique{i}@gmail.com"] for i in range(1, 21)
] + [
    [f"Client {i}", "client", f"client{i}@gmail.com"] for i in range(1, 21)
]

# Column names for the DataFrame
columns = ['ID_Responsable', 'Nom Responsable', 'Fonction', 'Email']

# Create a DataFrame from the data
df = pd.DataFrame(data, columns=columns[1:])  # Exclude 'ID' column for now

# Generate IDs for each row and add them to the DataFrame
ids = range(1, len(df) + 1)
df.insert(0, 'ID', ids)  # Insert 'ID' column at the beginning

# Generate random phone numbers that start with "06" and add them to the DataFrame
random_phone_numbers = [f"06{random.randint(10000000, 99999999)}" for _ in range(len(df))]
df['Téléphone'] = random_phone_numbers

# Save the DataFrame to an Excel file
file_name = "responsables.xlsx"
df.to_excel(file_name, index=False)

print(f"The Excel file '{file_name}' containing the information has been generated successfully.")
