import pandas as pd

# Data for the Excel file
data = [
    [f"{i}", f"Action_{i}"] for i in range(1, 13)
]
# Column names for the DataFrame
columns = ['ID', 'Nom Action']

# Create the DataFrame
df = pd.DataFrame(data, columns=columns)

# Save the DataFrame to an Excel file
file_name = "Action.xlsx"
df.to_excel(file_name, index=False)

print(f"The Excel file '{file_name}' containing the information has been generated successfully.")
