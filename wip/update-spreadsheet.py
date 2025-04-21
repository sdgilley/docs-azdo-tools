import os
import pandas as pd
import helpers.azdo as a

# Local file path to the Excel file
file_path = r"C:\Users\sgilley\OneDrive - Microsoft\doc-updates-build2025.xlsx"
df = pd.read_excel(file_path)

# Ensure Work Item column is of string type before using .str.replace
df['Work Item'] = df['Work Item'].astype(str)

# Remove prefix from the Work Item column
prefix = "https://dev.azure.com/msft-skilling/Content/_queries/edit/"
df['Work Item'] = df['Work Item'].str.replace(prefix, '', regex=False)

# Filter out rows with NaN or invalid Work Item values
df = df[df['Work Item'] != 'nan']

# Define the Azure DevOps query
project_name = "Content"
columns = ['System.Id', 'System.State', 'System.AssignedTo']
query = f"""
    SELECT {','.join(columns)}
    FROM workitems 
    WHERE [System.TeamProject] = 'Content' AND [System.Id] IN ({','.join(df['Work Item'])})
    """
# print("WIQL Query:", query)

# Query work items
work_items = a.query_work_items(query, columns)

# Debug: Print the results to verify the fields
# print(work_items.head())

# Replace the existing State and AssignedTo columns with the updated values
df.update(work_items.rename(columns={'Id': 'Work Item'}))

# Save the updated DataFrame back to the Excel file
df.to_excel(file_path, index=False)
print(f"Updated spreadsheet saved to {file_path}")