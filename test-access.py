# Updates Excel spreadsheet 
# Queries from a Work Item column, updates the State and AssignedTo columns in the file
# Preserves formulas in the file

# !IMPORTANT - sign in with az login --use-device-code before running this script
import pandas as pd
from openpyxl import load_workbook
import helpers.azdo as a

# Local file path to the Excel file
file_path = r"C:\Users\sgilley\OneDrive - Microsoft\AI Foundry\doc-updates-build2025.xlsx"
tab="Articles"  # Specify the tab name here
# tab = "Images"
df = pd.read_excel(file_path, sheet_name=tab)
print(df.shape)

# Save the updated DataFrame back to the Excel file
# Load the original workbook with openpyxl
wb = load_workbook(file_path)
ws = wb[tab]  # Specify the tab to update
print(f"Loaded {tab} tab in {file_path}...")