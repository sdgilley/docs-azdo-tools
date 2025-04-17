# Updates all the tabs of the build spreadsheet

from update_excel import update_spreadsheet

# Local file path to the Excel file
file_path = r"C:\Users\sgilley\OneDrive - Microsoft\AI Foundry\doc-updates-build2025.xlsx"
tabs = ["ai-foundry", "ai-services", "Images"] 
for tab in tabs: 
    update_spreadsheet(file_path, tab)
