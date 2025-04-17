### NOT YET WORKING USE WITH CARE AND SMALL TESTS!
import pandas as pd
import sys
import os
# Add the parent directory to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import helpers.azdo as a

# read spreadsheet
file_path = r"C:\Users\sgilley\OneDrive - Microsoft\AI Foundry\doc-updates-build2025.xlsx"
tab = "ai-foundry"
df = pd.read_excel(file_path, sheet_name=tab)
# only keep columns with work item and notes
work_items_df = df[['Work Item', 'Notes']].copy()
# filter out blank or nan work items
work_items_df = work_items_df[work_items_df['Work Item'].notna() & (work_items_df['Work Item'] != 'nan')].copy()
# filter out blank notes
work_items_df = work_items_df[work_items_df['Notes'].notna() & (work_items_df['Notes'] != '')].copy()
# Add a prefix and a new line to the Notes column
work_items_df['Notes'] = f"Notes from spreadsheet: {work_items_df['Notes']}"
print(work_items_df)
exit()
# Update the work items with the new values
a.add_to_discussion(work_items_df['Work Item'], work_items_df['Notes'])