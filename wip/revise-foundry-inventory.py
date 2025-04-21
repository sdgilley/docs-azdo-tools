# shows files that are no longer in the repo but were there in (local version of) file-inventory.xlsx file
import pandas as pd
import os

# SharePoint link to the Excel file
# find script directory
script_dir = os.path.dirname(os.path.realpath(__file__))
file_path = "foundry-file-inventory.xlsx"
file_path = os.path.join(script_dir, file_path)

current = pd.read_excel(file_path, sheet_name="file-inventory")
# change ai-studio to ai-foundry in the file path:
current["File Path"] = current["File Path"].str.replace("ai-studio\\", "")
# Normalize the paths to use forward slashes
current["File Path"] = current["File Path"].str.replace("\\", "/")
# Print the columns of the DataFrame
print(current.columns)
print(f"Files in current: {len(current)}")
print(current["File Path"].head(10))

# now list the files in the local directory
local_repo = r"c:\GitPrivate\azure-ai-docs-pr\articles\ai-foundry"

# Recursively list all .md files in the local_repo directory
md_files = []
for root, dirs, files in os.walk(local_repo):
    for file in files:
        if file.endswith(".md"):
            md_files.append(os.path.relpath(os.path.join(root, file), local_repo).replace("\\", "/"))

# print the .md files in the local directory
print(f"Markdown files in local repo: {len(md_files)}")
print(md_files[:10])

# if the file name is not in the local_repo, don't keep it in current
# remove the files that are not in the local 
removed = current[~current['File Path'].isin(md_files)]
current = current[current['File Path'].isin(md_files)]
print(f"Files in current after removing not in local repo: {len(current)}")
print(f"Files removed from current: {len(removed)}")
print(removed["File Path"].sort_values())