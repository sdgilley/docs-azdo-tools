'''
Create work items for each row in a CSV or Excel file.
use the foundry-file-inventory.xlsx file for input
(Save it here when you're ready to run the script.)

If you aren't already signed in, run `az login --use-device-code` to authenticate.

Work items are created in the sprint specified in the inputs section, 
and assigned to the MSAuthor found in the spreadsheet.
'''
# Read the read_file and create a work item for each row
# Start with engagement report excel file
# Remove rows until it contain only those you want to create work items for
import pandas as pd
import helpers.azdo as a
import os

#################### Inputs ####################
# all files are read and created from the same directory as this script
# Input Excel file name and sheet name - each row contains the data for a work item
# NOTE: you need to set the Sensitivity label to General if you are using an Excel file
# if you use Excel, set sheet_name.  If it came from engagement report, it's "Export"
# read_file = "feb-work-items.xlsx"
# sheet_names = ["Export"]
# # or input a csv file and set sheet_names to CSV.
read_file = "foundry-file-inventory.xlsx"
sheet_names = ["file-inventory"]  # set to CSV if you are using a csv file
# ADO parameters
ado_url = "https://dev.azure.com/msft-skilling"
project_name = "Content"
item_type = "User Story"
area_path = r"Content\Production\Core AI\AI Foundry"
# iteration_path = r"Content\Selenium\FY25Q3"
iteration_path = r"Content\Bromine\FY25Q4\04 Apr" #the sprint you want to assign to
assignee = ''
parent_item = "375929"  # the ADO parent feature to link the new items to. Empty string if there is none.
default_title = "Foundry 1RP | "
# Set mode to help set the fields that are saved into the work items

#################### End of inputs ####################
# add the path to the excel file:
script_dir = os.path.dirname(__file__)
read_file = os.path.join(script_dir, read_file)

### CHECK WHICH TAGS WE WANT
tags = ['1RP', 'Build 2025', 'Scripted']
default_description = ("This auto-generated item was created from the foundry-file-inventory.xlsx file. "
                       "<br/>Update to reflect the new 1RP project.  If functionality differs between old and new, keep old project info as well."
                       "<br/>Check with Sheri for includes to use if article applies to old project only or to new project only."
                       "<br/>Make all changes on the BUILD release branch."
                       "<br/><br/>The learn URL to update is: ")

# Read the Excel file
all_rows = []
if sheet_names == ["CSV"]:
    df = pd.read_csv(read_file)
    all_rows.extend(df.to_dict(orient='records'))
else:
    for sheet_name in sheet_names:
        df = pd.read_excel(read_file, sheet_name=sheet_name)
        all_rows.extend(df.to_dict(orient='records'))

print(f"Total rows in {read_file}: {len(all_rows)}")
# Print the keys of the first row for debugging
# if all_rows:
#     print("Keys in the first row:", all_rows[0].keys())
# exit()

connection = a.authenticate_ado()
wit_client = connection.clients.get_work_item_tracking_client()
print("Connected to Azure DevOps")


# Create work items
for index, row in enumerate(all_rows):
    if row.get('Needs 1RP revision') != 'Yes':
        continue
        # Skip rows that don't need 1RP revision
    print(f"Processing row {row.get('File Path', 'N/A')}")
    description = default_description
    description += f"<br/><a href={row.get('URL', '#')} target=_new>{row.get('URL', 'N/A')}</a><br/>"
    description += f"<br/>Additional Notes: <br/>{row.get('Notes', 'N/A')}<br/>"
    assignee = f"{row.get('Author', 'N/A')}@microsoft.com"
    points = row.get('Story Points', '1')

    # Ensure no `nan` values in the payload
    description = description if pd.notna(description) else "No description provided."
    assignee = assignee if pd.notna(assignee) else ""
    points = points if pd.notna(points) else "1"

    work_item = [
        {
            'op': 'add',
            'path': '/fields/System.Title',
            'value': default_title + row["File Path"]
        },
        {
            'op': 'add',
            'path': '/fields/System.AreaPath',
            'value': area_path
        },
        {
            'op': 'add',
            'path': '/fields/System.IterationPath',
            'value': iteration_path
        },
        {
            'op': 'add',
            'path': '/fields/System.Tags',
            'value': ','.join(tags)
        },
        {
            'op': 'add',
            'path': '/fields/System.Description',
            'value': description
        },
        {
            'op': 'add',
            'path': '/fields/System.AssignedTo',
            'value': assignee
        },
                {
            'op': 'add',
            'path': '/fields/Microsoft.VSTS.Scheduling.StoryPoints',
            'value': points
        }
    ]

    # Debugging: Print the payload before sending
    # print("Work item payload:", work_item)

    try:
        created_item = wit_client.create_work_item(document=work_item, project=project_name, type=item_type)
    except Exception as e:
        print(f"Failed to create work item for row {index}: {e}")
        continue
    # Link to a parent item if there's a number to link to
    if parent_item:
        wit_client.update_work_item(
            document=[
                {
                    "op": "add",
                    "path": "/relations/-",
                    "value": {
                        "rel": "System.LinkTypes.Hierarchy-Reverse",
                        "url": f"{ado_url}/{project_name}/_apis/wit/workItems/{parent_item}"
                    }
                }
            ],
            id=created_item.id
        )
    
    created_link = f"{ado_url}/{project_name}/_queries/edit/{created_item.id}"
    print(f"Created work item: {created_link}")
    # add the link to the work item in the excel file
    df.at[index, 'ADO Link'] = created_link

print("Done!")
 # save the file with the new link
df.to_csv("work-items-created.csv", index=False)
print(f"Saved the file with the new links to work-items-created.csv")
