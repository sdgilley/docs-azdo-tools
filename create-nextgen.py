'''
Create work items for each row in a CSV or Excel file.
Use `find-stale-items.py` to create the CSV file with the list of items to be refreshed.
Or edit the engagement report file to keep only the rows that need a refresh.
If you use the engagement file, make sure to set the Sensitivity label to General.

To use this script, fill out the inputs section below. 
If you aren't already signed in, run `az login --use-device-code` to authenticate.

Work items are created in the sprint specified in the inputs section, 
and assigned to the MSAuthor found in the spreadsheet.
This script reads the nextgen-items.csv file and creates NextGen work items.
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
read_file = "nextgen-items.csv"
sheet_names = ["CSV"]  # set to CSV if you are using a csv file
# ADO parameters
ado_url = "https://dev.azure.com/msft-skilling"
project_name = "Content"
item_type = "User Story"
area_path = r"Content\Production\Core AI\AI Foundry Core"
# iteration_path = r"Content\Selenium\FY25Q3"
iteration_path = r"Content\FY26\Q2\10 Oct" #the sprint you want to assign to
assignee = ''
parent_item = "492082"  # the ADO parent feature to link the new items to. Empty string if there is none.


#################### End of inputs ####################
# add the path to the excel file:
script_dir = os.path.dirname(__file__)
read_file = os.path.join(script_dir, read_file)


# ADO Values 

tags = ['Ignite', 'Scripted']
default_title = "NextGen | update "  # space at the end to separate from filename

default_description = ("This auto-generated item was created to track article updates for NextGen.")
article_prefix = '<br/><br/>Article to update for NextGen:  '
article_suffix ='<br/>Use the release branch <b>release-ignite-foundry-nextgen</b> for your updates.'
instructions1 = "<br/>1. Update the article, adding monikerRange and NextGen information. See <a href='https://aka.ms/foundry-content-instructions' target=_new>https://aka.ms/foundry-content-instructions</a> for additional details."
instructions2 = "<br/>2. Add to NextGen TOC. (Suggestion, modify as you see fit): "

instructions_final = "<br/>If you change the TOC value, (or if above suggestion is blank), update the value in the spreadsheet: <a href='https://aka.ms/foundry-spreadsheet' target=_new>https://aka.ms/foundry-spreadsheet</a>."

# Read the Excel file
all_rows = []
if sheet_names == ["CSV"]:
    df = pd.read_csv(read_file)
    all_rows.extend(df.to_dict(orient='records'))
else:
    for sheet_name in sheet_names:
        df = pd.read_excel(read_file, sheet_name=sheet_name)
        all_rows.extend(df.to_dict(orient='records'))

# Print the keys of the first row for debugging
# if all_rows:
#     print("Keys in the first row:", all_rows[0].keys())
# exit()

connection = a.authenticate_ado()
wit_client = connection.clients.get_work_item_tracking_client()

# List to store created work item details
created_items = []

# Create work items
for row in all_rows:

    print(f"Processing row {row.get('filename', 'N/A')}")

    description = default_description
    # add article link and suffix
    description += article_prefix + f"<a href={row.get('URL', '#')} target=_new>{row.get('URL', 'N/A')}</a><br/>" + article_suffix 
    # add notes if present
    description += f"<br/>Notes: {row.get('Notes', 'N/A')}<br/>"
    # add instructions
    description += instructions1
    description += instructions2 + f"<b>&nbsp;{row.get('NextGen TOC', 'N/A')}</b>" + instructions_final
    assignee = f"{row.get('ms.author', 'N/A')}@microsoft.com"


    work_item = [
        {
            'op': 'add',
            'path': '/fields/System.Title',
            'value': default_title + row["filename"]
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
            'path': '/fields/Microsoft.VSTS.Scheduling.StoryPoints',
            'value': 2
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
        }
    ]

    created_item = wit_client.create_work_item(document=work_item, project=project_name, type=item_type)

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

    # Store work item details for CSV
    created_items.append({
        'Work_Item_ID': created_item.id,
        'Article_URL': row.get('URL', 'N/A'),
        'Title': default_title + row["filename"],
        'ADO_URL': f"{ado_url}/{project_name}/_workitems/edit/{created_item.id}",
        'Assignee': assignee,
        'Filename': row["filename"],
        'Area_Path': area_path,
        'Iteration_Path': iteration_path,
        'Tags': ','.join(tags)
    })

    print(f"Created work item: {ado_url}/{project_name}/_workitems/edit/{created_item.id}")

# Create CSV file with all created work items
output_csv = os.path.join(script_dir, "created_work_items.csv")
created_items_df = pd.DataFrame(created_items)
created_items_df.to_csv(output_csv, index=False)

print(f"Done! Created {len(created_items)} work items.")
print(f"Details saved to: {output_csv}")