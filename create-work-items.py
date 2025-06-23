'''
Create work items for each row in a CSV or Excel file.
Use `find-stale-items.py` to create the CSV file with the list of items to be refreshed.
Or edit the engagement report file to keep only the rows that need a refresh.
If you use the engagement file, make sure to set the Sensitivity label to General.

To use this script, fill out the inputs section below. 
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
read_file = "Apr-foundry-work-items.csv"
sheet_names = ["CSV"]  # set to CSV if you are using a csv file
# ADO parameters
ado_url = "https://dev.azure.com/msft-skilling"
project_name = "Content"
item_type = "User Story"
area_path = r"Content\Production\Core AI\AI Foundry"
# iteration_path = r"Content\Selenium\FY25Q3"
iteration_path = r"Content\Bromine\FY26Q1\07 Jul" #the sprint you want to assign to
assignee = ''
parent_item = "414211"  # the ADO parent feature to link the new items to. Empty string if there is none.
freshness_title = "Freshness - over 90:  "
# Set mode to help set the fields that are saved into the work items
mode = "freshness"  # or "engagement or "empty"
# mode ="empty"  # or "engagement" or "freshness" 

#################### End of inputs ####################
# add the path to the excel file:
script_dir = os.path.dirname(__file__)
read_file = os.path.join(script_dir, read_file)

# ADO values for Content Engagement
if mode == "engagement":
    tags = ['content-engagement', 'Scripted']
    default_description = ("This auto-generated item was created to improve content engagement. "
                           "Review <a href='https://review.learn.microsoft.com/en-us/help/contribute/troubleshoot-underperforming-articles?branch=main'>"
                           "Troubleshoot lower-engaging articles</a> for tips. <br/><br/>The learn URL to improve is: ")
    default_title = "Improve engagement: "

# ADO Values for Freshness
if mode == "freshness" or mode == "empty":
    tags = ['content-health', 'freshness', 'Scripted']
    default_description = ("This auto-generated item was created to track a Freshness review. "
                           "Review <a href='https://review.learn.microsoft.com/en-us/help/contribute/freshness?branch=main'>"
                           "the freshness contributor guide page</a> for tips. <br/><br/>The learn URL to freshen up is: ")
    default_title = freshness_title

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

# Create work items
for row in all_rows:
    if mode == "empty":
        print(f"Processing row {row.get('filename', 'N/A')}")
        description = default_description
        description += f"<br/>{row.get('filename', '#')}<br/>"
        assignee = f"{row.get('ms.author', 'N/A')}@microsoft.com"

    else:
        print(f"Processing row {row.get('Url', 'N/A')}")

        description = default_description
        description += f"<br/><a href={row.get('Url', '#')} target=_new>{row.get('Url', 'N/A')}</a><br/>"
        description += "<table style='border: 1px solid black; border-collapse: collapse;'>"
        assignee = f"{row.get('MSAuthor', 'N/A')}@microsoft.com"

    if mode == "freshness":
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Freshness', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('LastReviewed', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>MSAuthor</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('MSAuthor', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>PageViews</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('PageViews', 'N/A')}</td></tr>"
    if mode == "engagement":
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Engagement</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Engagement', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Flags</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Flags', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>BounceRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('BounceRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>ClickThroughRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('ClickThroughRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>CopyTryScrollRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('CopyTryScrollRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Freshness', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('LastReviewed', 'N/A')}</td></tr>"
    if mode == "freshness" or mode == "engagement":
        # no table if mode is empty
        description += "</table><br/>"
        description += "Other page properties:<br/>"
        description += "<table style='border: 1px solid black; border-collapse: collapse;'>"

        for keyname in sorted(row.keys()):
            if keyname in ["Drilldown", "Trends", "GitHubOpenIssuesLink"]:
                description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>{keyname}</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'>"
                if row[keyname]:
                    description += f"<a href='{row[keyname]}' target='_new'>URL</a></td></tr>"
                else:
                    description += "</td></tr>"
            else:
                description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>{keyname}</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row[keyname]}</td></tr>"

        description += "</table>"

    work_item = [
        {
            'op': 'add',
            'path': '/fields/System.Title',
            'value': default_title + row["Title"]
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

    print(f"Created work item: {ado_url}/{project_name}/_workitems/edit/{created_item.id}")

print("Done!")