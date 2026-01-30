'''
Create work items for each row in a CSV or Excel file.
Use `find-stale-items.py` to create the CSV file with the list of items to be refreshed.
Or edit the engagement report file to keep only the rows that need a refresh.
If you use the engagement file, make sure to set the Sensitivity label to General.

To use this script, fill out the inputs section below. 
If you aren't already signed in, run `az login --use-device-code` to authenticate.

Work items are created in the sprint specified in the inputs section, 
and assigned to the MSAuthor found in the spreadsheet.
Instructions to create vendor work items are included in the description.
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
read_file = r"C:\Users\sgilley\OneDrive - Microsoft\AI Foundry\Freshness\work-items-feb.csv"
sheet_names = ["CSV"]  # set to CSV if you are using a csv file
# ADO parameters
ado_url = "https://dev.azure.com/msft-skilling"
project_name = "Content"
item_type = "User Story"
area_path = r"Content\Production\Core AI\AI Foundry Core"
iteration_path = r"Content\FY26\Q3\02 Feb" #the sprint you want to assign to
assignee = ''
parent_item = "414211"  # the ADO parent feature to link the new items to. Empty string if there is none.
freshness_title_90 = "Freshness - over 90:  "
freshness_title_180 = "Freshness - over 180:  "
# Set mode to help set the fields that are saved into the work items
mode = "freshness"  # or "engagement or "empty"
# mode ="empty"  # or "engagement" or "freshness" 

#################### End of inputs ####################
# add the path to the excel file:
script_dir = os.path.dirname(__file__)
read_file = os.path.join(script_dir, read_file)

vendor_description = "Leverage the vendors for freshness work wherever it's possible and makes sense to do so. Create <a href='https://dev.azure.com/msft-skilling/Content/_workitems/create/Feature?templateId=dcf45a43-b6de-4660-950e-fdaa62dc4a30&ownerId=c4a28f90-17ae-4384-b514-7273392b082b' target=_new>one work item</a> for 15 or less articles." 

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
                           "the freshness contributor guide page</a> for tips." 
                           f"<br>{vendor_description}<br/><br/>The learn URL to freshen up is: ")
    default_title = freshness_title_90  # Will be overridden per-row based on Freshness value

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
    # Determine effective mode - use empty if PageViews is NA/missing in freshness mode
    effective_mode = mode
    if mode == "freshness" and str(row.get('PageViews', 'NA')).strip() == 'NA':
        effective_mode = "empty"
    
    if effective_mode == "empty":
        print(f"Processing row {row.get('filename', row.get('Url', 'N/A'))}")
        description = default_description
        description += f"<br/>{row.get('filename', row.get('Url', '#'))}<br/>"
        author = row.get('ms.author', 'N/A')
        author = author.strip() if isinstance(author, str) else author
        assignee = f"{author}@microsoft.com"

    else:
        print(f"Processing row {row.get('Url', 'N/A')}")

        description = default_description
        description += f"<br/><a href={row.get('Url', '#')} target=_new>{row.get('Url', 'N/A')}</a><br/>"
        description += "<table style='border: 1px solid black; border-collapse: collapse;'>"
        author = row.get('ms.author', 'N/A')
        author = author.strip() if isinstance(author, str) else author
        assignee = f"{author}@microsoft.com"

    if effective_mode == "freshness":
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Freshness', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('LastReviewed', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>MSAuthor</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('MSAuthor', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>PageViews</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('PageViews', 'N/A')}</td></tr>"
    if effective_mode == "engagement":
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Engagement</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Engagement', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Flags</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Flags', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>BounceRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('BounceRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>ClickThroughRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('ClickThroughRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>CopyTryScrollRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('CopyTryScrollRate', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('Freshness', 'N/A')}</td></tr>"
        description += f"<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {row.get('LastReviewed', 'N/A')}</td></tr>"
    if effective_mode == "freshness" or effective_mode == "engagement":
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

    # Determine the title prefix based on Freshness value
    if effective_mode == "freshness" or effective_mode == "empty":
        try:
            freshness_days = int(row.get('review_cycle', 0))
            if freshness_days >= 180:
                title_prefix = freshness_title_180
            else:
                title_prefix = freshness_title_90
        except (ValueError, TypeError):
            title_prefix = freshness_title_90
    else:
        title_prefix = default_title

    # # TEST MODE: Print titles and skip work item creation
    # print(f"[TEST] Title: {title_prefix + row['Title']} | review_cycle: {row.get('review_cycle', 'N/A')}")
    # continue  # Skip work item creation
    # # END TEST MODE

    work_item = [
        {
            'op': 'add',
            'path': '/fields/System.Title',
            'value': title_prefix + row["Title"]
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

    # Store work item details for CSV
    work_item_data = {
        'Work_Item_ID': created_item.id,
        'Title': default_title + row.get("Title", row.get("filename", "N/A")),
        'URL': f"{ado_url}/{project_name}/_workitems/edit/{created_item.id}",
        'Assignee': assignee,
        'Mode': effective_mode,
        'Area_Path': area_path,
        'Iteration_Path': iteration_path,
        'Tags': ','.join(tags)
    }
    
    # Add mode-specific fields
    if effective_mode == "empty":
        work_item_data['Article_Filename'] = row.get('filename', row.get('Url', 'N/A'))
        author_value = row.get('ms.author', 'N/A')
        author_value = author_value.strip() if isinstance(author_value, str) else author_value
        work_item_data['MS_Author'] = author_value
    else:
        work_item_data['Article_URL'] = row.get('Url', 'N/A')
        author_value = row.get('ms.author', 'N/A')
        author_value = author_value.strip() if isinstance(author_value, str) else author_value
        work_item_data['MS_Author'] = author_value
        if effective_mode == "freshness":
            work_item_data['Freshness'] = row.get('Freshness', 'N/A')
            work_item_data['LastReviewed'] = row.get('LastReviewed', 'N/A')
            work_item_data['PageViews'] = row.get('PageViews', 'N/A')
        elif effective_mode == "engagement":
            work_item_data['Engagement'] = row.get('Engagement', 'N/A')
            work_item_data['BounceRate'] = row.get('BounceRate', 'N/A')
            work_item_data['ClickThroughRate'] = row.get('ClickThroughRate', 'N/A')
    
    created_items.append(work_item_data)

    print(f"Created work item: {ado_url}/{project_name}/_workitems/edit/{created_item.id}")

# Create CSV file with all created work items
output_csv = os.path.join(script_dir, f"created_work_items_{mode}.csv")
created_items_df = pd.DataFrame(created_items)
created_items_df.to_csv(output_csv, index=False)

print(f"Done! Created {len(created_items)} work items.")
print(f"Details saved to: {output_csv}")