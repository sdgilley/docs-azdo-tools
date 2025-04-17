'''
Functions to work with Azure DevOps (ADO) API
This module contains functions to authenticate with Azure DevOps and perform operations
!IMPORTANT - sign in with az login --use-device-code before running this script
''' 

# Import necessary libraries
from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
from azure.identity import DefaultAzureCredential
import os
import pandas as pd
from azure.devops.v7_0.work_item_tracking.models import Wiql

## Function to authenticate with Azure DevOps
## and return a connection object
def authenticate_ado(ado_url="https://dev.azure.com/msft-skilling"):  
    try:
        # Authenticate with Azure Active Directory (Entra ID)
        credential = DefaultAzureCredential()
        access_token = credential.get_token("499b84ac-1321-427f-aa17-267ca6975798/.default").token

        # Create a BasicAuthentication object with the access token
        credentials = BasicAuthentication('', access_token)

        # Connect to Azure DevOps
        connection = Connection(base_url=ado_url, creds=credentials)
        return connection
    except Exception as e:
        print("Error connecting to Azure DevOps:")
        print("Run `az login --use-device-code` before running this script.")
        exit(1)


def query_work_items(query, columns):
    # Authenticate with Azure DevOps
    connection = authenticate_ado()
    wit_client = connection.clients.get_work_item_tracking_client()

    # Define the WIQL query
    wiql_query = Wiql(query)

    # Execute the WIQL query
    wiql_result = wit_client.query_by_wiql(wiql=wiql_query)

    # # Debug: Print the raw response
    # print("WIQL Result:", wiql_result)

    if not wiql_result.work_items:
        print("No work items found.")
        return pd.DataFrame()

    # Fetch work item details
    work_item_ids = [item.id for item in wiql_result.work_items]
    # print("Work Item IDs:", work_item_ids)  # Debug: Print the IDs

    work_items = wit_client.get_work_items(ids=work_item_ids)

    # # Debug: Print the raw work items
    # for work_item in work_items:
    #     print("Work Item Object:", work_item)  # Debug: Print the full work item object

    # Create a DataFrame from the work items
    work_items_df = pd.DataFrame([{
        'Id': work_item.id,  # Explicitly include the System.Id from the work item reference
        'State': work_item.fields.get('System.State', ''),  # Map System.State
        'AssignedTo': work_item.fields.get('System.AssignedTo', {}).get('displayName', ''),  # Map System.AssignedTo
        **{
            column.replace('System.', ''): (
                work_item.fields.get(column, '') if column not in ['System.AssignedTo'] else ''
            )
            for column in columns if column not in ['System.Id', 'System.State', 'System.AssignedTo']
        }
    } for work_item in work_items])

    return work_items_df   

## function to query for freshness, supply the title you're looking for
def query_freshness(title_string, area_path, cutoff):
    project_name = "Content"

    # Define the WIQL query
    columns = ['System.Id', 'System.Title', 'System.State', 
               'System.CreatedDate', 'System.IterationPath', 'System.AssignedTo']
    # limit to work items created in the last 365 days, otherwise will be too large for ML queries
    # Also, keep closed items only if they were closed after the cutoff date
    # Ones closed prior to that need to be done again so should not be included in the results
    query=f"""
        SELECT {','.join(columns)}
        FROM workitems
        WHERE [System.TeamProject] = '{project_name}'
        AND [System.AreaPath] = '{area_path}'
        AND [System.CreatedDate] >= @StartOfMonth('-365d')
        AND [System.Title] CONTAINS '{title_string}'
        AND ([System.State] <> 'Removed')
        AND (
            [System.State] <> 'Closed' OR [Microsoft.VSTS.Common.ClosedDate] >= '{cutoff}'
        )

        """

    # execute the query
    work_items_df = query_work_items(query, columns)
    return work_items_df

# Function to add to the discussion of a work item
def add_to_discussion(work_items, new_values):
    # Authenticate with Azure DevOps
    connection = authenticate_ado()
    wit_client = connection.clients.get_work_item_tracking_client()

    # Add to the Discussion section of the work items
    for work_item_id, new_value in zip(work_items, new_values):
        # Debug: Print the current work item ID and new value
        # print(f"Processing work item ID: {work_item_id}, with value: {new_value}")

        # Create a JSON patch document to update the System.History field
        patch_document = [
            {
                "op": "add",
                "path": "/fields/System.History",
                "value": new_value
            }
        ]

        # Update the work item
        try:
            response = wit_client.update_work_item(
                document=patch_document,
                id=int(work_item_id)  # Ensure work_item_id is passed as an integer
            )
            print(f"Successfully added to Discussion for work item {work_item_id}: {new_value}")
        except Exception as e:
            print(f"Failed to update work item {work_item_id}: {e}")

# test it out
if __name__ == "__main__":
    work_items_df = pd.DataFrame([
        {
            'Id': 411019,
            'Notes': 'New discussion added by script'
        },
        {
            'Id': 375937,
            'Notes': 'Another new discussion added by script'
        }
    ])

    # Update the work items with the new values
    add_to_discussion(work_items_df['Id'], work_items_df['Notes'])
    exit()
    
    # columns = ['System.Id', 'System.State', 'System.AssignedTo']
    # query = f"""
    #     SELECT {','.join(columns)}
    #     FROM workitems 
    #     WHERE [System.TeamProject] = 'Content' AND [System.Id] IN (408127,408133)
    # """
    # print("WIQL Query:", query)

    # work_items = query_work_items(query, columns)
    # print(work_items)
    # exit()

# `   # Try update discussion:
#     project_name = "Content"
#     area_path = r"Content\Production\Core AI\AI Foundry" # where to query to find existing work items
#     # provide vars needed for query
#     title_string = "Freshness - over 90:  "
#     # call the query_freshness function:
#     work_items_df = query_freshness(title_string, area_path, cutoff="3/1/2025")
#     print(f"Freshness query, total work items found: {work_items_df.shape[0]}")
#     # save the resulting work items to a csv file
#     work_items_df.to_csv("freshness3.csv", index=False)
#     # check the results
#     # Convert cutoff to datetime for comparison
#     cutoff_date = pd.to_datetime("3/1/2025", format="%m/%d/%Y").tz_localize('UTC')
#     # save the resulting work items to a csv file
#     work_items_df.to_csv("debug-freshness.csv", index=False)