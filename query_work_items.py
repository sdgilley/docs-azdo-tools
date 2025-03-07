# !IMPORTANT - sign in with az login --use-device-code before running this script
# query work items from Azure DevOps using the Python SDK
# return a dataframe with the work items
# filters out work items created before freshness time period that are closed
# since they need freshness again
# see bottom for example usage
import os
import pandas as pd
import authenticate_ado as a
from azure.devops.v7_0.work_item_tracking.models import Wiql

def query_work_items(title_string, days=90):
    project_name = "Content"
    
    # Authenticate with Azure Active Directory (Entra ID)
    connection = a.authenticate_ado()
    wit_client = connection.clients.get_work_item_tracking_client()

    # Define the WIQL query
    # limit to work items created in the last 365 days, otherwise will be too large for ML queries
    wiql_query = Wiql(
        query=f"""
        SELECT [System.Id], [System.Title], [System.State], [System.CreatedDate], [System.IterationPath], [System.AssignedTo]
        FROM workitems
        WHERE [System.TeamProject] = '{project_name}'
        AND [System.CreatedDate] >= @StartOfMonth('-365d')
        AND [System.Title] CONTAINS '{title_string}'
        AND [System.State] <> 'Removed'
        """
    )

    # Execute the WIQL query
    wiql_result = wit_client.query_by_wiql(wiql=wiql_query)

    if not wiql_result.work_items:
        print("No work items found.")
        return pd.DataFrame()
    # Fetch the details of each work item
    work_item_ids = [item.id for item in wiql_result.work_items]
    work_items = wit_client.get_work_items(ids=work_item_ids)


    # Create a dataframe from the work items
    work_items_df = pd.DataFrame([{
        'ID': work_item.id, 
        'Title': work_item.fields['System.Title'], 
        'State': work_item.fields['System.State'],
        'Sprint': work_item.fields.get('System.IterationPath'),  # Use .get() to handle missing fields
        'CreatedDate': work_item.fields.get('System.CreatedDate'),  # Use .get() to handle missing fields
        'AssignedTo': work_item.fields.get('System.AssignedTo', {}).get('displayName', '') if work_item.fields.get('System.AssignedTo') else ''
    } for work_item in work_items])
    
    work_items_df['CreatedDate'] = pd.to_datetime(work_items_df['CreatedDate'], errors='coerce')
    # Remove closed items Created before the freshness time period - need freshness again.
    now_utc = pd.Timestamp.now(tz='UTC')
    cutoff_date = now_utc - pd.Timedelta(days=days)
    work_items_df = work_items_df[~((work_items_df['State'] == 'Closed') & (work_items_df['CreatedDate'] < cutoff_date))]

    return work_items_df

# Call the function
if __name__ == "__main__":
    title_string = "Freshness - over 90:  "
    work_items_df = query_work_items(title_string)
    script_dir = os.path.dirname(__file__)
    csv_file = os.path.join(script_dir, 'temp-mar-freshness_items.csv')
    work_items_df.to_csv(csv_file, index=False)
    print(f"Saved to {csv_file}")