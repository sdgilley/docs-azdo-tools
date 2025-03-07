# !IMPORTANT - sign in with az login --use-device-code before running this script
# query work items from Azure DevOps using the Python SDK
# return a dataframe with the work items
# see bottom for example usage
import os
import pandas as pd
import helpers.azdo as a

project_name = "Content"
# provide vars needed for quer
title_string = "Freshness - over 90:  "
# make sure you list all the columns you want returned from the query
columns = ['System.Id', 'System.Title', 'System.State', 
            'System.AssignedTo', 'System.IterationPath', 'System.CreatedDate']
query = f"""
    SELECT {','.join(columns)}
    FROM workitems
    WHERE [System.TeamProject] = '{project_name}'
    AND [System.Title] CONTAINS '{title_string}'
    AND [System.State] <> 'Removed'
    """
work_items = a.query_work_items(query, columns)
script_dir = os.path.dirname(__file__)
csv_file = os.path.join(script_dir, 'work_items.csv')
work_items.to_csv(csv_file, index=False)
print(f"{work_items.shape[0]} items saved to {csv_file}")
