# shows state of any work items which are for files that were deleted from PR 3360
# (Uses engagement report to find the titles since they're not in the repo anymore.)
# prints the work item title and state.  All should be removed or closed.
import requests
import os
import pandas as pd
import sys

# Add the helpers directory to the system path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'helpers')))

import azdo as a

def get_removed_files_from_pr(owner, repo, pr_number, token):
    url = f"https://api.github.com/repos/{owner}/{repo}/pulls/{pr_number}/files"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    files = response.json()

    removed_files = [file['filename'] for file in files if file['status'] == 'removed']
    return removed_files

if __name__ == "__main__":
    owner = "MicrosoftDocs"
    repo = "azure-ai-docs-pr"
    pr_number = 3360
    
    token = os.getenv("GH_ACCESS_TOKEN")  # Get the GitHub token from environment variable

    if not token:
        raise ValueError("GitHub token not found in environment variables. Please set GH_ACCESS_TOKEN.")

    removed_files = get_removed_files_from_pr(owner, repo, pr_number, token)

    # remove .md from the file names
    removed_files = [file.replace(".md", "") for file in removed_files]
    removed_files = [file.replace("articles/", "") for file in removed_files]
    removed_files = [file.replace(" - Azure AI Foundry", "") for file in removed_files]
    # print(f"Removed files in PR #{pr_number}:")
    # for file in removed_files:
    #     print(file)
        
    eng_file = "Feb-Foundry-Engagement.xlsx" # Engagement file to read
    article = pd.read_excel(eng_file, sheet_name="Export", usecols=['Title', 'Url'])
    
    # Print the Title if the URL contains one of the removed files
    with_titles = []
    for file in removed_files:
        for index, row in article.iterrows():
            url = row['Url']
            if pd.notna(url) and file in url:
                # add to with_titles
                with_titles.append((file, row['Title']))
    
    with_titles_df = pd.DataFrame(with_titles, columns=['File', 'Title'])
    # now query for the work items
    project_name = "Content"
    # provide vars needed for quer
    title_string = "Freshness - over 90:  "
    # make sure you list all the columns you want returned from the query
    columns = ['System.Title', 'System.State', 
                'System.IterationPath']
    query = f"""
        SELECT {','.join(columns)}
        FROM workitems
        WHERE [System.TeamProject] = '{project_name}'
        AND [System.Title] CONTAINS '{title_string}'
        """
    work_items = a.query_work_items(query, columns)
    print(f"Existing work items found: {work_items.shape[0]}")
    print(work_items['Title'].head(10))
    # Find the work items where the title contains one of the removed files
    for rtitle in with_titles_df['Title']:
        for index, row in work_items.iterrows():
            title = row['Title']

            if pd.notna(title) and rtitle in title:
                # print the work item
                print(f"Work item found for {rtitle}")
                print(row["State"])
    