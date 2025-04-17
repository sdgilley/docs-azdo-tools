# docs-azdo-tools

Tools for querying Azure DevOps,  finding articles in need of freshness review in Microsoft Learn docs, and creating work items in Azure DevOps.

!IMPORTANT - Sign in first with `az login --use-device-code` before running these scripts to authenticate with Azure DevOps. It doesn't matter which subscription you choose, just that you sign in with your Microsoft account.

!IMPORTANT - Also, when reading/writing to Excel, the file must have General permissions.  Anything more restrictive than General will cause an error.

## Prerequisites

Open in Codespaces or clone the repo to your local machine.  You can run the scripts in either environment, but Codespaces is a great way to get started quickly.

* [![Open in Codespace](https://github.com/codespaces/badge.svg)](https://github.com/codespaces/new?repo=https://github.com/sdgilley/docs-azdo-tools)

    If you [Open in Codespace](https://github.com/codespaces/new?repo=https://github.com/sdgilley/docs-azdo-tools), upload an Engagement report, filtered to the documents of interest, into the main directory once the Codespace starts. Also, the first time you use it, you may be prompted to create a virtual environment and install the requirements.  Select any version of Python, then select the .devcontainer/requirements.txt when prompted.

* Local execution.  If running locally, you'll need:

    * Python 3.8 or later installed on your machine.  You can check this by running `py -3 --version` in a command prompt.
    * Azure CLI installed and authenticated with `az login --use-device-code`
    * Create a python virtual environment and install requirements:

        ```bash
        py -3 -m venv .venv
        .venv\scripts\activate
        pip install azure-devops msrest azure-identity pandas openpyxl
        ```

## Freshness scripts

Use these scripts to find articles in need of a freshness review, and to create work items in Azure DevOps for items that need to be updated.
Open each script and fill in the inputs before running.

* `find-stale-items.py`: Find the items that need to be added as work items.
    * Reads an engagement report that contains the documents you're interested in updating.
    * Finds the files that need to be refreshed by the end of this month or next month, depending on the input value.
    * Reads the ms.date from a local version of the repo.  This picks up changes made since the engagement report was run.  
    * Queries to see if a work item is already present for each file it finds.  
    * Outputs a csv file with the items that need a work item for freshness review for the given month.
    It's a good idea to look over this file before using the next script to create the items in DevOps. See full instructions for running in the script.

* `create-work-items.py`: 
    * Reads an Excel or csv file. One way to use this is to start with an export from the Engagement report, then remove rows that are not needed.  Or, use `find-stale-items.py` which creates a .csv instead, and filters out items that already have a work item associated with it.  
    * Creates a work item in Azure DevOps for each row in the file.

    The following columns are expected in the input file:
    * Url
    * MSAuthor
    * Freshness
    * LastReviewed
    * PageViews
    * Engagement
    * Flags
    * BounceRate
    * ClickThroughRate
    * CopyTryScrollRate

* `devops_query.py`: An example of executing a generic DevOps query. 
    * Runs any query specification and any columns to get a devops query result in a dataframe, 
    * Writes query results to a csv file. You get the same results as if you did the query in DevOps and saved the results, but this way you can do it in code instead!
    * Note: This is not used in the other scripts, but is a good example of how to query DevOps using Python.  It uses the `azdo.py` helper functions to authenticate and run the query.

## Helper functions

You don't need to run these directly, but they are used by the scripts above.  They are in the **helpers** directory.

* `azdo.py`: Functions to query work items in DevOps. Contains these functions:
    * `authenticate_ado`: Call to get a connection to the database.  This is used by other functions.
    * `query_work_items`: A generic query.  Input the columns you want returned and the query you want to run.  Returns a pandas dataframe with the results.
    * `freshness_query`: A specific query for freshness. Call with a title string to search for, and a days argument. The days argument is used to remove items that were created this many days ago, as it's time to re-do this file again.  Returns a pandas dataframe with the results. (Note: only finds the last year's worth of data, to avoid errors with too many results.)
* `fix_titles.py`: Function to standardize the text and format of titles to be used prior to a merge. Supply the title string to be fixed, the suffix that is added to the title found in the file metadata, and optionally, a prefix.  Both prefix and suffix will be removed from the output title.
* `get_filelist.py`: Function to get a list of all files in the local repository. Call with arguments to get the metadata, then merge by the file URL. Used in the initial version of the `find-stale-items.py` script to get the most recent date for each file.

## Spreadsheet tools (spreadsheet-tools/ folder)

* `update_excel.py`: This script updates a spreadsheet that has a `Work Item` column with the status of work items.  It uses the `azdo.py` helper functions to query DevOps for the status of each work item. The script is set up to run on a specific spreadsheet, but you can modify it to run on any spreadsheet you want, as long as it contains the `Work Item` column, which is used to find the work items in DevOps. 
* `update-build-spreadsheets.py`: Runs the `update_excel.py` script for each of the tabs in the spreadsheet that requires updating. 

## Archive files

* `CreateWorkitemsFromExcelFile.ps1`: PowerShell script to create work items in Azure DevOps. Requires a PAT from DevOps. I developed the `create-work-items.py` script based on this script.  (And by "I", I mean mostly "Copilot".) I'm leaving the script here for reference, but I was unable to use it.  It requires a PAT for DevOps that I couldn't figure out how to create anymore.  Also, I've made lots of little modifications to the Python version. The Python version works with Entra ID, and does not require a PAT.  

* `find-stale-items-2step.py`: This version of the script doesn't look at local dates.  For awhile, I thought that was not needed. But there are cases where it is needed, so the `find-stale-items.py` script is the one to use.  
