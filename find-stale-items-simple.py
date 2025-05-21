# This simply looks at dates in the local repo, lists all files and their dates.
# # Finds the list of files that need to be refreshed for either this month or next month.
# creates a csv file with the list of files that need to be refreshed.

import helpers.get_filelist as h
import helpers.fix_titles as f
import helpers.azdo as a
import pandas as pd
import os

################################## inputs
repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry" # your local repo
repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry" # your local repo
offset = 2  # for items going stale NEXT month, use offset = 2. 
            # for the CURRENT month's items use offset = 1
req = 90 # required freshness in days

################################## end of inputs
# Calculate the cutoff date - items older than this will be stale at the end of the period
now = pd.Timestamp.now()
end_of_month = now + pd.offsets.MonthEnd(offset)
cutoff_date = end_of_month - pd.Timedelta(days=req)
cutoff_date = cutoff_date.replace(day=1) # set to beginning of month
cutoff = cutoff_date.strftime('%#m/%#d/%Y')
print(f"Finding articles stale by end of {end_of_month.strftime('%B')}, LastReviewed before {cutoff}")
cutoff_date = pd.Timestamp(cutoff_date)


# get script directory to read/write all files to same directory
script_dir = os.path.dirname(os.path.realpath(__file__))
csvfile = os.path.join(script_dir, f"stale_items_{end_of_month.strftime('%Y-%m')}.csv")

#### Step 1 - Get dates from the local repo - this is the most recent date, 
# since engagement is a month old.  Helps to cut out ones already updated.

# get most recent dates from local repo
# Only do this step if the local path exists:
if os.path.exists(repo_path):
    # Checkout the branch and pull latest changes if needed...
    # h.checkout(repo_path, "main")
    dates_df = h.get_filelist(repo_path, "ms.date")
    titles_df = h.get_filelist(repo_path, "title")
    # merge the dates and titles
    articles = pd.merge(dates_df, titles_df, on='filename')
    authors_df = h.get_filelist(repo_path, "ms.author")
    # merge the authors
    articles = pd.merge(articles, authors_df, on='filename')
    # remove includes
    articles = articles[~articles['filename'].str.contains("includes")]


#### Step 2 - filter out articles that are not stale

# filter out if date is later the cutoff date, these are not stale yet
# convert LastReviewed to datetime
articles['ms.date'] = pd.to_datetime(articles['ms.date'], errors='coerce')
articles = articles[articles['ms.date'] < cutoff_date]
print(f"Work items for {end_of_month.strftime('%B')}, LastReviewed before {cutoff_date.strftime('%m/%d/%Y')}: {articles.shape[0]}")
# save the resulting articles to a csv file -these need work items created

### Step 3 - write the list of articles to a csv file
articles.to_csv(csvfile, index=False)
print(f"Saved to {csvfile}")