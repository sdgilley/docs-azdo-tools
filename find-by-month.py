# This simply looks at dates in the local repo, lists all files and their dates.
# # Finds the list of files that need to be refreshed for the given month, based
# on the ms.date field. Looks only at the month.

import helpers.get_filelist as h
import helpers.fix_titles as f
import helpers.azdo as a
import pandas as pd
import os
import calendar

################################## inputs
repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry" # your local repo
suffix = " - Azure AI Foundry" # title suffix for your docs. Crucial for merging correctly.
for_month = 7 # month you are preparing for, 1-12
eng_file = "C:/Git/docs-azdo-tools/sprint-planning/Foundry-engagement-May.xlsx" # your engagement file
freshness = 3 # freshness in months
##############################################

for_month_name = calendar.month_name[for_month]
freq_month = freshness - 1 # subtract 1 to get the month review needs to take place.

# get script directory to read/write all files to same directory
script_dir = os.path.dirname(os.path.realpath(__file__))
csvfile = os.path.join(script_dir, f"stale_items_{for_month_name}.csv")

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
    service_df = h.get_filelist(repo_path, "ms.service")
    # merge the service data
    articles = pd.merge(articles, service_df, on='filename')    
    # remove includes
    articles = articles[~articles['filename'].str.contains("includes")]
    # only keep ms.service: azure-ai-foundry
    articles = articles[articles['ms.service'].str.contains("azure-ai-foundry", case=False)]

#### Step 2 - Find articles that need review in the month of interest
review_month = for_month - freq_month if for_month > freq_month else for_month + 12 - freq_month

# switch ms.date to datetime
articles['ms.date'] = pd.to_datetime(articles['ms.date'], errors='coerce')
# get the month from the ms.date:
articles['month'] = articles['ms.date'].dt.month
# keep articles where month is the review_month
articles = articles[articles['month'] == review_month]

print(f"Work items for {for_month_name}: {articles.shape[0]}")

### Step 3 - read the engagement stats
# keep the columns that are needed in the create-work-items script
engagement = pd.read_excel(eng_file, sheet_name="Export",
                           usecols=['Title', 'PageViews',
                                            'Url', 'MSAuthor', 'Freshness', 
                                            'LastReviewed', 'Engagement',
                                            'Flags', 'BounceRate', 'ClickThroughRate', 
                                            'CopyTryScrollRate'])
# Merge the engagement stats with the articles
# fix the titles so that they match the metadata from the repo
articles['title'] = articles['title'].apply(lambda x:f.fix_titles(x, suffix))
# change the name from title to Title to match the engagement file
articles.rename(columns={'title': 'Title'}, inplace=True)
# fix the titles in the engagement file
engagement['Title'] = engagement['Title'].apply(lambda x: f.fix_titles(x, suffix))
articles = pd.merge(articles, engagement, left_on='Title', right_on='Title', how='left')    

### Step 4 - write the list of articles to a csv file
articles.to_csv(csvfile, index=False)
print(f"Saved to {csvfile}")