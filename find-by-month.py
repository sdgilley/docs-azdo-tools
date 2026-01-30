# Use this to generate freshness items for a given month
# # Finds the list of files that need to be refreshed for the given month, based
# on the ms.date field. Determines freshness based on ms.update-cycle if present,
# otherwise uses a default value.

import helpers.get_filelist as h
import helpers.utilities as f
# import helpers.azdo as a
import pandas as pd
import os
import calendar

################################## inputs
repo_path = "C:/git/azure-ai-docs-pr/articles/ai-foundry" # your local repo
suffix = " - Microsoft Foundry" # title suffix for your docs. Crucial for merging correctly.
eng_suffix = " - Microsoft Foundry" # suffix used in engagement file
for_month = 2 # month you are preparing for, 1-12
eng_file = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness/foundry-dec.csv" # your engagement file
default_cycle = 90 # default review cycle in days if not specified in the metadata
# todo next month: now that work item titles are changing, can also check if 
# current work item open for a given filename.  This month I did it manually.

##############################################

for_month_name = calendar.month_name[for_month]

# get script directory to read/write all files to same directory
# freshness_dir = os.path.dirname(os.path.realpath(__file__))
freshness_dir = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness"
csvfile = os.path.join(freshness_dir, f"stale_items_{for_month_name}.csv")

#### Step 1 - Get dates from the local repo - this is the most recent date, 
# since engagement is a month old.  Helps to cut out ones already updated.

# get most recent dates and other metadata from local repo
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
    # if ms.update-cycle is present, add it to the articles
    ms_update_cycle_df = h.get_filelist(repo_path, "ms.update-cycle")
    if not ms_update_cycle_df.empty:
        articles = pd.merge(articles, ms_update_cycle_df, on='filename', how='left')

    # remove includes
    articles = articles[~articles['filename'].str.contains("includes")]
    # only keep ms.service: azure-ai-foundry
    articles = articles[articles['ms.service'].str.contains("azure-ai-foundry", case=False)]
    #turn back slashes to forward in filenames
    articles['filename'] = articles['filename'].str.replace("\\", "/", regex=False)

else:
    print(f"Repository path {repo_path} does not exist. Please check the path.")
    exit(1)

#### Step 2 - Find articles that need review in the month of interest

# Convert ms.date to datetime
articles['ms.date'] = pd.to_datetime(articles['ms.date'], errors='coerce')

# Calculate last day of the review month in the current year
from datetime import datetime, timedelta
review_year = datetime.now().year
current_month = datetime.now().month

# If the target month is earlier in the year, it must be next year
if for_month < current_month:
    review_year += 1

last_day_of_month = calendar.monthrange(review_year, for_month)[1]
review_month_end = datetime(review_year, for_month, last_day_of_month)

# Determine review cycle for each article
def get_review_cycle(row):
    val = row.get('ms.update-cycle', None)
    if pd.notnull(val):
        try:
            return int(str(val).replace('-days','').replace('days','').strip())
        except:
            return default_cycle
    return default_cycle
articles['review_cycle'] = articles.apply(get_review_cycle, axis=1)

# Find articles that will be stale by the end of the review month
articles['stale_by_review_month'] = articles['ms.date'] + pd.to_timedelta(articles['review_cycle'], unit='d') <= review_month_end
articles = articles[articles['stale_by_review_month']]

print(f"Work items for {for_month_name}: {articles.shape[0]}")

### Step 3 - read the engagement stats
# keep the columns that are needed in the create-work-items script
engagement_columns = ['Title', 'PageViews', 'Url', 'MSAuthor', 'Freshness', 
                     'LastReviewed', 'Engagement', 'Flags', 'BounceRate', 'ClickThroughRate', 
                     'CopyTryScrollRate']
if eng_file.lower().endswith('.csv'):
    engagement = pd.read_csv(eng_file, usecols=engagement_columns)
else:
    engagement = pd.read_excel(eng_file, sheet_name="Export", usecols=engagement_columns)
# Merge the engagement stats with the articles
# build the URL from the filename
articles['Url'] = articles['filename'].apply(lambda x: f.build_url(x))
# fix the titles so that they match the metadata from the repo
articles['title'] = articles['title'].apply(lambda x:f.fix_titles(x, suffix))
# change the name from title to Title to match the engagement file
articles.rename(columns={'title': 'Title'}, inplace=True)
# fix the titles in the engagement file - use eng_suffix
engagement['Title'] = engagement['Title'].apply(lambda x: f.fix_titles(x, eng_suffix))

articles = pd.merge(articles, engagement, left_on='Url', right_on='Url', how='left')    

# Handle articles not found in engagement report
# Fill missing engagement data with "(not found)"
engagement_columns = ['PageViews', 'MSAuthor', 'Freshness', 'LastReviewed', 'Engagement',
                     'Flags', 'BounceRate', 'ClickThroughRate', 'CopyTryScrollRate']
for col in engagement_columns:
    articles[col] = articles[col].fillna("NA")

# Create URLs for articles not found in engagement report
# For missing URLs, construct from filename using the build_url function
articles['Url'] = articles['Url'].fillna(
    articles['filename'].apply(lambda x: f.build_url(x))
)

### Step 4 - write the list of articles to a csv file
articles.to_csv(csvfile, index=False)
print(f"Saved to {csvfile}")