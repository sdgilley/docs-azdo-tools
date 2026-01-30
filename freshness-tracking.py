# This simply looks at dates in the local repo, lists all files and their dates.
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
tracking_file = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness/FreshnessTracking.xlsx"
eng_file = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness/foundry-dec.csv" # your engagement file
freshness_dir = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness"
##############################################

csvfile = os.path.join(freshness_dir, f"tracking_engagement.csv")
freshness = pd.read_excel(tracking_file, sheet_name="Sheet1")
print(freshness.head()  )

# keep the columns that are needed in the create-work-items script
engagement_columns = ['Title', 'PageViews', 'Url', 'MSAuthor', 'Freshness', 
                     'LastReviewed', 'Engagement', 'Flags', 'BounceRate', 'ClickThroughRate', 
                     'CopyTryScrollRate']
if eng_file.lower().endswith('.csv'):
    engagement = pd.read_csv(eng_file, usecols=engagement_columns)
else:
    engagement = pd.read_excel(eng_file, sheet_name="Export", usecols=engagement_columns)
# Merge the engagement stats with the freshness tracking
# build the URL from the filename
freshness['Url'] = freshness['Article'].apply(lambda x: f.build_url(x))

freshness = pd.merge(freshness, engagement, left_on='Url', right_on='Url', how='left')    

### Step 4 - write the list of articles to a csv file
print(freshness.head())
freshness.to_csv(csvfile, index=False)
print(f"Saved to {csvfile}")