'''
Find articles in need of freshness pass, create csv file with the list
To use this script:
1. Download the engagement report, filtered to the articles of interest
2. Save it in the same directory as this script
3. Open the report, and set the Sensitivity label to General. Save and close.
4. Fill in your details in the inputs section below
5. Sign in with az login --use-device-code
6. Run the script
7. Check the csv file created in the same directory as this script
8. Use the csv file as input to create-work-items.py
'''

# Finds the list of files that need to be refreshed for either this month or next month.
# creates a csv file with the list of files that need to be refreshed.
# !IMPORTANT - sign in with az login --use-device-code before running this script

import helpers.get_filelist as h
import helpers.fix_titles as f
import helpers.azdo as a
import pandas as pd
import os

################################## inputs
offset = 2  # for items going stale NEXT month, use offset = 2. 
            # for the CURRENT month's items use offset = 1
req = 90 # required freshness in days
freshness_title = f"Freshness - over {req}:"  # modify if you want a different title
suffix = " - Azure AI Foundry" # title suffix for your docs. Crucial for merging correctly.
eng_file = "Feb-Foundry-Engagement.xlsx" # Engagement file to read`
area_path = r"Content\Production\Core AI\AI Foundry" # where to query to find existing work items
# NOTE: you need to set the Sensitivity label to General on the Excel file!
csvfile = "Apr-foundry-work-items2.csv" # This is the file that will be created with the work items

################################## end of inputs

# Calculate the cutoff date - items older than this will be stale at the end of the period
now = pd.Timestamp.now()
end_of_month = now + pd.offsets.MonthEnd(offset)
cutoff_date = end_of_month - pd.Timedelta(days=req)
cutoff_date = cutoff_date.replace(day=1) # set to beginning of month
cutoff = cutoff_date.strftime('%#m/%#d/%Y')
print(f"Finding articles stale by end of {end_of_month.strftime('%B')}, LastReviewed {cutoff} or earlier")
# get script directory to read/write all files to same directory
script_dir = os.path.dirname(os.path.realpath(__file__))
csvfile = os.path.join(script_dir, csvfile)
eng_file = os.path.join(script_dir, eng_file)

#### Step 1 - read the engagement stats
# keep the columns that are needed in the create-work-items script
articles = pd.read_excel(eng_file, sheet_name="Export",
                           usecols=['Title', 'PageViews',
                                            'Url', 'MSAuthor', 'Freshness', 
                                            'LastReviewed', 'Engagement',
                                            'Flags', 'BounceRate', 'ClickThroughRate', 
                                            'CopyTryScrollRate'])
print(f"Total articles in engagement report: {articles.shape[0]}")
#### Step 2 - find existing work items and merge by title
print("Starting query for current work items...")
work_items = a.query_freshness(freshness_title, area_path, cutoff)

# fix the titles so that they match the metadata from the repo``
articles['Title'] = articles['Title'].apply(lambda x:f.fix_titles(x, suffix))
work_items['Title'] = work_items['Title'].apply(lambda x:f.fix_titles(x, suffix, freshness_title))

print(f"Existing work items found: {work_items.shape[0]}")
# # DEBUG: save work items to csv to figure out what happened...
# work_items.to_csv(os.path.join(script_dir,"debug-work-items.csv"), index=False)

# now merge to articles
articles = articles.merge(work_items, how='left', left_on='Title', right_on='Title')
# # DEBUG: save all items to csv to figure out what happened...
# articles.to_csv(os.path.join(script_dir,"debug-all-files.csv"), index=False)

# filter out ones with a work item already
articles = articles[articles['Id'].isnull()]
print(f"All articles without current work items, not filtered: {articles.shape[0]}")

#### Step 3 - filter out articles that are not stale
cutoff_date = pd.Timestamp(cutoff_date)

# filter out if date is later the cutoff date, these are not stale yet
articles = articles[articles['LastReviewed'] < cutoff_date]
print(f"Work item for {end_of_month.strftime('%B')}, LastReviewed before {cutoff_date.strftime('%m/%d/%Y')}: {articles.shape[0]}")
# save the resulting articles to a csv file -these need work items created
articles.to_csv(csvfile, index=False)
print(f"Saved to {csvfile}")