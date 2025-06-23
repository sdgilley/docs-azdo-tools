# look for my files to see what the ms.date is, from my local repo
# 
# # !IMPORTANT - sign in with az login --use-device-code before running this script
# Finds the list of files that need to be refreshed for either this month or next month.
# creates a csv file with the list of files that need to be refreshed.
import helpers.get_filelist as h
import helpers.fix_titles as f
import pandas as pd
import os

################################## inputs
repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry" # your local repo

if os.path.exists(repo_path):
    # Checkout the branch and pull latest changes if needed...
    # h.checkout(repo_path, "main")
    dates_df = h.get_filelist(repo_path, "ms.date")
    author_df = h.get_filelist(repo_path, "author")
    # merge the dates and titles
    # merge the author
    articles = pd.merge(dates_df, author_df, on='filename')
    # # find mine
    articles = articles[articles['author'].str.contains("sdgilley")]
    # remove if path starts with "includes"
    articles = articles[~articles['filename'].str.contains("includes")]
    # sort by date
    articles = articles.sort_values(by=['ms.date'], ascending=False)
    # save to csv
    articles.to_csv("foundry-articles.csv", index=False)
    print("foundry-articles.csv created")

