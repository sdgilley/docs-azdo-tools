import helpers.get_filelist as h

repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry" # your local repo

dates_df = h.get_filelist(repo_path, "ms.date")\
# find names of files that contain ai-foundry
dates_df = dates_df[dates_df['filename'].str.contains("ai-foundry", na=False)]

print(dates_df)