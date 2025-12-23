# Merge freshness items from temp file with stale items
# Delete items that are Committed or New in temp file from stale items
# Merge the two by Title after stripping the prefix

# to run: first run query_work_items to create temp file
# run find-by-month to create stale items file
# then run this script to filter out items that are already in progress.

import pandas as pd
import os

# Get script directory
script_dir = os.path.dirname(__file__)

# File paths
temp_file = os.path.join(script_dir, 'temp-Oct-freshness_items.csv')
stale_file = os.path.join(script_dir, 'stale_items_November.csv')
output_file = os.path.join(script_dir, 'merged_freshness_items_November.csv')

# Read the CSV files
print(f"Reading {temp_file}...")
temp_df = pd.read_csv(temp_file)
print(f"Found {len(temp_df)} items in temp file")

print(f"Reading {stale_file}...")
stale_df = pd.read_csv(stale_file)
print(f"Found {len(stale_df)} items in stale file")

# Strip the prefix from temp file's Title column
prefix = "Freshness - over 90:  "
temp_df['Title_Cleaned'] = temp_df['Title'].str.replace(prefix, '', regex=False)

# Filter temp items that are Committed or New
committed_or_new = temp_df[temp_df['State'].isin(['Committed', 'New'])].copy()
print(f"Found {len(committed_or_new)} items that are Committed or New in temp file")

# Get the cleaned titles to exclude from stale_df
titles_to_exclude = set(committed_or_new['Title_Cleaned'])

# Filter out items from stale_df that match the titles to exclude
print(f"Removing items from stale file that are Committed or New in temp file...")
stale_df_filtered = stale_df[~stale_df['Title'].isin(titles_to_exclude)].copy()
print(f"Removed {len(stale_df) - len(stale_df_filtered)} items")
print(f"Remaining items in stale file: {len(stale_df_filtered)}")

# Save the result
stale_df_filtered.to_csv(output_file, index=False)
print(f"Saved merged results to {output_file}")
print(f"\nSummary:")
print(f"  Original stale items: {len(stale_df)}")
print(f"  Items removed (Committed or New): {len(stale_df) - len(stale_df_filtered)}")
print(f"  Final item count: {len(stale_df_filtered)}")
