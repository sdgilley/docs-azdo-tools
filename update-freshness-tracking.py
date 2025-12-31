# Monthly script to track engagement data - adds each month as a new row
# Each row contains: Article, UpdateDate, Month, PageViews, Engagement, etc.

import helpers.utilities as f
import pandas as pd
import os
from datetime import datetime
import calendar
import re

################################## inputs
tracking_file = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness/FreshnessTracking.xlsx"
eng_file = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness/foundry-nov.csv"  # current month's engagement file
freshness_dir = "C:/Users/sgilley/OneDrive - Microsoft/AI Foundry/Freshness"
output_file = os.path.join(freshness_dir, "FreshnessTrackingEngagement.csv")

# Define month and year of the engagement data to add to the tracking spreadsheet
target_month = 11 # 1-12, or None for current 
target_year = 2025  # or None for current

##############################################

# Get month/year
month_str = f"{calendar.month_name[target_month]} {target_year}"

print(f"Processing engagement data for {month_str}")

# Load tracking file (articles to track)
if os.path.exists(tracking_file):
    tracking = pd.read_excel(tracking_file, sheet_name="Sheet1")
    print(f"Loaded {len(tracking)} articles from tracking file")
else:
    print(f"Tracking file not found: {tracking_file}")
    exit(1)

# Load current month's engagement data
engagement_columns = ['Title', 'PageViews', 'Url', 'MSAuthor', 
                     'Engagement', 'Flags', 'BounceRate', 'ClickThroughRate', 
                     'CopyTryScrollRate']

if eng_file.lower().endswith('.csv'):
    engagement = pd.read_csv(eng_file, usecols=engagement_columns)
else:
    engagement = pd.read_excel(eng_file, sheet_name="Export", usecols=engagement_columns)

print(f"Loaded {len(engagement)} articles from engagement file")


# Build URLs from article filenames in tracking file
# Remove leading slashes from Article names before building URL
tracking['Url'] = tracking['Article'].apply(lambda x: f.build_url(x.lstrip('/') if isinstance(x, str) else x))

# Normalize engagement URLs to handle potential mismatches (e.g., with /default/)
def normalize_url(url):
    # Remove /default/ from URLs to match tracking URLs
    if pd.isna(url):
        return url
    return re.sub(r'/default/', '/', str(url))

engagement['Url'] = engagement['Url'].apply(normalize_url)

# Convert Date column to datetime if it exists
if 'Date' in tracking.columns:
    tracking['Date'] = pd.to_datetime(tracking['Date'], errors='coerce')
    tracking.rename(columns={'Date': 'UpdatedDate'}, inplace=True)

# Merge engagement data with tracking articles
merged = pd.merge(tracking, engagement, left_on='Url', right_on='Url', how='left')

# Add Month column
merged['Month'] = month_str

# Calculate M-value (months since update)
def get_m_value(update_date, target_month, target_year):
    if pd.isna(update_date):
        return None
    months_since = (target_year - update_date.year) * 12 + (target_month - update_date.month)
    # M0 = month of update, M-1 = month before, M1 = month after, etc.
    return f"M{months_since}"

merged['FromUpdate'] = merged['UpdatedDate'].apply(lambda x: get_m_value(x, target_month, target_year))

# Select columns to keep: Article, UpdatedDate, Month, FromUpdate, and engagement stats
columns_to_keep = ['Article', 'UpdatedDate', 'Month', 'FromUpdate'] + [col for col in engagement_columns if col not in ['Title', 'Url']]
result = merged[columns_to_keep].copy()

# If output file exists, append to it; otherwise create new
if os.path.exists(output_file):
    existing = pd.read_csv(output_file)
    result = pd.concat([existing, result], ignore_index=True)
    print(f"Appended to existing file with {len(existing)} rows")

result.to_csv(output_file, index=False)
print(f"Saved {len(result)} total rows to {output_file}")

# Show summary
print(f"\nMatched articles this month: {result['PageViews'].notna().sum()}")
print(f"Unmatched articles this month: {result['PageViews'].isna().sum()}")
print(f"\nFromUpdate distribution:")
print(result['FromUpdate'].value_counts().sort_index())