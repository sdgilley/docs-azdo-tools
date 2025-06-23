# This script searches for specific terms in the MicrosoftDocs GitHub organization
# and saves the results to a CSV file.
import os
import requests
import pandas as pd

token = os.getenv("GH_ACCESS_TOKEN")
headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github+json"
}

# Define your search terms
search_term1 = "Azure OpenAI Service"
search_term2 = "Azure AI Model Inference API"
search_term3 = "Azure AI Agent Service"

org = "MicrosoftDocs"

def github_search(term):
    url = f"https://api.github.com/search/code?q={term}+org:{org}&per_page=100"
    response = requests.get(url, headers=headers)
    return response.json().get("items", [])

# Get results for both terms
results1 = github_search(search_term1)
results2 = github_search(search_term2)
results3 = github_search(search_term3)  

# Build a dict for quick lookup
def build_lookup(results):
    return {
        (item["name"], item["path"], item["repository"]["full_name"]): item
        for item in results
    }

lookup1 = build_lookup(results1)
lookup2 = build_lookup(results2)
lookup3 = build_lookup(results3)  

# Union of all unique files from both searches
all_keys = set(lookup1.keys()) | set(lookup2.keys() | set(lookup3.keys()))

rows = []
for key in all_keys:
    item = lookup1.get(key) or lookup2.get(key) or lookup3.get(key)
    file = item["name"]
    path = item["path"]
    repo = item["repository"]["full_name"]
    url = item["html_url"]
    term1 = "Y" if key in lookup1 else "N"
    term2 = "Y" if key in lookup2 else "N"
    term3 = "Y" if key in lookup3 else "N"
    rows.append({
        "file": file,
        "path": path,
        "repository": repo,
        "url": url,
        "term1": term1,
        "term2": term2,
        "term3": term3
    })

results_df = pd.DataFrame(rows)
script_dir = os.path.dirname(os.path.realpath(__file__))
csv_file_path = os.path.join(script_dir, "search_results.csv")
results_df.to_csv(csv_file_path, index=False)