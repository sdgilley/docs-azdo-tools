# repo location
repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry"  # your local repo
toc_file = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-foundry/toc.yml"  # your local repo
url_path = "https://learn.microsoft.com/azure/ai-foundry/"  # base URL for the articles
import pandas as pd
import yaml
import os

# Function to flatten the TOC structure with full parent hierarchy
def flatten_toc(items, parent_path=None):
    rows = []
    for item in items:
        name = item.get("name", "")
        href = item.get("href", "")
        parent = parent_path if parent_path else ""
        current_path = f"{parent} > {name}" if parent else name
        otherToc = ""

        # Generate URL based on the rules
        if not href:
            url = f"{url_path}{name.replace(' ', '-').lower()}"
        elif href.startswith(".."):
            url = f"https://learn.microsoft.com/{href[3:].replace('.md', '').replace('.yml', '')}"
            otherToc = "true"
        elif href.startswith("/"):
            url = f"https://learn.microsoft.com{href.replace('.md', '').replace('.yml', '')}"
            otherToc = "true"
        else:
            url = f"https://learn.microsoft.com/azure/ai-foundry/{href.replace('.md', '').replace('.yml', '')}"
            otherToc = "false"

        rows.append({"Parent Path": parent, "Name": name, "Href": href, "OtherTOC": otherToc, "URL": url})
        if "items" in item:
            rows.extend(flatten_toc(item["items"], current_path))
    return rows

# Read the TOC YAML file
with open(toc_file, 'r', encoding='utf-8') as file:
    toc = yaml.safe_load(file)

# Flatten the TOC structure
toc_items = toc.get("items", [])
flattened_toc = flatten_toc(toc_items)

# Convert the flattened TOC to a DataFrame
toc_df = pd.DataFrame(flattened_toc)
# remove rows without an href value or blank href value
toc_df = toc_df[toc_df['Href'].str.strip() != ""]
toc_df = toc_df[toc_df['Href'].notna()]
# Write to a CSV file in the current directory
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "toc.csv")
toc_df.to_csv(file_path, index=False)

print(f"TOC with URLs exported to {file_path}")