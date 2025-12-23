# Read work items 
import re
def build_url(filename):
    # remove .md extension and remove /default/ if present
    path = re.sub(r'/default/$', '/', re.sub(r'\.md$', '', filename))
    url = f"https://learn.microsoft.com/en-us/azure/ai-foundry/{path.replace('\\','/').replace(' ','-')}"
    return url
def fix_titles(title, suffix=None, freshness_title=None):
    if not isinstance(title, str):
        return title
    if freshness_title:
        pattern = f"^{re.escape(freshness_title)}"
        title = re.sub(pattern, '', title)
    if suffix:
        title = re.sub(re.escape(suffix) + '$', '', title)
    # remove quotes
    title = title.replace('"', '').replace("'", '')
    # remove leading & trailing blanks
    title = title.strip()
    return title

if __name__ == "__main__":
    import os
    import pandas as pd

    script_dir = os.path.dirname(__file__)
    work_items = pd.read_excel(os.path.join(script_dir, "Jan-engagement.xlsx"))
    work_items['before'] = work_items['Title']
    work_items = fix_titles(work_items)
    csvfile = os.path.join(script_dir, 'fixing.csv')
    temp = work_items[['before','Title']]
    temp.to_csv(csvfile)