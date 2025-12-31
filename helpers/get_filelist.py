# find metadata from file in a local repo

def get_filelist(repo_path, fstr):
    import pandas as pd
    import subprocess
    # cd to the repo, checkout main, and pull to get latest version

    command1 = f"cd {repo_path}"
    subprocess.check_output(command1, shell=True, text=True)

    # find the metadata in the md and yml files (excluding toc.yml)

    command1 = f'findstr /S "{fstr}" *.md *.yml'
    print(f"Running command: {command1}")
    output = subprocess.check_output(command1, shell=True, text=True, cwd=repo_path, encoding='utf-8', errors='ignore')    
    lines = output.strip().split('\n')
    # Handle both markdown and yaml formats - split on colon that separates filename from content
    data = []
    for line in lines:
        # The format is: filename:metadata-field: value
        # We need to find where the metadata field starts by looking for the first colon in the line
        # that separates the filename from the line content
        if ':' in line and f'{fstr}:' in line:
            # Find the first colon - this separates filename from content
            first_colon_pos = line.find(':')
            # The filename is everything before the first colon
            filename = line[:first_colon_pos]
            # Now split the rest by the metadata field
            rest = line[first_colon_pos+1:]
            if f'{fstr}:' in rest:
                parts = rest.split(f'{fstr}:', 1)
                value = parts[1].lstrip()
                data.append([filename, value])
            else:
                data.append(['', ''])
        else:
            data.append(['', ''])
    df = pd.DataFrame(data, columns=['filename', f'{fstr}'])
    # exclude toc.yml files
    df = df[~df['filename'].str.endswith('toc.yml')]
    # if fstr is title, fix the titles
    if fstr == "title":
        # remove quotes from titles
        df['title'] = df['title'].str.replace(r'"', '')
        df['title'] = df['title'].str.replace(r"'", '')


    return df

# separate out the checkout command.  use when you need to checkout and pull
def checkout(repo_path, branch="main"):
    import subprocess
    # cd to the repo, checkout main, and pull to get latest version
    command1 = f"cd {repo_path} && git checkout {branch} && git pull upstream {branch}"
    subprocess.check_output(command1, shell=True, text=True)
    print(f"Checked out {branch} and pulled latest changes")
    return True

if __name__ == "__main__":
    repo_path = "C:/GitPrivate/azure-ai-docs-pr/articles/ai-studio"
    articles = get_filelist(repo_path, "ms.author")
    
    print(f" Total articles: {articles.shape[0]}")
    print(articles.head())



