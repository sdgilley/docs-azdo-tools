# find metadata from file in a local repo

def get_filelist(repo_path, fstr):
    import pandas as pd
    import subprocess
    # cd to the repo, checkout main, and pull to get latest version

    command1 = f"cd {repo_path}"
    subprocess.check_output(command1, shell=True, text=True)

    # find the metadata in the md files

    command1 = f'findstr /S "{fstr}" *.*'
    print(f"Running command: {command1}")
    output = subprocess.check_output(command1, shell=True, text=True, cwd=repo_path, encoding='utf-8', errors='ignore')    
    lines = output.strip().split('\n')
    data = [line.split(f':{fstr}: ') for line in lines]
    df = pd.DataFrame(data, columns=['filename', f'{fstr}'])
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



