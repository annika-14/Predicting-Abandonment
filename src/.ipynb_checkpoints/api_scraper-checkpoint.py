import requests
from requests.auth import HTTPBasicAuth 
import pandas as pd
import time
import os
import math
import datetime
import subprocess
import sys

# Get the command line arguments
if len(sys.argv) != 6:
  print("Usage: [export_file] [projects_file] [low_stars] [high_stars] [access_token]")
  exit()
 
export_file = str(sys.argv[1])
projects_file = str(sys.argv[2])
low_stars = int(sys.argv[3])
high_stars = int(sys.argv[4])
access_token = str(sys.argv[5])

projects = pd.read_excel(projects_file)
project_list = projects.iloc[:, 0].tolist()
project_set = set(project_list)

# Declare lists to store feature data
repo_url= []
repo_stars = []
repo_wiki = []
repo_open_issues = []
repo_forks = []
repo_last_update = []
repo_size = []
repo_created_date = []
repo_last_push = []
repo_language = []
repo_discussions = []
repo_pages = []
repo_license = []
repo_archived = []
repo_projects = []
repo_homepage = []
repo_org = []
repo_topics = []
repo_ssh_url = []

# Function for collecting SBOMs
def collect_sbom(project_list, dir_path):
    for project in project_list:
        owner_repo = project[19:]
        sbom_url = f"https://api.github.com/repos/{owner_repo}/dependency-graph/sbom"
        file_name = owner_repo.split('/')
        file_name = f"{file_name[0]}_{file_name[1]}_sbom.json"
        print(file_name)
        
        headers = {
            'Accept': 'application/vnd.github+json',
            'Authorization': 'Bearer ' + access_token,
            'X-GitHub-Api-Version': '2022-11-28',
        }
        response = requests.get(sbom_url, headers=headers)
        if response.status_code == 200:
            with open(f"{dir_path}/{file_name}", "wb") as file:
                file.write(response.content)
            print(f"{project}: SBOM downloaded")
        else:
            print(f"{project}: SBOM download failed")

# Collecting SBOMs and storing them in a sbom directory
current_dir = os.getcwd()
date = datetime.date.today()
dir_name = f"sbom_{date}"
if not os.path.exists(dir_name): 
    os.makedirs(dir_name)
sbom_dir = f"{current_dir}/{dir_name}"
collect_sbom(project_list, sbom_dir)

# Function used to scrape the data using the Github API
def get_github_repo_info(search_filter, page_number, project_set):
    api_url = f"https://api.github.com/search/repositories?q="+str(search_filter)+"&page="+str(page_number)+"&per_page=100"

    headers = {
    "Authorization": "Bearer " + access_token,
    "Accept": "application/vnd.github.v3+json"
    }

    # Get number of repos to be used to determine number of pages in the calling cell below
    num_repos = 0

    response = requests.get(api_url, headers=headers)

    while response.status_code != 200:
        print(f"Request failed on page {page_number}")
        delay_seconds = 60  # default delay
        time.sleep(delay_seconds)
        response = requests.get(api_url, headers=headers)

    #response = requests.get(api_url)
    if response.status_code == 200:
        # Parse the JSON response
        repo_list = response.json()['items']
        num_repos = response.json()['total_count']

        for repo_info in repo_list:

            # Only scrapes repo if it is in project_set
            if not project_set:
                break
            repo_name = repo_info.get("full_name", "Name not found")
            if f"https://github.com/{repo_name}" in project_set:
                project_set.remove(f"https://github.com/{repo_name}")
            else:
                continue

            # Extract and print relevant information
            repo_url.append(repo_info.get("html_url", "URL not found"))
            repo_stars.append(repo_info.get("stargazers_count", "Stargazers count not found"))
            #repo_watches.append(repo_info.get("subscribers_count", "Watchers count not found"))
            repo_wiki.append(repo_info.get("has_wiki", "Wiki not found"))
            repo_open_issues.append(repo_info.get("open_issues_count", "Open issues count not found"))
            repo_forks.append(repo_info.get("forks_count", "Forks count not found"))
            repo_last_update.append(repo_info.get("updated_at", "Last update not found"))
            repo_size.append(repo_info.get("size", "size not found"))
            repo_created_date.append(repo_info.get("created_at", "Created date not found"))
            repo_last_push.append(repo_info.get("pushed_at", "Last push not found"))
            repo_language.append(repo_info.get("language", "Language not found"))
            repo_discussions.append(repo_info.get("has_discussions", "Discussions not found"))
            repo_pages.append(repo_info.get("has_pages", "Pages not found"))
            repo_archived.append(repo_info.get("archived", "Archived not found"))
            repo_projects.append(repo_info.get("has_projects", "Projects not found"))
            repo_topics.append(len(repo_info.get("topics", "No Topics")))
            repo_ssh_url.append(repo_info.get("ssh_url", "Projects not found"))
            repo_org.append(repo_info['owner'].get("type", "No type"))

            # Conditional statements are to avoid possible errors
            license = repo_info.get("license", "None")
            if license == "None" or license is None:
                repo_license.append("None")
            else:
                repo_license.append(license["spdx_id"])


            homepage = repo_info.get("homepage", "No Homepage")
            if homepage is None or len(homepage) == 0:
               repo_homepage.append("None")
            else:
                repo_homepage.append(homepage)


    else:
        # If the request was not successful, print an error message
        print("Error:", response.status_code)
        print("Response:", response.text)

    return num_repos



# Function used to facilitate scraping by changing search filters for number of stars, and some of them created date
def get_projects(low, high, project_set):
    # Variable for determining range of projects
    decrement = 500
    # While loop to go through each range from low to high
    while high != low:
        # Change ranges accordingly to get <1000 projects
        if high == 400000:
            decrement = 375000
        elif high == 25000:
            decrement = 5000
        elif high == 15000:
            decrement = 500
        elif high == 5000:
            decrement = 100
        elif high == 3000 or high == 1000:
            decrement = 10
        elif high == 680:
            decrement = 5
        elif high == 500:
            decrement = 1


        # Search URL just in case => q=stars%3A120..120+created%3A2021-01-01..2021-12-31&

        # Add the 'created:' parameter for <178 stars
        if high <= 179:
            decrement = 1
            # Value of 9 goes down to year 2016
            for i in range(9):
                year = 2024 - i
                created_date = "+created%3A" + str(year) + "-01-01.." + str(year) + "-12-31"
                print(high-decrement, high-1, year, 1)
                return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1)+created_date, 1, project_set)
                # For loop to run function to get features
                for page_number in range(2,math.ceil(return_value/100)+1):
                    print(high-decrement, high-1, page_number)
                    return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1)+created_date, page_number, project_set)
            # One more request for all projects <=2015
            created_date = "+created%3A<=2015-12-31"
            print(high-decrement, high-1, 2015, 1)
            return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1)+created_date, 1, project_set)
            for page_number in range(2,math.ceil(return_value/100)+1):
                print(high-decrement, high-1, 2015, page_number)
                return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1)+created_date, page_number, project_set)
        else:
            print(high-decrement, high-1, 1)
            return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1), 1, project_set)
            for page_number in range(2,math.ceil(return_value/100)+1):
                # For loop to run function to get features
                print(high-decrement, high-1, page_number)
                return_value = get_github_repo_info("stars%3A"+str(high-decrement)+'..'+str(high-1), page_number, project_set)
        high -= decrement



# Call the above function to begin scraping
get_projects(low_stars,high_stars,project_set)



# Create pandas DataFrame to store the data
projects_df = pd.DataFrame({'Project URL':repo_url,
                            'Clone SSH URL':repo_ssh_url,
                            'Organization':repo_org,
                            'Homepage':repo_homepage,
                            'Last Update':repo_last_update,
                            'Last Push':repo_last_push,
                            'Created Date':repo_created_date,
                            'Archived':repo_archived,
                            'Size':repo_size,
                            'Number of Stars':repo_stars,
                            #'Number of Watches':repo_watches,
                            'Open Issues + Open Pull Requests':repo_open_issues,
                            'Number of forks':repo_forks,
                            'Has a Wiki':repo_wiki,
                            'Has Discussions':repo_discussions,
                            'Has Projects':repo_projects,
                            'Has Pages':repo_pages,
                            'License':repo_license,
                            'Language':repo_language,
                            'Topics': repo_topics})

# For printing the data frame
# projects_df


# Export to Excel
# CHANGE THE EXPORTED FILE NAME ACCORDINGLY
try:
    with pd.ExcelWriter(
        export_file,
        mode="a",
        engine="openpyxl",
        if_sheet_exists="overlay",
    ) as writer:
         projects_df.to_excel(writer,sheet_name="Sheet1", startrow=writer.sheets["Sheet1"].max_row, index = False,header= False)
except FileNotFoundError:
    projects_df.to_excel(export_file, index=False)
