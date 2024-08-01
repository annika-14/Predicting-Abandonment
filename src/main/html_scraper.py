# Import all the Python libraries
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException
import pandas as pd
import time
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
import requests 

# Chrome options to be added for Selenium Driver to speed up data collection speed
# Includes: Headless mode and no images
chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-renderer-backgrounding")
chrome_options.add_argument("--disable-background-timer-throttling")
chrome_options.add_argument("--disable-backgrounding-occluded-windows")
chrome_options.add_argument("--disable-client-side-phishing-detection")
chrome_options.add_argument("--disable-crash-reporter")
chrome_options.add_argument("--disable-oopr-debug-crash-dump")
chrome_options.add_argument("--no-crash-upload")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-low-res-tiling")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--silent")
chrome_options.add_argument("--blink-settings=imagesEnabled=false")

projects = []

# BELOW: Defining individual functions to isolate pages used to scrape features
def save_html_to_file(content, url):
    file_name = "log"
    with open(file_name, 'a', encoding='utf-8') as file:
        file.write(f"{url}\n")
        

# Scrapes pull requests page
def scrape_prs(project_url, driver):
    project = project_url[19:]
    # Pull Requests
    pull_url = project_url + "/pulls"

    for i in range(0,10):
        driver.get(pull_url)
        # Wait for the document to be in 'complete' state
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
        )
        html = driver.page_source
        save_html_to_file(html, pull_url)
        soup = BeautifulSoup(html,"html.parser")

        open_prs = soup.find(href=f"/{project}/pulls?q=is%3Aopen+is%3Apr")
        close_prs = soup.find(href=f"/{project}/pulls?q=is%3Apr+is%3Aclosed")
        if open_prs != None and close_prs != None:
            open_prs = open_prs.text.split()[0]
            close_prs = close_prs.text.split()[0]
            print(f"{project_url}: {open_prs} open_prs and {close_prs} close_prs")
            return [open_prs, close_prs]
        else:
            time.sleep(10)

    print(f"{project_url}: open_prs and close_prs not found")
    return [None, None]

# Scrape owner's page for info like followers, members etc.
def scrape_owner(project_url, driver):
    project = project_url[19:]
    creator = project.split('/')[0]

    # Verified Repo Owner
    owner_url = f"https://github.com/{creator}"
    driver.get(owner_url)

    # Wait for the document to be in 'complete' state
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.TAG_NAME, 'body'))
    )

    html = driver.page_source
    save_html_to_file(html, owner_url)
    soup = BeautifulSoup(html, "html.parser")
    #print(soup.prettify())

    verified = soup.find('summary', {'title': 'Label: Verified'})
    if verified != None:
        verified = "TRUE"
    else:
        verified = "FALSE"

    print(f"{project_url} owner status: {verified}")

    # Number of Owner Followers
    followers = soup.find('a', class_='Link--secondary no-underline no-wrap')
    if followers != None:
        followers = followers.text.split()[0]
    else:
        followers = 0
    print(f"{project_url} followers: {followers}")

    # Number of Owner Members
    members = soup.find('span', class_='Counter js-profile-member-count')
    if members != None:
        while members.text == "":
            driver.get(owner_url)
            WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
            )
            time.sleep(5)
            html = driver.page_source
            save_html_to_file(html, owner_url)
            soup = BeautifulSoup(html, "html.parser")
            members = soup.find('span', class_='Counter js-profile-member-count')
            if members == None:
                break
        if members != None:
            members = members.text.split()[0]
      
    if members == None:
        members = 0

    print(f"{project_url} members: {members}")

    # Number of Other Repositories by Owner
    repositories = soup.find('span', class_='Counter js-profile-repository-count')
    if repositories == None:
        repositories = soup.find_all('span', class_='Counter')[0]

    if repositories != None:
        repositories = repositories.text.split()[0]
    print(f"{project_url} repositories: {repositories}")

    return [verified, followers, members, repositories]


# Scrape insights page: active prs and active issues
def scrape_insight(project_url, driver):
    project = project_url[19:]
    creator = project.split('/')[0]

    # Active prs and active issues
    insight_url = f"{project_url}/pulse"
    driver.get(insight_url)

    # Wait for the document to be in 'complete' state
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.TAG_NAME, 'body'))
    )

    html = driver.page_source 
    save_html_to_file(html, insight_url)
    soup = BeautifulSoup(html, "html.parser")

    active = soup.find_all('div', class_='mt-2')
    active_prs = active[0]
    active_issues = active[1]

    if active_prs != None:
        active_prs = active_prs.text.split()[0]

    if active_issues != None:
        active_issues = active_issues.text.split()[0]

    print(f"{project_url}: {active_prs} Active pull requests, {active_issues} Active issues")
    return [active_prs, active_issues]


# Scrape issues page
def scrape_issues(project_url, driver):
    project = project_url[19:]
    # Issues
    issue_url = project_url + "/issues"

    open_issues = None
    closed_issues = None
    num_labels = None
    num_milestones = None

    for i in range(0,10):
        driver.get(issue_url)

        # Wait for the document to be in 'complete' state
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
        )

        html = driver.page_source
        save_html_to_file(html, issue_url)
        soup = BeautifulSoup(html,"html.parser")

        if open_issues == None:
            open_issues = soup.find(href=f"/{project}/issues?q=is%3Aopen+is%3Aissue")
            if open_issues != None:
                open_issues = open_issues.text.split()[0]

        if closed_issues == None:
            closed_issues = soup.find(href=f"/{project}/issues?q=is%3Aissue+is%3Aclosed")
            if closed_issues != None:
                closed_issues = closed_issues.text.split()[0]

        if num_labels == None:
            num_labels = soup.find(href=f"/{project}/labels")
            if num_labels != None:
                num_labels = num_labels.find("span").text

        if num_milestones == None:
            num_milestones = soup.find(href=f"/{project}/milestones")
            if num_milestones != None:
                num_milestones = num_milestones.find("span").text
                
            if open_issues == None or closed_issues == None or num_labels == None or num_milestones == None:
                time.sleep(10)
            else:
                break
    
    if open_issues == None:
        open_issues = 0
    if closed_issues == None:
        closed_issues = 0

    if type(num_labels) != int:
        # labels
        driver.get(project_url + '/labels')
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
        )
        html = driver.page_source
        save_html_to_file(html, project_url + '/labels')
        soup = BeautifulSoup(html,"html.parser")
        num_labels = soup.find('span', class_='js-labels-count')
        num_labels = num_labels.text.split()[0]

        # milestones
        driver.get(project_url + '/milestones')
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
        )
        html = driver.page_source
        save_html_to_file(html, project_url + '/milestones')
        soup = BeautifulSoup(html,"html.parser")
        num_milestones = soup.find('a', class_='btn-link selected').text.split()[0]

    print(f"{project_url}, Open issues: {open_issues}, Closed issues: {closed_issues}, Labels: {num_labels}, Milestones: {num_milestones}")
    return [open_issues, closed_issues, num_labels, num_milestones]


# Scrape main/code page of repository
def scrape_page(project_url, driver):

    project_features = []

    # Get the OWNER/REPO
    project = project_url[19:]

    # Set up Web Driver
    driver.get(project_url)

    # Get number of watches and sponsors
    # Wait for the document to be in 'complete' state
    count = 0
    while count < 10:
        count += 1
        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.TAG_NAME, 'body'))
            )
        except TimeoutException:
            print(f"TimeoutException: Retrying (Attempt {count})...")
            

    # Parse HTML
    html = driver.page_source
    save_html_to_file(html, project_url)
    soup = BeautifulSoup(html,"html.parser")

    num_watches = soup.find(href=f"/{project}/watchers").find("strong").text

    creator = project.split('/')[0]
    sponsored = "TRUE" if soup.find(href=f"/sponsors/{creator}") != None else "FALSE"

    if sponsored == "TRUE":
        driver.get(f"https://github.com/sponsors/{creator}")
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, 'body'))
        )
        html = driver.page_source
        save_html_to_file(html, "https://github.com/sponsors/{creator}")
        soup = BeautifulSoup(html,"html.parser")

        current_sponsors = soup.find(lambda tag: tag.name == 'h4' and 'Current sponsors' in tag.get_text())
        past_sponsors = soup.find(lambda tag: tag.name == 'h4' and 'Past sponsors' in tag.get_text())

        if current_sponsors == None:
            current_sponsors = soup.find('p', class_='f3-light color-fg-muted mb-3')
            if current_sponsors != None:
                current_sponsors = current_sponsors.text.split()[0]
        else:
            current_sponsors = current_sponsors.text.split()[2]
            
        if past_sponsors == None:
            past_sponsors = 0
        else:
            past_sponsors = past_sponsors.text.split()[2]

    else:
        current_sponsors = 0
        past_sponsors = 0

    print(f"{project_url}: {current_sponsors} Current sponsors, {past_sponsors} Past sponsors")
    project_features.append(sponsored)
    project_features.append(current_sponsors)
    project_features.append(past_sponsors)
    project_features.append(num_watches)

    # Number of Workflow Runs
    workflow_url = project_url + "/actions"
    driver.get(workflow_url)

    # Wait for the document to be in 'complete' state
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.TAG_NAME, 'body'))
    )

    html = driver.page_source
    save_html_to_file(html, workflow_url)
    soup = BeautifulSoup(html, "html.parser")

    workflow = soup.find(lambda tag: tag.name == 'strong' and 'workflow runs' in tag.get_text())
    if workflow != None:
        workflow = workflow.text.split()[0]

    project_features.append(workflow)

    # Number of Dependent Repos
    dependent_url = project_url + "/network/dependents"
    driver.get(dependent_url)

    # Wait for the document to be in 'complete' state
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.TAG_NAME, 'body'))
    )

    html = driver.page_source
    save_html_to_file(html, dependent_url)
    soup = BeautifulSoup(html, "html.parser")

    dependents = soup.find('a', class_='btn-link selected')
    if dependents != None:
        dependents = dependents.text.split()[0]

    project_features.append(dependents)
    print(f"{project_url} {dependents} dependents, {workflow} workflows")

    return project_features

def is_repo_empty(repo_url):
    api_url = f"https://api.github.com/repos/{repo_url[19:]}/commits"
    response = requests.get(api_url)
    
    if response.status_code == 409:
        return True  # Repository is empty
    return False  # Repository is non-empty

def scrape_project(project_url):
    if is_repo_empty(project_url):
        print(f"Repository {project_url} is empty.")
        return [project_url, "Empty Repository"]
# Define function to be used to scrape each of the above pages and combine their result
    project = [project_url]

    driver = webdriver.Chrome(options=chrome_options)
    prs = scrape_prs(project_url, driver)
    time.sleep(1)
    driver = webdriver.Chrome(options=chrome_options)
    owner = scrape_owner(project_url, driver)
    time.sleep(1)
    driver = webdriver.Chrome(options=chrome_options)
    insight = scrape_insight(project_url, driver)
    time.sleep(1)
    driver = webdriver.Chrome(options=chrome_options)
    issues = scrape_issues(project_url, driver)
    time.sleep(1)
    driver = webdriver.Chrome(options=chrome_options)
    page = scrape_page(project_url, driver)

    driver.quit()

    project = project + prs + owner + insight + issues + page
    print(f"PROJECT: {project[0]}")
    print(" ".join(map(str, project)))    
    return project


# Define function that will use threads to scrape each project in the sublist passed to it
def scrape_project_list(project_list):
    with ThreadPoolExecutor(max_workers=10) as p:
        features = p.map(scrape_project, project_list)
        for f in features:
            projects.append(f)
    return projects

def convertToDataFrame():
    projects_df = pd.DataFrame(projects, columns=['Project URL',
                                                  'Open Pull Requests',
                                                  'Closed Pull Requests',
                                                  'Verified Owner',
                                                  'Followers of Owner',
                                                  'Members of Owner',
                                                  'Repos of Owner',
                                                  'Active Pull Requests',
                                                  'Active Issues',
                                                  'Open Issues',
                                                  'Closed Issues',
                                                  'Number of Labels',
                                                  'Number of Milestones',
                                                  'Sponsored',
                                                  'Current Sponsors',
                                                  'Past Sponsors',
                                                  'Number of Watches',
                                                  'Number of Workflow Runs',
                                                  'Number of Dependents'
                                                 ])
    return projects_df

