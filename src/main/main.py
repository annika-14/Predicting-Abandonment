# Import the python scripts and another libraries
import api_modified as api_m
import api_scraper as api
import html_scraper as html
import pandas as pd
import argparse
from pathlib import Path
import subprocess
import sys
import os

# Adding ArgumentParse object with description
parser = argparse.ArgumentParser(description='Scrapes features from Github projects.')

# Adding command-line arguments --mode, and other required arguments used in every mode
parser.add_argument('mode', choices=['scrape', 'rescrape', 'subscrape'], help='Specify mode for scraping: scrape, rescrape, subscrape.')
parser.add_argument('access_token', help='Github access token for authorization.')
parser.add_argument('export_file', help='Name of excel file that stores the collected features.')

# Parse command-line arguments
args = parser.parse_args()

# Access mode specified 
mode = args.mode
access_token = args.access_token
export_file = args.export_file

# Initialize Dataframe to be exported
df = ""

# Execute based on mode
if mode == 'scrape':
  # Prompt user to provide star range for scraping
  high = int(input("Please enter an upper limit for the star range you are collecting: ")) 
  low = int(input("Please enter a lower limit for the star range you are collecting: ")) 

  # Scraping features using html and api scrapers
  api.get_projects(low, high, access_token)
  html.scrape_project_list(api.repo_url)
  api_m.collect_sbom_list(api.repo_url, access_token)

  # Converting features to pandas dataframe
  api_df = api.convertToDataFrame()
  html_df = html.convertToDataFrame()

  # Merging dataframes
  df = pd.merge(api_df, html_df, how='outer', on='Project URL')

  # Retrieve the SSH urls for the bash script
  #df['Clone SSH URL'].to_csv('clone_urls.txt', header=False, index=False, line_terminator='\n')

  # Create a file to store the bash data, will be later deleted in order to avoid duplicate data
  #export_bash_csv = "clone_data.csv"
  #open(export_bash_csv, 'a').close()

  # Run the bash script scraper, using subprocess.run()
  #command = f"./clone_scraper.sh ./clone_urls.txt {export_bash_csv}"
  #subprocess.run(command, shell=True, capture_output=True, text=True)

  # Converting features to pandas dataframe
  #bash_df = pd.read_csv(export_bash_csv, header=None, names=['Clone SSH URL','Number of Files','Depth of Files','Number of Contributors','Number of Commits','Number of Merges','Number of Branches','Number of Tags','Number of Links','Has README','Has SECURITY','Has Conduct','Has Contributing','Has ISSUE_TEMPLATE','Has PULL_TEMPLATE']) 

  # Ask the user for the file path to the export file for the bash script
  # Merge the dataframes again
  #df = pd.merge(df, bash_df, how='outer', on='Clone SSH URL')

  # Delete the export file provided
  #os.remove(export_bash_csv)

elif mode == 'rescrape':
  # Prompt user for the name of the import file to be used
  import_file = str(input("Please enter the name of the import file containing the list of projects you've collected: ")) 
  
  # Scraping features using html and api scrapers
  api_m.scrape_project('https://github.com/freeCodeCamp/freeCodeCamp', access_token)
  api_m.collect_sbom('https://github.com/freeCodeCamp/freeCodeCamp', access_token)
  html.scrape_project('https://github.com/freeCodeCamp/freeCodeCamp')

  # Converting features to pandas dataframe
  api_df = api_m.convertToDataFrame()
  html_df = html.convertToDataFrame()

  # Merging dataframes
  df = pd.merge(api_df, html_df, how='outer', on='Project URL')

  # Retrieve the SSH urls for the bash script
  # df['Clone SSH URL'].to_csv('clone_urls.txt', header=False, index=False, line_terminator='\n')
  
  # Create a file to store the bash data, will be later deleted in order to avoid duplicate data
  export_bash_csv = "clone_data.csv"
  open(export_bash_csv, 'a').close()

  # Run the bash script scraper, using subprocess.run()
  # NOTE: clone_urls.txt should still exist if it wasn't deleted or if project lists haven't changed
  command = f"./clone_scraper.sh ./clone_urls.txt {export_bash_csv}"
  subprocess.run(command, shell=True, capture_output=True, text=True)

  # Converting features to pandas dataframe
  bash_df = pd.read_csv(export_bash_csv, header=None, names=['Clone SSH URL','Number of Files','Depth of Files','Number of Contributors','Number of Commits','Number of Merges','Number of Branches','Number of Tags','Number of Links','Has README','Has SECURITY','Has Conduct','Has Contributing','Has ISSUE_TEMPLATE','Has PULL_TEMPLATE']) 
  
  # Merging dataframes
  df = pd.merge(df, bash_df, how='outer', on='Clone SSH URL')
  
  # Delete the export file provided
  os.remove(export_bash_csv)

elif mode == 'subscrape':
  # Prompt user to provide star range for scraping
  high = int(input("Please enter an upper limit for the star range you are collecting: ")) 
  low = int(input("Please enter a lower limit for the star range you are collecting: "))
  
  # Scrape a smaller amount of projects, only a list
  api.get_projects(low, high)

  # Converting features to pandas dataframe
  api_df = api_m.convertToDataFrame()
  df = api_df

# Export merged dataframe as Excel file
if Path(export_file).is_file():
    with pd.ExcelWriter(export_file, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
         df.to_excel(writer,sheet_name="Sheet1", startrow=writer.sheets["Sheet1"].max_row, index = False,header= False)
else:
    df.to_excel(export_file, index=False)
