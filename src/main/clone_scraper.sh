#!/bin/bash

# Listed below are the headers to be copied to the csv file when needed
# Number of Files,Max Depth of Files,Number of Contributors,Number of Commits,Number of Merge Commits,Number of Branches,Number of Tags,Number of Links,Has README,Has SECURITY,Has CODE_OF_CONDUCT,Has CONTRIBUTING,Has ISSUE_TEMPLATE,Has PULL_REQUEST_TEMPLATE

if [[ $# -ne 2 ]]; then
  echo "usage: ./clone_scraper [fullpath_to_import_file] [fullpath_to_export_file]"
  exit 
fi

# Store command line argument variables
import_file=$1
export_file=$2

DATE=$(date)
repo_list=()

# While loop to get all the ssh_links for cloning from the import file
# Reading line by line
while read -r line
do
  repo_list+=("$line")
done < "$import_file"


for i in "${repo_list[@]}"; do
  echo "$i"
  git clone "$i"

  # Get directory name based off of SSH clone link
  directory=$(echo "$i" | cut -d '/' -f 2 | rev | cut -c 5- | rev)
  cd $directory

  # Gets all files, including ones in sub-directories
  # Doesn't count anything in the .git directory, as those files are not modified or created by the users directly
  num_files=$(find . -type f 2> /dev/null | grep -vwE ".git" | wc -l | tr -d "[:blank:]")

  # Get the max depth of files
  depth=$(find . -type f | sed 's/[^/]//g' | sort | tail -1 | awk '{ print length }' | tr -d "[:blank:]")
  
  # Number of Contributors, people who have made commits, could be more as remote commits are listed as different people
  num_contributors=$(git shortlog -s -n | wc -l | tr -d "[:blank:]")

  # Number of Commits
  num_commits=$(git log --pretty=oneline | wc -l | tr -d "[:blank:]")

  # Number of Merge Commits, is included in number of commits above
  num_merges=$(git log --pretty=oneline --merges | wc -l | tr -d "[:blank:]")

  # Number of Branches
  num_branches=$(git branch -a | wc -l | tr -d "[:blank:]")

  # Number of tags AKA releases
  num_tags=$(git tag | sort | uniq | wc -l | tr -d "[:blank:]")

  # Number of Links on the README (including images) will be set later
  num_links=0

  # Presence of README and other files
  README=$(ls | grep -w "README" | head -1)
  SECURITY=$(find . -name "SECURITY.md")
  CONDUCT=$(find . -name "CODE_OF_CONDUCT.md")
  CONTRIBUTING=$(find . -name "CONTRIBUTING.md")
  ISSUE_TEMPLATE=$(find . -name "ISSUE_TEMPLATE.md")
  PULL_TEMPLATE=$(find . -name "PULL_REQUEST_TEMPLATE.md")

  # Check if empty
  if [[ -z "$README" ]]; then
    README="NO"
  else
    # Change num_links accordingly if README file is present
    COUNT_ONE=$(cat $README | grep -o '(https' | wc -l | tr -d "[:blank:]")
    COUNT_TWO=$(cat $README | grep -o 'href' | wc -l | tr -d "[:blank:]")
    num_links=$(($COUNT_ONE + $COUNT_TWO))
    README="YES"
  fi

  if [[ -z "$SECURITY" ]]; then
    SECURITY="NO"
  else
    SECURITY="YES"
  fi

  if [[ -z "$CONDUCT" ]]; then
    CONDUCT="NO"
  else
    CONDUCT="YES"
  fi

  if [[ -z "$CONTRIBUTING" ]]; then
    CONTRIBUTING="NO"
  else
    CONTRIBUTING="YES"
  fi

  if [[ -z "$ISSUE_TEMPLATE" ]]; then
    ISSUE_TEMPLATE="NO"
  else
    ISSUE_TEMPLATE="YES"
  fi

  if [[ -z "$PULL_TEMPLATE" ]]; then
    PULL_TEMPLATE="NO"
  else
    PULL_TEMPLATE="YES"
  fi

  # THESE LAST COMMANDS ARE COMMENTED OUT
  # The data can be collected but is formatted in a way that makes it hard to store
  # List of commits for days
  # day_trends=$(git log --date=short --pretty=format:%cd | sort | uniq -c)
  # List of commits for months
  # month_trends=$(git log --date=short --pretty=format:%cd | cut -d '-' -f 1-2 | sort | uniq -c)
  # List of commits for years
  # year_trends=$(git log --date=short --pretty=format:%cd | cut -d '-' -f 1 | sort | uniq -c)

  # Declare empty list to store file types and counts
  file_types=(`find . -type f | rev | grep '.' | cut -d '.' -f 1 | rev | sort | uniq -c | awk '{$1=$1};1' | tr "[:blank:]" ":"`)
  
  # Switch to parent directory
  cd ..

  # Export filetypes to a separate file as data is awkward to store in pandas dataframe
  FILE_NAME="file_types -> $DATE"
  echo "$i,${file_types[@]}" >> "$FILE_NAME"

  # Output data to the csv file
  echo "$i,$num_files,$depth,$num_contributors,$num_commits,$num_merges,$num_branches,$num_tags,$num_links,$README,$SECURITY,$CONDUCT,$CONTRIBUTING,$ISSUE_TEMPLATE,$PULL_TEMPLATE,${file_types[@]}" >> "$export_file"

  # CODE LEFT HERE JUST IN CASE: 
  # Remove the cloned repository's directory
  # cd ..
  echo -e "y\ny\n" | rm -r $directory

  # Compress the cloned repository
  # -c creates an archive, -z tells tar to use gzip, -f specifies file name of compressed file
  # cd ..
  # tar -czf "$directory".tar.gz ./"$directory"

  # To decompress the repository use the following command:
  # -x extracts archive, other two options same as above command
  # tar -xzf "$directory".tar.gz
done 