import argparse
import os
import pandas as pd
from datetime import datetime
from prompt_toolkit import prompt
from prompt_toolkit.completion import PathCompleter

def validate_file_path(file_path):
    # Check if the file exists
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    # Check if the file is readable as an Excel file
    try:
        pd.read_excel(file_path)
    except Exception as e:
        raise ValueError(f"The file {file_path} is not a readable Excel file: {e}")

def validate_range(start_day, end_day, step_value):
    # Check if the values are integers
    if not isinstance(start_day, int):
        raise ValueError("start_day must be an integer.")
    if not isinstance(end_day, int):
        raise ValueError("end_day must be an integer.")
    if not isinstance(step_value, int):
        raise ValueError("step_value must be an integer.")
    # Validate that start_day is positive
    if start_day < 0:
        raise ValueError("start_day must be a positive integer.")
    # Validate that end_day is greater than or equal to start_day
    if end_day < start_day:
        raise ValueError("end_day must be equal to or larger than start_day.")
    # Validate that step_value is positive
    if step_value < 1:
        raise ValueError("step_value must be a positive integer greater or equal to 1.")


def confirm_step_values(start_day, end_day, step_value):
    while True:
        step_values = list(range(start_day, end_day + 1, step_value))
        if step_values[-1] != end_day:
            step_values.append(end_day)
        print(f"The steps from {start_day} to {end_day} by {step_value} are: {step_values}")
        confirmation = input("Are these values okay? (y/n): ").strip().lower()
        if confirmation == 'y':
            return step_values
        else:
            try:
                start_day = int(input("Enter the beginning day (positive integer): "))
                end_day = int(input("Enter the end day (positive integer, equal or larger to beginning day): "))
                step_value = int(input("Enter the step value (positive integer greater or equal to 1): "))
                validate_range(start_day, end_day, step_value)
            except ValueError as e:
                print(e)

def clean_excel(file_path):
    df = pd.read_excel(file_path)

    initial_row_count = len(df)
    df_cleaned = df.dropna()
    blanks_count = initial_row_count - len(df_cleaned)

    cleaned_file_path = os.path.splitext(file_path)[0] + "_cleaned.xlsx"
    df_cleaned.to_excel(cleaned_file_path, index=False)
    print(f"Cleaned data saved to {cleaned_file_path}")
    print(f"Number of rows with blanks removed: {blanks_count}")

def change_date(row, current_timestamp):
    print("made it to date")
    # Convert 'Last Update' to Unix timestamp and calculate days since last update
    last_update_dt = datetime.strptime(row['Last Update'], "%Y-%m-%dT%H:%M:%SZ")
    row['Days Since Last Update'] = round((current_timestamp - last_update_dt.timestamp()) / (60 * 60 * 24), 2)
    
    # Convert 'Last Push' to Unix timestamp and calculate days since last push
    last_push_dt = datetime.strptime(row['Last Push'], "%Y-%m-%dT%H:%M:%SZ")
    row['Days Since Last Push'] = round((current_timestamp - last_push_dt.timestamp()) / (60 * 60 * 24), 2)
    
    # Convert 'Created Date' to Unix timestamp and calculate repo age in days
    created_date_dt = datetime.strptime(row['Created Date'], "%Y-%m-%dT%H:%M:%SZ")
    row['Repo Age (Days)'] = round((current_timestamp - created_date_dt.timestamp()) / (60 * 60 * 24), 2)
    
    return row

def remove_k(row):
    print("made it to k")
    columns_to_check = ['Followers of Owner', 'Members of Owner', 'Repos of Owner', 'Number of Watches']
    for col in columns_to_check:
        print(f"{col}")
        value = row[col]
        if isinstance(value, str) and 'k' in value:
            value = value.replace('k', '')
            value = value.split('.')
            converted_value = int(value[0]) * 1000 + int(value[1]) * 100
            row[f'{col} (int)'] = converted_value
        else:
            row[f'{col} (int)'] = value  # No change if not a 'k' string
    return row

def process_excel(file_path, step_values, timestamp):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Check if the required column exists
    if "Last Commit" not in df.columns:
        raise KeyError("Column 'Last Commit' does not exist in the Excel file.")
    if "Execution Timestamp" not in df.columns:
        raise KeyError("Column 'Execution Timestamp' does not exist in the Excel file.")
    
    # Convert days to seconds and add a new column for each day
    for day in step_values:
        print(f"{day}")
        df[f'Abandoned Within {day} Days'] = (timestamp - df['Last Commit']) > (day * 24 * 60 * 60)
    df['Days Since Last Commit (now)'] = (timestamp - df['Last Commit']) / (24 * 60 * 60)
    df['Days Since Last Commit (time of collection))'] = (df['Execution Timestamp'] - df['Last Commit']) / (24 * 60 * 60)


    # Parse the date columns into day values, add them as new columns
    df = df.apply(change_date, axis=1, current_timestamp=timestamp)

    # get rid of k for thousands
    # Followers of Owner , Members of Owner, Repos of Owner , Number of Watches
    df = df.apply(remove_k, axis=1)

    print("done")
    # Save the modified DataFrame back to the Excel file
    df.to_excel(file_path + "_processed", index=False)

    processed_file_path = os.path.splitext(file_path)[0] + "_processed.xlsx"
    df_cleaned.to_excel(processed_file_path, index=False)
    print(f"Processed data saved to {processed_file_path}")

if __name__ == "__main__":
    # Argument parser setup
    parser = argparse.ArgumentParser(description='Process an Excel file.')
    parser.add_argument('file', type=str, help='Path to the Excel file.')
    
    # Parse the arguments
    args = parser.parse_args()
    
    # Validate the file path
    try:
        validate_file_path(args.file)
    except (FileNotFoundError, ValueError) as e:
        print(e)
        exit(1)
    
    mode = prompt("Would you like to clean a file or add post-processing columns? (clean/process): ")
    if mode == "clean":
        clean_excel(args.file)
    

    else:
        # Get the current time and convert it to a Unix timestamp
        timestamp = int(datetime.now().timestamp())
        
        # Get and confirm step values with error checking
        while True:
            try:
                start_day = int(input("Enter the beginning day (positive integer): "))
                end_day = int(input("Enter the end day (positive integer, equal or larger to beginning day): "))
                step_value = int(input("Enter the step value (positive integer greater or equal to 1): "))
                validate_range(start_day, end_day, step_value)
                step_values = confirm_step_values(start_day, end_day, step_value)
                break
            except ValueError as e:
                print(e)
        
        # Process the Excel file with the confirmed step values and timestamp
        try:
            process_excel(args.file, step_values, timestamp)
        except Exception as e:
            print(f"An error occurred while processing the Excel file: {e}")
