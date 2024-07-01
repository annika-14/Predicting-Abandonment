import argparse
import os
import pandas as pd
from datetime import datetime

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

def process_excel(file_path, step_values, timestamp):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Check if the required column exists
    if "Last Commit" not in df.columns:
        raise KeyError("Column 'Last Commit' does not exist in the Excel file.")
    
    
    # Convert days to seconds and add a new column for each day
    for day in step_values:
        print(f"{day}")
        df[f'Abandoned Within {day} Days'] = (timestamp - df['Last Commit']) > (day * 24 * 60 * 60)
    
    
    # Save the modified DataFrame back to the Excel file
    df.to_excel(file_path, index=False)
    print(f"Processed data saved to {file_path}")

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
