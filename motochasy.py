import re
from datetime import timedelta
import pandas as pd
import zipfile
import glob
import os



def parse_time_str(time_str):
    """
    Function to parse time strings in formats like:
    "7 дней 11:03:21", "1 день 6:19:13", "10 дней 5:14:38", "3 дня 5:16:27"
    or simply "11:03:21" when days are not mentioned.
    The word "день" can be in different cases: день, дня, дней.
    """
    # Regular expression to capture optional days and time
    pattern = re.compile(r'^(?:(\d+)\s+дн(?:ей|я|ь))?\s*(\d{1,2}:\d{2}:\d{2})$')

    match = pattern.match(time_str.strip())
    if not match:
        # If the format doesn't match, check if it's only time without days
        time_str = time_str.strip()
        if time_str == "":
            # Empty string interpreted as zero time
            return timedelta()
        # Check if it's in HH:MM:SS format
        time_pattern = re.compile(r'^(\d{1,2}:\d{2}:\d{2})$')
        tm_match = time_pattern.match(time_str)
        if not tm_match:
            # If the format is invalid, return zero
            return timedelta()
        else:
            # Time without days
            h, m, s = [int(x) for x in time_str.split(':')]
            return timedelta(hours=h, minutes=m, seconds=s)

    days_str, time_str_only = match.groups()
    days = int(days_str) if days_str else 0
    h, m, s = [int(x) for x in time_str_only.split(':')]
    return timedelta(days=days, hours=h, minutes=m, seconds=s)


def process_csv_from_zip(zip_path, type_abbr,month_abbr):
    """
    Extracts the CSV file from the given ZIP archive, processes it,
    and returns the processed DataFrame with an added 'Month' column.

    Parameters:
    - zip_path: Path to the ZIP archive.
    - month_abbr: Three-letter abbreviation of the month (e.g., 'jan').

    Returns:
    - Processed pandas DataFrame.
    """
    # Construct the expected CSV filename inside the ZIP
    csv_filename = f"{type_abbr}_{month_abbr}_Моточасы.csv"

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # Check if the expected CSV exists in the ZIP
        if csv_filename not in zip_ref.namelist():
            print(f"Warning: {csv_filename} not found in {zip_path}. Skipping this archive.")
            return None

        # Read the CSV file into a pandas DataFrame
        with zip_ref.open(csv_filename) as csv_file:
            df = pd.read_csv(csv_file, sep=';', encoding='utf-8')

    # Apply the parse_time_str function to relevant columns
    df['InMotion_timedelta'] = df['В движении'].apply(parse_time_str)
    df['Idle_timedelta'] = df['Холостой ход'].apply(parse_time_str)

    # Calculate the difference (InMotion - Idle)
    df['Difference'] = df['InMotion_timedelta'] - df['Idle_timedelta']

    # Calculate the percentage of idle time relative to in-motion time
    df['Idle_Percentage_of_InMotion'] = df.apply(
        lambda row: (row['Idle_timedelta'].total_seconds() / row['InMotion_timedelta'].total_seconds() * 100)
        if row['InMotion_timedelta'].total_seconds() > 0 else None,
        axis=1
    )

    # Add a 'Month' column based on the month abbreviation
    df['Month'] = month_abbr.capitalize()

    return df


def main():
    # List to hold DataFrames for each month
    monthly_data = []

    # Define the mapping of month abbreviations if necessary
    # Assuming English three-letter abbreviations. Adjust if using Russian.
    month_abbrs = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                   'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    # Alternatively, if using Russian abbreviations, define them here
    # month_abbrs = ['янв', 'фев', 'мар', 'апр', 'май', 'июн',
    #               'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']

    type_abbr = "gruz"

    # Iterate over each month's abbreviation
    for abbr in month_abbrs:
        # Construct the expected ZIP filename
        zip_filename = f"{type_abbr}_{abbr}.zip"

        # Check if the ZIP file exists in the current directory
        if not os.path.isfile(zip_filename):
            print(f"Warning: {zip_filename} does not exist. Skipping.")
            continue

        print(f"Processing {zip_filename}...")
        processed_df = process_csv_from_zip(zip_filename, type_abbr, abbr)

        if processed_df is not None:
            monthly_data.append(processed_df)

    if not monthly_data:
        print("No data processed. Exiting.")
        return

    # Concatenate all monthly DataFrames into one large DataFrame
    combined_df = pd.concat(monthly_data, ignore_index=True)

    # Optionally, save the combined DataFrame to a new CSV file
    combined_df.to_excel(f"Моточасы_{type_abbr} 20204.xlsx")

    print("All data processed and combined successfully.")
    # print(f"Combined data saved to 'combined_motochasy_data.csv'.")

    # Optionally, display the first few rows of the combined DataFrame
    print(combined_df.head())


if __name__ == "__main__":
    main()
