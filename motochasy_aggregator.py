import pandas as pd
import os
import re
from datetime import timedelta

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

def load_and_merge_xlsx(directory, output_file):
    all_data = []

    # Iterate over all files in the directory
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            file_path = os.path.join(directory, file)
            try:
                # Load the second sheet
                df = pd.read_excel(file_path, sheet_name=1, usecols=['Группировка', 'Моточасы'])
                # Parse "Моточасы" column
                df['Моточасы'] = df['Моточасы'].apply(parse_time_str)
                all_data.append(df)
            except Exception as e:
                print(f"Error processing file {file}: {e}")

    # Combine all data into one DataFrame
    combined_df = pd.concat(all_data, ignore_index=True)

    # Save the combined DataFrame to an Excel file
    combined_df.to_excel(output_file, index=False)
    print(f"Combined data saved to {output_file}")

# Example usage
directory_path = r"C:\Users\delxps\PycharmProjects\excelCollector\mch"
output_path = "combined_output.xlsx"
load_and_merge_xlsx(directory_path, output_path)