import pandas as pd
import os
from datetime import datetime, timedelta
import re



# Dictionary for month names
month_names = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}


def parse_motochasy(moto_string):
    if pd.isna(moto_string) or not isinstance(moto_string, str):
        return timedelta()

    # Pattern for days and time (e.g., "2 дня 15:30:45" or "15:30:45")
    pattern = r'(?:(\d+)\s*(?:день|дня|дней))?\s*(\d{1,2}:\d{2}:\d{2})?'
    match = re.match(pattern, moto_string.strip())

    total_duration = timedelta()
    if match:
        days, time_str = match.groups()

        # Add days if present
        if days:
            total_duration += timedelta(days=int(days))

        # Add time if present
        if time_str:
            h, m, s = map(int, time_str.split(':'))
            total_duration += timedelta(hours=h, minutes=m, seconds=s)

    return total_duration


def process_excel_files(directory):
    # Get all xlsx files in specified directory
    excel_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]

    # Sort files by creation time
    excel_files.sort(key=lambda x: os.path.getctime(x))

    all_data = []

    # Process each file
    for month_num, file in enumerate(excel_files, 1):
        # Read the second sheet "Сводка"
        df = pd.read_excel(file, sheet_name="Сводка")

        # Add month name and number
        df['month'] = month_names[month_num]
        df['month_number'] = month_num

        # Parse "Моточасы" column and create "Длительность"
        df['Длительность'] = df['Моточасы'].apply(parse_motochasy)

        all_data.append(df)

    # Concatenate all dataframes
    final_df = pd.concat(all_data, ignore_index=True)

    # Define output filename with current datetime
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = os.path.join(directory, f"fake_autobus-ms-probeg-{current_time}.xlsx")

    # Save to Excel
    final_df.to_excel(output_filename, index=False)
    print(f"Data saved to {output_filename}")

    return final_df


if __name__ == "__main__":

    # Directory containing the Excel files - change this to your directory path
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\AUTOBUS\autob-mch-prob2024"
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\AUTOBUS\_fake_busses_2024"
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\AUTOBUS\autobus mercenary tranco fact 2024"
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\GRUZ\gruz_tranco_own_2024"
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\GRUZ\gruz_mch-probeg_2024"
    INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\SPETZ\mch-probeg_spec_own"
    # INPUT_DIRECTORY = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\SPETZ\2024_full\mch-probeg-gruz_merc"
    try:
        # Check if directory exists
        if not os.path.exists(INPUT_DIRECTORY):
            print(f"Error: Directory {INPUT_DIRECTORY} does not exist")
        else:
            result_df = process_excel_files(INPUT_DIRECTORY)
            # Print some basic info about the result
            print("\nColumns in the final dataframe:")
            print(result_df.columns.tolist())
            print("\nFirst few rows:")
            print(result_df.head())
    except Exception as e:
        print(f"An error occurred: {str(e)}")