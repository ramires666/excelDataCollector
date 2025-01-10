import pandas as pd
import zipfile
import os

def process_bus_csv_from_zip(zip_path, type_abbr, month_abbr, month_digit):
    """
    Extracts the bus CSV file from the given ZIP archive, processes it,
    and returns the processed DataFrame with added 'Month' and 'Month_digit' columns.

    Parameters:
    - zip_path: Path to the ZIP archive.
    - type_abbr: Type abbreviation (e.g., 'bus').
    - month_abbr: Three-letter English abbreviation of the month (e.g., 'jan').
    - month_digit: Numerical representation of the month (e.g., 1 for January).

    Returns:
    - Processed pandas DataFrame or None if file not found or no valid data.
    """
    # Construct the expected CSV filename inside the ZIP
    csv_filename = f"{type_abbr}_{month_abbr}_Сливы.csv"

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Check if the expected CSV exists in the ZIP
            if csv_filename not in zip_ref.namelist():
                print(f"Warning: '{csv_filename}' not found in '{zip_path}'. Skipping this archive.")
                return None

            # Read the CSV file into a pandas DataFrame
            with zip_ref.open(csv_filename) as csv_file:
                df = pd.read_csv(csv_file, sep=';', encoding='utf-8')

    except zipfile.BadZipFile:
        print(f"Error: '{zip_path}' is not a valid ZIP file. Skipping.")
        return None
    except Exception as e:
        print(f"Error processing '{zip_path}': {e}")
        return None

    # Display initial rows for debugging (optional)
    # print(df.head())

    # Filter out rows that do not have actual data
    # Assuming that rows with '-----' or '0' indicate no data
    # Modify the conditions based on actual data patterns
    # Here, we assume that a valid row has 'Слито' not equal to '-----' and 'Пробег' not '0.00 км' or '-----'

    # First, ensure that the relevant columns exist
    required_columns = ['Слито', 'Пробег']
    for col in required_columns:
        if col not in df.columns:
            print(f"Warning: Column '{col}' not found in '{csv_filename}'. Skipping this file.")
            return None

    # Apply filtering conditions
    df_filtered = df[
        (df['Слито'].astype(str).str.strip() != "-----") &
        (df['Слито'].astype(str).str.strip() != "") &
        (df['Пробег'].astype(str).str.strip() != "0.00 км") &
        (df['Пробег'].astype(str).str.strip() != "-----") &
        (df['Пробег'].astype(str).str.strip() != "")
    ]

    if df_filtered.empty:
        print(f"No valid data found in '{csv_filename}'.")
        return None

    # Add 'Month' column based on the month abbreviation
    df_filtered['Month'] = month_abbr.capitalize()

    # Add 'Month_digit' column based on the month abbreviation
    df_filtered['Month_digit'] = month_digit

    # Optionally, convert 'Пробег' from string to float by removing ' км' and handling commas
    # Also, handle possible thousand separators (e.g., '1,234 км')
    df_filtered['Пробег_km'] = df_filtered['Пробег'].str.replace(' км', '').str.replace(',', '').astype(float)

    # Optionally, handle 'Слито' column (assuming it's in liters, e.g., '39.42 л')
    df_filtered['Слито_liters'] = df_filtered['Слито'].str.replace(' л', '').str.replace(',', '.').astype(float)

    # You can add more data processing steps here as needed

    return df_filtered

def main():
    # List to hold DataFrames for each month
    monthly_data = []

    # Define the English three-letter abbreviations for months
    month_abbrs = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                  'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    # Mapping of month abbreviations to their corresponding digit
    month_mapping = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,
        'may': 5, 'jun': 6, 'jul': 7, 'aug': 8,
        'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }

    type_abbr = "bus"

    # Directory containing the ZIP files (current directory)
    zip_dir = os.getcwd()

    # Iterate over each month's abbreviation
    for abbr in month_abbrs:
        # Get the corresponding month digit
        month_digit = month_mapping.get(abbr.lower())

        # Construct the expected ZIP filename
        zip_filename = f"{type_abbr}_{abbr}.zip"
        zip_path = os.path.join(zip_dir, zip_filename)

        # Check if the ZIP file exists
        if not os.path.isfile(zip_path):
            print(f"Warning: '{zip_filename}' does not exist in '{zip_dir}'. Skipping.")
            continue

        print(f"Processing '{zip_filename}'...")
        processed_df = process_bus_csv_from_zip(zip_path, type_abbr, abbr, month_digit)

        if processed_df is not None:
            monthly_data.append(processed_df)
            print(f"Successfully processed '{zip_filename}'.")
        else:
            print(f"No data appended for '{zip_filename}'.")

    if not monthly_data:
        print("No data processed from any ZIP archives. Exiting.")
        return

    # Concatenate all monthly DataFrames into one large DataFrame
    combined_df = pd.concat(monthly_data, ignore_index=True)

    # Define the output Excel filename
    output_filename = "Сливы_Автобусы_2024.xlsx"

    try:
        # Save the combined DataFrame to an Excel file
        combined_df.to_excel(output_filename, index=False)
        print(f"All data successfully combined and saved to '{output_filename}'.")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

    # Optionally, display the first few rows of the combined DataFrame
    print("Preview of the combined data:")
    print(combined_df.head())

if __name__ == "__main__":
    main()
