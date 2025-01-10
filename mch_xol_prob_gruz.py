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
    Returns total hours as a float.
    """
    if not isinstance(time_str, str):
        print(f"Warning: Expected string for time parsing but got '{time_str}'. Setting to 0.0")
        return 0.0

    # Normalize whitespace and lower case
    normalized_str = ' '.join(time_str.strip().lower().split())
    # Debug: Print the normalized string
    # print(f"Normalized time string: '{normalized_str}'")

    # Regular expression to capture optional days and time
    # Updated to match 'день', 'дня', 'дней'
    pattern = re.compile(
        r'^(?:(\d+)\s+д(?:ень|ня|ней))?\s*(\d{1,2}):(\d{2}):(\d{2})$',
        re.IGNORECASE
    )

    match = pattern.match(normalized_str)
    if not match:
        # If the format doesn't match, attempt to parse manually
        print(f"Warning: Unrecognized time format '{time_str}'. Setting to 0.0")
        return 0.0

    days_str, h_str, m_str, s_str = match.groups()
    days = int(days_str) if days_str else 0
    h, m, s = map(int, [h_str, m_str, s_str])
    total_hours = days * 24 + h + m / 60 + s / 3600
    return total_hours

def parse_kilometers(km_value):
    """
    Function to parse the "Пробег" column by handling both numeric and string inputs.
    - If the input is a string, it removes ' км' and converts it to float.
    - If the input is already a numeric type, it returns it as float.
    - If parsing fails, it returns 0.0.
    """
    try:
        if pd.isna(km_value):
            return 0.0
        if isinstance(km_value, (int, float)):
            return float(km_value)
        if isinstance(km_value, str):
            # Remove ' км', 'км', and any surrounding whitespace
            km_clean = km_value.replace(' км', '').replace('км', '').replace(' ', '').strip()
            # Replace comma with dot if necessary
            km_clean = km_clean.replace(',', '.')
            return float(km_clean)
        # If the value is neither string nor numeric, attempt to convert to float
        return float(km_value)
    except (ValueError, AttributeError) as e:
        print(f"Warning: Could not parse 'Пробег' value '{km_value}'. Setting to 0.0. Error: {e}")
        return 0.0

def load_and_merge_xlsx(directory, output_file):
    all_data = []

    # Check if directory exists
    if not os.path.isdir(directory):
        print(f"Error: The directory '{directory}' does not exist.")
        return

    # Iterate over all files in the directory
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            file_path = os.path.join(directory, file)
            try:
                # Load the second sheet (sheet index starts at 0)
                df = pd.read_excel(
                    file_path,
                    sheet_name=1,
                    usecols=['Группировка', 'Моточасы', 'В движении', 'Холостой ход', 'Пробег']
                )

                # Parse time columns
                time_columns = ['Моточасы', 'В движении', 'Холостой ход']
                for col in time_columns:
                    df[col] = df[col].apply(parse_time_str)

                # Parse "Пробег" column
                df['Пробег'] = df['Пробег'].apply(parse_kilometers)

                all_data.append(df)
            except ValueError as ve:
                print(f"ValueError processing file {file}: {ve}")
            except KeyError as ke:
                print(f"KeyError processing file {file}: Missing column {ke}")
            except Exception as e:
                print(f"Error processing file {file}: {e}")

    if not all_data:
        print("No data to combine.")
        return

    # Combine all data into one DataFrame
    combined_df = pd.concat(all_data, ignore_index=True)

    # Debug: Check if 'Пробег' has been parsed correctly
    if combined_df['Пробег'].isnull().any():
        print("Warning: Some 'Пробег' values are NaN.")

    # Aggregate data by 'Группировка'
    aggregated_df = combined_df.groupby('Группировка').agg({
        'Моточасы': 'sum',    # Total hours
        'В движении': 'sum',   # Total hours
        'Холостой ход': 'sum', # Total hours
        'Пробег': 'sum'        # Total kilometers
    }).reset_index()

    # Optional: Round the time columns to 2 decimal places
    for col in ['Моточасы', 'В движении', 'Холостой ход']:
        aggregated_df[col] = aggregated_df[col].round(2)

    # Optional: Round "Пробег" to 2 decimal places
    aggregated_df['Пробег'] = aggregated_df['Пробег'].round(2)

    # Save the aggregated DataFrame to an Excel file
    try:
        aggregated_df.to_excel(output_file, index=False)
        print(f"Aggregated data saved to '{output_file}'")
    except Exception as e:
        print(f"Error saving aggregated data to '{output_file}': {e}")

# Example usage
if __name__ == "__main__":
    directory_path = r"C:\Users\delxps\PycharmProjects\excelCollector\mch"  # Update this path as needed
    output_path = "aggregated_output.xlsx"  # You can specify a full path if needed
    load_and_merge_xlsx(directory_path, output_path)
