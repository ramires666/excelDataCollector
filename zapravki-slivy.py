import pandas as pd
import os


def extract_month_info(df, time_col='Время'):
    """
    Extracts month number and month name from the time column in the DataFrame.
    Assumes European date format (day-month-year).
    """
    df['Время_parsed'] = pd.to_datetime(df[time_col], errors='coerce', dayfirst=True)

    month_names = {
        1: 'январь', 2: 'февраль', 3: 'март',
        4: 'апрель', 5: 'май', 6: 'июнь',
        7: 'июль', 8: 'август', 9: 'сентябрь',
        10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'
    }

    df['MonthNumber'] = df['Время_parsed'].dt.month
    df['MonthName'] = df['MonthNumber'].map(month_names)
    return df


def process_sheet_zapravki(file_path):
    """
    Processes the 'Заправки' sheet: reads the sheet, filters out summary rows,
    extracts month info, and returns the processed DataFrame.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Заправки', dtype={'№': str})
        print(f"Read {len(df)} rows from {file_path} (Заправки)")
    except Exception as e:
        print(f"Error reading 'Заправки' from {file_path}: {e}")
        return None

    if '№' in df.columns:
        initial_count = len(df)
        df = df[~df['№'].astype(str).str.contains(r'\.')]
        print(f"Filtered out {initial_count - len(df)} summary rows in Заправки from {file_path}")
    else:
        print(f"Warning: '№' column not found in {file_path} for Заправки.")

    df = extract_month_info(df, time_col='Время')
    before_filter = len(df)
    df = df[df['Время_parsed'].notna()]
    print(f"Filtered out {before_filter - len(df)} invalid date rows in Заправки from {file_path}")
    return df


def process_sheet_slivi(file_path):
    """
    Processes the 'Сливы' sheet: reads the sheet, filters out summary rows if applicable,
    extracts month info, and returns the processed DataFrame.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Сливы', dtype={'№': str})
        print(f"Read {len(df)} rows from {file_path} (Сливы)")
    except Exception as e:
        print(f"Error reading 'Сливы' from {file_path}: {e}")
        return None

    if '№' in df.columns:
        initial_count = len(df)
        df = df[~df['№'].astype(str).str.contains(r'\.')]
        print(f"Filtered out {initial_count - len(df)} summary rows in Сливы from {file_path}")
    else:
        print(f"Warning: '№' column not found in {file_path} for Сливы.")

    df = extract_month_info(df, time_col='Время')
    before_filter = len(df)
    df = df[df['Время_parsed'].notna()]
    print(f"Filtered out {before_filter - len(df)} invalid date rows in Сливы from {file_path}")
    return df


def load_aggregate_and_merge(directory, output_file):
    all_zapravki = []
    all_slivi = []

    if not os.path.isdir(directory):
        print(f"Error: The directory '{directory}' does not exist.")
        return

    # Process each file in the directory for both sheets
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            file_path = os.path.join(directory, file)

            df_z = process_sheet_zapravki(file_path)
            if df_z is not None and not df_z.empty:
                all_zapravki.append(df_z)

            df_s = process_sheet_slivi(file_path)
            if df_s is not None and not df_s.empty:
                all_slivi.append(df_s)

    # Aggregate Заправки
    aggregated_zapravki = pd.DataFrame()
    if all_zapravki:
        combined_zapravki = pd.concat(all_zapravki, ignore_index=True)
        print(f"Combined Заправки data shape: {combined_zapravki.shape}")
        aggregated_zapravki = combined_zapravki.groupby(
            ['Группировка', 'MonthNumber', 'MonthName'],
            as_index=False
        ).agg({'Заправлено': 'sum'})

    # Aggregate Сливы
    aggregated_slivi = pd.DataFrame()
    if all_slivi:
        combined_slivi = pd.concat(all_slivi, ignore_index=True)
        print(f"Combined Сливы data shape: {combined_slivi.shape}")
        aggregated_slivi = combined_slivi.groupby(
            ['Группировка', 'MonthNumber', 'MonthName'],
            as_index=False
        ).agg({'Слито': 'sum'})

    # Merge the two aggregated DataFrames
    merged_df = None
    if not aggregated_zapravki.empty and not aggregated_slivi.empty:
        merged_df = pd.merge(aggregated_zapravki, aggregated_slivi,
                             on=['Группировка', 'MonthNumber', 'MonthName'], how='outer')
    elif not aggregated_zapravki.empty:
        merged_df = aggregated_zapravki.copy()
    elif not aggregated_slivi.empty:
        merged_df = aggregated_slivi.copy()

    # Write results to Excel
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            if not aggregated_zapravki.empty:
                aggregated_zapravki.to_excel(writer, sheet_name='Заправки_aggregated', index=False)
            if not aggregated_slivi.empty:
                aggregated_slivi.to_excel(writer, sheet_name='Сливы_aggregated', index=False)
            if merged_df is not None and not merged_df.empty:
                merged_df.to_excel(writer, sheet_name='Merged', index=False)
        print(f"Results saved to '{output_file}'")
    except Exception as e:
        print(f"Error saving results to '{output_file}': {e}")


if __name__ == "__main__":
    directory_path = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\Autobus_zapravki-slivy"
    directory_path = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\Gruz-zapravki-slivy"
    directory_path = r"C:\Users\delxps\PycharmProjects\excelCollector\mch\Spec-zapravki-slivy"
    output_path = "aggregated_output.xlsx"
    load_aggregate_and_merge(directory_path, output_path)
