import os
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from glob import glob
from pymorphy2 import MorphAnalyzer
import openpyxl
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore', category=FutureWarning)
import numpy as np

morph = MorphAnalyzer()

import re


def extract_month_from_filename(filename):
    # Получаем базовое имя файла без расширения
    basename = os.path.basename(filename)
    basename_no_ext = os.path.splitext(basename)[0]

    # Приводим к нижнему регистру
    basename_no_ext = basename_no_ext.lower()

    # Убираем знаки препинания и заменяем символы подчеркивания на пробелы
    basename_no_ext = re.sub(r'[^\w\s]', ' ', basename_no_ext)
    basename_no_ext = basename_no_ext.replace('_', ' ')

    # Разбиваем на слова
    words = re.findall(r'\w+', basename_no_ext)

    # Словарь месяцев
    months = {
        "январь": 1, "февраль": 2, "март": 3, "апрель": 4, "май": 5,
        "июнь": 6, "июль": 7, "август": 8, "сентябрь": 9,
        "октябрь": 10, "ноябрь": 11, "декабрь": 12
    }

    for word in words:
        # Получаем нормальную форму слова
        parsed = morph.parse(word)[0]
        normal_form = parsed.normal_form
        if normal_form in months:
            return months[normal_form], normal_form

    # Если месяц не найден
    return None, None


def get_yellow_columns(file_path, sheet_name):
    # Загружаем книгу и выбираем лист
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook[sheet_name]

    yellow_columns = []

    # Проход по всем колонкам и строкам листа
    for column in sheet.iter_cols():
        yellow_count = 0
        for cell in column:
            # Пропускаем объединенные ячейки
            if isinstance(cell, openpyxl.cell.cell.MergedCell) or sheet.row_dimensions[cell.row].hidden:
                # continue
                continue
            # Проверяем, есть ли заливка у ячейки
            if cell.fill:
                fill_color = cell.fill.fgColor
                # print(f"{cell.column_letter}-{cell.coordinate} {fill_color.rgb}")
                # Проверка RGB значения на желтый цвет, если тип цвета 'rgb'
                if fill_color and fill_color.type == "rgb" and fill_color.rgb == "FFFFFF00":
                    yellow_count += 1

        # Если в колонке более 15 строк окрашены в желтый
        if yellow_count > 9:
            column_number = column[0].column
            yellow_columns.append(column_number-1)
            continue

    # Выводим список букв колонок, в которых более 5 ячеек с желтой заливкой
    if yellow_columns:
        yellow_column_letters = [get_column_letter(col) for col in yellow_columns]
        print(f"Колонки с большиством ячеек с желтой заливкой: {', '.join(yellow_column_letters)}")
    else:
        print("Нет колонок с более чем 15 ячейками с желтой заливкой")

    # Возвращаем список номеров колонок с желтой заливкой
    return yellow_columns


# Функция для проверки строки
def is_valid_row(row):
    # Указание индексов столбцов для проверки
    columns_to_check = ['Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg']  # Или 'C', 'D', 'R', 'F', 'G', 'H'

    # Проверяем только указанные столбцы
    for col in columns_to_check:
        value = row[col]
        if isinstance(value, int) and value != 0:
            continue
        return True  # Если хотя бы одно значение не целое или равно 0
    return False  # Если все значения являются целыми и не равны 0


def fill_merged_cells_in_first_yellow_column(sheet, first_column):
    """
    Заполняет объединённые ячейки в указанной колонке.

    :param sheet: объект листа Excel.
    :param first_column: индекс первой жёлтой колонки (1-based).
    """
    for merged_range in sheet.merged_cells.ranges:
        if merged_range.min_col == first_column and merged_range.max_col == first_column:  # Только указанная колонка
            start_row = merged_range.min_row
            end_row = merged_range.max_row
            value = sheet.cell(row=start_row, column=first_column).value  # Читаем значение верхней левой ячейки

            # Разъединяем объединённые ячейки
            sheet.unmerge_cells(str(merged_range))
            for row in range(start_row, end_row + 1):
                if not sheet.row_dimensions[row].hidden:  # Пропускаем скрытые строки
                    sheet.cell(row=row, column=first_column).value = value  # Устанавливаем значение


def load_excel_data_with_flex(folder_path):
    """Load Excel data from a folder considering month flexions in file names."""
    all_files = glob(os.path.join(folder_path, "*.xlsx"))
    final_columns = ['Panel', 'Shtrek', 'Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg', 'Uchastok', 'month']
    final_df = pd.DataFrame(columns=final_columns)

    for file_path in all_files:
        file_name = os.path.basename(file_path)
        month_num, month_name = extract_month_from_filename(file_name)
        if not month_num:
            print(f"нет месяца - пропускаю файл: {file_name}")
            continue  # Skip files without recognizable month

        normalized_path = os.path.normpath(folder_path)
        # Извлекаем имя последней папки
        last_folder_name = os.path.basename(normalized_path)

        wb = load_workbook(file_path, data_only=True, read_only=False)
        # sheet_names = wb.sheetnames[1:]  # Exclude the first sheet

        # Список всех видимых листов
        visible_sheets = [sheet_name for sheet_name in wb.sheetnames if wb[sheet_name].sheet_state == "visible"]

        # Пропуск первого видимого листа
        visible_sheets = visible_sheets[1:] if len(visible_sheets) > 1 else []

        for sheet_name in visible_sheets:
            ws = wb[sheet_name]



            print(f"\n>>>>>>>>>>>>>>>>   {last_folder_name} == {month_name} ===  {sheet_name}  <<<<<<<<<<<<<<<<<<<<<<<<<<")
            # print(file_name)


            # Skip empty sheets
            if ws.max_row == 0 or ws.max_column == 0:
                print(f"Sheet '{sheet_name}' is empty.")
                continue

            # Get yellow columns using the provided function
            yellow_cols = get_yellow_columns(file_path, sheet_name)
            if not yellow_cols:
                print(f"No yellow columns found in sheet '{sheet_name}' in file '{file_name}'.")
                continue

            # Берём первую жёлтую колонку
            first_yellow_column = yellow_cols[0]  # Первую жёлтую колонку (1-based)

            # Заполняем объединённые ячейки только в первой жёлтой колонке =+1 так как нумерация в ексель с 1
            fill_merged_cells_in_first_yellow_column(ws, first_yellow_column+1)


            # Try to load data and set columns
            data_rows = []
            for idx, row in enumerate(ws.iter_rows(values_only=False), start=ws.min_row):
                if ws.row_dimensions[idx].hidden:
                    # Пропускаем скрытые строки
                    continue
                # Извлекаем значения ячеек
                cell_values = [cell.value for cell in row]
                data_rows.append(cell_values)

            if not data_rows:
                print(f"Sheet '{sheet_name}' has no data after removing hidden rows.")
                continue

            # Replace None in header with empty string
            header = [cell if cell is not None else '' for cell in data_rows[0]]
            data = pd.DataFrame(data_rows[1:], columns=header)

            # Filter yellow columns
            data = data.iloc[:, yellow_cols]

            # Set appropriate columns based on the sheet name
            if 'ОГР' in sheet_name.upper():
                ogr_columns = ['Panel', 'Ruda', 'Cu', 'fRuda', 'fCu', 'Ag', 'fAg']
                if len(data.columns) == len(ogr_columns):
                    data.columns = ogr_columns
                    data['Shtrek'] = ''  # Add empty 'Shtrek' column
                else:
                    print(f"Warning: Sheet '{sheet_name}' has unexpected number of columns for OGR sheet.")
                    continue
            else:
                standard_columns = ['Panel', 'Shtrek', 'Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg']
                if len(data.columns) == len(standard_columns):
                    data.columns = standard_columns
                else:
                    print(f"Warning: Sheet '{sheet_name}' has unexpected number of columns.")
                    continue

            # Fill down the 'Panel' values
            # data['Panel'].ffill(inplace=True)

            # Filter out rows where 'Shtrek' is empty (only if not OGR sheet)
            if 'ОГР' not in sheet_name.upper():
                data = data[data['Shtrek'].notnull() & (data['Shtrek'] != '')]

            # if 'ОГР' in sheet_name.upper():
            # убираем строки с пустой панелью
            data = data[data['Panel'].notnull() & (data['Panel'] != '')]

            # Удаление первой строки
            data = data.iloc[1:]  # Пропускаем первую строку

            # Применение фильтра только целые числа
            data = data[data.apply(is_valid_row, axis=1)]

            data['Uchastok'] = sheet_name
            data['month'] = month_name

            # Ensure all columns are in data
            for col in final_columns:
                if col not in data.columns:
                    data[col] = ''

            data = data[final_columns]

            final_df = pd.concat([final_df, data], ignore_index=True)

            final_df = add_calculated_columns(final_df)

    # Save the final DataFrame to a single Excel file
    # Use folder name as output file name
    t = datetime.now().microsecond
    output_file_name = os.path.basename(folder_path.strip('/\\')) +str(t)+ '.xlsx'
    output_file = os.path.join(folder_path, output_file_name)

    # Создаем дополнительные таблицы
    aggregated_df = aggregate_and_calculate(final_df)
    monthly_avg_df = calculate_monthly_averages(aggregated_df)
    monthly_percent_df = calculate_monthly_average_percentages(final_df)

    # Пишем все таблицы в один файл
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Data", index=False)
        aggregated_df.to_excel(writer, sheet_name="Aggregated Data", index=False)
        monthly_avg_df.to_excel(writer, sheet_name="Monthly Averages", index=False)
        monthly_percent_df.to_excel(writer, sheet_name="Monthly Percentages", index=False)

    return final_df

    # final_df.to_excel(output_file, index=False)
    # return final_df


def add_calculated_columns(df):
    """
    Добавляет расчетные столбцы в DataFrame:
    - dRuda, dCu, dAg
    - %Ruda, %Cu, %Ag

    :param df: pandas DataFrame, содержащий исходные столбцы
    :return: pandas DataFrame с добавленными расчетными столбцами
    """

    df['Ruda'] = df['Ruda'].fillna(0)
    df['Cu'] = df['Cu'].fillna(0)
    df['Ag'] = df['Ag'].fillna(0)
    df['fRuda'] = df['fRuda'].fillna(0)
    df['fCu'] = df['fCu'].fillna(0)
    df['fAg'] = df['fAg'].fillna(0)

    # Расчет dRuda, dCu, dAg (с обработкой ошибок, замена NaN на 0)
    df['diff-Ruda'] = abs(df['Ruda'] - df['fRuda'])
    df['diff-Cu'] = abs(df['Cu'] - df['fCu'])
    df['diff-Ag'] = abs(df['Ag'] - df['fAg'])

    # Замена NaN на 0, если расчет невозможен
    df['diff-Ruda'] = df['diff-Ruda'].fillna(0)
    df['diff-Cu'] = df['diff-Cu'].fillna(0)
    df['diff-Ag'] = df['diff-Ag'].fillna(0)

    # Расчет %Ruda, %Cu, %Ag (с обработкой деления на ноль или NaN)
    # df['%rel-Ruda'] = abs(np.where(df['Ruda'] != 0, df['diff-Ruda'] / df['Ruda'], 1.0) * 100)  # 100% если Ruda == 0
    # df['%rel-Cu'] = abs(np.where(df['Cu'] != 0, df['diff-Cu'] / df['Cu'], 1.0) * 100)  # 100% если Cu == 0
    # df['%rel-Ag'] = abs(np.where(df['Ag'] != 0, df['diff-Ag'] / df['Ag'], 1.0) * 100)  # 100% если Ag == 0

    # Расчет %Ruda, %Cu, %Ag (с обработкой деления на ноль или NaN и ограничением максимум 100)
    df['%rel-Ruda'] = np.clip(abs(np.where(df['Ruda'] != 0, df['diff-Ruda'] / df['Ruda'], 1.0) * 100), 0, 100)
    df['%rel-Cu'] = np.clip(abs(np.where(df['Cu'] != 0, df['diff-Cu'] / df['Cu'], 1.0) * 100), 0, 100)
    df['%rel-Ag'] = np.clip(abs(np.where(df['Ag'] != 0, df['diff-Ag'] / df['Ag'], 1.0) * 100), 0, 100)

    # Замена NaN на 100%, если расчет невозможен
    df['%rel-Ruda'] = df['%rel-Ruda'].fillna(100)
    df['%rel-Cu'] = df['%rel-Cu'].fillna(100)
    df['%rel-Ag'] = df['%rel-Ag'].fillna(100)

    return df

def add_calculated_columns(df):
    """
    Добавляет расчетные столбцы в DataFrame:
    - dRuda, dCu, dAg
    - %Ruda, %Cu, %Ag

    :param df: pandas DataFrame, содержащий исходные столбцы
    :return: pandas DataFrame с добавленными расчетными столбцами
    """
    # Преобразование колонок в числовой формат
    numeric_columns = ['Ruda', 'fRuda', 'Cu', 'fCu', 'Ag', 'fAg']
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)  # Преобразование и замена NaN на 0

    # Расчет dRuda, dCu, dAg
    df['diff-Ruda'] = abs(df['Ruda'] - df['fRuda'])
    df['diff-Cu'] = abs(df['Cu'] - df['fCu'])
    df['diff-Ag'] = abs(df['Ag'] - df['fAg'])

    # Расчет %Ruda, %Cu, %Ag
    df['%rel-Ruda'] = abs(np.where(df['Ruda'] != 0, df['diff-Ruda'] / df['Ruda'], 1.0) * 100)
    df['%rel-Cu'] = abs(np.where(df['Cu'] != 0, df['diff-Cu'] / df['Cu'], 1.0) * 100)
    df['%rel-Ag'] = abs(np.where(df['Ag'] != 0, df['diff-Ag'] / df['Ag'], 1.0) * 100)

    return df



def aggregate_and_calculate(df):
    """
    Группирует DataFrame по месяцам и штрекам, суммирует указанные колонки
    и добавляет расчетные колонки.

    :param df: Исходный DataFrame
    :return: Новый DataFrame с расчетными колонками
    """
    # Указываем, какие колонки нужно суммировать
    sum_columns = ['Ruda', 'fRuda', 'Cu', 'fCu', 'Ag', 'fAg']

    # Группировка по месяцам и панелям, вычисление суммы
    grouped = df.groupby(['month', 'Panel'], as_index=False)[sum_columns].sum()

    # Добавляем расчетные колонки
    grouped['diff-Ruda'] = grouped['Ruda'] - grouped['fRuda']
    grouped['diff-Cu'] = grouped['Cu'] - grouped['fCu']
    grouped['diff-Ag'] = grouped['Ag'] - grouped['fAg']

    # Замена ошибок на 0
    grouped['diff-Ruda'] = grouped['diff-Ruda'].fillna(0)
    grouped['diff-Cu'] = grouped['diff-Cu'].fillna(0)
    grouped['diff-Ag'] = grouped['diff-Ag'].fillna(0)

    # Добавляем расчетные колонки в процентах с ограничением максимум 100
    grouped['%rel-Ruda'] = np.clip(
        abs(np.where(grouped['Ruda'] != 0, grouped['diff-Ruda'] / grouped['Ruda'], 1.0) * 100), 0, 100
    )
    grouped['%rel-Cu'] = np.clip(
        abs(np.where(grouped['Cu'] != 0, grouped['diff-Cu'] / grouped['Cu'], 1.0) * 100), 0, 100
    )
    grouped['%rel-Ag'] = np.clip(
        abs(np.where(grouped['Ag'] != 0, grouped['diff-Ag'] / grouped['Ag'], 1.0) * 100), 0, 100
    )

    # # Замена NaN на 100%
    # grouped['%rel-Ruda'] = grouped['%rel-Ruda'].fillna(100)
    # grouped['%rel-Cu'] = grouped['%rel-Cu'].fillna(100)
    # grouped['%rel-Ag'] = grouped['%rel-Ag'].fillna(100)

    return grouped





def calculate_monthly_averages(df):
    """
    Группирует DataFrame по месяцам, вычисляет абсолютное среднее значение для %tRuda, %tCu, %tAg
    и записывает вычитание из 1 в виде процентов в колонки Ruda, Cu, Ag.
    Исключает февраль и сентябрь из расчета.

    :param df: DataFrame, содержащий колонки month, %tRuda, %tCu, %tAg.
    :return: Новый DataFrame с агрегированными данными.
    """

    # Средние значения по месяцам с расчётом абсолютных значений
    monthly_avg = df.groupby('month', as_index=False)[['%rel-Ruda', '%rel-Cu', '%rel-Ag']].apply(
        lambda x: np.clip(x.abs().mean(), 0, 100)
    )

    # Вычитание из 1, перевод в проценты и ограничение максимум 100
    monthly_avg['aver-Ruda'] = np.clip((1 - monthly_avg['%rel-Ruda'] / 100) * 100, 0, 100)
    monthly_avg['aver-Cu'] = np.clip((1 - monthly_avg['%rel-Cu'] / 100) * 100, 0, 100)
    monthly_avg['aver-Ag'] = np.clip((1 - monthly_avg['%rel-Ag'] / 100) * 100, 0, 100)

    # Оставляем только нужные колонки
    monthly_avg = monthly_avg[['month', 'aver-Ruda', 'aver-Cu', 'aver-Ag']]
    return monthly_avg



def calculate_monthly_average_percentages(df):
    """
    Группирует исходный DataFrame по месяцам, вычисляет абсолютное среднее значение для %Ruda, %Cu, %Ag,
    а затем вычитает это значение из 1, переводит в проценты и ограничивает максимум 100.

    :param df: Исходный DataFrame с колонками %Ruda, %Cu, %Ag.
    :return: Новый DataFrame с агрегированными данными.
    """
    # Группировка по месяцам и вычисление абсолютных средних значений
    monthly_avg = df.groupby('month', as_index=False)[['%rel-Ruda', '%rel-Cu', '%rel-Ag']].apply(
        lambda x: np.clip(x.abs().mean(), 0, 100)
    )



    # Вычитание средних значений из 1, взятие абсолютного значения, перевод в проценты и ограничение результата максимум 100
    monthly_avg['Ruda'] = np.clip(abs((1 - monthly_avg['%rel-Ruda'] / 100) * 100), 0, 100)
    monthly_avg['Cu'] = np.clip(abs((1 - monthly_avg['%rel-Cu'] / 100) * 100), 0, 100)
    monthly_avg['Ag'] = np.clip(abs((1 - monthly_avg['%rel-Ag'] / 100) * 100), 0, 100)

    # Оставляем только нужные колонки
    monthly_avg = monthly_avg[['month', 'Ruda', 'Cu', 'Ag']]
    return monthly_avg




# Пример использования:
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\Жиланды ИПГ 2024'
# directory =r'C:\Users\delxps\Documents\Kazakhmys\_alibek\ЗР ИПГ 2024'
directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\Жомарт ИПГ 2024'
directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\ВЖР ИПГ 2024'
final_data = load_excel_data_with_flex(directory)
