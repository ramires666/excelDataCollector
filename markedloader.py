import os
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from glob import glob
from pymorphy2 import MorphAnalyzer
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
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
        #add horizonts to panel
        yellow_columns.insert(0, yellow_columns[:1][0] - 1)
        yellow_column_letters = [get_column_letter(col+1) for col in yellow_columns]
        print(f"Колонки с большиством ячеек с желтой заливкой: {', '.join(yellow_column_letters)}")
    else:
        print("Нет колонок с более чем 15 ячейками с желтой заливкой")

    # Возвращаем список номеров колонок с желтой заливкой
    return yellow_columns


# Функция для проверки строки
def is_valid_row(row):
    # Указание индексов столбцов для проверки
    columns_to_check = ['Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg']  # Или 'C', 'D', 'R', 'F', 'G', 'H'
    # columns_to_check = ['Ruda', 'Cu', 'fRuda', 'fCu', 'fAg']  # Или 'C', 'D', 'R', 'F', 'G', 'H'

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


def convert_xls_to_xlsx(xls_file_path, xlsx_file_path):
    import pandas as pd
    # Читаем файл .xls с помощью xlrd
    df_dict = pd.read_excel(xls_file_path, sheet_name=None, engine='xlrd')
    # Пишем в файл .xlsx с помощью openpyxl
    with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def convert_xlsb_to_xlsx_with_formatting(xlsb_file_path, xlsx_file_path):
    """
    Конвертирует файл .xlsb в .xlsx с сохранением форматирования.
    Требуется установленный Microsoft Excel.
    """
    import win32com.client as win32
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        # Открываем файл .xlsb
        workbook = excel.Workbooks.Open(xlsb_file_path)
        # Сохраняем как .xlsx
        workbook.SaveAs(xlsx_file_path, FileFormat=51)  # 51 - это код для .xlsx
        workbook.Close()
    except Exception as e:
        print(f"Ошибка при конвертации файла {xlsb_file_path}: {e}")
    finally:
        excel.Quit()



def convert_xls_to_xlsx_with_formatting(xls_file_path, xlsx_file_path):
    """
    Конвертирует файл .xls в .xlsx с сохранением форматирования.
    Требуется установленный Microsoft Excel.
    """
    import win32com.client as win32
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(xls_file_path)
        workbook.SaveAs(xlsx_file_path, FileFormat=51)  # 51 - это формат .xlsx
        workbook.Close()
    except Exception as e:
        print(f"Ошибка при конвертации файла {xls_file_path}: {e}")
    finally:
        excel.Quit()



def load_excel_data_with_flex(folder_path):
    """Загружает данные из Excel файлов, учитывая возможные вариации в названиях месяцев в именах файлов."""
    # Шаг 1: Конвертируем все .xls файлы в .xlsx
    xls_files = glob(os.path.join(folder_path, "*.xls"))
    for xls_file in xls_files:
        # Генерируем путь для .xlsx файла
        xlsx_file = os.path.splitext(xls_file)[0] + '.xlsx'
        # Проверяем, существует ли уже конвертированный файл, чтобы избежать повторной конвертации
        if not os.path.exists(xlsx_file):
            try:
                convert_xls_to_xlsx_with_formatting(xls_file, xlsx_file)
                print(f"Конвертирован файл {xls_file} в {xlsx_file}")
            except Exception as e:
                print(f"Ошибка при конвертации файла {xls_file}: {e}")
                continue  # Переходим к следующему файлу, если конвертация не удалась

    xlsb_files = glob(os.path.join(folder_path, "*.xlsb"))
    for xlsb_file in xlsb_files:
        # Генерируем путь для .xlsx файла
        xlsx_file = os.path.splitext(xlsb_file)[0] + '.xlsx'
        # Проверяем, существует ли уже конвертированный файл, чтобы избежать повторной конвертации
        if not os.path.exists(xlsx_file):
            try:
                convert_xlsb_to_xlsx_with_formatting(xlsb_file, xlsx_file)
                print(f"Конвертирован файл {xlsb_file} в {xlsx_file}")
            except Exception as e:
                print(f"Ошибка при конвертации файла {xlsb_file}: {e}")
                continue  # Переходим к следующему файлу, если конвертация не удалась



    # Шаг 2: Обрабатываем только .xlsx файлы
    xlsx_files = glob(os.path.join(folder_path, "*.xlsx"))

    final_columns = ['Horizont','Panel', 'Shtrek', 'Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg', 'Uchastok', 'month']
    final_df = pd.DataFrame(columns=final_columns)


    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        month_num, month_name = extract_month_from_filename(file_name)
        if not month_num:
            print(f"-> -> ->  нет месяца - пропускаю файл: {file_name}")
            continue  # Skip files without recognizable month

        normalized_path = os.path.normpath(folder_path)
        # Извлекаем имя последней папки
        last_folder_name = os.path.basename(normalized_path)

        wb = load_workbook(file_path, data_only=True, read_only=False)
        # sheet_names = wb.sheetnames[1:]  # Exclude the first sheet

        # Список всех видимых листов
        visible_sheets = [sheet_name for sheet_name in wb.sheetnames if wb[sheet_name].sheet_state == "visible"]

        if len(visible_sheets) > 1:
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
            # и во второй
            fill_merged_cells_in_first_yellow_column(ws, first_yellow_column + 2)


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
                standard_columns = ['Horizont','Panel', 'Shtrek', 'Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg']   # STANDART
                standard_columns = ['Horizont','Panel', 'Shtrek', 'Ruda', 'Cu', 'fRuda', 'fCu', 'Ag', 'fAg']       # Nurkazgan
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
            # Filter rows where both Ruda and fRuda are 0
            data = data[~((data['Ruda'] == 0) & (data['fRuda'] == 0))]

            final_df = pd.concat([final_df, data], ignore_index=True)

            final_df = add_calculated_columns(final_df)

    final_df = final_df[~((final_df['Ruda'] == 0) & (final_df['fRuda'] == 0))]


    # Save the final DataFrame to a single Excel file
    # Use folder name as output file name
    t = datetime.now().microsecond
    # output_file_name = os.path.basename("_report "+folder_path.strip('/\\')) +str(t)+ '.xlsx'
    # output_file = os.path.join(folder_path, output_file_name)

    generate_report_with_charts(folder_path, final_df)
    return final_df



def add_summary_and_formula_rows1(output_file):
    # Load the workbook and select the first sheet
    workbook = load_workbook(output_file)
    sheet = workbook["Данные"]  # Adjust this to your first sheet name if different

    # Find the last row with data
    last_row = sheet.max_row

    # Define the columns of interest
    columns = {'p>0': 'Q', 'f>0': 'R', 'p=0': 'S', 'f=0': 'T'}

    # Add the sum row
    sum_row = last_row + 1
    for col_name, col_letter in columns.items():
        formula = f"=SUM({col_letter}2:{col_letter}{last_row})"
        sheet[f"{col_letter}{sum_row}"] = formula

    # Add the count row with the formula
    count_row = last_row + 2
    for col_name, col_letter in columns.items():
        if col_name in ['p>0', 'f>0']:  # Columns with ">0" condition
            formula = f'=COUNTIF({col_letter}2:{col_letter}{last_row},">0")'
            sheet[f"{col_letter}{count_row}"] = formula
        elif col_name in ['p=0', 'f=0']:  # Columns with "=0" condition
            formula = f'=COUNTIF({col_letter}2:{col_letter}{last_row},"=0")'
            sheet[f"{col_letter}{count_row}"] = formula


    # Add percentage calculations to the left of the sums row
    perc_column = 'U'  # Insert percentages in column P
    sheet[f"{perc_column}{sum_row}"] = f"=R{sum_row}/Q{sum_row}"  # R361/Q361
    sheet[f"{perc_column}{count_row}"] = f"=R{count_row}/Q{count_row}"  # R362/Q362

    # Format percentage cells
    percent_style = NamedStyle(name="percent_style", number_format="0.00%")
    workbook.add_named_style(percent_style)  # Ensure the style is added once
    sheet[f"{perc_column}{sum_row}"].style = "percent_style"
    sheet[f"{perc_column}{count_row}"].style = "percent_style"

    # Force recalculation in Excel
    workbook.properties.calcMode = "auto"

    # Save the workbook
    workbook.save(output_file)



def add_summary_and_formula_rows(output_file):
    # Load the workbook and select the first sheet
    workbook = load_workbook(output_file)
    sheet = workbook["Данные"]  # Adjust this to your first sheet name if different

    # Find the last row with data
    last_row = sheet.max_row

    # Define the columns for existing data
    columns = {'p': 'Q', 'f': 'R'}  # Assuming 'p' in column Q and 'f' in column R

    # Add titles for new columns
    sheet['S1'] = 'Bbln (p=1 and f=1)'
    sheet['T1'] = 'HeBbln (p=1 and f=0)'
    sheet['U1'] = 'BHE (p=0 and f=1)'

    # Add columns for PP, VN, NP
    pp_col, vn_col, np_col = 'S', 'U', 'T'  # Assign new columns for PP, VN, NP

    for row in range(2, last_row + 1):
        # PP = 1 when p=1 and f=1, otherwise 0
        sheet[f"{pp_col}{row}"] = f"=IF(AND({columns['p']}{row}=1,{columns['f']}{row}=1),1,0)"
        # VN = 1 when p=0 and f=1, otherwise 0
        sheet[f"{vn_col}{row}"] = f"=IF(AND({columns['p']}{row}=0,{columns['f']}{row}=1),1,0)"
        # NP = 1 when p=1 and f=0, otherwise 0
        sheet[f"{np_col}{row}"] = f"=IF(AND({columns['p']}{row}=1,{columns['f']}{row}=0),1,0)"

    # Add the sum row for new columns
    sum_row = last_row + 1
    for col in ["q","r",pp_col, vn_col, np_col]:
        formula = f"=SUM({col}2:{col}{last_row})"
        sheet[f"{col}{sum_row}"] = formula

    # Add the count row for new columns
    count_row = last_row + 2
    for col in [pp_col, vn_col, np_col]:
        formula = f'=COUNTIF({col}2:{col}{last_row}, ">0")'
        sheet[f"{col}{count_row}"] = formula

    # Add percentage calculations for new columns
    perc_column = 'V'  # Choose a column for percentage calculations

    # Ensure "percent_style" is added only once
    if "percent_style" not in workbook.named_styles:
        percent_style = NamedStyle(name="percent_style", number_format="0.00%")
        workbook.add_named_style(percent_style)


    sheet[f"{perc_column}{sum_row}"] = f"=SUM({vn_col}{sum_row},{np_col}{sum_row})/Q{sum_row}"
    sheet[f"{perc_column}{sum_row}"].style = "percent_style"
    # for i, col in enumerate([pp_col, vn_col, np_col], start=1):
    #     sheet[f"{perc_column}{sum_row}"] = f"=SUM({vn_col}{sum_row}:{vn_col}{sum_row + i})/{pp_col}{sum_row}"
    #
    #     # sheet[f"{perc_column}{sum_row + i}"] = f"={col}{sum_row}/Q{sum_row}"  # Percentage of each column with respect to p (column Q)
    #     # sheet[f"{perc_column}{count_row + i}"] = f"={col}{count_row}/Q{count_row}"
    #
    #     # Apply the percent style
    #     sheet[f"{perc_column}{sum_row + i}"].style = "percent_style"
    #     sheet[f"{perc_column}{count_row + i}"].style = "percent_style"

    # Force recalculation in Excel
    workbook.properties.calcMode = "auto"

    # Save the workbook
    workbook.save(output_file)




def generate_report_with_charts(folder_path, final_df):
    # Функции для агрегации и обработки данных
    aggregated_df = aggregate_and_calculate(final_df)
    monthly_avg_df = calculate_monthly_averages(aggregated_df)
    monthly_percent_df = calculate_monthly_average_percentages(final_df)

    monthly_avg_df = process_monthly_avg(monthly_avg_df)
    monthly_percent_df = process_monthly_avg(monthly_percent_df)

    # Генерация имени выходного файла
    t = datetime.now().microsecond
    output_file_name = os.path.basename("_report " + folder_path.strip('/\\')) + str(t) + '.xlsx'
    output_file = os.path.join(folder_path, output_file_name)

    # Создаем Excel с данными и графиками
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Записываем таблицы в Excel
        final_df.to_excel(writer, sheet_name="Данные", index=False)
        aggregated_df.to_excel(writer, sheet_name="Сумм. по панелям", index=False)

        # Вместо записи отдельных листов для "Групп. по панелям" и "Средн. по штрекам",
        # вызываем функцию для создания одного листа с графиками
        create_excel_with_charts_on_one_sheet(monthly_avg_df, monthly_percent_df, writer)

    # Freeze cell C2
    workbook = load_workbook(output_file)
    sheet_name = "Данные"  # Specify the sheet where you want to freeze C2
    sheet = workbook[sheet_name]
    sheet.freeze_panes = "C2"  # Set the freeze panes to C2
    workbook.save(output_file)
    # Add summary and formula rows to the first sheet
    add_summary_and_formula_rows(output_file)

    return final_df


from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

from openpyxl.utils import get_column_letter

def create_excel_with_charts_on_one_sheet1(df1, df2, writer):
    """
    Добавляет графики и данные для двух DataFrame на один лист Excel,
    размещая графики справа от таблиц и обеспечивая отступы между элементами.

    Args:
        df1 (pd.DataFrame): Первый DataFrame для первого графика.
        df2 (pd.DataFrame): Второй DataFrame для второго графика.
        writer (pd.ExcelWriter): Экземпляр ExcelWriter для записи в файл.
    """
    # Форматируем числа в DataFrame до 2 знаков после запятой
    df1 = df1.round(2)
    df2 = df2.round(2)

    # Загружаем рабочую книгу и создаем лист для графиков
    wb = writer.book
    ws = wb.create_sheet("Графики")

    # Добавляем данные для первого DataFrame
    headers_df1 = list(df1.columns)
    ws.append(headers_df1)  # Добавляем заголовки для df1
    data_start_row_df1 = ws.max_row + 1  # Начало данных для df1
    for row in df1.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df1 = ws.max_row  # Конец данных для df1
    data_max_col_df1 = len(df1.columns)  # Количество колонок в df1

    # Создаём первый график
    chart1 = BarChart()
    chart1.title = "Средние отклонения сгруппированные по панелям в разрезе месяцев"
    chart1.y_axis.title = "Значения"
    chart1.x_axis.title = "Месяцы"
    chart1.height = 15  # Регулировка высоты для лучшего отображения
    chart1.width = 25   # Регулировка ширины для лучшего отображения

    # Включаем метки данных
    chart1.dLbls = DataLabelList()
    chart1.dLbls.showVal = True

    # Источник данных для первого графика
    data_range1 = Reference(ws, min_col=2, min_row=data_start_row_df1 ,
                            max_col=data_max_col_df1, max_row=data_end_row_df1)
    categories1 = Reference(ws, min_col=1, min_row=data_start_row_df1 ,
                            max_row=data_end_row_df1)
    chart1.add_data(data_range1, titles_from_data=True)
    chart1.set_categories(categories1)

    # Позиционируем первый график справа от таблицы данных
    # Вычисляем начальный столбец для графика (добавляем 2 столбца для отступа)
    chart_start_col1 = data_max_col_df1 + 3
    chart_start_cell1 = f"{get_column_letter(chart_start_col1)}{data_start_row_df1 - 1}"
    ws.add_chart(chart1, chart_start_cell1)

    # Добавляем пустые строки для отступа между таблицами
    ws.append([])
    ws.append([])

    # Добавляем данные для второго DataFrame
    headers_df2 = list(df2.columns)
    ws.append(headers_df2)  # Добавляем заголовки для df2
    data_start_row_df2 = ws.max_row + 1  # Начало данных для df2
    for row in df2.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df2 = ws.max_row  # Конец данных для df2
    data_max_col_df2 = len(df2.columns)  # Количество колонок в df2

    # Создаём второй график
    chart2 = BarChart()
    chart2.title = "Средние отклонения по всем штрекам в разрезе месяцев"
    chart2.y_axis.title = "Значения"
    chart2.x_axis.title = "Месяцы"
    chart2.height = 15  # Регулировка высоты для лучшего отображения
    chart2.width = 25   # Регулировка ширины для лучшего отображения

    # Включаем метки данных
    chart2.dLbls = DataLabelList()
    chart2.dLbls.showVal = True

    # Источник данных для второго графика
    data_range2 = Reference(ws, min_col=2, min_row=data_start_row_df2 - 1,
                            max_col=data_max_col_df2, max_row=data_end_row_df2)
    categories2 = Reference(ws, min_col=1, min_row=data_start_row_df2 - 1,
                            max_row=data_end_row_df2)
    chart2.add_data(data_range2, titles_from_data=True)
    chart2.set_categories(categories2)

    # Позиционируем второй график справа от второй таблицы данных
    # Вычисляем начальный столбец для графика (добавляем 2 столбца для отступа)
    chart_start_col2 = data_max_col_df2 + 3
    chart_start_cell2 = f"{get_column_letter(chart_start_col2)}{data_start_row_df2 - 1}"
    ws.add_chart(chart2, chart_start_cell2)

def create_excel_with_charts_on_one_sheet2(df1, df2, writer):
    """
    Добавляет графики и данные для двух DataFrame на один лист Excel,
    размещая графики справа от таблиц и обеспечивая отступы между элементами.

    Args:
        df1 (pd.DataFrame): Первый DataFrame для первого графика.
        df2 (pd.DataFrame): Второй DataFrame для второго графика.
        writer (pd.ExcelWriter): Экземпляр ExcelWriter для записи в файл.
    """
    # Форматируем числа в DataFrame до 2 знаков после запятой
    df1 = df1.round(2)
    df2 = df2.round(2)

    # Загружаем рабочую книгу и создаем лист для графиков
    wb = writer.book
    ws = wb.create_sheet("Графики")

    # Добавляем данные для первого DataFrame
    headers_df1 = list(df1.columns)
    ws.append(headers_df1)  # Добавляем заголовки для df1
    data_start_row_df1 = ws.max_row + 1  # Начало данных для df1
    for row in df1.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df1 = ws.max_row  # Конец данных для df1
    data_max_col_df1 = len(df1.columns)  # Количество колонок в df1

    # Создаём первый график
    chart1 = BarChart()
    chart1.title = "Средние отклонения сгруппированные по панелям в разрезе месяцев"
    chart1.y_axis.title = "Значения"
    chart1.x_axis.title = "Месяцы"
    chart1.height = 15  # Регулировка высоты для лучшего отображения
    chart1.width = 25   # Регулировка ширины для лучшего отображения

    # Включаем метки данных
    chart1.dLbls = DataLabelList()
    chart1.dLbls.showVal = True

    # Источник данных для первого графика
    data_range1 = Reference(ws, min_col=2, min_row=data_start_row_df1,
                            max_col=data_max_col_df1, max_row=data_end_row_df1)
    categories1 = Reference(ws, min_col=1, min_row=data_start_row_df1,
                            max_row=data_end_row_df1)
    chart1.add_data(data_range1, titles_from_data=True)
    chart1.set_categories(categories1)

    # Позиционируем первый график справа от таблицы данных
    chart_start_col1 = data_max_col_df1 + 3
    chart_start_row1 = data_start_row_df1 - 1  # Начальная строка для графика
    chart_start_cell1 = f"{get_column_letter(chart_start_col1)}{chart_start_row1}"
    ws.add_chart(chart1, chart_start_cell1)

    # Добавляем пустые строки для отступа между таблицами
    ws.append([])
    ws.append([])

    # Добавляем данные для второго DataFrame
    headers_df2 = list(df2.columns)
    ws.append(headers_df2)  # Добавляем заголовки для df2
    data_start_row_df2 = ws.max_row + 1  # Начало данных для df2
    for row in df2.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df2 = ws.max_row  # Конец данных для df2
    data_max_col_df2 = len(df2.columns)  # Количество колонок в df2

    # Создаём второй график
    chart2 = BarChart()
    chart2.title = "Средние отклонения по всем штрекам в разрезе месяцев"
    chart2.y_axis.title = "Значения"
    chart2.x_axis.title = "Месяцы"
    chart2.height = 15  # Регулировка высоты для лучшего отображения
    chart2.width = 25   # Регулировка ширины для лучшего отображения

    # Включаем метки данных
    chart2.dLbls = DataLabelList()
    chart2.dLbls.showVal = True

    # Источник данных для второго графика
    data_range2 = Reference(ws, min_col=2, min_row=data_start_row_df2,
                            max_col=data_max_col_df2, max_row=data_end_row_df2)
    categories2 = Reference(ws, min_col=1, min_row=data_start_row_df2,
                            max_row=data_end_row_df2)
    chart2.add_data(data_range2, titles_from_data=True)
    chart2.set_categories(categories2)

    # Позиционируем второй график справа от второй таблицы данных
    chart_start_col2 = data_max_col_df2 + 3

    # Вычисляем приблизительное количество строк, занимаемых первым графиком
    chart_height_in_rows = int(chart1.height * 4)  # Примерное преобразование высоты графика в строки

    # Определяем начальную строку для второго графика
    chart_start_row2 = max(ws.max_row, chart_start_row1 + chart_height_in_rows + 5)

    chart_start_cell2 = f"{get_column_letter(chart_start_col2)}{chart_start_row2}"
    ws.add_chart(chart2, chart_start_cell2)

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
import pandas as pd

def create_excel_with_charts_on_one_sheet(df1, df2, writer):
    """
    Adds charts and data for two DataFrames on one Excel sheet,
    placing charts to the right of the tables and ensuring offsets between elements.

    Args:
        df1 (pd.DataFrame): First DataFrame for the first chart.
        df2 (pd.DataFrame): Second DataFrame for the second chart.
        writer (pd.ExcelWriter): ExcelWriter instance for writing to the file.
    """
    # Format numbers in DataFrame to 2 decimal places
    df1 = df1.round(2)
    df2 = df2.round(2)

    # Load the workbook and create a sheet for charts
    wb = writer.book
    ws = wb.create_sheet("Графики")

    ### Adding Data for the First DataFrame ###

    # Append headers for df1
    headers_df1 = list(df1.columns)
    ws.append(headers_df1)
    headers_row_df1 = ws.max_row  # Row number where headers are inserted

    # Append data for df1
    data_start_row_df1 = ws.max_row + 1  # Start row for data (after headers)
    for row in df1.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df1 = ws.max_row  # End row for data
    data_max_col_df1 = len(df1.columns)  # Number of columns in df1

    # Create the first chart
    chart1 = BarChart()
    chart1.title = "Средние отклонения сгруппированные по панелям в разрезе месяцев"
    chart1.y_axis.title = "Значения"
    chart1.x_axis.title = "Месяцы"
    chart1.height = 15  # Adjust height for better display
    chart1.width = 25   # Adjust width for better display

    # Enable data labels
    chart1.dLbls = DataLabelList()
    chart1.dLbls.showVal = True

    # Define data range for the first chart (include headers)
    data_range1 = Reference(
        ws,
        min_col=2,
        min_row=headers_row_df1,  # Include headers for titles
        max_col=data_max_col_df1,
        max_row=data_end_row_df1
    )
    categories1 = Reference(
        ws,
        min_col=1,
        min_row=data_start_row_df1,
        max_row=data_end_row_df1
    )
    chart1.add_data(data_range1, titles_from_data=True)
    chart1.set_categories(categories1)

    # Position the first chart to the right of the data table
    chart_start_col1 = data_max_col_df1 + 3
    chart_start_row1 = headers_row_df1  # Start row for the chart
    chart_start_cell1 = f"{get_column_letter(chart_start_col1)}{chart_start_row1}"
    ws.add_chart(chart1, chart_start_cell1)

    # Add empty rows for spacing between tables
    ws.append([])
    ws.append([])

    ### Adding Data for the Second DataFrame ###

    # Append headers for df2
    headers_df2 = list(df2.columns)
    ws.append(headers_df2)
    headers_row_df2 = ws.max_row  # Row number where headers are inserted

    # Append data for df2
    data_start_row_df2 = ws.max_row + 1  # Start row for data (after headers)
    for row in df2.itertuples(index=False, name=None):
        ws.append(row)
    data_end_row_df2 = ws.max_row  # End row for data
    data_max_col_df2 = len(df2.columns)  # Number of columns in df2

    # Create the second chart
    chart2 = BarChart()
    chart2.title = "Средние отклонения по всем штрекам в разрезе месяцев"
    chart2.y_axis.title = "Значения"
    chart2.x_axis.title = "Месяцы"
    chart2.height = 15  # Adjust height for better display
    chart2.width = 25   # Adjust width for better display

    # Enable data labels
    chart2.dLbls = DataLabelList()
    chart2.dLbls.showVal = True

    # Define data range for the second chart (include headers)
    data_range2 = Reference(
        ws,
        min_col=2,
        min_row=headers_row_df2,  # Include headers for titles
        max_col=data_max_col_df2,
        max_row=data_end_row_df2
    )
    categories2 = Reference(
        ws,
        min_col=1,
        min_row=data_start_row_df2,
        max_row=data_end_row_df2
    )
    chart2.add_data(data_range2, titles_from_data=True)
    chart2.set_categories(categories2)

    # Position the second chart to the right of the second data table
    chart_start_col2 = data_max_col_df2 + 3

    # Calculate approximate number of rows occupied by the first chart
    chart_height_in_rows = int(chart1.height * 4)  # Approximate conversion

    # Determine the start row for the second chart
    chart_start_row2 = headers_row_df2+20  # Align with the second table
    chart_start_cell2 = f"{get_column_letter(chart_start_col2)}{chart_start_row2}"
    ws.add_chart(chart2, chart_start_cell2)




def add_calculated_columns(df):
    """
    Добавляет расчетные столбцы в DataFrame:
    - dRuda, dCu, dAg
    - %Ruda, %Cu, %Ag

    :param df: pandas DataFrame, содержащий исходные столбцы
    :return: pandas DataFrame с добавленными расчетными столбцами
    """
    df['Ruda'] = df['Ruda'].astype(str)  # Приводим к строковому типу
    df = df[~df['Ruda'].str.contains('Руда', na=False)]

    df['Ruda'] = df['Ruda'].fillna(0)
    df['Cu'] = df['Cu'].fillna(0)
    df['Ag'] = df['Ag'].fillna(0)
    df['fRuda'] = df['fRuda'].fillna(0)
    df['fCu'] = df['fCu'].fillna(0)
    df['fAg'] = df['fAg'].fillna(0)

    # Заменяем некорректные строки (например, 'None') на NaN
    df.replace('None', np.nan, inplace=True)


    # Преобразуем все указанные колонки в float
    columns_to_convert = ['Ruda', 'Cu', 'Ag', 'fRuda', 'fCu', 'fAg']

    # for column in columns_to_convert:
    #     # Convert to numeric, keeping NaN and replacing invalid values with NaN
    #     df[column] = df[column].apply(lambda x: pd.to_numeric(x, errors='coerce'))
    #
    #     # Replace NaN with 0 only for rows that are not explicitly missing (e.g., None)
    #     df[column] = df[column].where(df[column].isna() | df[column].notna(), 0)
    # for column in columns_to_convert:
    #     # Replace None, empty strings, and NaN with 0
    #     df[column] = df[column].replace([None, '', np.nan], 0)
    #
    #     # Clean up the column to handle commas and trailing non-numeric characters
    #     df[column] = df[column].apply(lambda x: str(x).replace(',', '').strip())
    #
    #     # Convert valid numbers to float, leave text untouched
    #     def convert_value(value):
    #         try:
    #             return float(value) if value != '' else 0  # Convert to float if possible
    #         except ValueError:
    #             return value  # Leave non-numeric text as-is
    #
    #     df[column] = df[column].apply(convert_value)

    for column in columns_to_convert:
        # Replace None and empty strings with NaN
        df[column] = df[column].replace([None, ''], np.nan)

        # Remove commas and strip whitespace
        df[column] = df[column].astype(str).str.replace(',', '').str.strip()

        # Convert to numeric, coercing errors to NaN
        df[column] = pd.to_numeric(df[column], errors='coerce')

        # Optional: Replace NaN with 0
        df[column] = df[column].fillna(0)

    # df[columns_to_convert] = df[columns_to_convert].astype(float)


    # Расчет dRuda, dCu, dAg (с обработкой ошибок, замена NaN на 0)
    # df['diff-Ruda'] = abs(df['Ruda'] - df['fRuda'])
    # df['diff-Cu'] = abs(df['Cu'] - df['fCu'])
    # df['diff-Ag'] = abs(df['Ag'] - df['fAg'])
    df['diff-Ruda'] = df['Ruda'] - df['fRuda']
    df['diff-Cu'] = df['Cu'] - df['fCu']
    df['diff-Ag'] = df['Ag'] - df['fAg']

    # Замена NaN на 0, если расчет невозможен
    df['diff-Ruda'] = df['diff-Ruda'].fillna(0)
    df['diff-Cu'] = df['diff-Cu'].fillna(0)
    df['diff-Ag'] = df['diff-Ag'].fillna(0)

    # Расчет %Ruda, %Cu, %Ag (с обработкой деления на ноль или NaN)
    # df['%rel-Ruda'] = abs(np.where(df['Ruda'] != 0, df['diff-Ruda'] / df['Ruda'], 1.0) * 100)  # 100% если Ruda == 0
    # df['%rel-Cu'] = abs(np.where(df['Cu'] != 0, df['diff-Cu'] / df['Cu'], 1.0) * 100)  # 100% если Cu == 0
    # df['%rel-Ag'] = abs(np.where(df['Ag'] != 0, df['diff-Ag'] / df['Ag'], 1.0) * 100)  # 100% если Ag == 0

    # Расчет %Ruda, %Cu, %Ag (с обработкой деления на ноль или NaN и ограничением максимум 100)
    df['(Ruda-fRuda)/Ruda'] = np.clip(np.where(df['Ruda'] != 0, 1-abs(df['diff-Ruda']) / df['Ruda'], 1.0) * 100, 0, 100)
    df['(Cu-fCu)/Cu'] = np.clip(np.where(df['Cu'] != 0, 1-abs(df['diff-Cu']) / df['Cu'], 1.0) * 100, 0, 100)
    df['(Ag-fAg)/Ag'] = np.clip(np.where(df['Ag'] != 0, 1-abs(df['diff-Ag']) / df['Ag'], 1.0) * 100, 0, 100)

    # Замена NaN на 100%, если расчет невозможен
    df['(Ruda-fRuda)/Ruda'] = df['(Ruda-fRuda)/Ruda'].fillna(100)
    df['(Cu-fCu)/Cu'] = df['(Cu-fCu)/Cu'].fillna(100)
    df['(Ag-fAg)/Ag'] = df['(Ag-fAg)/Ag'].fillna(100)

    # Добавляем дополнительные колонки p>0, f>0, p=0, f=0
    df['p>0'] = np.where(df['Ruda'] > 0, 1, 0)
    df['f>0'] = np.where(df['fRuda'] > 0, 1, 0)
    # df['p=0'] = np.where(df['Ruda'] == 0, 1, 0)
    # df['f=0'] = np.where(df['fRuda'] == 0, 1, 0)

    return df

# def add_calculated_columns(df):
#     """
#     Добавляет расчетные столбцы в DataFrame:
#     - dRuda, dCu, dAg
#     - %Ruda, %Cu, %Ag
#
#     :param df: pandas DataFrame, содержащий исходные столбцы
#     :return: pandas DataFrame с добавленными расчетными столбцами
#     """
#     # Преобразование колонок в числовой формат
#     numeric_columns = ['Ruda', 'fRuda', 'Cu', 'fCu', 'Ag', 'fAg']
#     for col in numeric_columns:
#         df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)  # Преобразование и замена NaN на 0
#
#     # Расчет dRuda, dCu, dAg
#     df['Ruda-fRuda'] = abs(df['Ruda'] - df['fRuda'])
#     df['Cu-fCu'] = abs(df['Cu'] - df['fCu'])
#     df['Ag-fAg'] = abs(df['Ag'] - df['fAg'])
#
#     # Расчет %Ruda, %Cu, %Ag
#     df['(Ruda-fRuda)/Ruda'] = abs(np.where(df['Ruda'] != 0, df['Ruda-fRuda'] / df['Ruda'], 1.0) * 100)
#     df['(Cu-fCu)/Cu'] = abs(np.where(df['Cu'] != 0, df['Cu-fCu'] / df['Cu'], 1.0) * 100)
#     df['(Ag-fAg)/Ag'] = abs(np.where(df['Ag'] != 0, df['Ag-fAg'] / df['Ag'], 1.0) * 100)
#
#     return df


def process_monthly_avg(df):
    """
    Преобразует DataFrame monthly_avg:
    - Переименовывает колонки.
    - Конвертирует числовые колонки в числовой формат.
    - Сортирует строки по порядку месяцев.

    Args:
        df (pd.DataFrame): Исходный DataFrame.

    Returns:
        pd.DataFrame: Обработанный DataFrame.
    """
    # Заменяем имена колонок
    df.columns = ['Месяц', 'Товарная руда (СМТ)', 'Cu в руде', 'Ag в руде']

    # Преобразуем последние три колонки в числовой формат
    numeric_columns = ['Товарная руда (СМТ)', 'Cu в руде', 'Ag в руде']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    # Создаем порядок месяцев
    month_order = [
        "январь", "февраль", "март", "апрель", "май", "июнь",
        "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
    ]

    # Сортируем строки по порядку месяцев
    df['Месяц'] = pd.Categorical(df['Месяц'], categories=month_order, ordered=True)
    df = df.sort_values(by='Месяц')

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
    grouped = df.groupby(['month', 'Horizont','Panel'], as_index=False)[sum_columns].sum()

    # Добавляем расчетные колонки
    grouped['Ruda-fRuda'] = grouped['Ruda'] - grouped['fRuda']
    grouped['Cu-fCu'] = grouped['Cu'] - grouped['fCu']
    grouped['Ag-fAg'] = grouped['Ag'] - grouped['fAg']

    # Замена ошибок на 0
    grouped['Ruda-fRuda'] = grouped['Ruda-fRuda'].fillna(0)
    grouped['Cu-fCu'] = grouped['Cu-fCu'].fillna(0)
    grouped['Ag-fAg'] = grouped['Ag-fAg'].fillna(0)

    # Добавляем расчетные колонки в процентах с ограничением максимум 100
    grouped['(Ruda-fRuda)/Ruda'] = np.clip(
        np.where(grouped['Ruda'] != 0, 1-abs(grouped['Ruda-fRuda'] / grouped['Ruda']), 1.0) * 100, 0, 100
    )
    grouped['(Cu-fCu)/Cu'] = np.clip(
        np.where(grouped['Cu'] != 0, 1-abs(grouped['Cu-fCu'] / grouped['Cu']), 1.0) * 100, 0, 100
    )
    grouped['(Ag-fAg)/Ag'] = np.clip(
        np.where(grouped['Ag'] != 0, 1-abs(grouped['Ag-fAg'] / grouped['Ag']), 1.0) * 100, 0, 100
    )

    return grouped





def calculate_monthly_averages(df):
    """
    Группирует DataFrame по месяцам, вычисляет абсолютное среднее значение для %tRuda, %tCu, %tAg
    и записывает вычитание из 1 в виде процентов в колонки Ruda, Cu, Ag.
    Исключает февраль и сентябрь из расчета.

    :param df: DataFrame, содержащий колонки month, %tRuda, %tCu, %tAg.
    :return: Новый DataFrame с агрегированными данными.
    """

    # # Средние значения по месяцам с расчётом абсолютных значений
    # monthly_avg = df.groupby('month', as_index=False)[['(Ruda-fRuda)/Ruda', '(Cu-fCu)/Cu', '(Ag-fAg)/Ag']].apply( lambda x: x.abs().mean() )
    monthly_avg = df.groupby('month', as_index=False)[['(Ruda-fRuda)/Ruda', '(Cu-fCu)/Cu', '(Ag-fAg)/Ag']].apply( lambda x: x.mean())

    # Вычитание из 1, перевод в проценты и ограничение максимум 100
    # monthly_avg['aver-relation_Ruda'] = (1 - monthly_avg['(Ruda-fRuda)/Ruda'])*100
    # monthly_avg['aver-relation_Cu'] = (1 - monthly_avg['(Cu-fCu)/Cu'] )*100
    # monthly_avg['aver-relation_Ag'] = (1 - monthly_avg['(Ag-fAg)/Ag'] )*100

    monthly_avg['aver-relation_Ruda'] = monthly_avg['(Ruda-fRuda)/Ruda']
    monthly_avg['aver-relation_Cu'] = monthly_avg['(Cu-fCu)/Cu']
    monthly_avg['aver-relation_Ag'] = monthly_avg['(Ag-fAg)/Ag']

    # Оставляем только нужные колонки
    monthly_avg = monthly_avg[['month', 'aver-relation_Ruda', 'aver-relation_Cu', 'aver-relation_Ag']]
    return monthly_avg



def calculate_monthly_average_percentages(df):
    """
    Группирует исходный DataFrame по месяцам, вычисляет абсолютное среднее значение для %Ruda, %Cu, %Ag,
    а затем вычитает это значение из 1, переводит в проценты и ограничивает максимум 100.

    :param df: Исходный DataFrame с колонками %Ruda, %Cu, %Ag.
    :return: Новый DataFrame с агрегированными данными.
    """
    # # Группировка по месяцам и вычисление абсолютных средних значений
    # monthly_avg = df.groupby('month', as_index=False)[['(diff-Ruda)/Ruda', '(diff-Cu)/Cu', '(diff-Ag)/Ag']].apply(
    #     lambda x: np.clip(x.abs().mean(), 0, 100)
    # )

    # Группировка по месяцам и вычисление абсолютных средних значений
    # monthly_avg = df.groupby('month', as_index=False)[['(Ruda-fRuda)/Ruda', '(Cu-fCu)/Cu', '(Ag-fAg)/Ag']].apply( lambda x: x.abs().mean() )
    monthly_avg = df.groupby('month', as_index=False)[['(Ruda-fRuda)/Ruda', '(Cu-fCu)/Cu', '(Ag-fAg)/Ag']].apply( lambda x: x.mean())


    # Вычитание средних значений из 1, взятие абсолютного значения, перевод в проценты и ограничение результата максимум 100
    # monthly_avg['otlonenie-Ruda'] = (1 - monthly_avg['(Ruda-fRuda)/Ruda'] )*100
    # monthly_avg['otlonenie-Cu'] = (1 - monthly_avg['(Cu-fCu)/Cu']  )*100
    # monthly_avg['otlonenie-Ag'] = (1 - monthly_avg['(Ag-fAg)/Ag'] )*100

    monthly_avg['otlonenie-Ruda'] = monthly_avg['(Ruda-fRuda)/Ruda']
    monthly_avg['otlonenie-Cu'] = monthly_avg['(Cu-fCu)/Cu']
    monthly_avg['otlonenie-Ag'] = monthly_avg['(Ag-fAg)/Ag']

    # Оставляем только нужные колонки
    monthly_avg = monthly_avg[['month', 'otlonenie-Ruda', 'otlonenie-Cu', 'otlonenie-Ag']]
    return monthly_avg




# Пример использования:

# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__ЮЖР ИПГ 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__Шатыркуль ИПГ 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__ИПГ Саяк 3 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__Жомарт ИПГ 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__ИПГ Жайсан 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__ВЖР ИПГ 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__Жиланды ИПГ 2024'
# directory =r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__ЗР ИПГ 2024'
# directory = r'C:\Users\delxps\Documents\Kazakhmys\_alibek\__Конырат ИПГ 2024'
# directory = r"C:\Users\delxps\Documents\Kazakhmys\_alibek\__Акбастау ИПГ 2024"
# directory = r"C:\Users\delxps\Documents\Kazakhmys\_alibek\__ИПГ 2024 С-1"
directory = r"C:\Users\delxps\Documents\Kazakhmys\_alibek\__Нурказган ИПГ 2024"   # ----N
# directory = r"C:\Users\delxps\Documents\Kazakhmys\_alibek\Хаджиконган ИПГ 2024"   # ----N
# directory = r"C:\Users\delxps\Documents\Kazakhmys\_alibek\__Абыз ИПГ 2024"
final_data = load_excel_data_with_flex(directory)

