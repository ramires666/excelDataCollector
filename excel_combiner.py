import pandas as pd
import os

# Путь к папке с файлами
folder_path = r"C:\Users\delxps\Documents\Kazakhmys\nurda"

# Создаем пустой DataFrame для объединения данных
all_data = pd.DataFrame()

# Обходим все файлы в папке
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx"):  # Проверяем, что файл Excel
        file_path = os.path.join(folder_path, file_name)
        # Читаем первый лист, начиная с 5-й строки и колонок B по O
        # data = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=4, usecols="B:O")
        data = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=0, usecols="A:O")
        # Присваиваем заголовки из первой строки данных
        data.columns = data.iloc[0]
        data = data[1:]  # Убираем строку с заголовком
        # Добавляем колонку с именем файла
        data["Имя файла"] = file_name
        # Объединяем данные
        all_data = pd.concat([all_data, data], ignore_index=True)

# Сохраняем объединенный DataFrame, если необходимо
output_path = os.path.join(folder_path, "объединенные_данные.xlsx")
all_data.to_excel(output_path, index=False)

print("Данные успешно объединены и сохранены в файл:", output_path)