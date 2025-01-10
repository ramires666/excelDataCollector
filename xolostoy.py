import re
from datetime import timedelta
import pandas as pd


name = "gruz"


def parse_time_str(time_str):
    """
    Функция для парсинга строки формата:
    "7 дней 11:03:21", "1 день 6:19:13", "10 дней 5:14:38", "3 дня 5:16:27"
    или просто "11:03:21" (когда нет упоминания о днях),
    при этом слово "день" может склоняться по-разному: день, дня, дней.
    """

    # Пример ожидаемого формата:
    # Опционально: число + слово "день/дня/дней" + пробел + "HH:MM:SS"
    # Или просто "HH:MM:SS"
    # Используем регулярное выражение для вычленения компонентов
    # Регулярка: ^(?:(\d+)\s+дн(?:ей|я|ь))?\s*(\d{1,2}:\d{2}:\d{2})$
    # 1-я группа (опциональная): количество дней + различные варианты слова "день"
    # 2-я группа: время в формате HH:MM:SS
    pattern = re.compile(r'^(?:(\d+)\s+дн(?:ей|я|ь))?\s*(\d{1,2}:\d{2}:\d{2})$')

    match = pattern.match(time_str.strip())
    if not match:
        # Если формат не подошел, может быть только время без дней,
        # проверим отдельно
        # Возможно строка - это пустая или просто пробелы
        time_str = time_str.strip()
        if time_str == "":
            # Пустая строка интерпретируем как 0 времени
            return timedelta()
        # Проверим формат только времени HH:MM:SS без дней
        time_pattern = re.compile(r'^(\d{1,2}:\d{2}:\d{2})$')
        tm_match = time_pattern.match(time_str)
        if not tm_match:
            # Если формат вообще не подходит, вернем 0
            return timedelta()
        else:
            # Есть только время без дней
            h, m, s = [int(x) for x in time_str.split(':')]
            return timedelta(hours=h, minutes=m, seconds=s)

    days_str, time_str_only = match.groups()
    days = int(days_str) if days_str else 0
    h, m, s = [int(x) for x in time_str_only.split(':')]
    return timedelta(days=days, hours=h, minutes=m, seconds=s)


# Пример использования:
# Допустим у нас есть CSV файл "data.csv" со столбцами:
# "В движении" и "Холостой ход"
# В них содержатся данные в формате, описанном выше.
# Мы хотим:
# 1. Прочитать файл
# 2. Преобразовать каждый столбец в timedelta
# 3. Посчитать разницу во времени и отношение холостого хода к времени в движении.
#
# Предположим, столбцы называются:
# E.g. В движении -> InMotion
#      Холостой ход -> Idle
#
# Допустим в CSV-файле заголовки: "В движении","Холостой ход"
# (Если названия столбцов отличаются, подставьте нужные)

# Ниже пример кода. Тут предполагается, что файл data.csv есть,
# и в нем есть соответствующие столбцы.
# Если данных о днях нет, то строка может быть просто "11:03:21" или пустая.

df = pd.read_csv(f"{name}_motohour.csv", sep=';', encoding='utf-8')

# Применим функцию преобразования к столбцам
df['В движении (timedelta)'] = df['В движении'].apply(parse_time_str)
df['Холостой ход (timedelta)'] = df['Холостой ход'].apply(parse_time_str)

# Теперь мы можем считать разницу и отношение
# Например, разница (В движении - Холостой ход):
df['Разница'] = df['В движении (timedelta)'] - df['Холостой ход (timedelta)']

# Отношение холостого хода к времени в движении (в %),
# если время в движении не ноль:
df['Процент_холостого_к_движению'] = df.apply(
    lambda row: (row['Холостой ход (timedelta)'].total_seconds() / row['В движении (timedelta)'].total_seconds() * 100)
    if row['В движении (timedelta)'].total_seconds() > 0 else None,
    axis=1
)
df.to_excel(rf"C:\Users\delxps\PycharmProjects\excelCollector\motochasy_{name}.xlsx")




print(df)
# Теперь в df есть новые столбцы с вычисленными значениями.
# Можно вывести результат или сохранить обратно в CSV:
# df.to_csv("processed_data.csv", sep=';', encoding='utf-8', index=False)
