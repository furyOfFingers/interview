import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

# Замените 'file.xlsx' на путь к вашему файлу Excel
file_path = 'interview.xlsx'

# Создаем объект ExcelFile для чтения всех листов в файле
excel_file = pd.ExcelFile(file_path)

# Список для хранения данных всех листов
data_for_excel = []

# Проходимся по всем листам в файле
for sheet_name in excel_file.sheet_names:
    # Читаем данные текущего листа
    data_frame = excel_file.parse(sheet_name)

    # Преобразование столбца "grade" в числовой формат
    data_frame['grade'] = pd.to_numeric(data_frame['grade'], errors='coerce')

    # Удаляем строки с отсутствующими значениями в столбце "grade"
    data_frame = data_frame.dropna(subset=['grade'])

    # Replace '\n' with space in the 'topic' column (replace with the actual column name for 'Topic')
    data_frame['topic'] = data_frame['topic'].str.replace('\n', ' ')

    # Группировка по столбцу "topic" и вычисление среднего значения "grade"
    average_grade_by_topic = data_frame.groupby('topic')['grade'].mean()

    # Преобразование результата в словарь
    average_grades_by_topic_dict = average_grade_by_topic.to_dict()

    # Округление значений в словаре average_grades_by_topic_dict до одного знака после запятой
    rounded_average_grades_by_topic_dict = {
        topic: round(grade, 1) for topic, grade in average_grades_by_topic_dict.items()
    }

    # Создание списка данных для текущего листа
    data_for_sheet = []
    for topic, average_grade in rounded_average_grades_by_topic_dict.items():
        data_for_sheet.append([sheet_name, topic, average_grade])

    # Добавляем строку с "sum" в конце последней строки текущего листа, только если есть данные
    if data_for_sheet:
        data_for_sheet.append(
            [f"{sheet_name} sum", "", round(data_frame['grade'].mean() - 0.3, 1)])

    # Добавляем данные текущего листа в общий список
    data_for_excel.extend(data_for_sheet)

# Создание DataFrame из списка данных
data_frame_for_excel = pd.DataFrame(
    data_for_excel, columns=['Sheet', 'Topic', 'Average Grade'])

# Создаем объект Workbook и Sheet
output_file_path = 'output_file.xlsx'
workbook = Workbook()
sheet = workbook.active

# Заполняем данные в Sheet
for row in data_frame_for_excel.itertuples(index=False):
    sheet.append(row)

# Выделяем жирным текст в строках, где значение в столбце 'Sheet' содержит 'sum'
for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
    if 'sum' in row[0].value:
        for cell in row:
            cell.font = Font(bold=True)

# Сохранение Workbook в Excel файл
workbook.save(output_file_path)

print(f"Результаты сохранены в файле: {output_file_path}")
