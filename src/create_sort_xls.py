import json
import os

import openpyxl

from config import USERS


# Загружаем данные из JSON файла
filepath = os.path.join("data", "file_finish.json")
with open(filepath, "r", encoding="UTF-8") as file:
    users_data = json.load(file)

# Создаем новый файл Excel
workbook = openpyxl.Workbook()
sheet = workbook.create_sheet(title="ITOG")
# Удаляем лист по умолчанию
workbook.remove(workbook['Sheet'])

# Заполняем заголовки столбцов
sheet['A1'] = 'DATE'
sheet['B1'] = 'ALL'
sheet['C1'] = 'IT'
row_index = 2

# Итерируемся по данным и создаем листы для каждой даты
for date, data in users_data.items():

    sheet[f'A{row_index}'] = date
    sheet[f'B{row_index}'] = data['total_tasks']
    # Заполняем лист данными
    
    tasks_count = sum([tasks_count for employee, tasks_count in data['employees'].items() if employee in USERS])
    sheet[f'C{row_index}'] = tasks_count   
    row_index += 1


# Сохраняем файл Excel
workbook.save("output_excel_file_1.xlsx")
