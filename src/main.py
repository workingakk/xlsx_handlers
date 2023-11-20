import os
import openpyxl
from datetime import datetime


filepath = os.path.join("data", "all_tasks.xlsx")
# print(filepath)
# print(os.path.isfile(filepath))

def get_analysis_xlsx(filepath):

    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active

    tasks_data = {}

    rows = list(sheet.iter_rows(min_row=2, max_col=sheet.max_column, max_row=sheet.max_row, values_only=True))
    sorted_rows = sorted(rows, key=lambda x: datetime.strptime(x[8], '%d.%m.%Y %H:%M'))

    # Обработка отсортированных данных
    for row in sorted_rows:
        _, _, _, executor, _, _, _, _, close_date = row
        close_date = datetime.strptime(close_date, '%d.%m.%Y %H:%M')

        # Подсчет общего количества задач
        tasks_data['total_tasks'] = tasks_data.get('total_tasks', 0) + 1

        # Подсчет общего количества задач для каждого сотрудника
        tasks_data[executor] = tasks_data.get(executor, 0) + 1
        break

    # Вывод результатов
    print("Общее количество задач:", tasks_data['total_tasks'])
    print("Общее количество задач по сотрудникам:")
    for employee, tasks_count in tasks_data.items():
        if employee != 'total_tasks':
            print(f"{employee}: {tasks_count}")
        break

    # Закрытие файла Excel
    workbook.close()

    


get_analysis_xlsx(filepath)
