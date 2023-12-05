import json, requests
import openpyxl

# запрос к api и получение файла (указать url,id)
# response = requests.get('..../portfolio/{portfolio_id}/holdings/{stated_at}')
# portfolio = response.json()
#
# сохранение данных в JSON
# with open('portfolio.json', 'w') as file:
#     json.dump(portfolio, file, indent=3)

"""Создаем новую рабочую книгу xlsx использую библиотеку openpyxl"""
book = openpyxl.Workbook()
sheet = book.active

# Открываем файл JSON с данными
with open('portfolio.json') as file:
    json_data = json.load(file)
    # Получаем список заголовков столбцов из словаря
    column = list(json_data['portfolio']['columns'])
# Записываем заголовки в ячейки A1 и C1
sheet['A1'] = column[0].upper()
sheet['C1'] = column[1].upper()

# Получаем список ценных бумаг и итерируемся по каждому элементу
securities_list = json_data['portfolio']['columns']['securities']
for i, security in enumerate(securities_list):
    sheet[f'A{i + 2}'] = security

# Получаем список значений ценных бумаг и итерируемся по каждому элементу
securities_values = json_data['portfolio']['securities']
for i, security_values in enumerate(securities_values):
    sheet[f'B{i + 2}'] = str(security_values)

# Получаем список инструментов и итерируемся по каждому элементу
instruments = json_data['portfolio']['columns']['instruments']
for i, instrument in enumerate(instruments):
    sheet[f'C{i + 2}'] = instrument

# Получаем список значений инструментов и итерируемся по каждому элементу
instruments_values = json_data['portfolio']['instruments']
for i, instrument_values in enumerate(instruments_values):
    sheet[f'D{i + 2}'] = str(instrument_values)

# сохраняем и закрываем файл в формате xlsx
book.save('portfolio.xlsx')
book.close()