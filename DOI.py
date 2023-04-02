"""Функция new_doi присваивает новый DOI, возвращает значение DOI(в виде ссылки), а так же записывает DOI
в файл 'doi.xlsx'"""

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font


def new_doi(arg1, arg2):
    isbn = arg1
    book_name = arg2
    # начинаем работу с документом 'doi.xlsx'
    excel_data_df = pd.read_excel('doi.xlsx', sheet_name='Лист1')
    # формируем список всех DOI
    doi_list = excel_data_df['DOI'].tolist()
    # вычисляем порядковый номер DOI
    doi_temp = doi_list[-1].replace('10.1000/m', '').split('.')
    doi_number = str(int(doi_temp[0]) + 1)
    # создаем новый DOI
    new_doi = '10.1000/m' + doi_number + '.' + isbn
    # представляем DOI в виде ссылки
    doi_url = 'https://doi.org/' + new_doi

    wb = openpyxl.load_workbook('doi.xlsx')
    sheet = wb.active

    a = sheet[f'A{len(doi_list) + 2}']
    a.value = doi_number + '.'
    a.font = Font(name="Times New Roman", size=12)
    a.alignment = Alignment(wrap_text=True, vertical="center")

    b = sheet[f'B{len(doi_list) + 2}']
    b.value = new_doi
    b.font = Font(name="Times New Roman", size=12)
    b.alignment = Alignment(wrap_text=True, vertical="center")

    d = sheet[f'D{len(doi_list) + 2}']
    d.value = book_name
    d.font = Font(name="Times New Roman", size=12)
    d.alignment = Alignment(wrap_text=True, vertical="center")

    # следим за тем, чтобы документ не был открыт в другой программе; сохраняем документ
    try:
        wb.save('doi.xlsx')
    except PermissionError:
        return "Закройте документ 'doi'!"

    return doi_url