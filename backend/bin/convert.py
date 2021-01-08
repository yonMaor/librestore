#!/usr/bin/python3

import csv
from openpyxl import load_workbook

def csv_from_excel(book, sheet):
    workbook = load_workbook(filename=book, read_only=True)
    sheet = workbook[sheet]
    csv_file = open('e.csv', 'w')
    record = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

    for row in sheet.rows:
        index = []
        for cell in row:
            index.append(cell.value)
        record.writerow(index)

    csv_file.close()

csv_from_excel('e.xlsx', 'Sheet1')
