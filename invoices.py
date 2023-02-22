import csv
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font

search = r'\s*F\s*\s*V\s*\s*S\s*\s*W\s*/?\s*\d{2}\s*/?\s*\d{2}\s*/?\s*\d{4}\s*/?\s*|\s*F\s*\s*S\s*\s*Z\s*\s*W\s*/?\s*\d{2}\s*/?\s*\d{2}\s*/?\s*\d{4}\s*/?\s*|\s*N\s*\s*K\s*/?\s*\d{2}\s*/?\s*\d{2}\s*/?\s*\d{4}\s*/?\s*'

excelRead = pd.read_excel('faktury-2023-02-17.xlsx')

workbook = openpyxl.Workbook()
worksheet = workbook.active

worksheet['A1'] = 'Nr faktury'
worksheet['B1'] = 'Do zapłaty'
worksheet['C1'] = 'Kwota przelewu'
worksheet['D1'] = 'Waluta przelewu'
worksheet['E1'] = 'Data przelewu'
worksheet['F1'] = 'Kontrahent'
worksheet['G1'] = 'Nr. przelewu'

bold_font = Font(bold=True)
for cell in worksheet[1]:
    cell.font = bold_font

with open('historia_2023-02-17_91109010690000000102586741.csv','r', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    rowId = 1
    print("Trwa generowanie xls...")
    for row in reader:
        transfer_title = row[2]
        company = row[3]
        amount = row[5]
        pid = row[7]
        day = row[0]
        matches = re.findall(search, transfer_title)
        if len(matches) > 0:
            for item in matches:
                format = item.replace(" ", "").rstrip(" /")
                if "/" in format:
                    result = format
                else:
                    formatted_str = f"{format[:4]}/{format[4:6]}/{format[6:8]}/{format[8:12]}"
                    result = formatted_str
                for index, row in excelRead.iterrows():
                    if row['Nr faktury'] == result:
                        rowId += 1
                        worksheet['A' + str(rowId)] = row['Nr faktury']
                        worksheet['B' + str(rowId)] = row['Do zapłaty']
                        worksheet['C' + str(rowId)] = amount
                        worksheet['D' + str(rowId)] = row['Waluta']
                        worksheet['E' + str(rowId)] = day
                        worksheet['F' + str(rowId)] = company
                        worksheet['G' + str(rowId)] = pid

print("Plik utworzony")

workbook.save('przelewy.xlsx')
               
              
            
        
