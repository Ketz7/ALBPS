import openpyxl
import random

wb = openpyxl.load_workbook('Data.xlsx')
sheet = wb.get_sheet_by_name('Sheet')
excel = open('it works.xls', 'r')
i = 2
for line in excel:
  for j in range(2,6):
    if sheet['A{}'.format(j)].value.strip() == line[13:].strip():
      print('IN')
      try:
        sheet['B{}'.format(j)] = int(sheet['B{}'.format(j)].value) + random.randint(1,11)
        print('Added')
      except TypeError:
        sheet['B{}'.format(j)] = random.randint(1,11)
  i = i + 1
wb.save('Data.xlsx')
