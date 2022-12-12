from checkFiles import check_files
from data_base import use_database
from employerList import emloyers_list
from sheets import InputSheet, OutputSheet


check_files(emloyers_list)

for i in emloyers_list:
    print(i['name'] + ' ' + i['surname'])

month = None
year = None

use_database()

input('Проверьте чтобы список работников был актуальным и нажмите ENTER ')

while type(month) != int or not 1 <= month <= 12:
    month = int(input('Введите месяц в числовом формате - "2" и нажмите ENTER   '))
while type(year) != int or not 2021 <= year <= 2025:
    year = int(input('Введите год в числовом формате - "2022" и нажмите ENTER   '))

wb = InputSheet()
wb.get_active_sheet()

wb.parse_input_sheet(emloyers_list, year, month)
l = wb.parse_input_sheet(emloyers_list, year, month)

for employer in l:
    e = OutputSheet(employer['name'], employer['surname'], employer['stage'], month, year)
    e.creating(l)

print('Успешно')
