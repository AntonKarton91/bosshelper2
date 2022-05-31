import os
import re
import openpyxl

from employers import *
from sheets import InputSheet, OutputSheet

emloyers_list = [
    {'name': 'Антон', 'surname': 'Киселев', 'stage': 'Конструктор'},
    {'name': 'Михаил', 'surname': 'Соловьев', 'stage': 'Конструктор'},
    {'name': 'Денис', 'surname': 'Хохин', 'stage': 'Конструктор'},
    {'name': 'Александр', 'surname': 'Никитин', 'stage': 'Конструктор'},
    {'name': 'Дмитрий', 'surname': 'Бычков', 'stage': 'Конструктор'},
    {'name': 'Элеонора', 'surname': 'Йер', 'stage': 'Дизайнер'},
    {'name': 'Лена', 'surname': 'Пухова', 'stage': 'Дизайнер'},
    {'name': 'Стас', 'surname': 'Гагров', 'stage': 'Дизайнер'},
    {'name': 'Егор', 'surname': 'Горячкин', 'stage': 'Мухожук'},
]

for i in emloyers_list:
    print(i['name']+' '+i['surname'])
month=None
year=None
input('Проверьте чтобы список работников был актуальным и нажмите ENTER')
while not month or type(month) not in (int,):
    month = input('Введите месяц в числовом формате - "2" и нажмите ENTER')
while not year or type(year) not in (int,):
    year = input('Введите год в числовом формате - "2022" и нажмите ENTER')
wb=InputSheet()
wb.get_active_sheet()

wb.parse_input_sheet(emloyers_list, year, month)
l=wb.parse_input_sheet(emloyers_list, year, month)

for employer in l:
    e=OutputSheet(employer['name'], employer['surname'], employer['stage'], month, year)
    e.creating(l)




