import os


def check_files(list):
    mainDir = os.getcwd()
    for employer in list:
        employerName = employer['name'] + ' ' + employer['surname'] + '.xlsx'
        if os.path.exists(os.path.join(mainDir, employerName)):
            continue
        else:
            raise FileNotFoundError('Отсутствует один или несколько файлов Excel в папке назначения')

