# импорт библиотеки для работы с файлами excel
import win32com.client
# импорт библиотеки для работы с операционной системой
import os
# импорт библиотеки для работы с временем
import time
# импорт остановки выполнения
import sys

excel = win32com.client.Dispatch("Excel.Application")
# время старта приложения
startTime = time.time()

# Описание работы приложения
if True:
    print('Творящий РКД приветствую тебя!')
    # print('Программа обновления файлов excel РКД проекта Т40В:')
    # print(r'-    ищет все файлы excel в папке проекта \n\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\01 T40B '
    #       r'ПДСП\04 Электрика\nс суффиксом ВК')
    # print('-    выводит список файлов (по запросу)')
    # print('-    проверяет каждый файл ВК на доступность (существует/не существует, открыт/закрыт')
    # print('-    проверяет каждый файл ВК на дублированность')
    # print('-    открывает каждый файл из списка, обновляет и закрывает файл')
    # print()
    # print('-    проверяет на доступность обобщенных файлов excel:')
    # print('     1) Т40В-Кабельные журналы для NX import')
    # print('     2) Т40В-Кабельные журналы для NX import POWER')
    # print('     3) Т40В-Кабельные журналы для NX import SIGNAL')
    # print('-    открывает эти файлы, обновляет и закрывает')
    print('_' * 100)

# Вводные данные
if True:
    # Путь к файлу Базы данных
    file_path_DB = r'\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\011 Перечень оборудования' \
                   r'\01_База данных' \
                   r'\Т40В_БД Оборудования.xlsx '

    # Указываем пути к файлам для работы с NX
    file_path_importNX = r'\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\02 РКД Т-40\05 ' \
                         r'NX\Import\Т40В-Кабельные ' \
                         r'журналы для NX import.xlsx '
    file_path_NXPOWER = r'\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\02 РКД Т-40\05 ' \
                        r'NX\Import\Т40В-Кабельные ' \
                        r'журналы для NX import POWER.xlsx '
    file_path_NXSIGNAL = r'\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\02 РКД Т-40\05 ' \
                         r'NX\Import\Т40В-Кабельные ' \
                         r'журналы для NX import SIGNAL.xlsx '
    # список путей файлов для работы с NX
    NX_list = (file_path_importNX, file_path_NXPOWER, file_path_NXSIGNAL)

    # путь к папке ПДСП проекта Т40В
    path_folder = r'\\192.168.1.98\electric group\T40B\03 ПРОЕКТ док T40B\01 T40B ПДСП\04 Электрика'

    # список, хранящий имена файлов ВК
    file_list = []

    # список, хранящий пути файлов ВК
    path_list = []


# функция получения списка имен файлов ВК
def get_name_VK(path_folder):
    for root, dirs, files in os.walk(path_folder):
        for file in files:
            if file.endswith('ВК.xlsx') and not file.startswith('~$') \
                    or file.endswith('ВК.xls') and not file.startswith('~$') \
                    or file.endswith('ВК.xlsm') and not file.startswith('~$'):
                file_list.append(os.path.join(file))
    return file_list


# Функция проверки дублей файлов
def double(file_list):
    visited = set()
    dup = [x for x in file_list if x in visited or (visited.add(x) or False)]
    if len(dup) > 0:
        for i in dup:
            print(f'Документ {i} присутствует несколько раз')
        print(input('Удалите неактуальный файл(ы) и перезапустите приложение'))
        sys.exit()


# Функция получения списка путей к файлам ВК
def get_path_VK(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('ВК.xlsx') and not '~$' in file  \
                    or file.endswith('ВК.xls') and not '~$' in file \
                    or file.endswith('ВК.xlsm') and not '~$' in file:
                path_list.append(os.path.join(root, file))
    return path_list


# Функция проверки файла на присутствие и незанятость
def file_check(file):
    valid = False
    while not valid:
        if os.path.exists(file):
            try:
                os.rename(file, file)
                valid = True
            except IOError:
                print(input(f'Файл {file} открыт.\nДля принятия текущих изменений в файле сохраните его и нажмите '
                            f'Enter.\n'))
                valid = False
        else:
            print('Файл', file, 'был перемещен или удален')
            print(input('Для подтверждения нажмите Enter.\n'))
            valid = False
    # return valid


# Функция печати списков
def print_list(lst):
    for i in lst:
        print(i)


# Функция обновления файлов excel
def refresh_files(path_list):
    counter = 1
    for i in path_list:
        file_check(i)
        print("Обновление файла", counter, "в списке")
        print(i)
        wb = excel.Workbooks.open(i)
        wb.Application.DisplayAlerts = False
        excel.Visible = False
        wb.RefreshAll()
        wb.Save()
        wb.Close()
        excel.Quit()
        counter += 1
        print('Файл', counter - 1, 'обновлен')


# открытие / обоновление / сохранение / обновление файла Базы данных
def open_DB(i):
    file_check(i)
    print("Обновление файла Базы данных...")
    wb = excel.Workbooks.open(i)
    wb.Application.DisplayAlerts = True
    excel.Visible = True
    wb.RefreshAll()
    time.sleep(2)
    wb.Save()
    wb.RefreshAll()
    time.sleep(2)
    wb.Save()
    wb.Close()
    excel.Quit()
    print('Файл Базы данных обновлен')


# Функция проверки ввода пользовательских данных
def dialog_yes_no(question):
    while True:
        answers = {'yes': 1, 'y': 1, 'no': 0, 'n': 0}
        print(question)
        user_answer = input().lower()
        if user_answer in answers:
            return user_answer
        else:
            print('Ожидалось Y или N')


# Тело программы
open_DB(file_path_DB)
print('_' * 100)

get_name_VK(path_folder)

double(file_list)

print(f'В папке с проектом найдено {len(file_list)} файлов для обновления')
user_answer = dialog_yes_no('Вывести список файлов для обновления?\nВведите Y или N')

if user_answer == 'y':
    print_list(file_list)

print('_' * 100)

get_path_VK(path_folder)

print('Файлы готовы к обновлению')
user_answer = dialog_yes_no('Обновить файлы?\nВведите Y или N')

if user_answer == 'y':
    refresh_files(path_list)
else:
    print('_' * 100)
    print('Выполнение приложения прервано пользователем')
    endTime = time.time()
    totalTime = endTime - startTime
    print('Работа завершена')
    print(input(f'Затраченное время = {int(totalTime)} секунд\n'))
    sys.exit()

print(f'Обновлено {len(path_list)} файлов в списке')

print('_' * 100)

print('Обновление файлов взаимодействия с NX')
refresh_files(NX_list)
print(f'Обновлено {len(path_list)} файлов в списке')

print('_' * 100)

open_DB(file_path_DB)

endTime = time.time()
totalTime = endTime - startTime
print('Программа завершена')
print(input(f'Затраченное время = {int(totalTime)} секунд\n'))
# input()
sys.exit()
