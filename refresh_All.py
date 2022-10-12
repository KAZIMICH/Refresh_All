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


# cписки имен и путей файлов ВК
def lists_name_path_files(i):
    for root, dirs, files in os.walk(i):
        for file in files:
            if file.endswith('ВК.xlsx') and not '~$' in file \
                    or file.endswith('ВК.xls') and not '~$' in file \
                    or file.endswith('ВК.xlsm') and not '~$' in file:
                file_list.append(os.path.join(file))
                path_list.append(os.path.join(root, file))
    if len(file_list) > 0:
        print(f'В папке с проектом найдено {len(file_list)} файлов для обновления')
        answer = dialog_yes_no('Вывести список файлов для обновления?\nВведите Y или N')
        if answer == 'y':
            print_list(file_list)
            print('_' * 100)
    else:
        print(input('Список пуст. Что-то пошло не так! Перезапустите программу...'))
        sys.exit()
    return file_list, path_list


# проверка дублей файлов
def double(i):
    visited = set()
    dup = [x for x in i if x in visited or (visited.add(x) or False)]
    if len(dup) > 0:
        for j in dup:
            print(f'Документ {j} присутствует несколько раз')
        print(input('Удалите неактуальный файл(ы) и перезапустите приложение'))
        sys.exit()


# проверка на открытые файлы ВК
def file_list_open(i):
    for j in i:
        file_check(j)


# Проверка файла на присутствие и незанятость
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
    return valid


# открытие / обоновление / сохранение / обновление файла Базы данных
def refresh_DB(i):
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
    print('_' * 100)


# обновление файлов excel
def refresh_files(i):
    counter = 1
    for j in i:
        file_check(j)
        print("Обновление файла", counter, "в списке")
        print(j)
        wb = excel.Workbooks.open(j)
        wb.Application.DisplayAlerts = False
        excel.Visible = False
        wb.RefreshAll()
        wb.Save()
        wb.Close()
        excel.Quit()
        counter += 1
        print('Файл', counter - 1, 'обновлен')
    print(f'Обновлено {len(path_list)} файлов в списке')


# проверка ввода пользовательских данных
def dialog_yes_no(i):
    while True:
        answers = {'yes': 1, 'y': 1, 'no': 0, 'n': 0}
        print(i)
        user_answer = input().lower()
        if user_answer in answers:
            return user_answer
        else:
            print('Ожидалось Y или N')

# печать списков
def print_list(i):
    for j in i:
        print(j)


# Тело программы
refresh_DB(file_path_DB)

lists_name_path_files(path_folder)

double(file_list)

file_list_open(path_list)

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

print('_' * 100)

print('Обновление файлов взаимодействия с NX')
refresh_files(NX_list)
print(f'Обновлено {len(path_list)} файлов в списке')

print('_' * 100)

refresh_DB(file_path_DB)

endTime = time.time()
totalTime = endTime - startTime
print('Программа завершена')
print(input(f'Затраченное время = {int(totalTime)} секунд\n'))
sys.exit()

# тестировать
