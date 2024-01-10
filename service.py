import time

from variables import app_version, app_version_file
import os
import datetime
import platform
from pathlib import Path
import messages as mes
import variables as vars


def get_log_file():
    # Если есть лог флаги, то подготавливаем файл
    next_row = '\n-------------------------------------------------------------\n'

    # Преобразуем в читаемый вид
    temp_data = ''
    temp_data += '\n------НАЧАЛО ЗАПИСИ------\n\n'

    # Записываем информацию о пользователе
    temp_data += 'Пользователь: ' + os.environ.get('USERNAME') + next_row + 'Версия программы: ' + vars.app_version + next_row + \
                 'Время операции: ' + datetime.datetime.now().strftime("%H:%M:%S - %d.%m.%Y") + next_row
    temp_data += 'os: ' + platform.platform() + ', ' + platform.machine() + next_row

    # Записываем пути обрабатываемых файлов
    temp_files_path = ''
    temp_files_path += 'Обрабатываемые файлы:\n' + 'у-П2 Стационар:\n' + vars.tab_hospital_path + '\n'
    if vars.tab_visit_path != '':
        temp_files_path += 'у-П2 Выезд:\n' + vars.tab_visit_path + '\n'
    temp_files_path += 'у-П3 Список доноров:\n' + vars.tab_donors_path + '\n'

    temp_data += temp_files_path
    temp_data += '\nВыходной файл:\n' + vars.tab_donors_path_new + next_row

    if vars.data_lose or vars.donors_size_lose or vars.donors_full_fio_lose:
        temp_data += '\n-----Предупреждения:\n\n'

        if vars.data_lose:
            for row in vars.bad_datas:
                for col in row:
                    temp_data += str(col)
            temp_data += next_row
        if vars.donors_size_lose:
            for row in vars.bad_sizes:
                for col in row:
                    temp_data += str(col)
            temp_data += next_row
        if vars.donors_full_fio_lose:
            temp_data += '\n!!! Необработанные записи:\n'
            for row in vars.bad_donors_fio:
                for col in row:
                    temp_data += str(col)
            temp_data += next_row
        temp_data += '\n------КОНЕЦ ЗАПИСИ------\n'
        temp_data += '\n\n'

        outp_file_path = vars.tab_donors_path_new.replace(os.path.basename(vars.tab_donors_path_new), '')
        file_path = outp_file_path + f'Отчет об ошибках - {os.environ.get("USERNAME")} - {datetime.date.today().strftime("%d.%m.%Y")}.txt'
        file = open(file_path, "w+")
        file.write(temp_data)
        file.close()

        write_flag = False
        try_cont = 0
        while not write_flag and try_cont < 3:
            # Заносим в общий лог
            try:
                path_file = str(Path(r"\\192.168.15.4\Soft\Утилиты\log\Списки страхования\log_all.txt"))
                vars.log_all_path = ''.join(path_file)

                file = open(vars.log_all_path, "a+")
                file.write(temp_data)
                file.close()
                write_flag = True
                print('Записал в лог на сервере!')
            except:
                print(f'Не записал в лог на сервере, жду 10 сек... в {try_cont + 1} раз.')
                time.sleep(10)
                try_cont += 1
        if not write_flag:
            mes.warning('Подключение к файловому серверу',
                            'Внимание!\nСообщаем, что не удалось записать отчет об ошибках на файловый сервер по'
                            ' причине:\n\nПревышено время ожидания разрешения на доступ к файлу от сервера.')
    else:
        temp_data += '\nРезультат обработки:\nОБРАБОТКА УСПЕШНО ЗАВЕРШЕНА БЕЗ ИСКЛЮЧЕНИЙ!\n'

        temp_data += '\n------КОНЕЦ ЗАПИСИ------\n'
        temp_data += '\n\n'

        write_flag = False
        try_cont = 0
        while not write_flag and try_cont < 3:
            # Заносим в общий лог
            try:
                path_file = str(Path(r"\\192.168.15.4\Soft\Утилиты\log\Списки страхования\log_all.txt"))
                vars.log_all_path = ''.join(path_file)

                file = open(vars.log_all_path, "a+")
                file.write(temp_data)
                file.close()
                write_flag = True
                print('Записал в лог на сервере!')
            except:
                print(f'Не записал в лог на сервере, жду 10 сек... в {try_cont+1} раз.')
                time.sleep(10)
                try_cont += 1
        if not write_flag:
            print('Не удалось записать отчет об ошибках на файловый сервер в общий лог')


def write_file(out_text):
    if out_text != '':
        try:
            outp_file_path = vars.tab_donors_path_new.replace(os.path.basename(vars.tab_donors_path_new), '')
            file_path = outp_file_path + 'Отчет по обработке таблиц.txt'
            file = open(file_path, "w+")
            file.write(out_text)
            file.close()
        except Exception as e:
            mes.warning('Создание отчета по обработке таблиц', f'Не удалось записать файл отчета.\nПричина:\n[{e}]')
    # Если текст пустой, то чтобы не вводить в заблуждение пробуем удалить старый файл
    else:
        try:
            outp_file_path = vars.tab_donors_path_new.replace(os.path.basename(vars.tab_donors_path_new), '')
            file_path = outp_file_path + 'Отчет по обработке таблиц.txt'
            os.remove(file_path)
        except Exception as e:
            pass


# Информация о донорах с несколькими донациями
def get_multiple_donors_text(multiple_donors_list):
    multiple_donors_text = f'Доноров с несколькими донациями: {vars.donors_with_multiple_dons}\n\n'

    counter = 1
    for donor in multiple_donors_list:
        el_counter = 0
        for i in donor:
            if el_counter == 0:
                multiple_donors_text += str(counter) + '. (№ п/п) - ' + str(i)
            elif el_counter == 1:
                multiple_donors_text += ', (№ донора) - ' + str(i)
            elif el_counter == 2:
                multiple_donors_text += ', ' + str(i)
            elif el_counter == 3:
                multiple_donors_text += ', ' + str(i)
            elif el_counter == 4:
                multiple_donors_text += ' = ' + str(i)
            elif el_counter == 5:
                multiple_donors_text += ' -> Из: ' + str(i) + '\n'
            el_counter += 1
        counter += 1

    return multiple_donors_text


# Информация о неуникальных донорах
def get_not_uniq_donor_numbers(donors_list_inp, donors_list_outp):
    donors_inp = ''
    donors_outp = ''

    for row in donors_list_inp:
        temp_row_in = ''
        counter = 0
        for col in row:
            if counter == 0 or counter == 1 or counter == 2 or counter == 4 or counter == 5:
                temp_row_in += (str(col) + ', ')
            else:
                continue
        donors_inp = donors_inp + '\n' + temp_row_in

    for row in donors_list_outp:
        temp_row_out = ''
        counter = 0
        for col in row:
            if counter == 0 or counter == 1 or counter == 2 or counter == 3 or counter == 4 or counter == 11:
                temp_row_out += (str(col) + ', ')
            else:
                continue
        donors_outp = donors_outp + '\n' + temp_row_out

    return donors_inp, donors_outp


# Получить итоговый текст для инофрмации по обработке
def get_full_text(dn_donors_inp, dn_donors_out, dn_inp_size, dn_out_size, year_donors_inp, year_donors_out,
                  year_inp_size, year_out_size, multiple_donors_text):
    full_text = ''

    if multiple_donors_text != '':
        full_text += multiple_donors_text

    if dn_donors_inp != '':
        full_text += f'\n\nПри обработке таблиц обнаружены доноры, совпадающие по ФИ.О., но отличающиеся' \
                     f' по "глобальному коду донора":' \
                     f'\n\nКоличество:' \
                     f'\nВ таблицах "410/у-П2": {dn_inp_size}' \
                     f'\nВ таблице "410/у-П3 - Стационар + Выезд": {dn_out_size}' \
                     f'\n\nСведения о доноре:' \
                     f'\nВ таблице "410/у-П2": {dn_donors_inp}' \
                     f'\n\nК ним в "410/у-П3 - Стационар + Выезд": {dn_donors_out}'

    if year_donors_inp != '':
        full_text += f'\n\nПри обработке таблиц обнаружены доноры, совпадающие по ФИ.О. и' \
                     f' "глобальному коду донора", но отличающиеся по году:' \
                     f'\n\nКоличество:' \
                     f'\nВ таблицах "410/у-П2": {year_inp_size}' \
                     f'\nВ таблице "410/у-П3 - Стационар + Выезд": {year_out_size}' \
                     f'\n\nСведения о доноре:' \
                     f'\nВ таблице "410/у-П2": {year_donors_inp}' \
                     f'\n\nК ним в "410/у-П3 - Стационар + Выезд": {year_donors_out}'

    return full_text


# Разбор полного ФИО на Фамилия И.О.
def get_fio(row):
    fio_list = []
    fam, name, otch = '', '', ''

    fio_list = str(row).split()

    # Если Ф.И.
    if len(fio_list) == 2:
        fam = fio_list[0]
        name = fio_list[1]
    # Если Ф.И.О
    elif len(fio_list) == 3:
        fam = fio_list[0]
        name = fio_list[1]
        otch = fio_list[2]
    # Если Ф.И.О + ...
    else:
        fam = fio_list[0]
        name = fio_list[1]
        otch = fio_list[2]
        counter = 0
        for x in fio_list:
            if counter < 3:
                continue
            else:
                otch += ' ' + x
    if otch == '':
        io = name[0] + '.'
    else:
        io = name[0] + '.' + otch[0] + '.'
    fio = fam + ' ' + io
    return fio


# Обработка строки и получение из неё города
def get_city_name_from_row(row):
    start_ind = 0
    end_ind = 0
    city_name = ''
    if 'г.' in row:
        start_ind = row.index('г.') + 3
        end_ind = row.index('\n')
        city_name = row[start_ind:end_ind]
        print('city_name', city_name)
        return city_name
    else:
        return ''


# Удаление файла по его пути
def del_file(path):
    try:
        os.remove(path)
    except Exception as e:
        mes.warning('Удаление временных файлов', f'Не удалось удалить файл {path}\n\nОшибка: [{e}]')


def get_version_from_file():
    version = ''
    file_path = app_version_file
    with open(file_path, 'r') as file:
        version = file.read()
        file.close()
    return version


def check_version():
    version = get_version_from_file()
    if app_version != version:
        error_str = f'У вас используется неактуальная версия программы!\nУ вас - {app_version}.\n' \
                    f'Актуальная - {version}.\n\nАктуальная версия размещена по адресу:\n' \
                    rf'"\\192.168.15.4\Soft\Утилиты\Списки страхования".'
        return error_str
    else:
        return True

