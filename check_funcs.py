import variables as vars
import os
import win32com.client as win32
import messages as mes
import psutil


# Проверка периодов выгрузки во входных таблицах
def check_upload_period():
    vars.bad_datas.clear()
    vars.bad_donors_fio.clear()
    vars.bad_sizes.clear()

    if vars.tab_data_row_hospital != '' and vars.tab_data_row_last != '' and vars.tab_data_row_hospital != vars.tab_data_row_last:
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц\n', 'у-П2 Стационар: ', vars.tab_data_row_hospital, '\nу-П3 Список доноров: ', vars.tab_data_row_last))
        mes.warning('Отличается период выгрузки', f'Внимание!\n\nОбнаружено несовпадение периодов выгрузки!\n\nУП2 Стационар - {vars.tab_data_row_hospital}\nУП3 Доноры - {vars.tab_data_row_last}')

    elif vars.tab_data_row_hospital == '':
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц - "Пустой период в у-П2 Стационар"\n',
                              'у-П2 Стационар: ',
                              vars.tab_data_row_hospital, '\nу-П3 Список доноров: ',
                              vars.tab_data_row_last))
        mes.warning('Некорректный период выгрузки',
                    f'Внимание!\n\nПериод выгрузки в таблице УП2 Стационар не найден!')

    elif vars.tab_data_row_last == '':
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц - "Пустой период в у-П3 Список доноров"\n',
                              'у-П2 Стационар: ',
                              vars.tab_data_row_hospital, '\nу-П3 Список доноров: ',
                              vars.tab_data_row_last))
        mes.warning('Некорректный период выгрузки',
                    f'Внимание!\n\nПериод выгрузки в таблице УП3 Доноры не найден!')

    if vars.tab_visit_path != '' and vars.tab_data_row_visit != '' and vars.tab_data_row_last != '' and vars.tab_data_row_visit != vars.tab_data_row_last:
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц\n',
                              'у-П2 Выезд: ',
                              vars.tab_data_row_visit, '\nу-П3 Список доноров: ',
                              vars.tab_data_row_last))

        mes.warning('Отличается период выгрузки',
                    f'Внимание!\n\nОбнаружено несовпадение периодов выгрузки!\n\nУП2 Выезд - {vars.tab_data_row_visit}\nУП3 Доноры - {vars.tab_data_row_last}')
    elif vars.tab_visit_path != '' and vars.tab_data_row_visit != '' and vars.tab_data_row_last != '' and vars.tab_data_row_visit != vars.tab_data_row_hospital:
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц\n',
                              'у-П2 Выезд: ',
                              vars.tab_data_row_visit, '\nу-П2 Стационар: ',
                              vars.tab_data_row_hospital))

        mes.warning('Отличается период выгрузки',
                    f'Внимание!\n\nОбнаружено несовпадение периодов выгрузки!\n\nУП2 Выезд - {vars.tab_data_row_visit}\nУП2 Стационар - {vars.tab_data_row_hospital}')
    elif vars.tab_visit_path != '' and vars.tab_data_row_visit == '':
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц - "Пустой период в у-П2 Выезд"\n',
                              'у-П2 Выезд: ',
                              vars.tab_data_row_visit, '\nу-П3 Список доноров: ',
                              vars.tab_data_row_last))

        mes.warning('Некорректный период выгрузки',
                    f'Внимание!\n\nПериод выгрузки в таблице УП2 Выезд не найден!')
    elif vars.tab_visit_path != '' and vars.tab_data_row_last == '':
        vars.data_lose = True
        vars.bad_datas.append(('!!! Не совпадает период выгрузки для таблиц - "Пустой период в у-П3 Список доноров"\n',
                              'у-П2 Стационар: ',
                              vars.tab_data_row_visit, '\nу-П3 Список доноров: ',
                              vars.tab_data_row_last))

        mes.warning('Некорректный период выгрузки',
                    f'Внимание!\n\nПериод выгрузки в таблице УП3 Доноры не найден!')


def check_uniq_path(path, name):
    if name == 'hospital':
        if path == vars.tab_visit_path or path == vars.tab_donors_path:
            return True
        else:
            return False
    elif name == 'visit':
        if path == vars.tab_hospital_path or path == vars.tab_donors_path:
            return True
        else:
            return False
    elif name == 'donors':
        if path == vars.tab_hospital_path or path == vars.tab_visit_path:
            return True
        else:
            return False


def check_tab_paths():
    if vars.tab_donors_path != '' and vars.tab_hospital_path != '':
        return True
    else:
        return False


def check_visit_tab_path():
    if vars.tab_visit_path != '':
        return True
    else:
        return False


def check_extension(file_extension):
    if file_extension == '.xls' or file_extension == '.xlsx':
        return True
    else:
        return False


def get_file_info(path):
    name, extension, file_path = '', '', ''
    file_info = []
    name_ex = get_filename_and_ext(path)

    # название
    name = name_ex[0]

    # расширение
    extension = name_ex[1]

    # абсолютный путь до файла
    file_path = path.replace(os.path.basename(path), '')

    file_info.append(file_path)
    file_info.append(name)
    file_info.append(extension)

    return file_info


def get_filename_and_ext(path):
    filename = os.path.basename(path)
    pathname, file_extension = os.path.splitext(path)
    file_info = [filename.replace(file_extension, ''), file_extension]
    return file_info


def check_xlsx(path):
    if get_filename_and_ext(path)[1] == '.xlsx':
        return True
    else:
        return False


def check_xlsx_ex(file_extension):
    if file_extension == '.xlsx':
        return True
    else:
        return False


def check_old_extension(path1, path2, path3):
    if get_filename_and_ext(path1)[1] == '.xls' or get_filename_and_ext(path2)[1] == '.xls' or get_filename_and_ext(path3)[1] == '.xls':
        return True
    else:
        return False


def check_run_excel():
    if "EXCEL.EXE" in (i.name() for i in psutil.process_iter()):
        return True
    else:
        return False


# Проверка запрашиваемого пути
def check_path(path):
    if os.path.exists(path):
        return True
    else:
        return False

