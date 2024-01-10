import datetime
import calendar
from pathlib import Path
import os

import variables as vars
import check_funcs as chf


# Получаем имя временного файла с расширением xlsx вместо xls
def get_temp_file_path(path):
    import locale
    locale.setlocale(category=locale.LC_ALL, locale="Russian")
    dt_obj = datetime.date.today()
    dt_string = dt_obj.strftime("%d.%m.%Y")

    path_file = ''
    file_path = chf.get_file_info(path)[0]
    file_name = chf.get_file_info(path)[1]
    file_ex = chf.get_file_info(path)[2]
    if file_ex == '.xls':
        path_file = ''.join([file_path, file_name + ' _temp_ ' + dt_string, file_ex + 'x'])
    else:
        path_file = ''.join([file_path, file_name + ' _temp_ ' + dt_string, file_ex])
    path_xlsx = ''.join(str(Path(path_file)))
    return path_xlsx


# Получаем имя выходного файла с расширением xlsx и итоговым названием
def get_out_file_path(path):
    import locale
    locale.setlocale(category=locale.LC_ALL, locale="Russian")
    dt_obj = datetime.date.today()

    month = ''
    if vars.data_for_name != '':
        temp_data = vars.data_for_name
        month_str = temp_data[3:5]
        month_int = int(month_str)
        if 0 < month_int < 13:
            month = calendar.month_name[month_int]

    file_path = chf.get_file_info(path)[0]
    if month != '':
        if vars.city_name != '':
            file_name = f'Список для страхования {vars.city_name}' + ' - ' + month + ' ' + dt_obj.strftime("%Y")
        else:
            file_name = 'Список для страхования' + ' - ' + month + ' ' + dt_obj.strftime("%Y")
    else:
        if vars.city_name != '':
            file_name = f'Список для страхования {vars.city_name}' + ' - ' + dt_obj.strftime("%B %Y")
        else:
            file_name = f'Список для страхования' + ' - ' + dt_obj.strftime("%B %Y")

    file_ex = chf.get_file_info(path)[2]
    file_ex_xlsx = ''
    if file_ex == '.xls':
        file_ex_xlsx = file_ex + 'x'
    else:
        file_ex_xlsx = file_ex

    path_xlsx = ''.join([file_path, file_name, file_ex_xlsx])

    return ''.join(str(Path(path_xlsx)))


def get_folder_path(path):
    send_data = path.replace(os.path.basename(path), '')
    return send_data