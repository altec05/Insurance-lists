import os

# Текущая версия программы
app_version = '1.10'
app_last_edit_version = '01.09.2023 г.'
# Путь до файла версии
app_version_file = r"\\192.168.15.4\Soft\Программирование\Py\Списки страхования\Config\version.txt"

# Путь до файла с изменениями версий
# path_changes = r"\\192.168.15.4\Soft\Программирование\Py\Списки страхования\ПО Списки страхования\change_log.txt"
path_changes = os.path.abspath('Change_log.txt')

# Пути для входных таблиц, форматированных временных таблиц и итоговых файлов
tab_hospital_path = ''
tab_visit_path = ''
tab_donors_path = ''
out_tab_path = ''
temp_hospital_tab_path = ''
temp_visit_tab_path = ''

# Пути для преобразованных файлов
tab_hospital_path_new = ''
tab_visit_path_new = ''
tab_donors_path_new = ''

# Инструкция для пользователя
readme_path = os.path.abspath("Readme.txt")
donors_double_path = os.path.abspath("Отчет по обработке таблиц.txt")

# Статистика
donors_not_uniq_inp = 0
donors_not_uniq_outp = 0
donors_get = 0
donors_with_multiple_dons = 0
donors_out = 0

# Учет необработанных записей
donors_with_dots = 0
donors_nums_indexes_list = []

# Таймеры производительности
tab_get_time_min = 0
tab_get_time_sec = 0
tab_fill_time_min = 0
tab_fill_time_sec = 0

# Папка "Документы" пользователя
usr_docs = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')
# Папка "Рабочий стол"
usr_desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
# Последняя открытая папка
last_folder = ''

# Проверка даты выгрузки
tab_data_row_last = ''
tab_data_row_hospital = ''
tab_data_row_visit = ''

# Для задания периода выгрузки
data_for_name = ''

# Первая строка для титула выходного файла
first_row_text = ''

# Флаги для логов
data_lose = False
donors_size_lose = False
donors_full_fio_lose = False

# Флаг для вывода в txt отчет/лог
multidons_out = False
log_out = False

# Переменные для логов
bad_datas = list()
bad_sizes = list()
bad_donors_fio = list()
log_all_path = ''

# Переменная состояния верхнего окна
toplevel_window = None

# Переменная дял записи статуса Филиал или нет
city_name = ''


# Выдача статистических данных по завершению обработки данных
def get_statistic():
    statistic_data = []
    statistic_data.append(tab_get_time_sec)
    statistic_data.append(tab_fill_time_sec)
    statistic_data.append(donors_get)
    statistic_data.append(donors_with_multiple_dons)
    statistic_data.append(donors_out)
    return statistic_data


