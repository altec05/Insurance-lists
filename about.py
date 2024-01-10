import customtkinter as CTk
from datetime import datetime

import changes_from as chs_from
import variables as vars
import check_funcs
import os
import messages as mes


class AboutWin(CTk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("700x400+500+300")
        self.title('О программе')
        self.resizable(width=False, height=False)
        self.iconbitmap('logo.ico')
        CTk.deactivate_automatic_dpi_awareness()

        label_text_up = f'Сведения о программе "Списки страхования"\n\n© Разработка и права: Домашенко Иван / ' \
                        f'Администратор ИБ ВС\n\n\nПрограмма была разработана в целях форматирования списков ' \
                        f'страхования донаций в виде электронных таблиц Excel в формате xls или xlsx\n' \
                        f'Программа написана с применением языка программирования Python v3.11'
        label_text_center = f'Версия программы - ver. {vars.app_version} от {vars.app_last_edit_version}'
        label_text_down = f'КГКУЗ "Красноярский краевой центр крови №1"\n\n2023 - {datetime.now().year}'

        self.label_up = CTk.CTkLabel(self, text=label_text_up, wraplength=550)
        self.label_up.pack(padx=20, pady=15)

        self.label_center = CTk.CTkLabel(self, text=label_text_center, wraplength=550, anchor='center')
        self.label_center.pack(padx=20, pady=20)

        self.show_changes_button = CTk.CTkButton(master=self, text='Изменения', width=125,
                                               command=self.open_changes)
        self.show_changes_button.pack(pady=5, padx=5)

        self.label_down = CTk.CTkLabel(self, text=label_text_down)
        self.label_down.pack(padx=20, pady=50)

        self.focus()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: self.dismiss())  # перехватываем нажатие на крестик

    def dismiss(self):
        vars.toplevel_window = None
        print(vars.toplevel_window)
        self.grab_release()
        self.destroy()

    def open_changes(self):
        if check_funcs.check_path(vars.path_changes):
            print(vars.path_changes)
            os.system(fr"explorer.exe {vars.path_changes}")
        else:
            try:
                outp_file_path = vars.path_changes.replace(os.path.basename(vars.path_changes), '')
                file_path = outp_file_path + 'Change_log.txt'
                file = open(file_path, "w+")
                file.write(chs_from.changes_row)
                file.close()

                if check_funcs.check_path(vars.path_changes):
                    print(vars.path_changes)
                    os.system(fr"explorer.exe {vars.path_changes}")
                
            except Exception as e:
                mes.warning('Создание файла изменений', f'Не удалось записать файл изменений.\nПричина:\n[{e}]')
            # mes.error('Открытие файла', 'Файл с изменениями не найден!')

