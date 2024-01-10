import customtkinter as CTk
import tkinter.filedialog as fd
import os
import threading
import psutil

from functools import partial
from tkinter import messagebox
import psutil

import about, instruction, service
import variables as vars
import messages as mes
import check_funcs as ch_foo
import data_edit
from path_funcs import get_folder_path


def change_appearance_mode_event(new_appearance_mode):
    print(CTk.get_appearance_mode())
    CTk.set_appearance_mode(new_appearance_mode)
    print(CTk.get_appearance_mode())


class App(CTk.CTk):
    def __init__ (self):
        super().__init__()

        self.geometry("530x422")
        self.title("Списки страхования")
        self.resizable (False, False)
        self.iconbitmap('logo.ico')
        CTk.set_default_color_theme("dark-blue")
        CTk.set_appearance_mode("system")
        CTk.deactivate_automatic_dpi_awareness()

        self.input_hospital_frame = CTk.CTkFrame(master=self, fg_color='transparent')
        self.input_hospital_frame.pack(fill='x', ipadx=10, ipady=10)

        self.input_visit_frame = CTk.CTkFrame(master=self, fg_color='transparent')
        self.input_visit_frame.pack(fill='x', ipadx=10, ipady=10)

        self.output_frame = CTk.CTkFrame(master=self, fg_color='transparent')
        self.output_frame.pack(fill='x', ipadx=10, ipady=10)

        self.start_frame = CTk.CTkFrame(master=self, fg_color='transparent')
        self.start_frame.pack(fill='x', ipadx=10, ipady=5)

        self.start_options_frame = CTk.CTkFrame(master=self, fg_color='transparent')
        self.start_options_frame.pack(fill='x', ipadx=10, ipady=5)

        self.options_frame = CTk.CTkFrame(master=self, fg_color='transparent', height=100)
        self.options_frame.pack(fill='x', ipadx=10, ipady=10)

        self.tab_hospital_button = CTk.CTkButton(master=self.input_hospital_frame, text='410/у-П2\nСтационар', width=180,
                                                 command=self.set_tab_hospital)
        self.tab_hospital_button.pack(side='left', pady=5, padx=25)

        self.tab_hospital_label = CTk.CTkLabel(master=self.input_hospital_frame, text='Таблица не выбрана', width=250,
                                               text_color='red', anchor='w')
        self.tab_hospital_label.pack(side='left', pady=5)

        self.tab_visit_button = CTk.CTkButton(master=self.input_visit_frame, text='410/у-П2\nВыезд (при наличии)', width=180,
                                              command=self.set_tab_visit)
        self.tab_visit_button.pack(side='left', pady=5, padx=25)

        self.tab_visit_label = CTk.CTkLabel(master=self.input_visit_frame, text='Таблица не выбрана', width=250,
                                            text_color='red', anchor='w')
        self.tab_visit_label.pack(side='left', pady=5)

        self.tab_donors_button = CTk.CTkButton(master=self.output_frame, text='410/у-П3\nСтационар + Выезд', width=180,
                                               command=self.set_tab_donors)
        self.tab_donors_button.pack(side='left', pady=5, padx=25)

        self.tab_donors_label = CTk.CTkLabel(master=self.output_frame, text='Таблица не выбрана', width=250,
                                             text_color='red', anchor='w')
        self.tab_donors_label.pack(side='left', pady=5)

        self.tab_start_button = CTk.CTkButton(master=self.start_frame, text='Подготовить список', width=180,
                                              fg_color='green',
                                              command=self.start_dons)
        self.tab_start_button.pack(side='left', pady=5, padx=25)

        self.wait_label = CTk.CTkLabel(master=self.start_frame, text='', width=250, text_color='green', anchor='w')
        self.wait_label.pack(side='left', ipadx=5, ipady=5, pady=15)

        def checkbox_event():
            print("checkbox toggled, current value:", self.check_var.get())

        self.check_var = CTk.StringVar(value="off")
        self.multidons_out_checkbox = CTk.CTkCheckBox(master=self.start_options_frame,
                                        text="Вывести список доноров с несколькими донациями в txt файл",
                                        command=checkbox_event, variable=self.check_var, onvalue="on", offvalue="off",
                                        corner_radius=6, border_width=1)
        self.multidons_out_checkbox.pack(side='top', ipadx=5, ipady=5, pady=5, padx=25, anchor='w')

        self.check_log_var = CTk.StringVar(value="on")
        self.log_out_checkbox = CTk.CTkCheckBox(master=self.start_options_frame,
                                                text="Вывод ошибок и предупреждений в txt файл",
                                                command=checkbox_event, variable=self.check_log_var, onvalue="on",
                                                offvalue="off", corner_radius=6, border_width=1)
        self.log_out_checkbox.pack(side='top', ipadx=5, ipady=5, pady=5, padx=25, anchor='w')

        self.appearance_mode_option_menu = CTk.CTkOptionMenu(master=self.options_frame, values=["System", "Light", "Dark"], command=change_appearance_mode_event, width=180)
        self.appearance_mode_option_menu.pack(side='left', pady=5, padx=25)

        self.show_about_button = CTk.CTkButton(master=self.options_frame, text='О программе', width=125, command=self.show_about)
        self.show_about_button.pack(side='left', pady=5, padx=5)

        self.show_instruction_button = CTk.CTkButton(master=self.options_frame, text='Инструкция', width=125, command=self.show_instruction)
        self.show_instruction_button.pack(side='left', pady=5, padx=25)

        # self.toplevel_window = None

    def clear_paths_tabs(self):
        self.tab_hospital_label.configure(text_color='red')
        self.tab_hospital_label.configure(text='Таблица не выбрана!')

        self.tab_visit_label.configure(text_color='red')
        self.tab_visit_label.configure(text='Таблица не выбрана')

        self.tab_donors_label.configure(text_color='red')
        self.tab_donors_label.configure(text='Таблица не выбрана!')

    def set_tab_hospital(self):
        filetypes = [("Excel files", ".xlsx .xls")]

        if vars.last_folder == '':
            vars.tab_hospital_path = fd.askopenfilename(title="Укажите таблицу 410/у-П2 Стационар", initialdir=f"{vars.usr_desktop}",
                                                     filetypes=filetypes)
            if vars.tab_hospital_path != '':
                vars.last_folder = get_folder_path(vars.tab_hospital_path)
        else:
            vars.tab_hospital_path = fd.askopenfilename(title="Укажите таблицу 410/у-П2 Стационар", initialdir=f"{vars.last_folder}",
                                                     filetypes=filetypes)
            if vars.tab_hospital_path != '':
                vars.last_folder = get_folder_path(vars.tab_hospital_path)

        if vars.tab_hospital_path != '':
            if not ch_foo.check_uniq_path(vars.tab_hospital_path, 'hospital'):
                self.tab_hospital_label.configure(text_color='green')
                self.tab_hospital_label.configure(text=os.path.basename(vars.tab_hospital_path))
            else:
                mes.error('Ошибка уникальности файла', f'Файл {os.path.basename(vars.tab_hospital_path)} уже был выбран в другом разделе!')
        else:
            self.tab_hospital_label.configure(text_color='red')
            self.tab_hospital_label.configure(text='Таблица не выбрана!')

    def set_tab_visit(self):
        filetypes = [("Excel files", ".xlsx .xls")]

        if vars.last_folder == '':
            vars.tab_visit_path = fd.askopenfilename(title="Укажите таблицу 410/у-П2 Выезд (при наличии)", initialdir=f"{vars.usr_desktop}",
                                                     filetypes=filetypes)
            if vars.tab_visit_path != '':
                vars.last_folder = get_folder_path(vars.tab_visit_path)
        else:
            vars.tab_visit_path = fd.askopenfilename(title="Укажите таблицу 410/у-П2 Выезд (при наличии)", initialdir=f"{vars.last_folder}",
                                                     filetypes=filetypes)
            if vars.tab_visit_path != '':
                vars.last_folder = get_folder_path(vars.tab_visit_path)

        if vars.tab_visit_path != '':
            if not ch_foo.check_uniq_path(vars.tab_visit_path, 'visit'):
                self.tab_visit_label.configure(text_color='green')
                self.tab_visit_label.configure(text=os.path.basename(vars.tab_visit_path))
            else:
                mes.error('Ошибка уникальности файла', f'Файл {os.path.basename(vars.tab_visit_path)} уже был выбран в другом разделе!')

        else:
            self.tab_visit_label.configure(text_color='red')
            self.tab_visit_label.configure(text='Таблица не выбрана')

    def set_tab_donors(self):
        filetypes = [("Excel files", ".xlsx .xls")]

        if vars.last_folder == '':
            vars.tab_donors_path = fd.askopenfilename(title="Укажите таблицу 410/у-П3 Стационар + Выезд 'Список приема'", initialdir=f"{vars.usr_desktop}",
                                                     filetypes=filetypes)
            if vars.tab_donors_path != '':
                vars.last_folder = get_folder_path(vars.tab_donors_path)
        else:
            vars.tab_donors_path = fd.askopenfilename(title="Укажите таблицу 410/у-П3 Стационар + Выезд 'Список приема'", initialdir=f"{vars.last_folder}",
                                                     filetypes=filetypes)
            if vars.tab_donors_path != '':
                vars.last_folder = get_folder_path(vars.tab_donors_path)

        if vars.tab_donors_path != '':
            if not ch_foo.check_uniq_path(vars.tab_donors_path, 'donors'):
                self.tab_donors_label.configure(text_color='green')
                self.tab_donors_label.configure(text=os.path.basename(vars.tab_donors_path))
            else:
                mes.error('Ошибка уникальности файла', f'Файл {os.path.basename(vars.tab_donors_path)} уже был выбран в другом разделе!')
        else:
            self.tab_donors_label.configure(text_color='red')
            self.tab_donors_label.configure(text='Таблица не выбрана!')

    def check_thread(self, thread):
        if thread.is_alive():
            if CTk.get_appearance_mode() == 'Dark':
                self.wait_label.configure(text='Ожидайте выполнения программы...', text_color='yellow')
            else:
                self.wait_label.configure(text='Ожидайте выполнения программы...', text_color='black')
            self.after(100, lambda: self.check_thread(thread))
        else:
            self.tab_start_button.configure(state='normal')
            self.tab_hospital_button.configure(state='normal')
            self.tab_visit_button.configure(state='normal')
            self.tab_donors_button.configure(state='normal')
            self.wait_label.configure(text='', text_color='green')
            self.show_about_button.configure(state='normal')
            self.show_instruction_button.configure(state='normal')
            self.appearance_mode_option_menu.configure(state='normal')
            self.multidons_out_checkbox.configure(state='normal')
            self.log_out_checkbox.configure(state='normal')

            self.clear_paths_tabs()

    def start_dons(self):
        # Проверка версии программы перед началом
        version_result = ''
        version_return = service.check_version()
        if version_return != True:
            # mes.warning('Проверка версии ПО', f'{version_return}')
            version_result = messagebox.askyesno('Проверка версии',
                                     f'{version_return}\n\nПродолжить?',
                                     icon='warning')
        if version_return != True and version_result or version_return == True:
            result = messagebox.askyesno('Предупреждение перед началом обработки',
                                         'Внимание!\n\nВо время работы программы будут закрыты все документы Excel.\n\nПродолжить?',
                                         icon='warning')
            if result:
                if ch_foo.check_run_excel():
                    while ch_foo.check_run_excel():
                        for proc in psutil.process_iter():
                            if proc.name() == 'EXCEL.EXE':
                                try:
                                    proc.kill()
                                except:
                                    continue
            else:
                mes.error('Отмена обработки',
                          'Операция отменена пользователем!')
                exit()

            if ch_foo.check_tab_paths():
                self.tab_start_button.configure(state='disabled')
                self.tab_hospital_button.configure(state='disabled')
                self.tab_visit_button.configure(state='disabled')
                self.tab_donors_button.configure(state='disabled')
                self.show_about_button.configure(state='disabled')
                self.show_instruction_button.configure(state='disabled')
                self.appearance_mode_option_menu.configure(state='disabled')
                self.multidons_out_checkbox.configure(state='disabled')
                self.log_out_checkbox.configure(state='disabled')
                if self.check_var.get() == 'off':
                    vars.multidons_out = False
                else:
                    vars.multidons_out = True

                if self.check_log_var.get() == 'off':
                    vars.log_out = False
                else:
                    vars.log_out = True

                if ch_foo.check_run_excel():
                    while ch_foo.check_run_excel():
                        for proc in psutil.process_iter():
                            if proc.name() == 'EXCEL.EXE':
                                try:
                                    proc.kill()
                                except:
                                    continue

                thread = threading.Thread(target=data_edit.get_input_data, daemon=True)
                thread.start()
                self.check_thread(thread)
                self.wait_label.configure(text='')
            else:
                self.wait_label.configure(text='')
                mes.error('Ошибка входных данных', 'Вы не указали запрашиваемые таблицы!')
        else:
            mes.error('Отмена обработки', 'Обработка отменена пользователем.')

    def show_about(self):
        print(vars.toplevel_window)
        if vars.toplevel_window is None:
            vars.toplevel_window = about.AboutWin(self)
            print(vars.toplevel_window)

    def show_instruction(self):
        print(vars.toplevel_window)
        if vars.toplevel_window is None:
            vars.toplevel_window = instruction.InstructionWin(self)
            print(vars.toplevel_window)

    def on_close(root):
        root.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", partial(app.on_close))
    app.mainloop()
