import customtkinter as CTk
from datetime import datetime
from check_funcs import check_path
from variables import readme_path
import variables as vars



def read_file_instr(path):
    file = open(path, mode="r")
    list_readme = file.read()
    file.close()
    return list_readme


class InstructionWin(CTk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("1300x450+500+300")
        self.minsize(600, 450)
        self.resizable(width=False, height=False)
        self.title('Инструкция к программе')
        self.iconbitmap('logo.ico')
        CTk.deactivate_automatic_dpi_awareness()

        label_text_up = f'Настоящая инструкция осветит основные моменты взаимодействия с программой "Списки страхования"'
        label_text_down = f'КГКУЗ "Красноярский краевой центр крови №1"\n\n2023 - {datetime.now().year}'

        self.label_up = CTk.CTkLabel(self, text=label_text_up, wraplength=550)
        self.label_up.pack(padx=20, pady=10)

        self.readme_box = CTk.CTkTextbox(self, corner_radius=0, wrap='word')
        self.readme_box.pack(padx=20, pady=10, expand=True, fill='both')
        self.create_file()
        self.readme_box.configure(state='disabled')

        self.label_down = CTk.CTkLabel(self, text=label_text_down)
        self.label_down.pack(padx=20, pady=10)

        self.focus()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: self.dismiss())  # перехватываем нажатие на крестик

    def dismiss(self):
        vars.toplevel_window = None
        print(vars.toplevel_window)
        self.grab_release()
        self.destroy()

    def create_file(self):
        if check_path(readme_path):
            list_readme = read_file_instr(readme_path)
            self.readme_box.insert('0.0', list_readme)
        else:
            readme_text = '\tКак работать с программой?\n\n\n' \
                          '1. В АИСТ переходим во вкладку Донорская.' \
                          '\n2. Переходим в раздел Отчеты по донорам > Списки принятых доноров. Формируем три разных ' \
                          'отчета в формате Excel:' \
                          '\n\tа) 410/уП2 За период (в настройках выбираем Стационар).' \
                          '\n\tб) 410/уП2 За период (в настройках выбираем Выезд) (!ПРИ ЕГО НАЛИЧИИ!).' \
                          '\n\tв) 410/уП3 Список принятых доноров (в настройках выбираем И Стационар И Выезд).' \
                          '\n3. В программу в соответствии с названием кнопок, вносим скачанные ранее отчеты.' \
                          '\n4. Закрываем открытые документы Excel (если открыты).' \
                          '\n5. Запускаем работу программы по кнопке "Подготовить список".' \
                          '\n6. Читаем уведомления.' \
                          '\n\t6.1. Если появится сообщение о том, что файл существует, то выберите "Да, заменить".' \
                          '\n7. По завершению программы вы получите итоговое сообщение и откроется папка с ' \
                          'готовым файлом формата:' \
                          '\n\t- "Список для страхования Город - Месяц год.xlsx".' \
                          '\n\nВнимание!' \
                          '\nПри работе программы используются ресурсы жесткого диска и загружаемые файлы.' \
                          ' В связи с этим обращаем ваше внимание на то, что:' \
                          '\n\t- файлы при их обработке открывать не рекомендуется, во избежание ошибок;' \
                          '\n\t- программа может временно не отвечать при работе, закрывать её не нужно, ' \
                          'она просто обрабатывает входящие файлы, ожидайте сообщения о завершении;' \
                          '\n\t- при работе программы, ресурсов жесткого диска может не хватать для других приложений' \
                          ' и они могут также временно не отвечать;' \
                          '\n\t- скорость обработки зависит от конкретного компьютера. (Средняя скорость 35 сек. на ' \
                          'форматирование таблиц, 60 сек на заполнение итогового файла). Общее время не должно ' \
                          'превышать более 3 мин.\n'

            file = open(readme_path, "w+")
            file.write(readme_text)
            file.close()

            list_readme = read_file_instr(readme_path)
            self.readme_box.insert('0.0', list_readme)
