changes_row = """v1.10 от 01.09.2023
- Исправлен некорректный контроль количества записей во входных таблицах УП-2 и УП-3.

v1.9 от 09.08.2023
- Изменена система именования временных и выходных файлов. Теперь имя выходного файла выглядит всегда единым образом, формата "Список для страхования Город - Месяц в таблице Год.xlsx".
- Исправлены некоторые опечатки в интерфейсе.
- Исправлено размещение элементов интерфейса.
- Скорректированы записи в log-файле.
- Изменена система формирования списка изменений версий программы.
- Реализован выбор необходимости вывода списка доноров с несколькими донациями в текстовый файл рядом с выходным файлом.
- Реализован выбор необходимости вывода ошибок и предупреждений при обработке данных в текстовый файл.

v1.8 от 17.07.2023
- Исправлен вывод исключений и предупреждений в log-файл.
- Реализован механизм ожидания доступа к log-файлу на файловом сервере.
- Добавлена фиксация успешных обработок в общем log-файле.

v1.7 от 13.06.2023
- Переработана структура программы в целях повышения производительности.
- Изменена система измерения времени обработки таблиц и его отображение в статистике по завершению программы.
- Изменено форматирование содержимого log записи.
- Изменено содержание отчета по обработке таблиц.
- Отключено уведомление о существовании файла и его перезаписи.
- Исправлена ошибка в работе программы при неудачной попытке завершения процессов Excel на устройстве.

v1.6 от 05.06.2023
- Стартовая страница при добавлении файла изменена на "Рабочий стол" с папки "Документы".
- Исправлено срабатывание контроля при проверке ФИО в итоговом файле для ФИО, содержащих точки.
- Добавлен механизм проверки версии программы.
- Добавлена фиксация версии программы в серверный log файл.
- Исправлена некорректная очистка глобальных переменных при непрерывной обработке разных входных файлов.

v1.5 от 02.05.2023
- Исправлено ошибочное уведомление о несовпадении количества записей во входных таблицах УП-2 и УП-3.
- В название итогового файла добавлено название города из заголовка таблицы для большего удобства при обработке нескольких списков.
- После успешной обработки, заполненные пути до таблиц автоматически очищаются.

v1.4 от 15.04.2023
- В наименовании итогового файла месяц исправлен с текущего месяца на месяц из заголовка таблицы по последней дате периода выгрузки.
- Исправлена ошибка отображения пустых дополнительных окон.

v 1.3 от 14.04.2023
- Исправлено накладывание главного окна на дополнительные окна при их открытии.
- Сообщение о статусе выполнения на главном окне оформлено в более заметные цвета.
- Наименование выходного файла теперь в формате "Список для страхования - Месяц год".
- Из итогового файла удалены заглавные строки, обозначающие источник данных, для более легкого копирования (Например, "ОЗК", "ВЫЕЗД").
- Создаваемый по завершению файл "Похожие доноры" заменен на "Отчет по обработке таблиц". Теперь файл содержит записи о донорах совершивших более 1 донации и записи о совпадении доноров при обработке.

v 1.2 от 06.03.2023
- Добавлены контроли на периоды выгрузки, на количество записей по таблицам, на обработанные записи.
- Добавлено сохранение зафиксированных ошибок в файл к итоговому результату.
- Добавлен вывод возможных доноров по ФИО и дате рождения, но пропущенных по коду донора в отдельный файл по пути итогового результата.
- Изменена политика назначения имен выходным файлам.
- Добавлена выгрузка ошибок на сервер
- Добавлена строка "Итого" в итоговый файл
- Добавлена очистка памяти при запуске повторной обработки.

v 1.1 от 03.03.2023
- Исправлено полное зависание интерфейса до конца обработки путем добавления "многопоточности".
- Исправлены элементы интерфейса.
- Добавлен раздел с инструкцией об использовании программы.
- Добавлен контроль искомых доноров по глобальному коду донора.

v 1.0 от 02.03.2023
- Программа протестирована в ОИТ и выгружена на файловый сервер."""