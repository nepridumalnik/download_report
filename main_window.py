# _*_ coding: utf-8 _*_

import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkcalendar import DateEntry

import openpyxl

from database import extract_dict

'''
Интерфейс и работа с ним
'''

NUM_CELL: str = 'A1'
FIAS_CELL: str = 'I1'
DESC_CELL: str = 'L1'


class MainWindow:
    '''
    Класс главного окна
    '''

    def __init__(self) -> None:
        '''
        Конструктор
        '''

        self.root = tk.Tk()
        self.root.title('Выбор даты и файла')

        # Создаем фрейм для первой строки
        top_frame = tk.Frame(self.root)
        top_frame.pack()

        # Создаем два элемента ввода даты
        date_label_start = tk.Label(top_frame, text='Дата начала:')
        date_label_start.pack(side=tk.LEFT)
        self.date_entry_start = DateEntry(top_frame, width=12, background='darkblue',
                                          foreground='white', borderwidth=2, locale='ru_RU')
        self.date_entry_start.pack(side=tk.LEFT, padx=5)
        date_label_end = tk.Label(top_frame, text='Дата конца:')
        date_label_end.pack(side=tk.LEFT)
        self.date_entry_end = DateEntry(top_frame, width=12, background='darkblue',
                                        foreground='white', borderwidth=2, locale='ru_RU')
        self.date_entry_end.pack(side=tk.LEFT, padx=5)

        # Создаем кнопку 'Выбрать файл' во второй строке
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10)
        select_button = tk.Button(
            bottom_frame, text='Выбрать файл', command=self.__select_file)
        select_button.pack()

    def run(self) -> None:
        '''
        Запуск основного цикла
        '''

        self.root.mainloop()

    def __is_sheet_valid(self, sheet) -> bool:
        '''
        Проверка валидности листа
        sheet - лист excel
        '''

        if not str(sheet[NUM_CELL].value).startswith('№'):
            return False

        if not str(sheet[FIAS_CELL].value) == 'ФИАС':
            return False

        if not str(sheet[DESC_CELL].value) == 'Расшифровка':
            return False

        return True

    def __clean_sheet(self, worksheet) -> None:
        '''
        Очистить лист
        '''

        for row in worksheet.iter_rows(min_row=2, min_col=11, max_col=12):
            for cell in row:
                cell.value = ''

    def __handle_sheet(self, worksheet, dictionary) -> None:
        '''
        Обработать лист
        worksheet - лист
        dictionary - словарь
        '''

        # Подготовка
        self.__clean_sheet(worksheet)

        # num_rows: int = len(list(worksheet['A']))

        totals: dict = dict()

    def __handle_dict(self, dictionary, workbook) -> None:
        '''
        Обработать словарь
        dictionary - словарь
        path - путь до файла
        '''

        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]

            if not self.__is_sheet_valid(worksheet):
                continue

            self.__handle_sheet(worksheet, dictionary)

    def __select_file(self) -> None:
        '''
        Обработка нажатия кнопки
        '''

        # Если дата конца интервала позже даты начала, выдаём ошибку
        if self.date_entry_start.get_date() > self.date_entry_end.get_date():
            start: str = str(self.date_entry_start.get_date()
                             ).replace('-', '.')
            end: str = str(self.date_entry_end.get_date()).replace('-', '.')
            error_text: str = f'Дата начала {start} больше даты конца {end}'
            mb.showerror('Ошибка', error_text)
            return

        try:
            path: str = fd.askopenfilename()
            workbook = openpyxl.load_workbook(
                path, read_only=False, keep_vba=True)

            start: str = str(self.date_entry_start.get_date()
                             ).replace('-', '.')
            end: str = str(self.date_entry_end.get_date()).replace('-', '.')
            dictionary = extract_dict(start, end)

            self.__handle_dict(dictionary, workbook)

            workbook.save(path)

        except Exception as e:
            mb.showerror('Ошибка', e)
            print(e)
