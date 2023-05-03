import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkcalendar import DateEntry


class MainWindow:
    def __init__(self) -> None:
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
        self.root.mainloop()

    def __select_file(self) -> None:
        if self.date_entry_start.get_date() > self.date_entry_end.get_date():
            mb.showerror(
                'Ошибка', f'Дата начала {str(self.date_entry_start.get_date()).replace("-", ".")} больше даты конца {str(self.date_entry_end.get_date())}'.replace("-", "."))
            return

        file_path = fd.askopenfilename()
        print('Выбранный файл: ', file_path)
