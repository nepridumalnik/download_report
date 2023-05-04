# _*_ coding: utf-8 _*_

import pyodbc

'''
Работа с базой данных
'''

FIAS_INDEX: int = 1
DESCRIPTION_INDEX: int = 2
STATUS_INDEX: int = 3
VALUE_INDEX: int = 4


class FiasData:
    '''
    Данные ФИАСа
    '''

    def __init__(self) -> None:
        '''
        Конструктор
        '''

        # Расшифровка
        self.__descriptions: dict = dict()

    def append(self, description: str, volume: float) -> None:
        '''
        Добавить расшифровку
        description - новая расшифровка
        volume - значение
        '''

        self.__descriptions[description] = volume

    def __iter__(self):
        '''
        Получаем итератор для описания
        '''

        return iter(self.__descriptions.items())

    def __str__(self) -> str:
        '''
        Получить расшифровку
        '''

        output: str = ''

        # Форматируем расшифровку
        for key, val in self:
            output += f'{key}: {val}\n'

        # Удаляем последний символ переноса строки
        if output and output.endswith('\n'):
            output = output[:-1]

        return output


def convert_rows_to_dict(rows) -> dict:
    '''
    Конвертировать строку из СУБД в словарь
    rows - строки
    '''

    dictionary: dict = dict()

    # Извлекаем данные из строк и записываем в словарь
    for row in rows:
        fias = row[FIAS_INDEX]

        # Пропускаем нулевые значения
        if not fias:
            continue

        desc = row[DESCRIPTION_INDEX]
        val = float(row[VALUE_INDEX])
        status = row[STATUS_INDEX]

        # Добавляем ФИАС, если не существует
        if not dictionary.__contains__(fias):
            dictionary[fias] = FiasData()

        # Добавляем расшифровку для ФИАСа
        dictionary[fias].append(f'{desc}.{status} м3', val)

    return dictionary


def extract_dict(data_begin: str, data_end: str) -> dict:
    '''
    Извлечь данные из базы данных в виде словаря, у которого ключи это фиасы, а значения это данные
    data_begin - дата начала
    data_end - дата конца
    '''

    # Конфигурация соединения
    CONNECTION_CONFIGURE: str = 'Driver={SQL Server};Server=10.1.1.7;Uid=dudnik; Pwd=tvk2022dud;'

    # Процедура
    PROCEDURE: str = f'EXEC Get_Obem_Doma @Data_Begin="{data_begin}", @Data_End="{data_end}"'

    # Соединение с СУБД
    con = pyodbc.connect(CONNECTION_CONFIGURE)

    # Курсор для работы с СУБД
    cursor = con.cursor()

    # Вызов процедуры
    rows = cursor.execute(PROCEDURE)

    return convert_rows_to_dict(rows)
