# _*_ coding: utf-8 _*_

import pyodbc

'''
Работа с базой данных
'''


class FiasData:
    '''
    Данные ФИАСа
    '''

    def __init__(self) -> None:
        '''
        Конструктор
        '''

        # Расшифровка
        self.descriptions: dict = dict()

    def append(self, description: str, volume: float) -> None:
        '''
        Добавить расшифровку
        description - новая расшифровка
        volume - значение
        '''

        self.descriptions[description] = volume

    def __iter__(self):
        '''
        Получаем итератор для описания
        '''

        return iter(self.descriptions.items())

    def __str__(self) -> str:
        '''
        Получить расшифровку
        '''

        output: str = ''

        for key, val in self:
            output += f'{key}: {val}\n'

        # Удаляем последний символ переноса строки
        if output and output.endswith('\n'):
            output = output[:-1]

        return output


def extract_value(input: str) -> float:
    '''
    Извлечь кубометры из input
    input - строка в формате Decimal('30.000')
    '''

    start_len: int = 9
    end_len: int = 2

    tmp: str = input[start_len: -end_len]
    value: float = float(tmp)

    return value


def convert_rows_to_dict(rows) -> dict:
    '''
    Конвертировать строку из СУБД в словарь
    rows - строки
    '''

    FIAS_INDEX: int = 1
    DESCRIPTION_INDEX: int = 2
    VALUE_INDEX: int = 4

    dictionary: dict = dict()

    for row in rows:
        fias = row[FIAS_INDEX]

        if not fias:
            continue

        desc = row[DESCRIPTION_INDEX]
        val = float(row[VALUE_INDEX])

        if not dictionary.__contains__(fias):
            dictionary[fias] = FiasData()

        dictionary[fias].append(desc, val)

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
