from main_window import MainWindow

import pyodbc
import openpyxl

CONNECTION: str = 'Driver={SQL Server};Server=10.1.1.7;Uid=dudnik; Pwd=tvk2022dud;'


XLSM_FILE: str = 'C:/Users/nepri/Рабочий стол/Водный баланс в. 2.xlsm'


def is_sheet_valid(sheet) -> bool:
    return sheet['A1'].value == '№\nп/п'


def connect() -> None:
    # con = pyodbc.connect(CONNECTION)

    # cursor = con.cursor()
    # for row in cursor.tables():
    #     print(row.table_name)

    # data_begin: str = '01.09.2022'
    # data_end: str = '02.09.2022'
    # call: str = f'EXEC Get_Obem_Doma @Data_Begin="{data_begin}", @Data_End="{data_end}"'

    # rows = cursor.execute(call)
    # with open(file='file.txt', mode='w') as f:
    #     for row in rows:
    #         f.write(str(row))
    #         f.write('\n')
    #         print(row)

    workbook = openpyxl.load_workbook(XLSM_FILE)
    worksheet = workbook.active

    # for row in worksheet.iter_rows(min_row=1, max_row=1, max_col=10):
    #     for cell in row:
    #         print(cell.value)

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]

        if not is_sheet_valid(worksheet):
            continue

        print(f'Sheet \'{sheet_name}\' is valid')


if __name__ == '__main__':
    try:
        window = MainWindow()
        window.run()

        # connect()
    except Exception as e:
        print(e)
