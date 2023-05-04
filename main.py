# _*_ coding: utf-8 _*_

from main_window import MainWindow

'''
Модуль запуска программы
'''


def main() -> None:
    '''
    Основная функция
    '''

    try:
        window = MainWindow()
        window.run()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    '''
    Запускать только если это основной модуль
    '''

    main()
