# import os
import datetime as DT

# os.system('cls')

# data = "27.03.1970"

def get_num_moon(txt='21.06.2015'):
    """Функция возвращает номер месяца
       :param txt: строка вида: '21.06.2015' """
    date = DT.datetime.strptime(txt, '%d.%m.%Y').date()
    # print(str(date))                  # 2018-08-19
    # print(date.strftime('%Y-%m-%d'))  # 2018-08-19
    return int(date.strftime('%m'))

def get_num_moon_long(txt='2015-06-21 00:00:00'):
    """Функция возвращает номер месяца
       :param txt: строка вида: '2015-06-21 00:00:00' """
    date = DT.datetime.strptime(txt, '%Y-%m-%d %H:%M:%S').date()
    # print(str(date))                  # 2018-08-19
    # print(date.strftime('%Y-%m-%d'))  # 2018-08-19
    return int(date.strftime('%m'))

def get_rus_str_moon(num=3):
    """Функция возвращает название месяца на русском языке
       :param num: целое число int() - номер месяца."""
    moonth = {1:'январь', 2:'февраль', 3:'март', 4:'апрель', 5:'май', 
              6:'июнь', 7:'июль', 8:'август', 9:'сентябрь', 10:'октябрь', 11:'ноябрь', 12:'декабрь'}
    return moonth[num]          


# print(get_rus_str_moon(get_num_moon(data)))