# Создаётся старый формат .xls - неудачная попытка... 
# coding: utf8
import xlwt
from xlwt import Workbook
from my_lib import my_func

rep_st_row = 0  # report_start_row - стартовая позиция строки для данных в report.xls

book_save = xlwt.Workbook('utf8')  # Создаем книгу
font_zag1 = xlwt.easyxf('font: height 440,name Cambria,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')  # Создаем шрифт для заголовка
font = xlwt.easyxf('font: height 240,name Cambria,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')  # Создаем шрифт для простого текста
font_bold = xlwt.easyxf('font: height 240,name Cambria,colour_index black, bold on,\
    italic off; align: wrap on, vert top, horiz center;\
    pattern: pattern solid, fore_colour white;')  # Создаем жирный шрифт для заголовков

sheet_save = book_save.add_sheet('Лист1')  # Добавляем лист
sheet_save.row(1).height = 2500  # Высота строки
sheet_save.col(0).width = 10000  # Ширина колонки
sheet_save.write(1,0,'Посетители веб-сайта', font_zag1)  # Заполняем ячейку (Строка, Колонка, Текст, Шрифт)
sheet_save.write_merge(3,3,2,14,"количество посещений", font_bold) # Объединение ячеек - start_row, end_row, start_col, end_col
sheet_save.write(4, 0,'Браузер', font)
sheet_save.write(4, 1, 'Тренд', font)
sheet_save.row(4).height = 250  # Высота строки
for i in range(0, 12):
    sheet_save.row(4).height = 250  # Высота строки
    sheet_save.col(i+2).width = 3000  # Ширина колонки
    sheet_save.write(4, i+2, my_func.get_rus_str_moon(i+1), font_bold)
book_save.save('.\\hw6\\report.xls')  # Сохраняем в файл
rep_st_row = 5
