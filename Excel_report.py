import openpyxl as xl
from openpyxl.chart import LineChart, Reference, BarChart, ScatterChart, PieChart, RadarChart
from openpyxl import Workbook, drawing
from openpyxl.chart.axis import TextAxis
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.series import Series
from openpyxl.chart.text import RichText
from openpyxl.descriptors import Typed
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
from openpyxl.styles import Font, Alignment, Side, Border, PatternFill
import openpyxl.styles.numbers

import win32com.client
import PIL
from PIL import ImageGrab, Image
import os
import sys

from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import random
import datetime
import matplotlib.pyplot as plt

# whatis = lambda obj: print(type(obj), "\n\t" + "\n\t".join(dir(obj)))

from openpyxl.styles.borders import BORDER_THIN

workbook = xl.load_workbook('Отчет ОПЭ.xlsx',
                            data_only=True)  # открываем базу с отчетами, data_only=True - только данные (без формул)
sheet_1 = workbook.active  # выбираем активный лист или sheet_1 = workbook['Отчет для АО НИПОМ']  # выбираем нужный лист
wb = Workbook()  # создаем рабочую книгу, в которую будем сохранять данные из workbook
ws = wb.active  # создаем рабочий лист
ws.sheet_properties.tabColor = "1072BA"  # задаем цвет вкладки
ws.title = "Графики"  # задаем имя вкладки

# задаем наименование столбцов
ws['A1'] = "Неделя"
ws['B1'] = "Wгпэг,\nкВт*ч"
ws['C1'] = "Wсм,\nкВт*ч"
ws['D1'] = "Wсумм,\nкВт*ч"
ws['E1'] = "Wпотр,\nкВт*ч"
ws['F1'] = "Wсн,\nкВт*ч"
ws['G1'] = "Моточасы\n(ГПЭГ)"
ws['H1'] = "Vст (Газ),\nм^3"
ws['I1'] = "Qкотла,\nкВт*ч"
ws['J1'] = "Vгпэг (Газ),\nм^3"
ws['K1'] = "Vкотел (Газ),\nм^3"
ws['L1'] = "Vгпэг/Wсумм,\nм3/кВт*ч"

i = 2  # переменная для итерации строк в последующем цикле for
# for row in range(10073, sheet_1.max_row + 1):  # цикл по строкам, начиная с нужной

week_OPE = [50, 51, 52, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
            27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49,
            50]  # количество недель ОПЭ

for row in range(5034, sheet_1.max_row + 1):  # цикл по строкам с данными, начиная с 50-й недели

    if sheet_1.cell(row, 46).value == week_OPE[i - 2] and sheet_1.cell(row,
                                                                       47).value == 0:  # если значение в ячейке равно номеру недели из списка и значение расхода не равно нулю
        v_kotel = ws.cell(i, 11)  # создаем переменную расхода котла без учета времени работы ГПЭГ
        if v_kotel.value is None:  # проверка на наличии в ячейке значения
            v_kotel.value = 0
        v_kotel.value += sheet_1.cell(row, 48).value  # суммирование значений в ячейках
        if ws.cell(i, 13).value is None:  # проверка на наличии в ячейке значения
            ws.cell(i, 13).value = 0
        ws.cell(i, 13).value += 1  # суммирование количества ячеек

    if sheet_1.cell(row, 1).value is not None:  # проверяем пустая ячейка первого столбца или нет
        print(sheet_1.cell(row, 1).value)
        week_cell = ws.cell(i, 1)  # создаем переменную строк для столбца с номерами недель
        power_GPEG = ws.cell(i, 2)  # создаем переменную строк для столбца с выработанной мощности ГПЭГ
        power_Sun = ws.cell(i, 3)  # создаем переменную строк для столбца с выработанной мощности солнечного модуля
        power_Sum = ws.cell(i, 4)  # создаем переменную строк для столбца с выработанной суммарной мощностью
        power_Potr = ws.cell(i, 5)  # создаем переменную строк для столбца с потребленной нагрузкой мощностью
        power_SN = ws.cell(i, 6)  # создаем переменную строк для столбца с собственными нуждами
        mototime_GPEG = ws.cell(i, 7)  # создаем переменную строк для столбца с моточасами ГПЭГ
        v_Sum = ws.cell(i, 8)  # создаем переменную строк для столбца с общим потребленным объемом газа

        week_cell.value = sheet_1.cell(row, 1).value  # присваиваем переменной значение из базовой таблицы
        week_cell.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[1]  # формат ячейки 0
        power_GPEG.value = sheet_1.cell(row, 29).value  # присваиваем переменной значение из базовой таблицы
        power_GPEG.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        power_Sun.value = sheet_1.cell(row, 30).value  # присваиваем переменной значение из базовой таблицы
        power_Sun.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        power_Sum.value = sheet_1.cell(row, 31).value  # присваиваем переменной значение из базовой таблицы
        power_Sum.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        power_Potr.value = sheet_1.cell(row, 32).value  # присваиваем переменной значение из базовой таблицы
        power_Potr.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        power_SN.value = sheet_1.cell(row, 33).value  # присваиваем переменной значение из базовой таблицы
        power_SN.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        mototime_GPEG.value = sheet_1.cell(row, 34).value  # присваиваем переменной значение из базовой таблицы
        mototime_GPEG.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        v_Sum.value = sheet_1.cell(row, 35).value  # присваиваем переменной значение из базовой таблицы
        v_Sum.number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00

        i += 1
for i in range(2, ws.max_row + 1):
    if ws.cell(i, 1).value is not None:
        ws.cell(i, 11).value = ws.cell(i, 11).value / (
                ws.cell(i, 13).value / 30) * 168  # столбец с данными расхода газа котлом
        ws.cell(i, 11).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        ws.cell(i, 10).value = ws.cell(i, 8).value - ws.cell(i, 11).value  # столбец с данными расхода газа ГПЭГ
        ws.cell(i, 10).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        ws.cell(i, 9).value = ws.cell(i, 11).value * 9.5  # столбец с данными выработки тепловой энергии котла
        ws.cell(i, 9).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[1]  # формат ячейки 0
        ws.cell(i, 12).value = ws.cell(i, 10).value / ws.cell(i,
                                                              4).value  # расчет эффективности выработки БКЭУ, выраженной через Vгпэг / Wсумм
        ws.cell(i, 12).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[2]  # формат ячейки 0.00
        ws.cell(i, 13).value = None  # удаляем временное значение количества ячеек
    else:
        ws.cell(i, 11).value = None  # удаляем значения выходящие за пределы расчетных недель
        ws.cell(i, 13).value = None  # удаляем значения выходящие за пределы расчетных недель

# задаем ширину столбцов
ws.column_dimensions['A'].width = 10
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 10
ws.column_dimensions['D'].width = 10
ws.column_dimensions['E'].width = 10
ws.column_dimensions['F'].width = 10
ws.column_dimensions['G'].width = 12
ws.column_dimensions['H'].width = 10
ws.column_dimensions['I'].width = 10
ws.column_dimensions['J'].width = 12
ws.column_dimensions['K'].width = 14
ws.column_dimensions['L'].width = 15

thin_border = Border(  # выделение границ ячеек
    left=Side(border_style=BORDER_THIN, color='00000000'),
    right=Side(border_style=BORDER_THIN, color='00000000'),
    top=Side(border_style=BORDER_THIN, color='00000000'),
    bottom=Side(border_style=BORDER_THIN, color='00000000')
)
# цикл для задания ячейкам заголовков свойств
for row in ws.iter_cols(min_col=1, max_col=12, min_row=1, max_row=1):
    for cel in row:
        cel.font = Font(size=12, bold=True)  # размер шрифта и жирное выделение
        cel.alignment = Alignment(horizontal="center", vertical="center",
                                  wrapText=True)  # выравнивание по центру и разрешение переноса строк
        cel.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
# цикл для выделения границ ячеек
for row in ws.iter_cols(min_col=1, max_col=12, min_row=1, max_row=ws.max_row):
    for cel in row:
        cel.border = thin_border

# построение графиков
# график "ДИАГРАММА ИЗМЕРЯЕМЫХ ПРАМЕТРОВ ПО НЕДЕЛЯМ"

cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=2, max_col=8)
# chart = LineChart()
chart = BarChart()
chart.y_axis.title = 'Параметры'
chart.x_axis.title = 'Недели'
chart.height = 10
chart.width = 30
chart.add_data(values, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "A{}".format(ws.max_row + 2))

# график выработки ГПЭГ, СМ и общий
ch1 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=2, max_col=4)
ch1.title = "ВЫРАБОТКА ЭЛЕКТРОЭНЕРГИИ"  # заголовок
ch1.style = 13  # шрифт
ch1.height = 10  # высота
ch1.width = 20  # ширина
ch1.x_axis.title = 'Недели'  # подпись оси х
ch1.y_axis.title = 'кВт*ч'  # подпись оси у
ch1.legend.position = 'r'  # позиция подписей данных справа
ch1.add_data(values, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch1.set_categories(cats)  # загрузка подписи оси х
ch1.series[0].graphicalProperties.line.solidFill = "85C1E9"  # цвет синий
ch1.series[1].graphicalProperties.line.solidFill = "F7DC6F"  # цвет желтый
ch1.series[2].graphicalProperties.line.solidFill = "EC7063"  # цвет красный
ch1.series[0].graphicalProperties.solidFill = "85C1E9"  # цвет синий
ch1.series[1].graphicalProperties.solidFill = "F7DC6F"  # цвет желтый
ch1.series[2].graphicalProperties.solidFill = "EC7063"  # цвет красный
ws.add_chart(ch1, "A{}".format(ws.max_row + 22))  # загрузка графика в ячейку

# график ПОТРЕБЛЕНИЕ ЭЛЕКТРОЭНЕРГИИ
ch2 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=4, max_col=6)
ch2.title = "ПОТРЕБЛЕНИЕ ЭЛЕКТРОЭНЕРГИИ"  # заголовок
ch2.style = 13  # шрифт
ch2.height = 10  # высота
ch2.width = 20  # ширина
ch2.x_axis.title = 'Недели'  # подпись оси х
ch2.y_axis.title = 'кВт*ч'  # подпись оси у
ch2.legend.position = 'r'  # позиция подписей данных справа
ch2.add_data(values, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch2.set_categories(cats)  # загрузка подписи оси х
ch2.series[0].graphicalProperties.line.solidFill = "EC7063"  # цвет линии красный
ch2.series[1].graphicalProperties.line.solidFill = "F7DC6F"  # цвет линии желтый
ch2.series[2].graphicalProperties.line.solidFill = "85C1E9"  # цвет линии синий
ch2.series[0].graphicalProperties.solidFill = "EC7063"  # цвет заливки красный
ch2.series[1].graphicalProperties.solidFill = "F7DC6F"  # цвет заливки желтый
ch2.series[2].graphicalProperties.solidFill = "85C1E9"  # цвет заливки синий
ws.add_chart(ch2, "A{}".format(ws.max_row + 42))  # загрузка графика в ячейку

# график ПОТРЕБЛЕНИЕ ГАЗА
ch3 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=10, max_col=11)
ch3.title = "ПОТРЕБЛЕНИЕ ГАЗА"  # заголовок
ch3.style = 13  # шрифт
ch3.height = 10  # высота
ch3.width = 20  # ширина
ch3.x_axis.title = 'Недели'  # подпись оси х
ch3.y_axis.title = 'м3'  # подпись оси у
ch3.legend.position = 'r'  # позиция подписей данных справа
ch3.add_data(values, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch3.set_categories(cats)  # загрузка подписи оси х
ch3.series[0].graphicalProperties.line.solidFill = "EC7063"  # цвет линии красный
ch3.series[1].graphicalProperties.line.solidFill = "85C1E9"  # цвет линии синий
ch3.series[0].graphicalProperties.solidFill = "EC7063"  # цвет заливки красный
ch3.series[1].graphicalProperties.solidFill = "85C1E9"  # цвет заливки синий

ch31 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values1 = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=8, max_col=8)
ch31.title = "ПОТРЕБЛЕНИЕ ГАЗА"  # заголовок
ch31.style = 13  # шрифт
ch31.height = 10  # высота
ch31.width = 20  # ширина
ch31.x_axis.title = 'Недели'  # подпись оси х
ch31.y_axis.title = 'м3'  # подпись оси у
ch31.legend.position = 'r'  # позиция подписей данных справа
ch31.add_data(values1, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch31.set_categories(cats)  # загрузка подписи оси х

ch3 += ch31
ws.add_chart(ch3, "A{}".format(ws.max_row + 62))  # загрузка графика в ячейку

# график ЭФФЕКТИВНОСТЬ ВЫРАБОТКИ ЭЛЕКТРОЭНЕРГИИ БКЭУ
ch4 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=12, max_col=12)
ch4.title = "ЭФФЕКТИВНОСТЬ ВЫРАБОТКИ ЭЛЕКТРОЭНЕРГИИ БКЭУ"  # заголовок
ch4.style = 13  # шрифт
ch4.height = 10  # высота
ch4.width = 20  # ширина
ch4.x_axis.title = 'Недели'  # подпись оси х
ch4.y_axis.title = 'м3/кВт*ч'  # подпись оси у
ch4.legend.position = 'r'  # позиция подписей данных справа
ch4.add_data(values, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch4.set_categories(cats)  # загрузка подписи оси х
ch4.series[0].graphicalProperties.line.solidFill = "28B463"  # цвет линии зеленый
ch4.series[0].graphicalProperties.solidFill = "28B463"  # цвет заливки зеленый

ws.add_chart(ch4, "A{}".format(ws.max_row + 82))  # загрузка графика в ячейку

# график ВЫРАБОТКА ТЕПЛОВОЙ ЭНЕРГИИ
ch5 = LineChart()
cats = Reference(ws, min_row=2, max_row=ws.max_row-1, min_col=1, max_col=1)
values = Reference(ws, min_row=1, max_row=ws.max_row-1, min_col=9, max_col=9)
ch5.title = "ВЫРАБОТКА ТЕПЛОВОЙ ЭНЕРГИИ"  # заголовок
ch5.style = 13  # шрифт
ch5.height = 10  # высота
ch5.width = 20  # ширина
ch5.x_axis.title = 'Недели'  # подпись оси х
ch5.y_axis.title = 'кВт*ч'  # подпись оси у
ch5.legend.position = 'r'  # позиция подписей данных справа
ch5.add_data(values, titles_from_data=True)  # загрузка данных с заголовками столбцов
ch5.set_categories(cats)  # загрузка подписи оси х
ch5.series[0].graphicalProperties.line.solidFill = "C0392B"  # цвет линии красный
ch5.series[0].graphicalProperties.solidFill = "C0392B"  # цвет заливки красный

ws.add_chart(ch5, "A{}".format(ws.max_row + 102))  # загрузка графика в ячейку

# расчет суммы выработки ГПЭГ и СМ для построения общей диаграммы
ws.cell(ws.max_row, 2).value = 0
ws.cell(ws.max_row, 3).value = 0
for i in range(2, ws.max_row+1):
    if ws.cell(i, 1).value is None:
        ws.cell(i, 1).value = "Итого:"
    if ws.cell(i, 2).value is not None:
        ws.cell(ws.max_row, 2).value += ws.cell(i, 2).value
        ws.cell(ws.max_row, 2).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[1]
    if ws.cell(i, 3).value is not None:
        ws.cell(ws.max_row, 3).value += ws.cell(i, 3).value
        ws.cell(ws.max_row, 3).number_format = openpyxl.styles.numbers.BUILTIN_FORMATS[1]

# построение общей диаграммы выработки ГПЭГ и СМ за отчетный период
chart_itog = PieChart()
labels = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=1)
data = Reference(ws, min_col=2, max_col=3, min_row=ws.max_row, max_row=ws.max_row)
chart_itog.title = "Выработка за отчетный период {} недель".format(ws.max_row-2)
chart_itog.style = 13  # шрифт
chart_itog.height = 10  # высота
chart_itog.width = 20  # ширина
chart_itog.add_data(data, titles_from_data=True)  # загрузка данных с заголовками столбцов
chart_itog.set_categories(labels)  # загрузка подписи оси х
slice = DataPoint(idx=0, explosion=20)  # разделение пирога и сдвиг на 20 пунктов
chart_itog.series[0].data_points = [slice]  # применение сдвига к первому значению
ws.add_chart(chart_itog, "A{}".format(ws.max_row + 122))

wb.save('Test.xlsx')  # сохранение таблицы

# whatis()
