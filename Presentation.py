import Excel_report

import win32com.client
from PIL import ImageGrab
from pptx import Presentation
from pptx.util import Inches, Pt


# whatis = lambda obj: print(type(obj), "\n\t" + "\n\t".join(dir(obj)))

def last_cell_value(sheet_2, ranges):
    value = 0
    last = 0
    for v in sheet_2.Range(ranges).Value:
        if v[0] is None:
            break
        value += v[0]
        last = v[0]
    if ranges == 'B2:B54' or ranges == 'C2:C54':
        value = value - last
    return str(round(value, 1))


def create_table(sheet_2, slide_table):
    # задание параметров будущей таблицы
    rows = 11
    cols = 3
    left = Inches(0.2)
    top = Inches(1)
    width = Inches(10.0)
    height = Inches(4)

    table = slide_table.shapes.add_table(rows, cols, left, top, width, height).table  # создание таблицы
    # задание ширины столбцов
    table.columns[0].width = Inches(5)
    table.columns[1].width = Inches(1.9)
    table.columns[2].width = Inches(2.8)
    # write head cells
    table.cell(0, 0).text = 'Наименование параметра, единица измерения'
    table.cell(0, 1).text = 'Всего с начала\n2-ого этапа ОПЭ'
    table.cell(0, 2).text = 'Примечание'

    # write body cells
    table.cell(1, 0).text = 'Выработка электроэнергии БКЭУш, всего - кВт*ч'
    table.cell(2, 0).text = 'Выработка электроэнергии СМ – кВт*ч'
    table.cell(3, 0).text = 'Выработка электроэнергии ГПЭГ – кВт*ч'
    table.cell(4, 0).text = 'Расход электроэнергии на СН БКЭУш, кВт*ч'
    table.cell(5, 0).text = 'Расход электроэнергии на нужды площадки РН-21-22, кВт*ч'
    table.cell(6, 0).text = 'Выработка тепловой энергии от ГВК – кВт*ч'
    table.cell(7, 0).text = 'Наработка ГПЭГ - моточасы'
    table.cell(8, 0).text = 'Расход газа на ГВК – н.куб.м.'
    table.cell(9, 0).text = 'Расход газа на ГПЭГ – н.куб.м'
    table.cell(10, 0).text = 'Расход газа на БКЭУш, всего – н.куб.м'
    table.cell(7, 2).text = 'Последнее ТО – 26.01.21 – 230 мч'

    table.cell(1, 1).text = last_cell_value(sheet_2, 'D2:D54')
    table.cell(2, 1).text = last_cell_value(sheet_2, 'C2:C54')
    table.cell(3, 1).text = last_cell_value(sheet_2, 'B2:B54')
    table.cell(4, 1).text = last_cell_value(sheet_2, 'F2:F54')
    table.cell(5, 1).text = last_cell_value(sheet_2, 'E2:E54')
    table.cell(6, 1).text = last_cell_value(sheet_2, 'I2:I54')
    table.cell(7, 1).text = last_cell_value(sheet_2, 'G2:G54')
    table.cell(8, 1).text = last_cell_value(sheet_2, 'K2:K54')
    table.cell(9, 1).text = last_cell_value(sheet_2, 'J2:J54')
    table.cell(10, 1).text = last_cell_value(sheet_2, 'H2:H54')

    # Формат шрифта ячеек
    for cell in table.iter_cells():
        cell.text_frame.paragraphs[0].font.size = Pt(12)
        cell.text_frame.paragraphs[0].font.name = "Franklin Gothic Book"
    # Формат шрифта заголовков ячеек первой строки
    table.cell(0, 0).text_frame.paragraphs[0].font.size = Pt(14)
    table.cell(0, 1).text_frame.paragraphs[1].font.size = Pt(14)
    table.cell(0, 1).text_frame.paragraphs[0].font.size = Pt(14)
    table.cell(0, 2).text_frame.paragraphs[0].font.size = Pt(14)
    table.cell(0, 0).text_frame.paragraphs[0].font.bold = True
    table.cell(0, 1).text_frame.paragraphs[1].font.bold = True
    table.cell(0, 1).text_frame.paragraphs[0].font.bold = True
    table.cell(0, 2).text_frame.paragraphs[0].font.bold = True
    table.cell(0, 1).text_frame.paragraphs[1].font.name = "Franklin Gothic Book"
    # Смещение данных относительно левого края ячейки
    table.cell(0, 0).margin_left = Inches(0.4)
    table.cell(0, 1).margin_left = Inches(0.3)
    table.cell(0, 2).margin_left = Inches(0.7)
    table.cell(0, 0).margin_top = Inches(0.15)
    table.cell(0, 2).margin_top = Inches(0.15)
    table.cell(1, 1).margin_left = Inches(0.7)
    table.cell(2, 1).margin_left = Inches(0.7)
    table.cell(3, 1).margin_left = Inches(0.7)
    table.cell(4, 1).margin_left = Inches(0.7)
    table.cell(5, 1).margin_left = Inches(0.7)
    table.cell(6, 1).margin_left = Inches(0.7)
    table.cell(7, 1).margin_left = Inches(0.7)
    table.cell(8, 1).margin_left = Inches(0.7)
    table.cell(9, 1).margin_left = Inches(0.7)
    table.cell(10, 1).margin_left = Inches(0.7)


def make_presentations():  # функция сохранения всех графиков в формате png и создания презентации
    # ------ СОХРАНЕНИЕ ГРАФИКОВ ИЗ ТАБЛИЦЫ EXCEL В ФОРМАТЕ PNG-----------------
    input_file = "//files.nipom.org/res/Razrab-09/Обмен/АИП/6707-Кузнецк/Тренды/Новокузнецк-2020/Графики.xlsx"
    output_image = "C:/Razrab-10/python/ExcelWordIntegration/"

    operation = win32com.client.Dispatch("Excel.Application")
    operation.Visible = 0
    operation.DisplayAlerts = 0

    workbook_2 = operation.Workbooks.Open(input_file)  # открытие таблицы Графики.xlsx
    sheet_2 = operation.Sheets(1)  # присвоение переменной sheet_2 параметров первого листа Графики.xlsx

    # создание рисунков из графиков и их сохранение по указанному пути в формате png
    for x, chart in enumerate(sheet_2.Shapes):
        chart.Copy()
        image = ImageGrab.grabclipboard()
        image.save(output_image + "{}.png".format(x), 'png')
        pass

    # установление даты начала и конца ОПЭ из таблицы Графики.xlsx
    date_start = sheet_2.Cells(2, 15).Value.strftime('%d.%m.%Y')
    date_end = 0
    for v in sheet_2.Range('P2:P54').Value:
        if v[0] is None:
            break
        date_end = v[0].strftime('%d.%m.%Y')

    # присвоение переменным наименований рисунков и их пути расположения
    img_path_0 = 'C:/Razrab-10/python/ExcelWordIntegration/0.png'
    img_path_1 = 'C:/Razrab-10/python/ExcelWordIntegration/1.png'
    img_path_2 = 'C:/Razrab-10/python/ExcelWordIntegration/2.png'
    img_path_3 = 'C:/Razrab-10/python/ExcelWordIntegration/3.png'
    img_path_4 = 'C:/Razrab-10/python/ExcelWordIntegration/4.png'
    img_path_5 = 'C:/Razrab-10/python/ExcelWordIntegration/5.png'
    img_path_6 = 'C:/Razrab-10/python/ExcelWordIntegration/6.png'
    img_path_7 = 'C:/Razrab-10/python/ExcelWordIntegration/7.png'

    img_path_all = [img_path_0, img_path_1, img_path_2, img_path_3, img_path_4, img_path_5, img_path_6,
                    img_path_7]  # создание списка картинок для слайдов с картинками

    title_name = "Итоги ОПЭ БКЭУ-СМ/ГПЭГ/ГВК в ООО «Газпром добыча Кузнецк»\nна площадке РН-21-22\nза период с {} по {} г.".format(
        date_start,
        date_end)  # создание наименования заголовка заглавного слайда с автоматической вставкой дат начала и конца текущего ОПЭ
    title_table = "СВОДНЫЙ ОТЧЕТ"  # создание наименования заголовка слайда с таблицей
    # создание наименований заголовков слайдов с картинками
    title_0 = "ДИАГРАММА ИЗМЕРЯЕМЫХ ПРАМЕТРОВ ПО НЕДЕЛЯМ"
    title_1 = "ВЫРАБОТКА ЭЛЕКТРОЭНЕРГИИ"
    title_2 = "ПОТРЕБЛЕНИЕ ЭЛЕКТРОЭНЕРГИИ"
    title_3 = "ПОТРЕБЛЕНИЕ ГАЗА"
    title_4 = "ЭФФЕКТИВНОСТЬ ВЫРАБОТКИ ЭЛЕКТРОЭНЕРГИИ БКЭУ"
    title_5 = "ВЫРАБОТКА ТЕПЛОВОЙ ЭНЕРГИИ"
    title_6 = "ВЫРАБОТКА ЗА ОТЧЕТНЫЙ ПЕРИОД"
    title_7 = "ТЕМПЕРАТУРА ВНУТРИ БКЭУ"

    title_all = [title_0, title_1, title_2, title_3, title_4, title_5, title_6,
                 title_7]  # создание списка заголовков слайдов с картинками

    prs = Presentation("temp_prs.pptx")  # открытие шаблона презентации
    blank_slide_layout = prs.slide_layouts[0]  # задание шаблона слайда из конструктора
    blank_slide_name = prs.slide_layouts[1]  # задание шаблона слайда из конструктора
    blank_slide_table = prs.slide_layouts[0]  # задание шаблона слайда из конструктора

    slide_name = prs.slides.add_slide(blank_slide_name)  # создание заглавного слайда
    slide_table = prs.slides.add_slide(blank_slide_table)  # создание слайда с таблицей
    slide_0 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_1 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_2 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_3 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_4 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_5 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_6 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой
    slide_7 = prs.slides.add_slide(blank_slide_layout)  # создание слайда с картинкой

    slide_all = [slide_0, slide_1, slide_2, slide_3, slide_4, slide_5, slide_6,
                 slide_7]  # создание списка слайдов с картинками
    # создание заголовка заглавному слайду
    title = slide_name.shapes.title
    title.text = title_name
    # создание заголовка слайду с таблицей
    title = slide_table.shapes.title
    title.text = title_table

    for i in range(0, 8):  # создание заголовков слайдов презентаций с картинками
        title = slide_all[i].shapes.title
        title.text = title_all[i]

    create_table(sheet_2, slide_table)  # создание слайда с таблицей

    # задание параметров для построения слайда 0 с картинкой
    top = Inches(1.5)
    left = Inches(0.1)
    height = Inches(3.3)

    for i in range(0, 8):  # создание слайдов презентаций с картинками
        if i > 0:  # изменение размеров картинок, начиная с 1-й
            top = Inches(1)
            left = Inches(0.5)
            height = Inches(4.5)
        pic = slide_all[i].shapes.add_picture(img_path_all[i], left, top, height=height)

    # whatis(sheet_2)
    prs.save("C:" + Excel_report.os.path.join(Excel_report.os.environ['HOMEPATH'],
                                              'Desktop') + '\\Промежуточные итоги ОПЭ БКЭУ.pptx')  # сохранение презентации на рабочем столе

    workbook_2.Close(True)  # закрытие рабочей книги
    operation.Quit()  # окончание работы с COM командами
    # ----------------------------------------------------------------------------
    # whatis(sheet_1.values)
