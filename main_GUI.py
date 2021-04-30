#!C:/msys64/mingw64/bin/python.exe
import os
import time

import gi

import Excel_report

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk
from gi.repository import Gdk, GLib
from threading import Timer

delay_in_sec = 2

whatis = lambda obj: print(type(obj), "\n\t" + "\n\t".join(dir(obj)))

week_OPE = [50, 51, 52, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
            27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49,
            50, 51, 52, 53]  # количество недель ОПЭ

def load_lStore(weeks):  # загрузка недель в список
    for week in range(len(weeks)):
        lStore.append([weeks[week]])


def do_pulse():  # формирует бегущую строку в поле вывода, пока формируется отчет
    entry_info.set_progress_pulse_step(0.1)
    while entry_info.get_text() == "Отчет формируется":  #time.time() - now < 60
        entry_info.progress_pulse()
        time.sleep(0.1)
        Gtk.main_iteration_do(False)


def report():  # формирование отчета
    text_return = Excel_report.create_a_report(file_location.get_filename(), save_place.get_filename(),
                                               week_OPE[cb_week.get_active()])
    if text_return != "Отчет успешно сформирован":
        entry_info.set_progress_pulse_step(0)
        entry_info.set_text(text_return)
        entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("red"))
        entry_info.set_progress_fraction(0)
    else:
        entry_info.set_progress_pulse_step(0)
        entry_info.set_progress_fraction(0)
        entry_info.set_text(text_return)
        entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("green"))


class Handler:

    def button_report_clicked_cb(self, button):  # обрабатывает нажатие кнопки "Сформировать отчет"
        if save_place.get_filename() is None:
            entry_info.set_text("Выберите папку для сохранения отчета!")
            save_place.set_filename("C:" + os.path.join(os.environ['HOMEPATH'], 'Desktop'))
            entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("red"))
        else:
            t = Timer(delay_in_sec, report)  # задержка delay_in_sec для формирования отчета, чтобы запустить бегунок do_pulse()
            t.start()  # возвращает None
            entry_info.set_text("Отчет формируется")
            do_pulse()

    def cb_week_changed_cb(self, button):  # обрабатывает выпадающий список
        pass

    def save_place_file_set_cb(self, button):  # обрабатывает выбор папки для сохранения данных
        entry_info.set_text("Папка для отчета: {}".format(save_place.get_filename()))
        entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("green"))

    def file_location_file_set_cb(self, button):  # обрабатывает выбор файла данных
        entry_info.set_text("")
        for name in file_location.get_filename().split("\\"):
            if name == "Отчет ОПЭ.xlsx":
                entry_info.set_text("Файл \"Отчет ОПЭ.xlsx\" успешно выбран")
                entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("green"))
                button_report.set_sensitive(True)
        if entry_info.get_text() != "Файл \"Отчет ОПЭ.xlsx\" успешно выбран":
            entry_info.set_text("Выберите файл \"Отчет ОПЭ.xlsx\"!")
            file_location.set_filename("C:/")
            entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("red"))


abuilder = Gtk.Builder()
abuilder.add_from_file("GUI.glade")
abuilder.connect_signals(Handler())

Window = abuilder.get_object("main_window")
Window.connect("destroy", Gtk.main_quit)

entry_info = abuilder.get_object("entry_info")

lStore = abuilder.get_object("list_store")

cb_week = abuilder.get_object("cb_week")
button_report = abuilder.get_object("button_report")  # кнопка формирования отчета
save_place = abuilder.get_object("save_place")  # путь места сохранения
file_location = abuilder.get_object("file_location")  # путь к файлу Отчет ОПЭ.xlsx
load_lStore(week_OPE)

renderer_text = Gtk.CellRendererText()
cb_week.pack_start(renderer_text, True)
cb_week.add_attribute(renderer_text, "text", 0)
cb_week.set_active(17)

Window.set_title("Формирование отчетов по ОПЭ БКЭУ")
Window.set_icon_from_file("icon.ico")
Window.show_all()
# whatis(entry_info.progress_pulse())
if __name__ == '__main__':
    button_report.set_sensitive(False)  # делает кнопку не чувствительной
    save_place.set_filename(
        "C:" + os.path.join(os.environ['HOMEPATH'], 'Desktop'))  # устанавливает путь по умолчанию "Рабочий стол"
    file_location.set_filename(
        "\\\\files.nipom.org\\res\Razrab-09\Обмен\АИП\\6707-Кузнецк\Тренды\Новокузнецк-2020\Отчет ОПЭ.xlsx")  # устанавливает путь по умолчанию "Рабочий стол"
    for name in file_location.get_filename().split("\\"):
        if name == "Отчет ОПЭ.xlsx":
            entry_info.set_text("Файл \"Отчет ОПЭ.xlsx\" успешно выбран")
            entry_info.modify_fg(Gtk.StateFlags.NORMAL, Gdk.color_parse("green"))
            button_report.set_sensitive(True)
    if not button_report.get_sensitive():
        entry_info.set_text("Выберите файл \"Отчет ОПЭ.xlsx\"")
    Gtk.main()
