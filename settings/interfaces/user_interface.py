import calendar
import sys
import time
from datetime import datetime
import locale
import win32com.client
import PySimpleGUI as sg

from settings.update.update import call_updater, check_version

locale.setlocale(locale.LC_TIME, 'ru')
sg.LOOK_AND_FEEL_TABLE['SamoletTheme'] = {
                                        'BACKGROUND': '#007bfb',
                                        'TEXT': '#FFFFFF',
                                        'INPUT': '#FFFFFF',
                                        'TEXT_INPUT': '#000000',
                                        'SCROLL': '#FFFFFF',
                                        'BUTTON': ('#FFFFFF', '#007bfb'),
                                        'PROGRESS': ('#354d73', '#FFFFFF'),
                                        'BORDER': 1, 'SLIDER_DEPTH': 0,
                                        'PROGRESS_DEPTH': 0, }

def main_panel():
    recentMonth = calendar.month_name[datetime.now().month-1    ]

    MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь',
              'Ноябрь', 'Декабрь']
    sg.theme('SamoletTheme')
    UPD_FRAME = [[
                  sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt', background_color='#007bfb', font='bold')),
                  sg.Push(),
                sg.pin(sg.Button('Обновить', key='upd_btn',  visible=False))]]
    DOC_FRAME = [
                [sg.Text('Месяц', font='bold'),sg.Push(), sg.Combo(MONTHS, default_value=recentMonth, key='--MONTH--')],
                [sg.Text('Ссылка на лист GoogleDoc', font='bold')],
                [sg.InputText(key='ref')],
                [sg.Text('Данные для графиков', font='bold')],
                [sg.Input(key='plot_data'), sg.FileBrowse(button_text='Выбрать')],
                [sg.Text('Папка для сохранения', font='bold')],
                [sg.Input(key='save_path'), sg.FolderBrowse(button_text='Выбрать')]
                ]

    layout = [
        [sg.Frame(layout=UPD_FRAME, title='Обновление', key='--UPD_FRAME--',
                  size=(400, 60))],
        [sg.Frame(layout=DOC_FRAME, title='Выбор документов', size=(400, 250))],
        [sg.OK(button_text='Далее'), sg.Cancel(button_text='Выход')]

    ]
    yeet = sg.Window(f'Сверка БИТ и CRM', layout=layout)
    check, upd_check = False, True
    while True:
        event, values = yeet.read(timeout=10)
        if check:
            upd_check = check_version()
            check = False
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        if event == 'check_upd':
            check = True
        if not upd_check:
            yeet['not_upd_txt'].Update(visible=False)
            yeet['upd_txt'].Update(visible=True)
            yeet['upd_btn'].Update(visible=True)
        if event == 'upd_btn':
            yeet.close()
            call_updater('pocket')
        elif event == 'Далее':
            break
    yeet.close()
    check_values = check_user_values(values)
    if check_values:
        return values
    else:
        check = input_error_panel()
        if check:
            return main_panel()
        else:
            sys.exit()


def check_user_values(user_values):
    keys = ['--MONTH--', 'ref', 'plot_data', 'save_path']
    if any(map(lambda x: user_values[x] == '', keys)):
        return False
    else:
        return True


def input_error_panel():
    event = sg.popup('Ошибка ввода', 'При вводе данных возникла ошибка.\nВы хотите повторить ввод данных?',
                     button_color=('white', '#007bfb'),
                     title='Ошибка', custom_text=('Да', 'Нет'))
    if event == 'Да':
        return True
    else:
        sys.exit()

def end_panel(path):
    event = sg.popup('Сверка завершена\nОткрыть обработанный файл?', background_color='#007bfb',
                     button_color=('white', '#007bfb'),
                     title='Завершение работы', custom_text=('Да', 'Нет'))
    if event == 'Да':
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = True
        Excel.Workbooks.Open(Filename = path)
        time.sleep(5)
        del Excel
    else:
        sys.exit()

def error_panel(exp_desc):
    event = sg.popup_ok(f'При обработке данных возникла следующая ошибка:\n{exp_desc}',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Внутренняя ошибка')
    if event == 'OK':
        sys.exit()