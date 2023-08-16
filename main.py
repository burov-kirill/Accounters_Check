from settings.classes.URLObject import GoogleDocFile
from settings.operations.excel_export import export_to_excel
from settings.operations.ppoint import create_presentation
from settings.functions.read_data import read_old_data
from settings.interfaces.user_interface import main_panel, end_panel, error_panel

try:
    values = main_panel()
    name = 'Отчет по закрытию.xlsx'
    save_path = fr'{values["save_path"]}\{name}'
    data = GoogleDocFile(values['ref'])
    data.tables = read_old_data(data.tables, values['--MONTH--'], values['plot_data'])
    export_to_excel(data, values['--MONTH--'], values['save_path'])
    create_presentation(values['save_path'], values['--MONTH--'], data)
except Exception as exp:
    error_panel(exp)
else:
    end_panel(save_path)


