import re
import sys
from collections import defaultdict, namedtuple
from copy import copy
from datetime import datetime
import datetime as dt
from io import BytesIO
from logs import log
import numpy as np
import pandas as pd
import requests

from settings.interfaces.user_interface import error_panel


class GoogleDocFile:
    HUMAN_DATES = {1: 'января',
                   2: 'февраля',
                   3: 'марта',
                   4: 'апреля',
                   5: 'мая',
                   6: 'июня',
                   7: 'июля',
                   8: 'августа',
                   9: 'сентября',
                  10: 'октября',
                  11: 'ноября',
                  12: 'декабря'
    }
    CURRENT_YEAR = 2023
    KEYWORDS = ['Организация', 'Закрываемость', 'Отв.', 'Главный бухгалтер', 'Бухгалтер по первичной документации', 'Ответственные']
    COLUMNS = ['Организация', 'Закрываемость', 'Ответственные']
    DATE_DICT = dict()
    LAST_COLUMN = 'Дата итогового закрытия'
    SUCCEED_TABLE = dict()
    AVERAGE_POINT_DICT = defaultdict(list)
    def __init__(self, url):
        log.info('Считывание файла Google Doc')
        self.raw_url = url
        self.url = self.edit_url()
        self.raw_report = self.filter_raw_data(self.get_data())
        self.report = self.edit_raw_report()
        self.tables = self.create_tables()

    def edit_url(self):
        log.info('Редактирование ссылки')
        export_param = 'export?format=xlsx'
        first_part = self.raw_url[:self.raw_url.rfind('/')]
        second_part = self.raw_url[self.raw_url.rfind('/')+1:]
        gid_part = re.match(r'.+(gid=\d+).*', second_part)
        if gid_part != None and len(gid_part.groups()) == 1:
            param = f'{export_param}&{gid_part.groups()[0]}'
            correct_url = first_part + '/' + param
            return correct_url
        else:
            error_panel('Некорректный URL адрес')

    def get_data(self):
        response = requests.get(self.url)
        if response.status_code == 200:
            data = response.content
            return pd.read_excel(BytesIO(data))
        else:
            error_panel('Невозможно получить ответ от URL адреса')

    def filter_raw_data(self, raw_frame):
        columns = []
        list_idx = []
        for keyword in self.KEYWORDS:
            for column in raw_frame.columns:
                if keyword in list(raw_frame[column]):
                    columns.append(column)
                    list_idx.append(list(raw_frame[column]).index(keyword))
        if all(map(lambda x: x == list_idx[0], list_idx)):
            idx = list_idx[0]
            self.DATE_DICT = self.create_date_dict(idx, raw_frame)
            # date_dict = self.create_date_dict(idx, raw_frame)
            columns = self.extend_list(raw_frame.columns, columns)
            raw_frame = raw_frame[columns]
            raw_frame.columns = raw_frame.loc[idx]
            raw_frame.drop(index=list(range(idx + 1)), inplace=True)
            self.LAST_COLUMN = raw_frame.columns[-1]
            if self.LAST_COLUMN == 'Дата итогового закрытия':
                drop_idx = raw_frame[raw_frame[self.LAST_COLUMN] == 'н/а'].index
                raw_frame.drop(drop_idx, inplace=True)
                # Обработка ошибок, может быть строка или пропуск на месте ячейки
                try:
                    raw_frame['Ответственные'] = raw_frame['Главный бухгалтер'] + '\n' + raw_frame[
                        'Бухгалтер по первичной документации']
                    raw_frame.drop(['Главный бухгалтер', 'Бухгалтер по первичной документации'], axis=1, inplace=True)
                except KeyError as key_err:
                    raw_frame['Ответственные'] = raw_frame['Отв.']
                    raw_frame.drop(['Отв.'], axis=1, inplace=True)
                raw_frame = self.edit_data_cells(raw_frame)
        return raw_frame

    def extend_list(self, raw_columns, filtered_columns):
        raw_date_columns = list(filter(lambda x: re.match(r'\d{4}-\d{2}-\d{2}',
                                                          str(x).split(' ')[0]) != None, raw_columns))
        filtered_columns.extend(raw_date_columns)
        return filtered_columns

    def create_date_dict(self, idx, raw_frame):
        date_dict = {}
        for col in raw_frame.columns:
            if re.match(r'\d{4}-\d{2}-\d{2}', str(col).split(' ')[0]) != None:
                date_dict[raw_frame[col][idx]] = datetime.strptime(str(col).split(" ")[0], '%Y-%m-%d').date()

        return date_dict

    def edit_raw_report(self):
        log.info('Оформление главной таблицы')
        report = copy(self.raw_report)
        report = self.date_coding(self.LAST_COLUMN, report)
        return report

    def edit_data_cells(self, report):
        for column in report.columns:
            if column not in self.KEYWORDS:
                # report[column] = report[column]).apply(self.convert_to_data, args=[self.CURRENT_YEAR])
                report[column] = pd.DataFrame(report[column]).apply(self.convert_to_data, args=[self.CURRENT_YEAR, column], axis = 1)
        return report

    @staticmethod
    def convert_to_data(row, current_year, column):
        string = row[column]
        idx = str(row)[str(row).find('Name'):str(row).find('dtype')-2].split(': ')[1]
        if re.match(r'\d{4}-\d{2}-\d{2}', str(string).split(' ')[0]) != None:
            closed_date = datetime.strptime(str(string).split(" ")[0], '%Y-%m-%d').date()
            if closed_date.year != current_year:
                log.info(f'В строке {idx} столбца {column} обнаружена некорректная дата')
                closed_date = closed_date.replace(year=current_year)
                return closed_date
            else:
                return closed_date
        elif string == 'н/а':
            return ''
        else:
            return string

    def date_coding(self, last_column, report):
        for column in report:
            if column not in self.KEYWORDS and column != last_column:
                report[column] = report[column].apply(self.set_point, args=[column, self.DATE_DICT])
            elif column == last_column:
                report[column] = report[column].apply(self.set_point, args=[column, self.DATE_DICT, True])
        return report
    @staticmethod
    def set_point(string, column, DATE_DICT, is_last_col = False):
        if isinstance(string, dt.date):
            check = DATE_DICT[column] - string
            if check.days == 0:
                return 5 if is_last_col else 2
            elif check.days > 0:
                return 5+check.days if is_last_col else 3
            else:
                return max([5-check.days, 1]) if is_last_col else 1
        elif string == 'Участок не закрыт':
            return 0

        else:
            return ''

    def count_point(self, opt = True):
        result_list = []
        self.report.reset_index(inplace=True)
        for idx in self.report.index:
            count, summary = 0, 0
            for col in self.DATE_DICT.keys():
                if str(self.report.iloc[idx][col]).isdigit() and col != self.LAST_COLUMN:
                    if opt:
                        self.AVERAGE_POINT_DICT[col].append(self.report.iloc[idx][col])
                    count+=1
                if str(self.report.iloc[idx][col]).isdigit():
                    summary+=self.report.iloc[idx][col]
            if opt:
                self.SUCCEED_TABLE[self.report['Организация'][idx]] = (self.report['Закрываемость'][idx], summary/(count*2+5))
            result_list.append(summary if opt else count)

        return pd.DataFrame(result_list)
    @staticmethod
    def is_succeed(point):
        if point > 0:
            return 'План перевыполнен'
        elif point == 0:
            return 'План выполнен'
        else:
            return 'План не выполнен'
    @staticmethod
    def set_human_date(string,date_dict):
        if type(string) == dt.date:
            return f'{string.day} {date_dict[string.month]}'
        else:
            return string

    def create_tables(self):
        log.info('Оформление дополнительных таблиц')
        tables_name = ['main_table', 'summary_percent', 'superior_percent', 'date_table', 'term_data',
                                       'plan_table', 'best_company', 'worst_company', 'average_table']
        main_table_columns = ['Количество участков','План','Факт','Выполнение','Отклонение','Результат']
        main_table = pd.DataFrame(columns=main_table_columns)
        main_table['Количество участков'] = self.count_point(False)
        main_table['План'] = main_table['Количество участков']*2 + 5
        main_table['Факт'] = self.count_point(True)
        main_table['Выполнение'] = main_table['Факт']/main_table['План']
        main_table['Отклонение'] = (main_table['Факт'] -main_table['План'])/ main_table['План']
        main_table['Результат'] = main_table['Отклонение'].apply(self.is_succeed)
        summary_row = pd.DataFrame([[sum(main_table['Количество участков']),
                                     sum(main_table['План']),
                                     sum(main_table['Факт']),
                                     '',
                                     '',
                                     '']], columns=main_table_columns)

        summary_percent = sum(main_table['Факт'])/sum(main_table['План'])
        superior_percent = len(main_table[main_table['Результат'] == 'План не выполнен'])/len(main_table)
        date_table = np.around(self.raw_report[[self.LAST_COLUMN, 'Организация']].groupby([self.LAST_COLUMN], as_index=False).count(), 0)
        date_table['human_date'] = date_table[self.LAST_COLUMN].apply(self.set_human_date, args=[self.HUMAN_DATES])
        date_table.sort_index(ascending = False, inplace=True)
        term_data = defaultdict(int)
        for i in range(len(date_table)):
            if isinstance(date_table[self.LAST_COLUMN][i], dt.date) and date_table[self.LAST_COLUMN][i] == self.DATE_DICT[self.LAST_COLUMN]:
                term_data['Закрыто ровно в срок']+=date_table['Организация'][i]
            elif isinstance(date_table[self.LAST_COLUMN][i], dt.date) and date_table[self.LAST_COLUMN][i] < self.DATE_DICT[self.LAST_COLUMN]:
                term_data['Закрыто раньше срока ']+=date_table['Организация'][i]
            else:
                term_data['Закрыто позже срока ']+=date_table['Организация'][i]
        term_data = pd.DataFrame(term_data.items(), columns=['Условие', 'Процент'])
        # term_data['Процент'] = term_data['Процент']
        common_count = sum(term_data['Процент'])
        term_data['Процент'] = np.around(term_data['Процент']/common_count,2)

        plan_table = main_table[['Результат', 'Количество участков']].groupby(['Результат'], as_index=False).count()
        common_count = sum(plan_table['Количество участков'])
        plan_table['Количество участков'] = np.around(plan_table['Количество участков'] / common_count,2)
        succeed_table = pd.DataFrame.from_dict(self.SUCCEED_TABLE, orient='index', columns=['Type', 'Coeff'])
        succeed_table.reset_index(inplace=True)
        succeed_table.columns = ['Company', 'Type', 'Coeff']
        worst_company = succeed_table[succeed_table['Type'] != 'ТЗО'].sort_values(by='Coeff').iloc[:10][['Company', 'Coeff']]
        best_company = succeed_table[succeed_table['Type'] == 'ТЗО'].sort_values(by='Coeff').iloc[-11:-1][['Company', 'Coeff']]
        average_table = pd.DataFrame(map(lambda x: (x[0], sum(x[1])/len(x[1])), self.AVERAGE_POINT_DICT.items()), columns = ['Участок','Средний бал'])
        average_table = average_table.sort_values(by=['Средний бал'], ascending=True)
        main_table = pd.concat([main_table, summary_row])
        tables_data = [main_table, summary_percent, superior_percent, date_table, term_data,
                                       plan_table, best_company,worst_company, average_table]
        tables = dict()
        for key, value in zip(tables_name, tables_data):
            tables[key] = value
        self.COLUMNS.extend(list(self.DATE_DICT.keys()))
        self.report = self.report[self.COLUMNS]
        return tables







