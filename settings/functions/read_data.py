import pandas as pd

def read_old_data(table_dict, month, path):
    summary_row = pd.DataFrame([[month, table_dict['summary_percent']]], columns=['month', 'percent'])
    superior_row = pd.DataFrame([[month, table_dict['superior_percent']]], columns=['month', 'percent'])
    df = pd.read_excel(path)
    first_table = df.iloc[:,[0, 1]]
    first_table = pd.concat([first_table, summary_row])
    first_table['percent'] = first_table['percent']
    second_table = df.iloc[:,[3, 4]]
    second_table.columns = first_table.columns
    second_table = pd.concat([second_table, superior_row])
    second_table['percent'] = second_table['percent']
    table_dict['summary_percent'] = first_table
    table_dict['superior_percent'] = second_table

    return table_dict
