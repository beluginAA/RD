import pandas as pd  # module for working with data sets (need for working inner classes and modules)
import traceback
import datetime
import subprocess
import warnings
import numpy as np
import os  # module for working with operating system catalog structure
import openpyxl  # module for working with Excel files
import time  # module for working with date and time
import pyodbc  # module for working with databases
import pyxlsb
import threading

from datetime import datetime  # Module for working with date and time data
from tkinter.filedialog import askopenfilename  # Module for open file with win gui

# Ignoring pandas version errors
warnings.simplefilter(action='ignore', category=(FutureWarning, UserWarning))

# show an "Open" dialog box and return the path to the selected file
flag = False
filename_comp = askopenfilename(title="Select file for compare", filetypes=[("Excel Files", "*.xlsx"), ("Excel Binary Workbook", "*.xlsb")])
database_root = askopenfilename(title="Select database", filetypes=[('*.mdb', '*.accdb')]).replace('/', '\\')
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={database_root};'
    )

table = input("Input table's name : ")
with pyodbc.connect(conn_str) as conn:
    query = f'''SELECT * FROM {table}'''
    new_df = pd.read_sql(query, conn)

# Function for changing code 
def changing_code(df):
    date_expected, date_release = df['Код KKS документа'], df['Код KKS документа_new']
    if  not pd.isna(date_expected) and not pd.isna(date_release):
        return f'Смена кода с <{date_expected}> на <{date_release}>'
    else:
        return date_expected

# Function for changing name
def changing_name(df):
    date_expected, date_release = df['Наименование объекта/комплекта РД'], df['Наименование объекта/комплекта РД_new']
    if  not pd.isna(date_expected) and not pd.isna(date_release) and date_expected != date_release:
        return f'Смена наименования с <{date_expected}> на <{date_release}>'
    else:
        return date_expected

# Function for converting date
def convert_date(row):
    if row in ['в производстве', 'В производстве', None]:
        return row
    else:
        if dayfirst is True:
            return pd.to_datetime(row, dayfirst=dayfirst).date()
        else: 
            return pd.to_datetime(row, format='%Y/%m/%d').date()

# Function for changing_developer
def change_developer(df):
    if ~pd.isna(df['Разработчики РД (актуальные)']):
        return df['Разработчик РД']
    else:
        return df['Разработчики РД (актуальные)']

# Function for changing status
def change_status(df):
    if ~pd.isna(df['Статус текущей ревизии_new']):
        return df['Статус РД в 1С']
    else:
        return df['Статус текущей ревизии_new']

#Function for changing data
def changing_data(row):
    if (row[col] == row[f'{col}_new']) or (pd.isna(row[col]) and pd.isna(row[f'{col}_new'])):
        return None
    else:
        return f'Смена {col.lower()} с <{row[col]}> на <{row[f"{col}_new"]}>'

#Columns for changing dataframe

convert_columns = ['Дата выпуска РД по договору подрядчика', 
                   'Дата выпуска РД по графику с Заказчиком', 
                   'Дата статуса Заказчика', 
                   'Ожидаемая дата выдачи РД в производство', 
                   'Дата выпуска РД по договору подрядчика_new',
                   'Дата выпуска РД по графику с Заказчиком_new',
                   'Дата статуса Заказчика_new',
                   'Ожидаемая дата выдачи РД в производство_new']

clmns = ['Пакет РД', 'Статус Заказчика', 'Текущая ревизия', 'Статус текущей ревизии', 
           'Дата выпуска РД по договору подрядчика', 'Дата выпуска РД по графику с Заказчиком', 
           'Дата статуса Заказчика', 'Ожидаемая дата выдачи РД в производство']

inf_columns = ['Наименование объекта/комплекта РД',
               'Коды работ по выпуску РД', 
               'Пакет РД', 
               'Код KKS документа',
               'Статус Заказчика', 
               'Текущая ревизия', 
               'Статус текущей ревизии',
               'Дата выпуска РД по договору подрядчика',
               'Дата выпуска РД по графику с Заказчиком',
               'Дата статуса Заказчика',
               'Ожидаемая дата выдачи РД в производство', 
               'Письма',
               'Источник информации', 
               'Разработчики РД (актуальные)', 
               'Статус РД в 1С']

changed_cols = ['Коды работ по выпуску РД',
        'Наименование объекта/комплекта РД',
       'Код KKS документа',
       'Пакет РД',
       'Статус Заказчика', 
       'Текущая ревизия', 
       'Статус текущей ревизии',
       'Дата выпуска РД по договору подрядчика',
       'Дата выпуска РД по графику с Заказчиком', 'Дата статуса Заказчика',
       'Ожидаемая дата выдачи РД в производство',
       'Разработчики РД (актуальные)']

base_columns = ['Система',
                'Наименование объекта/комплекта РД',
                'Коды работ по выпуску РД',
                'Тип',
                'Пакет РД',	'Код KKS документа',
                'Статус Заказчика',
                'Текущая ревизия',
                'Статус текущей ревизии',
                'Дата выпуска РД по договору подрядчика',
                'Дата выпуска РД по графику с Заказчиком',
                'Дата статуса Заказчика',	
                'Ожидаемая дата выдачи РД в производство',	
                'Письма',	
                'Источник информации',	
                'Разработчики РД (актуальные)',	
                'Объект',	
                'WBS',	
                'КС',	
                'Примечания']

# read Excel files with current and new data
print('Read excel files with current and new data')
if '.xlsb' in filename_comp:
    with pyxlsb.open_workbook(filename_comp) as wb:
        with wb.get_sheet(1) as sheet:
            data = []
            for row in sheet.rows():
                data.append([item.v for item in row])
    base_df = pd.DataFrame(data[1:], columns=data[0])
    flag = True
else: 
    base_df = pd.read_excel(filename_comp)
new_df.columns = base_columns

#  Clear the data in both dataframe
print('Clear the empty rows in both dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df['Разработчики РД (актуальные)'] = base_df.apply(change_developer, axis = 1)

# Removing unnecessary data
print('Clear the unnecessary data in base dataframe')
base_df = base_df.loc[~base_df['Коды работ по выпуску РД'].str.contains('.C.')]
base_df = base_df.loc[~base_df['Код KKS документа'].isin(['.KZ.', '.EK.', '.TZ.', '.KM.', '.GR.'])]

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

#  Merging two dataframes dddd
print('Merging two dataframes')
missed_code = (new_df.merge(base_df,
                           how='left',
                           on=['Коды работ по выпуску РД'],
                           suffixes=['', '_new'], 
                           indicator=True))

m_df_1 = (pd.merge(new_df_copy, base_df_copy,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Код KKS документа'],
                           suffixes=['', '_new'], 
                           indicator=True))
tmp_df = m_df_1[m_df_1['_merge'] == 'left_only'][new_df_copy.columns]

m_df_2 = (tmp_df.merge(base_df_copy,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД'],
                           suffixes=['', '_new'],
                           indicator=True))

# на данном этапе можно выполнить копирование и запуск второго потока

m_df_1 = m_df_1[m_df_1['_merge'] == 'both']
m_df_1['Наименование объекта/комплекта РД'] = m_df_1.apply(lambda row: changing_name(row), axis = 1)
m_df_2 = m_df_2[m_df_2['_merge'] == 'both']
m_df_2['Код KKS документа'] = m_df_2.apply(lambda row: changing_code(row), axis = 1)

# Converting types
print('Converting types')
m_df_1['Статус текущей ревизии_new'] = m_df_1.apply(change_status, axis = 1)
m_df_2['Статус текущей ревизии_new'] = m_df_2.apply(change_status, axis = 1)

if flag:  
    for col in convert_columns:
        dayfirst = False
        if 'new' in col:
            m_df_1[col] = m_df_1[col].apply(lambda row: row if type(row) is str or row in ['в производстве', 'В производстве', None] else pd.to_datetime(row, unit='D', origin='1900-01-01').date() - pd.Timedelta(days=2))
            m_df_2[col] = m_df_2[col].apply(lambda row: row if type(row) is str or row in ['в производстве', 'В производстве', None] else pd.to_datetime(row, unit='D', origin='1900-01-01').date() - pd.Timedelta(days=2))
        else:
            dayfirst = True
            m_df_1[col] = m_df_1[col].apply(convert_date)
            m_df_2[col] = m_df_2[col].apply(convert_date)
else:
    for col in convert_columns:
        dayfirst = False
        if 'new' in col:
            m_df_1[col] = m_df_1[col].apply(convert_date)
            m_df_2[col] = m_df_2[col].apply(convert_date)
        else:
            dayfirst = True
            m_df_1[col] = m_df_1[col].apply(convert_date)
            m_df_2[col] = m_df_2[col].apply(convert_date)

for col in clmns:
    m_df_1[col] = m_df_1.apply(changing_data, axis = 1)
    m_df_2[col] = m_df_2.apply(changing_data, axis = 1)

m_df = pd.concat([m_df_1[changed_cols], m_df_2[changed_cols]])
m_df = m_df.reset_index()[changed_cols]
missed = missed_code[missed_code['_merge'] == 'left_only'].reset_index()[['Система', 'Коды работ по выпуску РД', 'Наименование объекта/комплекта РД']]
max_len_row = [max(m_df[row].apply(lambda x: len(str(x)) if x else 0)) for row in m_df.columns]
max_len_name = [len(row) for row in changed_cols]
max_len = [col_len if col_len > row_len else row_len for col_len, row_len in zip(max_len_name, max_len_row)]

comf_ren = input('Use standard file name for log file (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
   output_filename = 'log-RD-' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
with open(f'{output_filename}.txt', 'w') as log_file:
    log_file.write('Список измененных значений:\n')
    log_file.write('\n')
    file_write = ' ' * (len(str(m_df.index.max())) + 3)
    for column, col_len in zip(changed_cols, max_len):
        file_write += f"{column:<{col_len}}|"
    log_file.write(file_write)
    log_file.write('\n')
    for index, row in m_df.iterrows():
        column_value = ''
        for i in range(len(changed_cols)):
            column_value += f"{str(row[changed_cols[i]]) if row[changed_cols[i]] else '-':<{max_len[i]}}|"
        log_file.write(f"{index: <{len(str(m_df.index.max()))}} | {column_value}\n")
    log_file.write('\n')
    log_file.write('Список отсутствующих кодов.\n')
    log_file.write('     Коды работ по выпуску РД' + '\t | \t' + 'Наименование объекта/комплекта РД\n')
    for index, row in missed.iterrows():
        log_file.write(f'{str(index)}\t{row["Коды работ по выпуску РД"]}\t | \t{row["Наименование объекта/комплекта РД"]}\n')

# Этот код полностью готов, принимает любые файлы и возвращает лог файл с измененными значениями
# Подумать над реализацией многопоточности

# import threading
# import time

# # Программа создает 10 потоков, доводит их до команды wait, проверяет условие, что количество потоков должно быть больше 10
# # Далее ставит значение event как True и идет дальше

# event = threading.Event() # На этом этапе значение автоматически устанавливается как False

# def image_handler():
#     thr_num = threading.currentThread().name
#     print(f'Идет подготовка изображения из потока [{thr_num}]')
#     event.wait()
#     print(f'Изображение отправлено')

# # def test():
# #     while True:
# #         event.wait() # Процесс продолжится только тогда, когда значение event = True
# #         print('test')
# #         time.sleep(2)

# # event.clear() # Сброс значения event до значения False
# for i in range(10):
#     threading.Thread(target=image_handler, name = str(i)).start()
#     print(f'Поток [{i}] запущен')
#     time.sleep(1)

# if threading.active_count() >= 10:
#     event.set() # Задается значение True этой командой

# time.sleep(10)
# event.set()