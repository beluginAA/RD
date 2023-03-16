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

# Event for new thread
event = threading.Event()

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

start_time = time.time()

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

# Function for result-file
def new_threading(df_1, df_2):

    def change(df):
        if ~pd.isna(df['Статус текущей ревизии']):
            return df['Статус РД в 1С']
        else:
            return df['Статус текущей ревизии']

    #  Preparation columns list with necessary information
    mdf1_columns = ['Система',
                'Наименование объекта/комплекта РД_new',
                'Коды работ по выпуску РД',
                'Тип',
                'Пакет РД_new',	'Код KKS документа',
                'Статус Заказчика_new',
                'Текущая ревизия_new',
                'Статус текущей ревизии_new',
                'Дата выпуска РД по договору подрядчика_new',
                'Дата выпуска РД по графику с Заказчиком_new',
                'Дата статуса Заказчика_new',	
                'Ожидаемая дата выдачи РД в производство_new',	
                'Письма_new',	
                'Источник информации_new',	
                'Разработчики РД (актуальные)_new',	
                'Объект_new',	
                'WBS_new',	
                'КС',	
                'Примечания',
                'Статус РД в 1С']
    
    mdf2_columns = ['Система',
                'Наименование объекта/комплекта РД',
                'Коды работ по выпуску РД',
                'Тип',
                'Пакет РД_new',	'Код KKS документа_new',
                'Статус Заказчика_new',
                'Текущая ревизия_new',
                'Статус текущей ревизии_new',
                'Дата выпуска РД по договору подрядчика_new',
                'Дата выпуска РД по графику с Заказчиком_new',
                'Дата статуса Заказчика_new',	
                'Ожидаемая дата выдачи РД в производство_new',	
                'Письма_new',	
                'Источник информации_new',	
                'Разработчики РД (актуальные)_new',	
                'Объект_new',	
                'WBS_new',	
                'КС',	
                'Примечания',
                'Статус РД в 1С']
    
    new_columns = ['Система',
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
                'Примечания',
                'Статус РД в 1С']

    #  Generate temporary dataframe for next appending for changed documents
    print('Initiate dataframe with changed documents')
    tmp_df = df_1[df_1['_merge'] == 'both'][mdf1_columns]
    tmp_df.columns = new_columns

    
    # Generate result dataframe, with remove document firstly
    print('Generating result dataframe, with remove document firstly')
    changed_df = df_2[df_2['_merge'] == 'left_only'][new_columns]
    tmp_df = pd.concat([changed_df, tmp_df])

    #  Generate result dataframe 
    print('Generating result dataframe ')
    changed_df = df_2[df_2['_merge'] == 'both'][mdf2_columns]
    changed_df.columns = new_columns
    changed_df = pd.concat([changed_df, tmp_df])

    # Changing dataframe
    print('Changing dataframe')
    changed_df['Дата выпуска РД по договору подрядчика'] = pd.to_datetime(changed_df['Дата выпуска РД по договору подрядчика'], dayfirst=True).dt.date
    changed_df['Дата выпуска РД по графику с Заказчиком'] = pd.to_datetime(changed_df['Дата выпуска РД по графику с Заказчиком'], dayfirst=True).dt.date
    changed_df['Дата статуса Заказчика'] = pd.to_datetime(changed_df['Дата статуса Заказчика'], dayfirst=True).dt.date
    changed_df['Ожидаемая дата выдачи РД в производство'] = changed_df['Ожидаемая дата выдачи РД в производство'].apply(lambda row: row if row in ['в производстве', 'В производстве', None]  else pd.to_datetime(row).date())
    changed_df['Объект'] = changed_df['Объект'].apply(lambda row: changed_df['Коды работ по выпуску РД'].str.slice(0, 5) if pd.isna(row) else row)
    changed_df['Статус текущей ревизии'] = changed_df.apply(change, axis = 1)
    changed_df['WBS'] = changed_df['WBS'].apply(lambda row: row if ~pd.isna(row) else changed_df['Коды работ по выпуску РД'].apply(lambda row: row[6 : row.find('.', 6)]))
    changed_df = changed_df.iloc[:, :-1]
    changed_df.columns = changed_columns

    event.wait()
    # comf_ren = input('Use standard file name (y/n): ')
    # while comf_ren not in 'YyNn':
    #     comf_ren = input('For next work choose <y> or <n> simbols): ')

    # if comf_ren in 'Yy':
    #     output_filename = 'result' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
    # else:
    #     output_filename = input('Input result file name: ')
    # changed_df.iloc[:, :-1].to_excel(f'./{output_filename}.xlsx', encoding='cp1251', index = False)

    cnxn = pyodbc.connect(conn_str)

    # Создаем курсор
    cursor = cnxn.cursor()
    table_name = input("Input tables's name: ")
    cursor.execute(f"DROP TABLE [{table_name}]")
    cnxn.commit()

    cursor = cnxn.cursor()

    create_table_query = '''CREATE TABLE [РД] ([Система] VARCHAR(200), 
                                            [Наименование] VARCHAR(200),
                                            [Код] VARCHAR(200),
                                            [Тип] VARCHAR(200),
                                            [Пакет] VARCHAR(200),
                                            [Шифр] VARCHAR(200),
                                            [Итог_статус] VARCHAR(200),
                                            [Ревизия] VARCHAR(200), 
                                            [Рев_статус] VARCHAR(200), 
                                            [Дата_план] VARCHAR(200), 
                                            [Дата_граф] VARCHAR(200), 
                                            [Рев_дата] VARCHAR(200), 
                                            [Дата_ожид] VARCHAR(200), 
                                            [Письмо] VARCHAR(200), 
                                            [Источник] VARCHAR(200), 
                                            [Разработчик] VARCHAR(200), 
                                            [Объект] VARCHAR(200), 
                                            [WBS] VARCHAR(200), 
                                            [КС] VARCHAR(200), 
                                            [Примечания] VARCHAR(200))'''
    cursor.execute(create_table_query)
    cnxn.commit()
    cursor = cnxn.cursor()
    for row in changed_df.itertuples(index=False):
        insert_query = f'''INSERT INTO [РД] ([Система], [Наименование], [Код], [Тип], [Пакет], [Шифр], [Итог_статус], [Ревизия], 
                                             [Рев_статус], 
                                             [Дата_план], 
                                             [Дата_граф], 
                                             [Рев_дата], 
                                             [Дата_ожид], 
                                             [Письмо], 
                                             [Источник], 
                                             [Разработчик], 
                                             [Объект], 
                                             [WBS], 
                                             [КС], 
                                             [Примечания]) VALUES ({",".join(f"'{x}'" for x in row)})'''
        cursor.execute(insert_query)
    cnxn.commit()
    cursor.close()
    cnxn.close()

    end_time = time.time()
    print(end_time - start_time)

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
        return date_release
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
changed_columns = new_df.columns
new_df.columns = base_columns

#  Clearing dataframes
print('Clearing dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df['Разработчики РД (актуальные)'] = base_df.apply(change_developer, axis = 1)
base_df = base_df[~base_df['Коды работ по выпуску РД'].str.contains('\.C\.', regex=False)]
base_df['Код KKS документа'] = base_df['Код KKS документа'].astype(str)
base_df = base_df.loc[~base_df['Код KKS документа'].str.contains('\.KZ\.|\.EK\.|\.TZ\.|\.KM\.|\.GR\.')]

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

#  Merging two dataframes dddd
print('Merging two dataframes')
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

#  Missed rows
missed_rows = m_df_2[m_df_2['_merge'] == 'left_only'].reset_index()
missed_rows[['Система', 'Коды работ по выпуску РД', 'Наименование объекта/комплекта РД']]

#Changing dates
if flag:  
    for col in convert_columns:
        dayfirst = False
        if 'new' in col:
            m_df_1[col] = m_df_1[col].apply(
                lambda row: row if type(row) is str or row in ['в производстве', 'В производстве', None] else pd.to_datetime(row, unit='D', origin='1900-01-01').date() - pd.Timedelta(days=2)
                )
            m_df_2[col] = m_df_2[col].apply(
                lambda row: row if type(row) is str or row in ['в производстве', 'В производстве', None] else pd.to_datetime(row, unit='D', origin='1900-01-01').date() - pd.Timedelta(days=2)
                )
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

#  Making copy of merging dataframes
copy_df_1 = m_df_1.copy()
copy_df_2 = m_df_2.copy()

# Making a new thread fot result-file
result_file = threading.Thread(target=new_threading, args = (copy_df_1, copy_df_2, ), name = 'result_file')
result_file.start()

# Changing dataframes
print('Changing dataframes')
for col in clmns:
    m_df_1[col] = m_df_1.apply(changing_data, axis = 1)
    m_df_2[col] = m_df_2.apply(changing_data, axis = 1)

m_df_1 = m_df_1[m_df_1['_merge'] == 'both']
m_df_1['Наименование объекта/комплекта РД'] = m_df_1.apply(lambda row: changing_name(row), axis = 1)
m_df_2 = m_df_2[m_df_2['_merge'] == 'both']
m_df_2['Код KKS документа'] = m_df_2.apply(lambda row: changing_code(row), axis = 1)
m_df_1['Статус текущей ревизии_new'] = m_df_1.apply(change_status, axis = 1)
m_df_2['Статус текущей ревизии_new'] = m_df_2.apply(change_status, axis = 1)

# Preparing log-file
m_df = pd.concat([m_df_1[changed_cols], m_df_2[changed_cols]])
m_df = m_df.reset_index()[changed_cols]
max_len_row = [max(m_df[row].apply(lambda x: len(str(x)) if x else 0)) for row in m_df.columns]
max_len_name = [len(row) for row in changed_cols]
max_len = [col_len if col_len > row_len else row_len for col_len, row_len in zip(max_len_name, max_len_row)]
cols = ['Система', 'Коды работ по выпуску РД', 'Наименование объекта/комплекта РД']
max_missed_row = [max(missed_rows[row].apply(lambda x: len(str(x)) if x else 0)) for row in  cols]
max_missed_name = [len(row) for row in  cols]
max_missed = [col_len if col_len > row_len else row_len for col_len, row_len in zip(max_missed_name,max_missed_row)]
output_filename = 'log-RD-' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
# with open(f'{output_filename}.txt', 'w') as log_file:
#     log_file.write('Список измененных значений:\n')
#     log_file.write('\n')
#     file_write = ' ' * (len(str(m_df.index.max())) + 3)
#     for column, col_len in zip(changed_cols, max_len):
#         file_write += f"{column:<{col_len}}|"
#     log_file.write(file_write)
#     log_file.write('\n')
#     for index, row in m_df.iterrows():
#         column_value = ''
#         for i in range(len(changed_cols)):
#             column_value += f"{str(row[changed_cols[i]]) if row[changed_cols[i]] else '-':<{max_len[i]}}|"
#         log_file.write(f"{index: <{len(str(m_df.index.max()))}} | {column_value}\n")
#     log_file.write('\n')
#     log_file.write('Список отсутствующих кодов.\n')
#     log_file.write('\n')
#     file_write = ' ' * (len(str(m_df.index.max())) + 2)
#     for column, col_len in zip(cols, max_missed):
#         file_write += f"{column:<{col_len}}|"
#     log_file.write(file_write)
#     log_file.write('\n')
#     for index, row in missed_rows.iterrows():
#         column_value = ''
#         for i in range(len(cols)):
#             column_value += f"{str(row[cols[i]]) if row[cols[i]] else '-':<{max_missed[i]}}|"
#         log_file.write(f"{index: <{len(str(missed_rows.index.max()))}} | {column_value}\n")

# Continue new thread
event.set()
