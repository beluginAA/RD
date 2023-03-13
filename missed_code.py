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
import win32com.client  # Module for generating MS access data base

from datetime import datetime  # Module for working with date and time data
from win32com.client import Dispatch  # Module for generating MS access data base
from tkinter.filedialog import askopenfilename  # Module for open file with win gui

# Ignoring pandas version errors
warnings.simplefilter(action='ignore', category=(FutureWarning, UserWarning))

# show an "Open" dialog box and return the path to the selected file
filename_comp = askopenfilename(title="Select file for compare", filetypes=[("excel files", "*.xlsx")])
database_root = askopenfilename(title="Select database", filetypes=[('*.mdb', '*.accdb')]).replace('/', '\\')
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={database_root};'
    )

table = input("Input table's name : ")
with pyodbc.connect(conn_str) as conn:
    query = f'''SELECT * FROM {table}'''
    new_df = pd.read_sql(query, conn)

columns = ['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД']

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
base_df = pd.read_excel(filename_comp)
new_df.columns = base_columns

#  Clear the data in both dataframe
print('Clear the empty rows in both dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df['Разработчики РД (актуальные)'] = base_df['Разработчики РД (актуальные)'].apply(
    lambda row: base_df['Разработчик РД'] if row is None else row
    )

# Removing unnecessary data
print('Clear the unnecessary data in base dataframe')
base_df = base_df.loc[(~base_df['Коды работ по выпуску РД'].str.contains('.C.'))]
base_df = base_df.loc[((~base_df['Код KKS документа'].isin(['.KZ.', '.EK.', '.TZ.', '.KM.', '.GR.'])))]

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

#  Merging two dataframes dddd
print('Merging two dataframes')
missed_code = (new_df.merge(base_df,
                           how='outer',
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
                           indicator=True
                           ))

m_df_2['Код KKS документа'] = m_df_2.apply(lambda row: row['Код KKS документа'] if (row['Код KKS документа'] == row['Код KKS документа_new'] or (pd.isna(row['Код KKS документа']) and pd.isna(row['Код KKS документа_new'])))
    else f'Смена кода с <{row["Код KKS документа"]}> на <{row["Код KKS документа_new"]}>', axis=1)

# Preparing log-file
print('Preparing log-file')
m_df = m_df_1.copy()
m_df = m_df[(m_df['_merge'] == 'both')]
m_df = pd.concat([m_df, m_df_2[changed_cols]])
m_df['Дата выпуска РД по договору подрядчика'] = m_df['Дата выпуска РД по договору подрядчика'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, dayfirst=True).date()
    )
m_df['Дата выпуска РД по графику с Заказчиком'] = m_df['Дата выпуска РД по графику с Заказчиком'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row  else pd.to_datetime(row, dayfirst=True).date()
    )
m_df['Дата статуса Заказчика'] = m_df['Дата статуса Заказчика'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, dayfirst=True).date()
    )
m_df['Ожидаемая дата выдачи РД в производство'] = m_df['Ожидаемая дата выдачи РД в производство'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, dayfirst=True).date()
    )
m_df['Дата выпуска РД по договору подрядчика_new'] = m_df['Дата выпуска РД по договору подрядчика_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
m_df['Дата выпуска РД по графику с Заказчиком_new'] = m_df['Дата выпуска РД по графику с Заказчиком_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
m_df['Дата статуса Заказчика_new'] = m_df['Дата статуса Заказчика_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
m_df['Ожидаемая дата выдачи РД в производство_new'] = m_df['Ожидаемая дата выдачи РД в производство_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )

m_df['Статус текущей ревизии_new'] = m_df['Статус текущей ревизии_new'].apply(
    lambda row: base_df['Статус РД в 1С'] if row is None else row
    )
m_df['Наименование объекта/комплекта РД'] = m_df[['Наименование объекта/комплекта РД', 'Наименование объекта/комплекта РД_new']].apply(
    lambda row: row[0] if (row[0] == row[1] or (not row[0] == row[0] and not row[1] == row[1]) or (not row[0] and not row[1]))
    else f'Смена наименования с <{row[0]}> на <{row[1]}>', axis=1)
m_df['Пакет РД'] = m_df.apply(lambda row: None 
        if row['Пакет РД'] == row['Пакет РД_new'] or (pd.isna(row['Пакет РД']) and pd.isna(row['Пакет РД_new'])) 
        else f'Смена пакета с <{row["Пакет РД"]}> на <{row["Пакет РД_new"]}>', axis=1)
m_df['Статус Заказчика'] = m_df.apply(lambda row: None 
        if row['Статус Заказчика'] == row['Статус Заказчика_new'] or (pd.isna(row['Статус Заказчика']) and pd.isna(row['Статус Заказчика_new']))
        else f'Смена статуса с <{row["Статус Заказчика"]}> на <{row["Статус Заказчика_new"]}>', axis=1)
m_df['Текущая ревизия'] = m_df.apply(lambda row: None 
    if row['Текущая ревизия'] == row['Текущая ревизия_new'] or (pd.isna(row['Текущая ревизия']) and pd.isna(row['Текущая ревизия_new']))
    else f'Смена ревизии с <{row["Текущая ревизия"]}> на <{row["Текущая ревизия_new"]}>', axis=1)
m_df['Статус текущей ревизии'] = m_df.apply(lambda row: None 
    if row['Статус текущей ревизии'] == row['Статус текущей ревизии_new'] or (pd.isna(row['Статус текущей ревизии']) and pd.isna(row['Статус текущей ревизии_new']))
    else f'Смена статуса ревизии с <{row["Статус текущей ревизии"]}> на <{row["Статус текущей ревизии_new"]}>', axis=1)
m_df['Дата выпуска РД по договору подрядчика'] = m_df.apply(lambda row: None 
    if row['Дата выпуска РД по договору подрядчика'] == row['Дата выпуска РД по договору подрядчика_new'] or (pd.isna(row['Дата выпуска РД по договору подрядчика']) and pd.isna(row['Дата выпуска РД по договору подрядчика_new']))
    else f'Смена даты с <{row["Дата выпуска РД по договору подрядчика"]}> на <{row["Дата выпуска РД по договору подрядчика_new"]}>', axis=1)
m_df['Дата выпуска РД по графику с Заказчиком'] = m_df.apply(lambda row: None 
    if row['Дата выпуска РД по графику с Заказчиком'] == row['Дата выпуска РД по графику с Заказчиком_new'] or (pd.isna(row['Дата выпуска РД по графику с Заказчиком']) and pd.isna(row['Дата выпуска РД по графику с Заказчиком_new']))
    else f'Смена даты с <{row["Дата выпуска РД по графику с Заказчиком"]}> на <{row["Дата выпуска РД по графику с Заказчиком_new"]}>', axis=1)
m_df['Дата статуса Заказчика'] = m_df.apply(lambda row: None 
    if row['Дата статуса Заказчика'] == row['Дата статуса Заказчика_new'] or (pd.isna(row['Дата статуса Заказчика']) and pd.isna(row['Дата статуса Заказчика_new']))
    else f'Смена даты с <{row["Дата статуса Заказчика"]}> на <{row["Дата статуса Заказчика_new"]}>', axis=1)
m_df['Ожидаемая дата выдачи РД в производство'] = m_df.apply(lambda row: None 
    if row['Ожидаемая дата выдачи РД в производство'] == row['Ожидаемая дата выдачи РД в производство_new'] or (pd.isna(row['Ожидаемая дата выдачи РД в производство']) and pd.isna(row['Ожидаемая дата выдачи РД в производство_new']))
    else f'Смена даты с <{row["Ожидаемая дата выдачи РД в производство"]}> на <{row["Ожидаемая дата выдачи РД в производство_new"]}>', axis=1)
m_df['Разработчики РД (актуальные)'] = m_df.apply(lambda row: None 
    if row['Разработчики РД (актуальные)'] == row['Разработчики РД (актуальные)_new'] or (pd.isna(row['Разработчики РД (актуальные)']) and pd.isna(row['Разработчики РД (актуальные)_new']))
    else f'Смена разработчика с <{row["Разработчики РД (актуальные)"]}> на <{row["Разработчики РД (актуальные)_new"]}>', axis=1)

m_df = m_df.reset_index()[changed_cols]
missed = missed_code[missed_code['_merge'] == 'left_only'].reset_index()[columns]
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


#  Columns with necessary information
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
               'Статус РД в 1С'
               ]

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

#  Use columns numbers for next actions
col_numb = len(inf_columns)

# read Excel files with current and new data
print('Read excel files with current and new data')
new_base_df = base_df.copy()
new_new_df = new_df.copy()
base_df_1 = pd.read_excel(filename_comp, sheet_name='блок 1')[inf_columns]
base_df_2 = pd.read_excel(filename_comp, sheet_name='блок 2')[inf_columns]
base_df_3 = pd.read_excel(filename_comp, sheet_name='блок 3')[inf_columns]
base_df_4 = pd.read_excel(filename_comp, sheet_name='блок 4')[inf_columns]
new_new_df.columns = base_columns

# Finding missed rows
print('Finding missed rows')
missed_df = pd.concat([new_base_df, base_df_1, base_df_1]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_2, base_df_2]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_3, base_df_3]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_4, base_df_4]).drop_duplicates(keep=False)

#  Merging two dataframes dddd
print('Merging two dataframes')
m_df_1 = (new_new_df.merge(new_base_df,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Код KKS документа'],
                           suffixes=['', '_new'], 
                           indicator=True))
# Это слияение старого фрейма и нового, то есть в данном случае помимо ключевых столбцов будут новыми все остальные

tmp_df = m_df_1[m_df_1['_merge'] == 'left_only'][new_new_df.columns]
m_df_2 = (tmp_df.merge(new_base_df,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД'],
                           suffixes=['', '_new'],
                           indicator=True
                           ))

#  Preparation columns list with necessary information
tmp = np.append(m_df_1.columns[0:20].values, m_df_1.columns[-2])
columns = np.concatenate((m_df_2.columns[0:4],
                          m_df_2.columns[20:32],
                          m_df_2.columns[16:20],
                          m_df_2.columns[-2]),
                          axis = None)
tmp_columns = np.concatenate((
    m_df_1.columns[0], 
    m_df_1.columns[20],
    m_df_1.columns[2:4], 
    m_df_1.columns[21], 
    m_df_1.columns[5], 
    m_df_1.columns[22:32], 
    m_df_1.columns[16:20], 
    m_df_1.columns[-2]),
    axis = None)
# в данном случае мне стоит проверить, какие слобцы стоит обновить, а какие - оставить
# вся суть заключается в том, чтобы взять верные столбцы, тогда можно будет реализовать многопоточность
# где основной поток работает с файлом, а второстепенный - с формированием log файла
# тогда мне лучше копировать сразу не исходные фреймы, а фреймы, которые получились после merge (m_df_1, m_df_2)


#  Generate temporary dataframe for next appending for changed documents
print('Initiate dataframe with changed documents')
tmp_df = m_df_1[m_df_1['_merge'] == 'both'][tmp_columns]
tmp_df.columns = tmp

# Generate result dataframe, with remove document firstly
print('Generating result dataframe, with remove document firstly')
changed_df = m_df_2[m_df_2['_merge'] == 'left_only'][tmp]
tmp_df = pd.concat([changed_df, tmp_df])

#  Generate result dataframe 
print('Generating result dataframe ')
changed_df = m_df_2[m_df_2['_merge'] == 'both'][columns]
changed_df.columns = tmp
changed_df = pd.concat([changed_df, tmp_df])

# Changing dataframe
print('Changing dataframe')
changed_df['Дата выпуска РД по договору подрядчика'] = pd.to_datetime(changed_df['Дата выпуска РД по договору подрядчика'], dayfirst=True).dt.date
changed_df['Дата выпуска РД по графику с Заказчиком'] = pd.to_datetime(changed_df['Дата выпуска РД по графику с Заказчиком'], dayfirst=True).dt.date
changed_df['Дата статуса Заказчика'] = pd.to_datetime(changed_df['Дата статуса Заказчика'], dayfirst=True).dt.date
changed_df['Ожидаемая дата выдачи РД в производство'] = changed_df['Ожидаемая дата выдачи РД в производство'].apply(
    lambda row: row if row in ['в производстве', 'В производстве', None]  else pd.to_datetime(row).date()
    )
changed_df['Объект'] = changed_df['Объект'].apply(
    lambda row: changed_df['Коды работ по выпуску РД'].str.slice(0, 5) if pd.isna(row) else row)
changed_df['Статус текущей ревизии'] = changed_df['Статус текущей ревизии'].apply(
    lambda row: base_df['Статус РД в 1С'] if pd.isna(row) else row)
changed_df['WBS'] = changed_df['WBS'].apply(
    lambda row: row if ~pd.isna(row) else changed_df['Коды работ по выпуску РД'].apply(lambda row: row[6 : row.find('.', 6)]))

comf_ren = input('Use standard file name (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
    output_filename = 'result' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
changed_df.to_excel(f'./{output_filename}.xlsx', encoding='cp1251', index = False)
