import pandas as pd  # module for working with data sets (need for working inner classes and modules)
import traceback
import datetime
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

# show an "Open" dialog box and return the path to the selected file
filename_comp = askopenfilename(title="Select file for compare", filetypes=[("excel files", "*.xlsx")])
database_root = askopenfilename(title="Select database", filetypes=[('*.mdb', '*.accdb')]).replace('/', '\\')
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={database_root};'
    )

table = input("Input table's name : ")
with pyodbc.connect(conn_str) as conn:
    # Создаем курсор для выполнения запросов
    query = f'''SELECT * FROM {table}'''
    new_df = pd.read_sql(query, conn)

# Ignoring pandas version errors
warnings.simplefilter(action='ignore', category=(FutureWarning, UserWarning))

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
base_df = pd.read_excel(filename_comp)[inf_columns]
base_df_1 = pd.read_excel(filename_comp, sheet_name='блок 1')[inf_columns]
base_df_2 = pd.read_excel(filename_comp, sheet_name='блок 2')[inf_columns]
base_df_3 = pd.read_excel(filename_comp, sheet_name='блок 3')[inf_columns]
base_df_4 = pd.read_excel(filename_comp, sheet_name='блок 4')[inf_columns]
new_df.columns = base_columns

# Finding missed rows
print('Finding missed rows')
missed_df = pd.concat([base_df, base_df_1, base_df_1]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_2, base_df_2]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_3, base_df_3]).drop_duplicates(keep=False)
missed_df = pd.concat([missed_df, base_df_4, base_df_4]).drop_duplicates(keep=False)

#  Clear the data in both dataframe
print('Clear the empty rows in both dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df['Разработчики РД (актуальные)'] = base_df['Разработчики РД (актуальные)'].apply(
    lambda row: base_df['Разработчик РД'] if row is None else row
    )

# Removing unnecessary data
print('Clear the unnecessary data in base dataframe')
base_df = base_df.loc[(base_df['Коды работ по выпуску РД'].str.contains('.C.') == False)]
base_df = base_df.loc[(~base_df['Код KKS документа'].isin(['.KZ.', '.EK.', '.TZ.', '.KM.', '.GR.']))]
# base_df.count()

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

#  Merging two dataframes dddd
print('Merging two dataframes')
m_df_1 = (new_df_copy.merge(base_df_copy,
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
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row  else pd.to_datetime(row).date()
    )
changed_df['Объект'] = changed_df['Объект'].apply(
    lambda row: changed_df['Коды работ по выпуску РД'].str.slice(0, 5) if row is None else row
    )
changed_df['Статус текущей ревизии'] = changed_df['Статус текущей ревизии'].apply(
    lambda row: base_df['Статус РД в 1С'] if row is None else row
    )
changed_df['WBS'] = changed_df['WBS'].apply(
    lambda row: row if row is not None else changed_df['Коды работ по выпуску РД'].apply(lambda row: row[6 : row.find('.', 6)])
    )

comf_ren = input('Use standard file name (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
    output_filename = 'result' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
changed_df.to_excel(f'./{output_filename}.xlsx', encoding='cp1251', index = False)
