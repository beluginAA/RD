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
filename_new = askopenfilename(title="Select new file", filetypes=[("excel files", "*.xlsx")])

# Ignoring pandas version errors
warnings.simplefilter(action='ignore', category=(FutureWarning, UserWarning))

# Function for creating contact_delay column
def contract_delay(df):
    date_expected = df['Ожидаемая дата выдачи РД в производство_new']
    date_release = df['Дата выпуска РД по договору подрядчика_old']
    if not pd.isna(date_expected) and not pd.isna(date_release) and date_expected != 'в производстве' and date_release != 'в производстве':
        return date_expected - date_release
    else:
        return None if pd.isna(date_expected) or pd.isna(date_release) else 'в производстве'

# Function for creating expected_date_difference column
def expected_date_difference(df):
    date_expected = df['Ожидаемая дата выдачи РД в производство_new']
    date_release = df['Ожидаемая дата выдачи РД в производство_old']
    if not pd.isna(date_expected) and not pd.isna(date_release) and date_expected != 'в производстве' and date_release != 'в производстве':
        return date_expected - date_release
    else:
        return None if pd.isna(date_expected) or pd.isna(date_release) else 'в производстве'

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
               'Статус РД в 1С']

columns = ['Система',
           'Наименование объекта/комплекта РД',
           'Коды работ по выпуску РД',
           'Ожидаемая дата выдачи РД в производство',
           'Дата выпуска РД по договору подрядчика',
           'Дата выпуска РД по графику с Заказчиком',
           'Ожидаемая дата выдачи РД в производство_new',
           'Дата выпуска РД по договору подрядчика_new',
           'Дата выпуска РД по графику с Заказчиком_new']

new_columns = ['Система',
               'Наименование объекта/комплекта РД',
               'Коды работ по выпуску РД',
               'Ожидаемая дата выдачи РД в производство_old',
               'Дата выпуска РД по договору подрядчика_old',
               'Дата выпуска РД по графику с Заказчиком_old',
               'Ожидаемая дата выдачи РД в производство_new',
               'Дата выпуска РД по договору подрядчика_new',
               'Дата выпуска РД по графику с Заказчиком_new']

#  Use columns numbers for next actions
col_numb = len(inf_columns)

# read Excel files with current and new data
print('Read excel files with current and new data')
base_df = pd.read_excel(filename_comp)[inf_columns]
new_df = pd.read_excel(filename_new)

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

#  Merging two dataframes
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

# Creating databases with dates
print('Creating databases with dates')
changed_df = pd.concat([m_df_1[m_df_1['_merge'] == 'both'][columns], m_df_2[m_df_2['_merge'] == 'left_only'][columns]])
changed_df = pd.concat([changed_df, m_df_2[m_df_2['_merge'] == 'both'][columns]])
changed_df =  changed_df[~changed_df['Система'].isin(['ГПМ', 'ИМ', "КИТС ФЗ", "МиЗ", "ОРС", "РЗА", "Связь", "СКУ ЭЧ ОС СН", "СКУ ЭЧ ЭБ", "Технология"])]
changed_df['Дата выпуска РД по договору подрядчика'] = changed_df['Дата выпуска РД по договору подрядчика'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, dayfirst=True).date()
    )
changed_df['Дата выпуска РД по графику с Заказчиком'] = changed_df['Дата выпуска РД по графику с Заказчиком'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row  else pd.to_datetime(row, dayfirst=True).date()
    )
changed_df['Дата выпуска РД по договору подрядчика_new'] = changed_df['Дата выпуска РД по договору подрядчика_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
changed_df['Дата выпуска РД по графику с Заказчиком_new'] = changed_df['Дата выпуска РД по графику с Заказчиком_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
changed_df['Ожидаемая дата выдачи РД в производство'] = changed_df['Ожидаемая дата выдачи РД в производство'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, dayfirst=True).date()
    )
changed_df['Ожидаемая дата выдачи РД в производство_new'] = changed_df['Ожидаемая дата выдачи РД в производство_new'].apply(
    lambda row: row if row == 'в производстве' or row == 'В производстве' or not row else pd.to_datetime(row, format = '%Y/%m/%d').date()
    )
changed_df.columns = new_columns

# Saving database
print('Saving database')
changed_df['Задержка договора'] = changed_df.apply(lambda row: contract_delay(row), axis=1)
changed_df['Разница ожидаемых дат new-old'] = changed_df.apply(lambda row: expected_date_difference(row), axis=1)

comf_ren = input('Use standard file name (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
    output_filename = 'dates' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
changed_df.to_excel(f'./{output_filename}.xlsx', encoding='cp1251', index = False)