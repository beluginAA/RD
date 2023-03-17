import pandas as pd  # module for working with data sets (need for working inner classes and modules)
import traceback
import os  # module for working with operating system catalog structure
import openpyxl  # module for working with Excel files
import time  # module for working with date and time
# import pyodbc  # module for working with databases
# import win32com.client  # Module for generating MS access data base

from datetime import datetime  # Module for working with date and time data
# from win32com.client import Dispatch  # Module for generating MS access data base
from tkinter.filedialog import askopenfilename  # Module for open file with win gui

# show an "Open" dialog box and return the path to the selected file
filename_comp = askopenfilename(title="Select file for compare", filetypes=[("excel files", "*.xlsx")])
filename_new = askopenfilename(title="Select new file", filetypes=[("excel files", "*.xlsx")])

#  Columns with necessary information
inf_columns = ['Наименование объекта/комплекта РД',
               'Коды работ по выпуску РД',
               'Код KKS документа',
               'Текущая ревизия',
               'Статус текущей ревизии',
               'Письма',
               'Статус Заказчика',
               'Дата статуса Заказчика',
               'Ожидаемая дата выдачи РД в производство',
               'Источник информации',
               'Разработчик РД',
               'WBS',
               'Статус РД в 1С'
               ]
#  Use columns numbers for next actions
col_numb = len(inf_columns)

# read Excel files with current and new data
print('Read excel files with current and new data')
base_df = pd.read_excel(filename_comp)
new_df = pd.read_excel(filename_new)

#  Clear the empty rows in both dataframes
print('Clear the empty rows in both dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
new_df = new_df.dropna(subset=['Коды работ по выпуску РД'])

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df[inf_columns].copy()
new_df_copy = new_df[inf_columns].copy()

#  Combine status from 1C and documentary department for both dataframes copies
print('Combining status from 1C and documentary department for both dataframes copies')
base_df_copy.loc[:, 'Статус текущей ревизии'] = base_df_copy[['Статус текущей ревизии', 'Статус РД в 1С']].apply(
    lambda row: row[0] if row[0] == row[0] else row[1], axis=1)

new_df_copy.loc[:, 'Статус текущей ревизии'] = new_df_copy[['Статус текущей ревизии', 'Статус РД в 1С']].apply(
    lambda row: row[0] if row[0] == row[0] else row[1], axis=1)

#  Merging two dataframes
print('Merging two dataframes')
m_df = (base_df_copy.merge(new_df_copy,
                           how='outer',
                           on=['Коды работ по выпуску РД'],
                           suffixes=['', '_new'],
                           indicator=True))

# display(m_df[m_df['_merge']=='right_only'])
# m_df['_merge'].value_counts()

# Generate result dataframe, with remove document firstly
print('Generating result dataframe, with remove document firstly')
changed_df = m_df[m_df['_merge'] == 'left_only'][m_df.columns[:col_numb]]
changed_df['Статус'] = 'РД отсутствует в новом списке'

#  Generate temporary dataframe for next appending for new documents

#  Preparation columns list with necessary information
tmp_columns = [m_df.columns[col_numb], m_df.columns[1]]
# print(tmp_columns)
tmp = m_df.columns[col_numb + 1:-1].values
# print(tmp)
tmp_columns.extend(tmp)

# print(tmp_columns)

#  Initiate dataframe of new documents
print('Initiate dataframe of new documents')
tmp_df = m_df[m_df['_merge'] == 'right_only'][tmp_columns]
tmp_df['Статус'] = 'РД отсутствует в изначальном списке'
tmp_df.columns = changed_df.columns
# tmp_df.info()

# Add dataframe with new documents to result dataframe
changed_df = changed_df.append(tmp_df)

#  Generate temporary dataframe for next appending for changed documents
print('Initiate dataframe with changed documents')
tmp_df = m_df[m_df['_merge'] == 'both']
tmp_df['Статус'] = tmp_df[['Статус Заказчика', 'Статус Заказчика_new']].apply(
    lambda row: None if (row[0] == row[1] or (not row[0] == row[0] and not row[1] == row[1]))
    else f'Смена статуса с <{row[0]}> на <{row[1]}>', axis=1)

#  Remove all non-changed rows
tmp_df = tmp_df.dropna(subset=['Статус'])

tmp_df['Статус Заказчика_new'] = tmp_df['Статус Заказчика_new'].apply(
    lambda x: x if ('ВК +' in str(x)) or ('Выдан в производство' in str(x)) else None)

tmp_df = tmp_df.dropna(subset=['Статус Заказчика_new'])

#  Preparation columns list with necessary information
tmp_columns = [tmp_df.columns[col_numb], tmp_df.columns[1]]
tmp = tmp_df.columns[col_numb + 1:-2].values
tmp_columns.extend(tmp)
tmp_columns.extend([tmp_df.columns[-1]])
# tmp_df.info()

#  Use only necessary columns
tmp_df = tmp_df[tmp_columns]

# Rename columns un temporary dataframe
tmp_df.columns = changed_df.columns
# tmp_df.info()

#  Add dataframe with changed documents to result dataframe
changed_df = changed_df.append(tmp_df)

print('Removing architectural documents')
changed_df['Коды работ по выпуску РД'] = changed_df['Коды работ по выпуску РД'].apply(
    lambda x: None if '.C.' in x else x
)
changed_df = changed_df.dropna(subset=['Коды работ по выпуску РД'])

comf_ren = input('Use standard file name (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
    output_filename = 'result' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
changed_df.to_excel(f'./{output_filename}.xlsx', encoding='cp1251')
