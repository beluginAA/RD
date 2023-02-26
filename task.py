import pandas as pd  # module for working with data sets (need for working inner classes and modules)
import traceback
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

#  Columns with necessary information
inf_columns = ['КС1',
               'КС2',
               'КСС3',
               'Данные 1',
               'Данные 2',
               'Данные 3'
               ]

#  Use columns numbers for next actions
col_numb = len(inf_columns)

# read Excel files with current and new data
print('Read excel files with current and new data')
base_df = pd.read_excel('Новый отчет.xlsx')
new_df = pd.read_excel('Книга1.xlsx')

#  Clear the empty rows in both dataframes
# print('Clear the empty rows in both dataframes')
# base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
# new_df = new_df.dropna(subset=['Коды работ по выпуску РД'])

# Removing unnecessary data
# print('Clear the unnecessary data in base dataframe')
# base_df = base_df.loc[(base_df['Коды работ по выпуску РД'].str.contains('.C.') == False)]

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df[inf_columns].copy()
new_df_copy = new_df.copy()

#  Merging two dataframes
print('Merging two dataframes')
m_df = (new_df_copy.merge(base_df_copy,
                           how='outer',
                           on=['КС1', 'КС2'],
                           suffixes=['', '_new'], # Определяет, какая подпись прибавится к столбцам
                           indicator=True))

# Generate result dataframe, with remove document firstly
print('Generating result dataframe, with remove document firstly')
changed_df = m_df[m_df['_merge'] == 'left_only'][new_df.columns]

#  Generate temporary dataframe for next appending for new documents

#  Preparation columns list with necessary information
tmp = np.append(m_df.columns[0:3].values, m_df.columns[4])
tmp_columns = np.append(tmp, m_df.columns[8:-1].values)

#  Initiate dataframe of new documents
print('Initiate dataframe of new documents')
tmp_df = m_df[m_df['_merge'] == 'right_only'][tmp_columns]
tmp_df.columns = changed_df.columns

# Add dataframe with new documents to result dataframe
changed_df = changed_df.append(tmp_df)

#  Generate temporary dataframe for next appending for changed documents
print('Initiate dataframe with changed documents')
tmp_df = m_df[m_df['_merge'] == 'both'][tmp_columns]
# tmp_df['Статус'] = tmp_df[['Статус Заказчика', 'Статус Заказчика_new']].apply(
#     lambda row: None if (row[0] == row[1] or (not row[0] == row[0] and not row[1] == row[1]))
#     else f'Смена статуса с <{row[0]}> на <{row[1]}>', axis=1)
tmp_df.columns = changed_df.columns

# Add dataframe with new documents to result dataframe
changed_df = changed_df.append(tmp_df)

comf_ren = input('Use standard file name (y/n): ')
while comf_ren not in 'YyNn':
    comf_ren = input('For next work choose <y> or <n> simbols): ')

if comf_ren in 'Yy':
    output_filename = 'result' + str(datetime.now().isoformat(timespec='minutes')).replace(':', '_')
else:
    output_filename = input('Input result file name: ')
changed_df.to_excel(f'./{output_filename}.xlsx', encoding='cp1251')
