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
               'Статус РД в 1С'
               ]

base_df = pd.read_excel(filename_comp)
new_df = pd.read_excel(filename_new)

base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df = base_df.loc[(~base_df['Коды работ по выпуску РД'].str.contains('.C.'))]
base_df = base_df.loc[((~base_df['Код KKS документа'].isin(['.KZ.', '.EK.', '.TZ.', '.KM.', '.GR.'])))]

base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

missed_code = (new_df.merge(base_df,
                           how='outer',
                           on=['Коды работ по выпуску РД'],
                           suffixes=['', '_new'], 
                           indicator=True))



missed = missed_code[missed_code['_merge'] == 'left_only'].reset_index()[columns]
with open('text.txt', 'w') as file:
    file.write('\n')
    file.write('Список отсутствующих кодов.\n')
    file.write('\n')
    file.write('     Коды работ по выпуску РД' + '\t | \t' + 'Наименование объекта/комплекта РД\n')
    for index, row in missed.iterrows():
        file.write(f'{str(index)}\t{row["Коды работ по выпуску РД"]}\t | \t{row["Наименование объекта/комплекта РД"]}\n')