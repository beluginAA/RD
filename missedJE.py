import pandas as pd
import time
import datetime
import pyodbc
import pyxlsb
import warnings

from tkinter.filedialog import askopenfilename

warnings.simplefilter(action = 'ignore', category=(UserWarning, FutureWarning))

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


filenameExcel = askopenfilename(title = 'Select file excel for compare', filetypes=[("Excel Files", "*.xlsx"), ("Excel Binary Workbook", "*.xlsb")])
filenameDB = askopenfilename(title = 'Select database for compare', filetypes = [('*.mdb', '*.accdb')]).replace('/', '\\')
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={filenameDB};'
    )

table = input("Input table's name: ")
with pyodbc.connect(conn_str) as connection:
    query = f'''SELECT * FROM {table}'''
    new_df = pd.read_sql(query, connection)

start_time = time.time()

# read Excel files with current and new data
print('Read excel files with current and new data')
if '.xlsb' in filenameExcel:
    with pyxlsb.open_workbook(filenameExcel) as wb:
        with wb.get_sheet(1) as sheet:
            data = []
            for row in sheet.rows():
                data.append([item.v for item in row])
    base_df = pd.DataFrame(data[1:], columns=data[0])
    flag = True
else: 
    base_df = pd.read_excel(filenameExcel)
new_df.columns = base_columns

#  Clearing dataframes
print('Clearing dataframes')
base_df = base_df.dropna(subset=['Коды работ по выпуску РД'])
base_df = base_df[~base_df['Коды работ по выпуску РД'].str.contains('\.C\.')]
base_df['Код KKS документа'] = base_df['Код KKS документа'].astype(str)
base_df = base_df[~base_df['Код KKS документа'].str.contains('\.KZ\.|\.EK\.|\.TZ\.|\.KM\.|\.GR\.')]
print(base_df.count())

#  Making copy of original dataframes
print('Making copy of original dataframes')
base_df_copy = base_df.copy()
new_df_copy = new_df.copy()

#  Merging two dataframes 
print('Merging two dataframes')
m_df_1 = (pd.merge(base_df_copy, new_df_copy,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Код KKS документа'],
                           suffixes=['', '_new'], 
                           indicator=True))
tmp_df = m_df_1[m_df_1['_merge'] == 'left_only'][base_df_copy.columns]

m_df_2 = (tmp_df.merge(new_df_copy,
                           how='left',
                           on=['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД'],
                           suffixes=['', '_new'],
                           indicator=True))

missedJE = m_df_2[m_df_2['_merge'] == 'left_only']
missedJE = missedJE[missedJE['Коды работ по выпуску РД'].str.contains('\.J\.|\.E\.')].reset_index()[['Наименование объекта/комплекта РД', 'Коды работ по выпуску РД']]
missedJE.to_excel('missedJE.xlsx', encoding='cp1251', index = False)