import pandas as pd  
import datetime
import warnings
import numpy as np
import time 
import threading 
import columns # module with the necessery columns
import sys

from datetime import datetime  
from tkinter.filedialog import askopenfilename  
from loguru import logger 
from processing import Preproccessing, PostProcessing, ResultFiles
from Functions import Functions


logger.remove()
warnings.simplefilter(action='ignore', category=(FutureWarning, UserWarning))
mainLogger = logger.bind(name = 'Mainlogger').opt(colors = True)
mainLogger.add(sink = sys.stdout, format = "<green>{time:HH:mm:ss}</green> | {message}", level = 'INFO')

mainLogger.info('The program starts.')
databaseRoot = askopenfilename(title='Select database', filetypes=[('*.mdb', '*.accdb')]).replace('/', '\\')
excelRoot = askopenfilename(title="Select file for compare", filetypes=[("Excel Files", "*.xlsx"), ("Excel Binary Workbook", "*.xlsb")])
xlsbFind = True if '.xlsb' in excelRoot else False

mainLogger.info('Read excel files with current and new data')
func = Functions()
connect = Preproccessing(databaseRoot, excelRoot)
msDatabase = connect.to_database()
excelDatabase = connect.to_excel()
changedColumns = msDatabase.columns

mainLogger.info('Clearing dataframes')
excelDatabase = excelDatabase.dropna(subset=['Коды работ по выпуску РД'])
excelDatabase['Разработчики РД (актуальные)'] = excelDatabase.apply(func.changing_developer, axis = 1)
excelDatabase = excelDatabase[~excelDatabase['Коды работ по выпуску РД'].str.contains('\.C\.', regex=False)]
excelDatabase['Объект'] = excelDatabase['Объект'].apply(lambda row: row[ : row.find(' ')])
excelDatabase['WBS'] = excelDatabase['WBS'].apply(func.changing_wbs)
excelDatabase['Код KKS документа'] = excelDatabase['Код KKS документа'].astype(str)
excelDatabase = excelDatabase.loc[~excelDatabase['Код KKS документа'].str.contains('\.KZ\.|\.EK\.|\.TZ\.|\.KM\.|\.GR\.')]
for column in columns.convert_columns[:4]:
    excelDatabase[column] = excelDatabase[column].apply(lambda row: '' if not isinstance(row, datetime) else row.strftime('%d-%m-%Y'))
    msDatabase[column] = msDatabase[column].apply(lambda row: '' if not isinstance(row, datetime) else row.strftime('%d-%m-%Y'))
 
mainLogger.info('Finding missing values ​​in a report.')
result = ResultFiles()
msDatabaseCopy = msDatabase.copy()
excelDatabaseCopy = excelDatabase.copy()
msDbJE = msDatabase.copy()
msDbJE = msDbJE[['Коды работ по выпуску РД']]
excelDbJE = excelDatabase.copy()
excelDbJE = excelDbJE[excelDbJE['Коды работ по выпуску РД'].str.contains('\.J\.|\.E\.')].reset_index()[['Коды работ по выпуску РД']]
excelDbJE = excelDbJE.apply(lambda df: func.missed_codes_excel(df, msDbJE), axis = 1)
result.to_logfile(excelDbJE.dropna().reset_index(drop = True), 'Пропущенные значения, которые есть в отчете, но нет в РД (J, E)')

mainLogger.info('Merging two dataframes')
rdKksDf = (pd.merge(excelDatabaseCopy, msDatabaseCopy, #m_df_1
                           how='outer',
                           on=['Коды работ по выпуску РД', 'Код KKS документа'],
                           suffixes=['', '_new'], 
                           indicator=True))
dfPath =rdKksDf[rdKksDf['_merge'] == 'right_only'][columns.mdf1_columns] #tmp_df
dfPath.columns = columns.new_columns

rdNameDf = (dfPath.merge(excelDatabaseCopy, # m_df_2
                           how='outer',
                           on=['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД'],
                           suffixes=['', '_new'],
                           indicator=True))
rdKksDf['Статус текущей ревизии_new'] = rdKksDf.apply(func.changing_status, axis = 1)
rdNameDf['Статус текущей ревизии_new'] = rdNameDf.apply(func.changing_status, axis = 1)

mainLogger.info('Finding missed rows.')
missedRows = rdNameDf[rdNameDf['_merge'] == 'left_only'].reset_index()[columns.mdf2_columns]
missedRows.columns = columns.new_columns
missedJE = missedRows[missedRows['Коды работ по выпуску РД'].str.contains('\.J\.|\.E\.')].reset_index()[['Коды работ по выпуску РД', 'Наименование объекта/комплекта РД']]
missedJE = missedJE.apply(lambda df: func.missed_codes(df, excelDatabase), axis = 1)
missedJE = missedJE.dropna().reset_index(drop = True)
result.to_logfile(missedJE, 'Пропущенные значения, которые есть в РД, но нет в отчете (J, E)')

mainLogger.info('Changing dataframes')
rdKksDf = rdKksDf[rdKksDf['_merge'] == 'both']
rdNameDf = rdNameDf[rdNameDf['_merge'] == 'both']
rdKksDf['Наименование объекта/комплекта РД'] = rdKksDf.apply(lambda row: func.changing_name(row), axis = 1)
rdNameDf['Код KKS документа'] = rdNameDf.apply(lambda row: func.changing_code(row), axis = 1)
for col in columns.clmns:
    rdKksDf[col] = rdKksDf.apply(lambda df: func.changing_data(df, col), axis = 1)
    rdNameDf[col] = rdNameDf.apply(lambda df: func.changing_data(df, col), axis = 1)
rdKksDfCopy = rdKksDf.copy()
rdNameDfCopy = rdNameDf.copy()
rdKksDf = rdKksDf[columns.mdf1_columns]
rdKksDf.columns = columns.new_columns
rdNameDf = rdNameDf[columns.mdf2_columns]
rdNameDf.columns = columns.new_columns

mainLogger.info('Preparing changed data for log-file.')
rdNameLogFile = rdNameDfCopy[columns.logFileColumns]
rdKksLogFile = rdKksDfCopy[columns.logFileColumns]
rdKks = rdKksLogFile.copy()
rdName = rdNameLogFile.copy()
rdName['Код KKS документа'] = rdName['Код KKS документа'].apply(func.find_row)
changedLogfile = pd.concat([rdName, rdKks])
resultThread = threading.Thread(name = 'resultThread', target = result.to_logfile, args = (changedLogfile.reset_index(drop = True), 'Измененные значения',))
resultThread.start()

mainLogger.info('Preparing the final files.')
summaryDf = pd.concat([rdKksDf, rdNameDf])
summaryDf = summaryDf[columns.base_columns]
summaryDf = summaryDf.reset_index(drop = True)
resultExcelDf = summaryDf.copy()
resultExcelDf['Объект'] = resultExcelDf['Объект'].apply(lambda row: resultExcelDf['Коды работ по выпуску РД'].str.slice(0, 5) if pd.isna(row) else row)
resultExcelDf['WBS'] = resultExcelDf['WBS'].apply(lambda row: row if ~pd.isna(row) else resultExcelDf['Коды работ по выпуску РД'].apply(lambda row: row[6 : row.find('.', 6)]))
resultExcelDf.columns = changedColumns
for col in resultExcelDf.columns:
    resultExcelDf[col] = resultExcelDf.apply(lambda df: func.finding_empty_rows(df, col), axis = 1)
    summaryDf[col] = summaryDf.apply(lambda df: func.finding_empty_rows(df, col), axis = 1)
result.to_resultfile(resultExcelDf)

mainLogger.info('Making changes to the database.')
step = PostProcessing(databaseRoot)
step.delete_table()
step.create_table()
step.insert_into_table(summaryDf)





