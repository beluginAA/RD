import pandas as pd


class Functions:

    def changing_code(self, df:pd.DataFrame) -> str:
        date_expected, date_release = df['Код KKS документа'], df['Код KKS документа_new']
        if  not pd.isna(date_expected) and not pd.isna(date_release):
            return f'Смена кода с <{date_expected}> на <{date_release}>'
        else:
            return '-'
        
    def finding_empty_rows(self, df:pd.DataFrame, column:str) -> str:
        if df[column] in ['nan', 'None', '0'] or pd.isna(df[column]):
            return ''
        else:
            return df[column]

    def changing_name(self, df:pd.DataFrame) -> str:
        date_expected, date_release = df['Наименование объекта/комплекта РД'], df['Наименование объекта/комплекта РД_new']
        if  not pd.isna(date_expected) and not pd.isna(date_release) and date_expected != date_release:
            return date_release
        else:
            return date_expected

    def changing_developer(self, df:pd.DataFrame) -> str:
        if ~pd.isna(df['Разработчики РД (актуальные)']):
            return df['Разработчик РД']
        else:
            return df['Разработчики РД (актуальные)']

    def changing_status(self, df:pd.DataFrame) -> str:
        if isinstance(df['Статус текущей ревизии_new'], float) or df['Статус текущей ревизии_new'] is None:
            return df['Статус РД в 1С']
        else:
            return df['Статус текущей ревизии_new']

    def changing_data(self, df:pd.DataFrame, column:str) -> str:
        if isinstance(df[column], float) or df[column] is None:
            df[column] = ''
        if isinstance(df[f'{column}_new'], float) or df[f'{column}_new'] is None:
            df[f'{column}_new'] = ''
        if (df[column] == df[f'{column}_new']) or (pd.isna(df[column]) and pd.isna(df[f'{column}_new'])):
            return None
        else:
            return f'Смена {column.lower()} с <{df[column]}> на <{df[f"{column}_new"]}>'
        
    def changing_wbs(self, row:str) -> str:
        split_row = row.split()
        if len(split_row) > 1:
            if split_row[0].upper() == split_row[0] and split_row[1] == '-':
                return split_row[0]
            else:
                return row
        else:
            return row

    def missed_codes(self, df:pd.DataFrame, anotherDf:pd.DataFrame) -> str:
        if df['Коды работ по выпуску РД'] not in list(anotherDf['Коды работ по выпуску РД']):
            return df
        else:
            return None

    def missed_codes_excel(self, df:pd.DataFrame, anotherDf:pd.DataFrame) -> str:
        if df['Коды работ по выпуску РД'] not in list(anotherDf['Коды работ по выпуску РД']):
            return df
        else:
            return None

    def find_row(self, row:str) -> str:
        if 'Смена' in row:
            return row
        else:
            return '-'