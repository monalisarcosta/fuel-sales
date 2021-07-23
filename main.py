import requests
import pandas as pd
import numpy as np
from datetime import datetime
from requests.exceptions import ConnectionError

def download_excel():
    try:
        print('LOG[INFO]: Starting file download.')
        url = 'http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls'
        r = requests.get(url, allow_redirects=True)
        open('resources/vendas-combustiveis-m3.xls', 'wb').write(r.content)
        print('LOG[INFO]: File downloaded successfully')

    except ConnectionError as e:
        print(f'LOG[ERRO]: Connection error, check url. \nErro: {e}.')

def read_excel(list_sheet_origem):
    try:
        print('LOG[INFO]: Starting to read the file.')
        df_result = pd.DataFrame()
        for sheet in list_sheet_origem:
            print(f'LOG[INFO]: Getting values from the sheet: {sheet}')
            df_file = pd.read_excel(r'resources/vendas-combustiveis-m3.xls', sheet_name=sheet)
            list_months = {name:num+1 for num, name in enumerate(df_file.keys()[4:-1])}

            for num in list_months.values():
                if len(str(num)) == 1: num = '0' + str(num)
                list_months = [*['ANO'], *['ESTADO'], *['COMBUSTÍVEL']]
                df_grouped = (df_file.groupby(list_months, as_index=False).mean().fillna(0))
                result = df_grouped[df_grouped.columns[:4]]
                result = result.set_axis(['year_month', 'uf', 'product', 'unit'], axis=1, inplace=False)
                result['created_at'] = datetime.now()
                result['year_month'] = result['year_month'].apply(lambda row: f"{row}-{num}")

                df_result = pd.concat([df_result, result], ignore_index=True)

        print('LOG[INFO]: Data extraction completed.')

        return df_result

    except FileNotFoundError as e:
        print(f'LOG[ERRO]: {e}')
    except ValueError as e:
        print(f'LOG[ERRO]: {e}')


def save_parquet(df_result):
    print('LOG[INFO]: Starting file save process .parquet.')

    try:
        list_types ={"year_month": "datetime64[ms]", "uf": "string", "product": "string", "unit": "double", 'created_at': 'datetime64[ms]'}
        df_result = df_result.astype(list_types)
        df_result.to_parquet('resources/result.parquet', engine='pyarrow')
    except Exception as e:
        print(f'LOG[ERRO]: {e}')

    print('LOG[INFO]: File saved successfully.')

if __name__ == '__main__':
    ## Method responsible for downloading the source file
    download_excel()
    ## Arquivo precisa ser aberto e salvo novamente, pois não está reconhecendo os outros sheets
    ## TODO: Ler arquivo pelo cache do excel

    ## List of sheets that data are stored. Later in Airflow it will be an environment variable
    list_sheet_origem = ['DPCache_m3', 'DPCache_m3_2']

    ## Method responsible for reading and extracting excel data.
    df_result = read_excel(list_sheet_origem)

    ## Method responsible for saving data in a .parquet file.
    save_parquet(df_result)

