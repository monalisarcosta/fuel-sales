import xml.etree.ElementTree as et
import pandas as pd
import requests
import pandas as pd
import numpy as np
from datetime import datetime
from requests.exceptions import ConnectionError
import win32com.client as win32
import zipfile

def download_xls():
    try:
        print('LOG[INFO]: Starting file download.')
        url = 'http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls'
        r = requests.get(url, allow_redirects=True)
        open('resources/vendas-combustiveis-m3.xls', 'wb').write(r.content)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(
            'C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\resources\\vendas-combustiveis-m3.xlsx')

        wb.SaveAs("C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\resources\\vendas-combustiveis-m3.xlsx", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
        wb.Close()
        excel.Application.Quit()
        print('LOG[INFO]: File downloaded successfully')
    except ConnectionError as e:
        print(f'LOG[ERRO]: Connection error, check url. \nErro: {e}.')

def convert_xlsx_to_xml():
    print('LOG[INFO]: Starting Conversion.')
    with zipfile.ZipFile('C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\resources\\vendas-combustiveis-m3.xlsx', 'r') as zip_ref:
        zip_ref.extractall('result')
    print('LOG[INFO]: Conversion completed.')

def get_info_definition():
    print('LOG[INFO]: Seeking definition information')
    ## Função responsável por pegar as informações de cada TAG e tranformar em cicionário
    xml_definition = 'C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\result\\xl\\pivotCache\\pivotCacheDefinition1.xml'
    xtree = et.parse(xml_definition)
    xroot = xtree.getroot()

    list_elements = []
    for element in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}s'):
        list_elements.append(element.attrib['v'])

    ## Informação do combustível
    dict_fuel= {}
    for index, value in enumerate(list_elements[:8]):
        dict_fuel[index] = value
    print('LOG[INFO]: Extracted fuel information.')

    ## informação de Região
    dict_region= {}
    for index, value in enumerate(list_elements[8:13]):
        dict_region[index] = value
    print('LOG[INFO]: Extracted region information.')

    ## informação de Estado
    dict_state= {}
    for index, value in enumerate(list_elements[13:]):
        dict_state[index] = value
    print('LOG[INFO]: Extracted state information.')

    ## Pega informação do ano e que está em outra TAG do xml

    list_year = []
    for element in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}n'):
        list_year.append(element.attrib['v'])

    dict_year= {}
    for index, value in enumerate(list_year):
        dict_year[index] = value
    print('LOG[INFO]: Extracted year information.')

    print('LOG[INFO]: Completed extraction of definition information')

    return dict_fuel, dict_region, dict_state, dict_year

def get_info_general(dict_fuel, dict_region, dict_state, dict_year):
    ## Função responsável por puxar as informaçoes de Combustível, Ano, Região e Estado da tag x do xml

    print('LOG[INFO]: Seeking additional information')

    xml_cache = 'C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\result\\xl\\pivotCache\\pivotCacheRecords1.xml'
    xtree = et.parse(xml_cache)
    xroot = xtree.getroot()

    list_info = []
    for value_tag in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x'):
        list_info.append(value_tag.attrib['v'])

    result = {"Combustivel": [], "ANO": [], "Região": [], "Estado": []}

    for i in range(0, len(list_info), 4):
        result['Combustivel'].append(dict_fuel[int(list_info[i])])
        result['ANO'].append(dict_year[int(list_info[i + 1])])
        result['Região'].append(dict_region[int(list_info[i + 2])])
        result['Estado'].append(dict_state[int(list_info[i + 3])])

    out_df = pd.DataFrame(result)

    print('LOG[INFO]: Completed extraction of additional information')

    return out_df

def get_info_month():
    ## Função responsável por puxar as informaçoes dos valores dos meses da tag n do xml
    print('LOG[INFO]: Searching month information')

    xml_cache = 'C:\\Users\\monal\\PycharmProjects\\anpFuelSales\\result\\xl\\pivotCache\\pivotCacheRecords1.xml'
    xtree = et.parse(xml_cache)
    xroot = xtree.getroot()

    list_info = []
    for value_tag in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}n'):
        list_info.append(value_tag.attrib['v'])

    # Lista com todas as tags dsponíveis no XML
    lista_tag = [elem.tag for elem in xroot.iter()]

    # Lista somente com as TAG m e n. Pois na TAG n vem o valor do mês, e quando TAG m há um mês sem valor
    new_list = []
    for i in lista_tag:
        if '}n' in i:
            new_list.append('1')
        elif '}m' in i:
            new_list.append('0')

    result = {"1": [], "2": [], "3": [], "4": [], "5": [], "6": [], "7": [], "8": [], "9": [], "10": [], "11": [], "12": [], "total": []}
    list_test = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]

    # Baseado na lista new_list, atribui 0 a lista list_info cujo mês não tem informação
    for index_new, value_new in enumerate(new_list):
        if value_new != '1':
            list_info.insert(index_new, 0)

    for i in range(0, len(new_list), 13):
        index_list = list_test.pop(-1)
        list_test.insert(0, index_list)

        result['1'].append(list_info[i + list_test[0]])
        result['2'].append(list_info[i + list_test[1]])
        result['3'].append(list_info[i + list_test[2]])
        result['4'].append(list_info[i + list_test[3]])
        result['5'].append(list_info[i + list_test[4]])
        result['6'].append(list_info[i + list_test[5]])
        result['7'].append(list_info[i + list_test[6]])
        result['8'].append(list_info[i + list_test[7]])
        result['9'].append(list_info[i + list_test[8]])
        result['10'].append(list_info[i + list_test[9]])
        result['11'].append(list_info[i + list_test[10]])
        result['12'].append(list_info[i + list_test[11]])
        result['total'].append(list_info[i + list_test[12]])

    out_df = pd.DataFrame(result)

    print('LOG[INFO]: Completed the extraction of information for the months')

    return out_df

def data_transformation(df_file):
    ## Função responsável por pegar todos os dados, remontar a tabela, extrair e gerar .parquet

    print('LOG[INFO]: Starting data transformation.')

    try:
        df_result = pd.DataFrame()

        #Lista que têm a realação de todos os meses
        list_months = {name: num + 1 for num, name in enumerate(df_file.keys()[4:-1])}

        #Caso o mês possua um dígito, acrescenta o 0 na frente
        for month in list_months.values():
            if len(str(month)) == 1:
                month_new = '0' + str(month)
            else:
                month_new = str(month)

            df_year = df_file[df_file.columns[1:2]]
            df = df_year.set_axis(['year_month'], axis=1, inplace=False)
            df['year_month'] = df['year_month'].apply(lambda row: f"{row}-{month_new}")
            df['uf'] = df_file['Estado']
            df['product'] = df_file['Combustivel']
            df['unit'] = df_file[str(month)]
            df['volume'] = df_file['total']
            df['created_at'] = datetime.now()

            df_result = pd.concat([df_result, df], ignore_index=True)

        print('LOG[INFO]: Transformation completed.')

        return df_result

    except FileNotFoundError as e:
        print(f'LOG[ERRO]: {e}')
    except ValueError as e:
        print(f'LOG[ERRO]: {e}')

def save_parquet(df_result):
    list_types ={"year_month": "datetime64[ms]", "uf": "string", "product": "string", "unit": "string", "volume": "double", 'created_at': 'datetime64[ms]'}
    df_result = df_result.astype(list_types)
    df_result.to_parquet('resources/result.parquet', engine='pyarrow')
    print('LOG[INFO]: .parquet file saved successfully.')


if __name__ == '__main__':

    download_xls()
    convert_xlsx_to_xml()
    dict_fuel, dict_region, dict_state, dict_year = get_info_definition()
    df_1 = get_info_general(dict_fuel, dict_region, dict_state, dict_year)
    df_2 = get_info_month()

    df_excel = pd.concat([df_1, df_2], axis=1)

    df_result = data_transformation(df_excel)
    save_parquet(df_result)

