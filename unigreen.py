import statistics

import pandas as pd
import requests
import urllib3
import xlwt

# Отключаем предупреждение о необходимости верификации
urllib3.disable_warnings()


def atsenergo_data(dict_url: dict) -> None:
    """Загрузка файлов EXCEL с данными для аналитики.

    Args:
        dict_url: dict
    """
    # Цикл по словарю dict_url, ключи уникальные наименования файлов.
    # Значения уникальные строки url.
    for key, value in dict_url.items():
        try:
            resp = requests.get(
                f'https://www.atsenergo.ru/nreport?fid={value}&region=eur',
                verify=False,  # Отключение верификации.
            )
        except requests.exceptions.ConnectionError as e:
            print(f'Ошибка загрузки файла {key}: {e}')
        try:
            # Сохранение на диск файл для аналитики.
            output = open(f'{key}_eur_big_nodes_prices_pub.xls', 'wb')
            output.write(resp.content)
            output.close
        except Exception as e:
            print(f'Ошибка сохранения файла{key}: {e}')
        else:
            print(f'Файл {key} успешно сохранён')


def write_tu_xls(out_dict: dict) -> None:
    """Сохранение данных в формате xls
    args:
        out_dict: dict
    """
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')
    for i in range(len(out_dict['value'])+1):
        if i == 0:
            sheet.write(0, 0, 'date')
            sheet.write(0, 1, 'value')
        else:
            sheet.write(i, 0, out_dict['date'][i-1])
            sheet.write(i, 1, out_dict['value'][i-1])
    workbook.save('out_xls_excel.xls')


def pandas_data(
        dict_url: dict, start: int, finish: int, region_of_the_RF: str,
        ) -> None:
    """Определение средних цен в указанные часы для заданного региона.

    Args:
        dict_url: dict
        start: int
        finish: int
        region_of_the_RF: str
    """
    try:
        # Словарь для вывода данных.
        out_dict: dict = {'date': [], 'value': []}
        # Цикл для обращения к файлам, в каждой итерации
        # обращение в следующему файлу.
        for key in dict_url:
            xls = pd.ExcelFile(f'{key}_eur_big_nodes_prices_pub.xls')
            # Список для хранения средних значений за каждый час.
            list_data: list[float] = []
            # Цикл для обращения к листам книги внутри файла.
            # В каждой итерации обращение к классу pandas.ExcelFile
            # в который уже загружен файл, повторного обращения к диску нет.
            for hour in range(start, finish + 1):
                # В каждой итерации создаётся новый датафрейм с 2 колонками.
                df = pd.read_excel(xls, hour, usecols='E:F')
                # Удаление пустых объектов (NaN).
                df.dropna(inplace=True, axis=0)
                # Фильтрация объектов датафрейм по региону.
                df = df.loc[
                    df['Unnamed: 4'] == region_of_the_RF, ['Unnamed: 5']]
                # Добавление в список list_data среднего значения
                # за каждый час.
                list_data.append(float(df['Unnamed: 5'].mean()))
            # Создание даты в списке date в словаре вывода out_dict.
            out_dict['date'].append(f'{key[6:]}.{key[4:6]}.{key[:4]}')
            # Вычисление среднего значения за сутки в списке list_data,
            # метод statistics.mean().
            # Добавление данных в список value словаря вывода out_dict.
            out_dict['value'].append(statistics.mean(list_data))
    except Exception as e:
        print(f'Ошибка при анализе данных {e}')
    try:
        # Сохранение на диск набора файлов с форматами из документации.
        pd.DataFrame(out_dict).to_excel('out_excel.xlsx', index=False)
        pd.DataFrame(out_dict).to_xml('out_xml.xml', index=False)
        pd.DataFrame(out_dict).to_csv('out_csv.csv', index=False)
        # Запись данных в формат xls
        write_tu_xls(out_dict)
    except Exception as e:
        print(f'Ошибка сохранения файлов: {e}')
    else:
        print('Файлы успешно сохранёны')


# Ключи - уникальные строки в названиях файлов.
# Значения - уникальные строки url.
dict_url: dict = {
    '20240902': '21106C4C4A250310E0630A4900E1BBD1',
    '20240903': '2124868598630102E0630A4900E1F5AC',
    '20240904': '2138B2AEC999038EE0630A4900E16528',
    '20240905': '214CD0A215090342E0630A4900E17296',
    '20240906': '2160DFEA1B8F02A0E0630A4900E19301',
    '20240907': '2174F3075435021CE0630A4900E1CFC7',
    '20240908': '21891809BF93030CE0630A4900E19D02',
    '20240909': '219D3D097A2A0362E0630A4900E124C2',
}
# Начало интервала (от 0 до 23).
start: int = 2
# Завершение интервала (от 0 до 23).
finish: int = 15
# Регион Российской Федерации.
region_of_the_RF: str = 'Республика Бурятия'


atsenergo_data(dict_url)
pandas_data(dict_url, start, finish, region_of_the_RF)
