# %% Определение функций
import requests
import json
import time
from datetime import date,datetime,timedelta
import pandas as pd
import shutil
import os
import glob
import csv
import numpy as np
from openpyxl import Workbook
import re
from loguru import logger
import getopt
import sys
pd.options.mode.chained_assignment = None

# Функция форматирования excel с планом акций, отдельный скрипт
from wb.scripts.format_actions_svod import format_excel_actions
# Функция выгрузки списка товаров
from wb.scripts.uploadDataFromWB import getWBProduct, get_prices_WB
# Файл с настройками для клиента
from wb.scripts.constants import(
    headers,
    client_number,
    client_name,
    marketplace_dir_name,
    catalog_action_columns,
    svod_actions_cols,
    net_cost_koef
)


# Функция чтения дат формирования отчета
def read_dates_file(date_report_created):
    report_dates = pd.read_csv(f"{marketplace_dir_name}/Clients/{client_name}/UploadFiles/UploadFiles_{date_report_created}/{date_report_created}_dates_from_to.csv", sep=';')
    for col in report_dates:
        report_dates[col] = pd.to_datetime(report_dates[col])
    return report_dates



# Функция создания диапазона дат
# GPT START ----
def generate_date_range(
        reference_date: str = None,
        start_date: str = None,
        end_date: str = None
    ):
    # Проверка на несовместимость параметров:
    # либо используем reference_date, либо вручную заданный диапазон
    if reference_date and (start_date or end_date):
        raise ValueError("Нельзя одновременно задавать reference_date и start_date/end_date.")

    # Если оба manual-диапазона заданы
    if start_date and end_date:
        # Парсим строки дат в объекты datetime
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    elif start_date or end_date:
        # Один из диапазонов не указан — ошибка
        raise ValueError("Нужно задать одновременно start_date и end_date.")
    else:
        # Ни диапазон, ни reference_date не заданы явно
        #  — значит работаем с reference_date (по умолчанию — сегодня)
        if reference_date is None:
            ref_dt = datetime.now()
        else:
            # Парсим reference_date из строки
            ref_dt = datetime.strptime(reference_date, "%Y-%m-%d")

        # Если дата — первое число месяца
        if ref_dt.day == 1:
            # Берём последний день прошлого месяца
            prev_month_last_day = ref_dt - timedelta(days=1)
            # Начало диапазона — первое число прошлого месяца
            start_dt = prev_month_last_day.replace(day=1)
            # Конец диапазона — последний день прошлого месяца
            end_dt = prev_month_last_day
        else:
            # Если не первое число месяца:
            # Начало диапазона — первое число текущего месяца
            start_dt = ref_dt.replace(day=1)
            # Конец диапазона — предыдущие сутки относительно reference_date
            end_dt = ref_dt - timedelta(days=1)

    # Форматы для строк и для ISO-даты по времени суток
    start_date_str = start_dt.strftime('%Y-%m-%d')
    end_date_str = end_dt.strftime('%Y-%m-%d')
    start_date_iso = start_dt.strftime('%Y-%m-%dT00:00:00Z')
    end_date_iso = end_dt.strftime('%Y-%m-%dT23:59:59Z')

    # Собираем результат в DataFrame
    df_dates = pd.DataFrame([{
        'date_start': start_date_str,       # начало диапазона (дд.мм.гггг)
        'date_end': end_date_str,           # конец диапазона (дд.мм.гггг)
        'datetime_start': start_date_iso,   # начало ISO-датой с 00:00:00
        'datetime_end': end_date_iso        # конец ISO-датой с 23:59:59
    }])

    return start_date_iso, end_date_iso
# GPT END ----


# Функция обработки данных каталога
def process_product_list(df_products):
    # Создаем копию для избежания изменений в оригинальном df
    df_products_processed = df_products.copy()
    # Для некоторых клиентов в размере заменяем точку на запятую
    if client_name in ['KU_And_KU', 'Soyuz']:
        df_products_processed['Размер'] =df_products_processed['Размер'].str.replace(",", ".", regex=False)
        df_products_processed['Размер'] = df_products_processed['Размер'].astype(str).str.replace('.0', '')
        df_products_processed['Размер'] = df_products_processed['Размер'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_products_processed['Размер'] = df_products_processed['Размер'].fillna(0)
    # Удаляем дубликаты
    df_products_processed = df_products_processed.drop_duplicates(subset=['Артикул продавца', 'Размер'])
    # Создаем колонку Артикул_Размер для мерджа
    df_products_processed['Артикул_Размер'] = df_products_processed[['Артикул продавца', 'Размер']].apply(lambda row: '_size_'.join(row.values.astype(str)), axis=1)
    # Считаем скидку
    df_products_processed['discount'] = np.round(1 - (df_products_processed['discount_price']/df_products_processed['price']), 3)

    return df_products_processed

# Функция выгрузки списка акций
def get_action_list(headers, date_start_actions, date_end_actions):
    # Параметры запроса
    params_action_list = {
                'startDateTime': date_start_actions,
                'endDateTime': date_end_actions,
                'allPromo': True,
                'limit': 1000,
                'offset': 0
        }
    # Получаем список акций
    resp_data_action_list = requests.get('https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions', headers=headers, params=params_action_list).json()
    # Переводим в df
    df_action_list = pd.DataFrame(resp_data_action_list['data']['promotions'])
    # Переименовываем некоторые колонки для удобства
    df_action_list = df_action_list.rename(columns={
        'id': 'action_id',
        'name': 'action_title'
        })
    # Получаем информацию по акциям
    # df, куда будем помещать информацию о кампаниях
    df_actions_info = pd.DataFrame()

    # Разбиваем на диапазоны по 100 акций
    df_action_list['id'] = np.arange(df_action_list.shape[0])
    step = 100
    df_action_list['chunks'] = df_action_list["id"].apply(lambda x: int(x/step) + 1)
    # Цикл по каждому диапазону
    for chunk in df_action_list['chunks'].unique():
        # Выбираем отдельный диапазон
        tmp_df_action_list = df_action_list.loc[df_action_list['chunks'] == chunk, :]
        # Формируем list с id акций из диапазона
        action_ids_list = tmp_df_action_list['action_id'].astype(str).to_list()
        # Параметры запроса
        params_actions_info = {'promotionIDs': action_ids_list}
        # Делаем запрос
        resp_data_actions_info = requests.get("https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions/details",
                                              headers=headers,
                                              params=params_actions_info).json()
        # Создаем df из ответа апи
        tmp_df_actions_info = pd.DataFrame(resp_data_actions_info['data']['promotions'])
        # Объединяем с предыдущим проходом цикла
        df_actions_info = pd.concat([df_actions_info, tmp_df_actions_info])

    # Переименовываем некоторые колонки для удобства
    df_actions_info = df_actions_info.rename(columns={
        'id': 'action_id',
        'name': 'action_title'
        })
    # Убираем акции, в которых нет товаров
    df_actions_list_final = (
        df_actions_info
        # .loc[df_actions_info['participationPercentage'] > 0, :]
        .reset_index(drop=True)
    )

    return df_actions_list_final



# Функция выгрузки информации по акциям
def get_actions_info(headers, df_action_list):
    # Создаем копию для избежания изменений в оригинальном df
    df_action_list_ = df_action_list.copy()
    # df, куда будем помещать итоговый результат
    df_actions_info = pd.DataFrame()

    # Разбиваем на диапазоны по 100 акций
    df_action_list_['id'] = np.arange(df_action_list_.shape[0])
    step = 100
    df_action_list_['chunks'] = df_action_list_["id"].apply(lambda x: int(x/step) + 1)
    # Цикл по каждому диапазону
    for chunk in df_action_list_['chunks'].unique():
        # Выбираем отдельный диапазон
        tmp_df_action_list = df_action_list_.loc[df_action_list_['chunks'] == chunk, :]
        # Формируем list с id акций из диапазона
        action_ids_list = tmp_df_action_list['action_id'].astype(str).to_list()
        # Параметры запроса
        params_actions_info = {'promotionIDs': action_ids_list}
        # Делаем запрос
        resp_data_actions_info = requests.get("https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions/details",
                                              headers=headers,
                                              params=params_actions_info).json()
        # Создаем df из ответа апи
        tmp_df_actions_info = pd.DataFrame(resp_data_actions_info['data']['promotions'])
        # Объединяем с предыдущим проходом цикла
        df_actions_info = pd.concat([df_actions_info, tmp_df_actions_info])

    # Переименовываем некоторые колонки для удобства
    df_actions_info = df_actions_info.rename(columns={
        'id': 'action_id',
        'name': 'action_title'
        })
    # Убираем акции, в которых нет товаров
    df_actions_list_final = df_actions_info.loc[df_actions_info['participationPercentage'] > 0, :]

    return df_actions_list_final


# Функция выгрузки списка товаров для участия в акции
def get_products_in_action(headers, df_action_list, inAction=True):
    # df, в который будем помещать список товаров в акциях
    df_products_actions = pd.DataFrame()
    # Лист с акциями, в которых были товары
    actions_with_products = []
    for i in range(df_action_list.shape[0]):
        # Переменные, определяющие переход к следующей итерации цикла (следующей акии)
        offset = 0
        nomenclatures = ['begin']
        while len(nomenclatures) > 0:
            # Параметры запроса
            params_in_action = {
                'promotionID': int(df_action_list['action_id'][i]),
                'inAction': inAction,
                'limit': 1000,
                'offset': offset
            }
            print(f"Выгрузка данных по акции {df_action_list['action_title'][i]}")
            resp_data_action_products = requests.get("https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions/nomenclatures", headers=headers, params=params_in_action)
            # Если есть данные по акции, получаем товары в этой акции
            if resp_data_action_products.status_code == 200:
                nomenclatures_api = resp_data_action_products.json()
                nomenclatures = nomenclatures_api['data']['nomenclatures']
                # Если пришел не пустой список товаров по акции, начинаем его обработку
                if len(nomenclatures) > 0:
                    print(f"Найдены товары по акции {df_action_list['action_id'][i]} {df_action_list['action_title'][i]}")
                    tmp_df = pd.DataFrame(nomenclatures)
                    # Добавляем в df id и имя акции
                    tmp_df = tmp_df.assign(
                        action_id=df_action_list['action_id'][i],
                        action_title=df_action_list['action_title'][i],
                        dt_start_action=df_action_list['startDateTime'][i],
                        dt_end_action=df_action_list['endDateTime'][i],
                    )
                    # Переименовываем колонку с ID артикула
                    tmp_df = tmp_df.rename(columns={'id': 'nmID'})
                    # Делим скидку на 100 для дальнейших расчетов
                    for col in ['discount','planDiscount']:
                        tmp_df[col] = tmp_df[col] / 100
                    # Объединяем с предыдущим проходом цикла
                    df_products_actions = pd.concat([df_products_actions, tmp_df], axis=0)
                    # Добавляем акцию в список акций, в которой были товары
                    action_name = df_action_list['action_title'][i] # df_action_list.iloc[i, df_action_list.columns.get_loc('action_title')].to_list()
                    actions_with_products.append(action_name)
                    # Увеличиваем offset на случай, если товаров больше 1000
                    offset = offset + len(nomenclatures)
            # Если нет данных по данной акции, переходим в цикле for к следующей акции
            else:
                # print(f"ОШИБКА выгрузки товаров по акции {df_action_list['action_title'][i]}: {resp_data_action_products.json()}")
                nomenclatures = []
        time.sleep(1)

    return df_products_actions, actions_with_products


# Функция объединения товаров, участвующих и не участвующих в акции
def merge_products_in_actions(df_products_in_action, df_products_not_in_action):
    df_products_actions_all = pd.concat([df_products_in_action, df_products_not_in_action], axis=0).reset_index(drop=True)
    return df_products_actions_all


# Функция разбивки товаров в акции по размерам
def unpack_sizes(df_products_actions_all, df_products):
    # Создаем копию для избежания изменений в оригинальном df
    # df_products_actions_all_ = df_products_actions_all.copy()
    # df_products_ = df_products.rename(columns={'price': 'price (prices api)',
    #                                            'discount': 'discount (prices api)',
    #                                            'discount_price': 'discount_price (prices api)'})
    # Мерджим список товаров в акции со списком товаров для разбивки по размерам
    df_products_actions_with_sizes = df_products_actions_all.merge(df_products[['nmID', 'Артикул продавца', 'Размер']],
                                                                   how='left',
                                                                   on='nmID')
    return df_products_actions_with_sizes


# Функция чтения и обработки справочной таблицы
def read_catalog():
    # Чтение справочной таблицы
    catalog = pd.read_excel(f"Clients/{client_name}/catalog/Справочная_таблица_{client_name}_WB.xlsx")
    # Удаляем товары, где не указан размер
    catalog_ = catalog.dropna(subset=['Размер'])
    # Если каких-то колонок не хватает, искуственно создаем их, чтобы не было ошибок
    for col in catalog_action_columns:
        if col not in catalog_.columns:
            catalog_[col] = np.nan
    # У некоторых клиентов делаем размер строковым типом
    if client_name in ['KU_And_KU', 'Soyuz']:
        catalog_['Размер'] = catalog_['Размер'].astype(str).str.replace('.0', '')
        catalog_['Размер'] = catalog_['Размер'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        catalog_['Размер'] = catalog_['Размер'].fillna(0)

    return catalog_


# Функция добавления столбца с РРЦ
def add_marketing_price_to_actions(catalog, df_products_actions_with_sizes):
    # Мерджим список товаров в акции, разбитый по размерам, со справочной таблицей
    df_products_actions_with_marketing_price = df_products_actions_with_sizes.merge(catalog[['Артикул продавца', 'Размер', 'РРЦ', 'Себестоимость']],
                                                                            on=['Артикул продавца', 'Размер'],
                                                                            how='left')

    return df_products_actions_with_marketing_price


# Получение колонок из файла с метриками
def add_columns_from_metrics(df_products):
    df_products_ = df_products.copy()
    # Чтение файла с метриками
    metrics = pd.read_excel(f"Clients/{client_name}/Metrics/{date_report_created}_МетрикиИтоги.xlsx", sheet_name='summary')
    if client_name in ['KU_And_KU', 'Soyuz']:
        metrics['Размер'] = metrics['Размер'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        metrics['Размер'] = metrics['Размер'].fillna(0)
        df_products_['Размер'] = df_products_['Размер'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_products_['Размер'] = df_products_['Размер'].fillna(0)
    df_products_with_metrics = df_products_.merge(metrics[['Артикул продавца', 'Размер', 'Остатки']],
                              how='left',
                              on=['Артикул продавца', 'Размер'])
    return df_products_with_metrics

# Получение колонок из справочной таблицы
def add_columns_from_catalog(catalog, df_products_with_metrics):
    # Мерджим со списком товаров с метриками
    df_products_with_catalog = df_products_with_metrics.merge(catalog[catalog_action_columns],
        # catalog[[col for col in catalog_action_columns if col not in ['Себестоимость']]],
                                                              on=['Артикул продавца', 'Размер'],
                                                              how='left')
    df_products_with_catalog['Скидка до РРЦ'] = (df_products_with_catalog['price'] - df_products_with_catalog['РРЦ']) / df_products_with_catalog['РРЦ']
    return df_products_with_catalog


# Функция расчета столбцов по акциям
def calc_action_columns(df_products_actions_with_marketing_price, df_products_with_catalog):
    # Создаем копию для избежания изменений в оригинальном df
    df_products_actions_with_marketing_price_ = df_products_actions_with_marketing_price.copy()
    df_products_for_actions = df_products_with_catalog.copy()
    # Для некоторых клиентов переводим размер в число
    if client_name in ['KU_And_KU', 'Soyuz']:
        df_products_actions_with_marketing_price_['Размер'] = df_products_actions_with_marketing_price_['Размер'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_products_actions_with_marketing_price_['Размер'] = df_products_actions_with_marketing_price_['Размер'].fillna(0)
    # Переименовываем некоторые колонки в списке товаров
    df_products_for_actions = df_products_for_actions.rename(columns={
        'barcode': 'Баркод',
        'price': 'Цена до скидки',
        'discount_price': 'Цена после скидки',
        'discount': 'Скидка WB'
    })
    # Создаем доп. колонки участия в акции (Да или Нет) и скидки по акции
    df_products_actions_with_marketing_price_['Участие в акции'] = np.where(
        df_products_actions_with_marketing_price_['inAction'] == True,
        'Да',
        'Нет')
    # df_products_actions_with_marketing_price_['Скидка по акции'] = np.where(df_products_actions_with_marketing_price_['inAction'] == True,
    #                                                df_products_actions_with_marketing_price_['discount'],
    #                                                df_products_actions_with_marketing_price_['planDiscount'])
    df_products_actions_with_marketing_price_['Скидка по акции'] = df_products_actions_with_marketing_price_['planDiscount']
    df_products_actions_with_marketing_price_['Цена по акции'] = df_products_actions_with_marketing_price_['planPrice']

    # Получаем список акций
    df_actions_from_products = (
        df_products_actions_with_marketing_price_
        .groupby(['action_id', 'action_title', 'dt_start_action', 'dt_end_action'])
        .size()
        .reset_index()
        .drop(columns=0)
        .rename(columns={
            'action_title': 'Наименование акции',
            'dt_start_action': 'Дата начала акции',
            'dt_end_action': 'Дата окончания акции',
        })
    )
    # list, куда будем помещать названия колонок с акциями
    action_cols = []
    # Цикл по списку акций
    for i in range(df_actions_from_products.shape[0]):
        # Получаем ID и Название акции
        action_id = df_actions_from_products['action_id'][i]
        action_name = df_actions_from_products['Наименование акции'][i]
        # df с товарами, участвующими в акции и кандитатами в акцию
        tmp_df_action_products = (
            df_products_actions_with_marketing_price_
            # .rename(columns={
            #     'Участие': f"Участие в акции {action_name}",
            #     'Скидка по акции': f"Скидка по акции {action_name}"
            # })
            .loc[df_products_actions_with_marketing_price_['action_id'] == action_id, ['nmID', 'Артикул продавца', 'Размер', 'РРЦ', 'Себестоимость', 'Участие в акции', 'Цена по акции', 'Скидка по акции']]
        )
        # Считаем разницу с РРЦ
        tmp_df_action_products['Скидка от РРЦ по акции'] = (tmp_df_action_products['РРЦ'] - tmp_df_action_products['Цена по акции']) / tmp_df_action_products['РРЦ']
        # Удаляем колонку с РРЦ, её мы возьмем из справочной таблицы
        tmp_df_action_products = tmp_df_action_products.drop(columns='РРЦ')
        # Для клиента KU_And_KU применяем доп. логику расчета участия в акциях
        if client_name in ['KU_And_KU', 'Soyuz']:
            # Маржинальность по акциям
            tmp_df_action_products['Расчетная маржа, руб по акции'] = tmp_df_action_products['Цена по акции'] - (net_cost_koef * tmp_df_action_products['Цена по акции']) - (tmp_df_action_products['Себестоимость'])
            tmp_df_action_products['Расчетная маржа, % по акции'] = tmp_df_action_products['Расчетная маржа, руб по акции'] / tmp_df_action_products['Цена по акции']
        # Удаляем колонку с Себестоимостью, её мы возьмем из справочной таблицы
        tmp_df_action_products = tmp_df_action_products.drop(columns='Себестоимость')
        # Добавляем названия акций в колонки по акциям
        current_action_cols = []
        # В какие колонки добавляем названия акции
        action_cols_to_rename = ['Участие в акции', 'Цена по акции', 'Скидка по акции', 'Скидка от РРЦ по акции']
        # Для некоторых клиентов добавляем названия акций в колонки с маржинальностью
        if client_name in ['KU_And_KU', 'Soyuz']:
            action_cols_to_rename = action_cols_to_rename + ['Расчетная маржа, руб по акции', 'Расчетная маржа, % по акции']
        # Переименование
        for col in action_cols_to_rename:
            new_col_name = f"{col} {action_name}"
            tmp_df_action_products = tmp_df_action_products.rename(columns={col: new_col_name})
            action_cols.append(new_col_name)
            current_action_cols.append(new_col_name)
        # Мерджим со списком товаров
        df_products_for_actions = df_products_for_actions.merge(tmp_df_action_products[['nmID', 'Размер'] + current_action_cols],
                                                                on=['nmID', 'Размер'],
                                                                how='left')

    # Переименовываем некоторые колонки для удобства
    df_products_for_actions = df_products_for_actions.rename(columns={'nmID': 'Артикул WB'})
    # Делим скидку WB на 100 для перевода в проценты
    df_products_for_actions['Скидка WB'] = df_products_for_actions['Скидка WB'] / 100
    # Выбираем нужные колонки
    df_products_actions_svod = df_products_for_actions.loc[:, ['Артикул продавца', 'Предмет', 'Наименование товара', 'Баркод', 'Размер', 'Цвет',
                                                              'Цена до скидки', 'Цена после скидки', 'Скидка WB', 'РРЦ', 'Скидка до РРЦ',
                                                              'Минимальная цена расчетная, руб.', 'Себестоимость',
                                                              'Остатки'] + action_cols]

    return df_products_actions_svod, action_cols, df_actions_from_products


def create_svod_excel_actions(date_report_created, df_products_actions_svod, action_cols):
    # Создаем копию для избежания изменений в оригинальном df
    df_actions_svod = df_products_actions_svod.copy()
    # Переводим дату выгрузки файлов в datetime
    date_report_created_ = datetime.strptime(date_report_created, '%Y-%m-%d')
    # Определяем колонки, которые должны быть в отчете
    svod_columns = svod_actions_cols + action_cols + ['Итоговая цена', 'Итоговая скидка']
    # Если какой-то колонки нет в своде по акциям, создаем её
    for col in svod_columns:
        if col not in df_actions_svod.columns:
           df_actions_svod[col] = np.nan

    # Порядок колонок
    df_actions_svod = df_actions_svod[svod_columns]
    # Переименовываем некоторые колонки для соответствия шаблону
    date_for_columns = date_report_created_.strftime('%d.%m')
    date_for_columns_next = (date_report_created_ + timedelta(days=1)).strftime('%d.%m')
    df_actions_svod = df_actions_svod.rename(columns={
        'Минимальная цена расчетная, руб.': 'Min цена, руб.',
        'Скидка WB': f"Скидка WB {date_for_columns}",
        'Остатки': f"Ост. {date_for_columns}",
        'Остатки след. день': f"Ост. {date_for_columns_next}"
    })

    return df_actions_svod


# Функция сохранения результатов в excel
def save_excel(
        client_name,
        date_report_created,
        df_actions_svod_excel,
        actions_with_products
    ):

    with pd.ExcelWriter(f"Clients/{client_name}/Actions/{date_report_created}_Таблица_по_акциям_{client_name}_WB.xlsx") as w:
        df_actions_svod_excel.to_excel(w, sheet_name='Акции', index=False, na_rep='')
        actions_with_products.to_excel(w, sheet_name='Названия акций', index=False, na_rep='')


# %% Вызов всех функций
if __name__ == '__main__':
    # date_report_created = '2025-01-23'
    date_report_created = str(date.today())
    # date_start_actions = '2025-05-01T00:00:00Z'
    # date_end_actions = '2025-05-23T23:59:59Z'
    # Чтение дат, за которые была сделана выгрузка по апи
    report_dates = read_dates_file(date_report_created)
    # Чтение списка товаров
    # df_products = read_products_file(date_report_created)
    # Добавление колонок из справочной таблицы
    # catalog = get_columns_from_catalog(df_products)
    # # Расчет доп. колонок для акций
    # catalog_with_stats = calc_additional_columns(catalog)
    # # Добавление данных из файла с метриками
    # catalog_metrics = get_columns_from_metrics(date_report_created, catalog_with_stats)

    # Генерируем начальную и конечную дату для метода списка акций
    date_start_actions, date_end_actions = generate_date_range(
        start_date=str(date.today() - timedelta(days=30)),
        end_date=str(date.today())
    )

    # Получаем список акций
    df_action_list = get_action_list(headers, date_start_actions, date_end_actions)
    # Выгружаем список товаров, которые участвуют в акции
    df_products_in_action, actions_with_participating_products = get_products_in_action(headers, df_action_list, inAction=True)
    # Выгружаем список товаров, которые не участвуют в акции
    df_products_not_in_action, actions_without_participating_products = get_products_in_action(headers, df_action_list, inAction=False)
    # Объединяем товары, которые участвуют в акции и которые не участвуют в акции, в один df
    df_products_actions_all = merge_products_in_actions(df_products_in_action, df_products_not_in_action)
    # Выгружаем список товаров и цен
    df_products_list = getWBProduct(headers, type_products='not_from_recycle', to_save=False)
    df_products = get_prices_WB(headers, df_products_list, to_save=False)
    df_products_processed = process_product_list(df_products)

    # Разбиваем товары в акциях по размерам
    df_products_actions_with_sizes = unpack_sizes(df_products_actions_all, df_products_processed)
    # Чтение справочной таблицы
    catalog = read_catalog()
    # Добавляем РРЦ в таблицу товаров по акциям
    df_products_actions_with_marketing_price = add_marketing_price_to_actions(catalog, df_products_actions_with_sizes)
    # Добавляем столбцы из файла с метриками
    df_products_with_metrics = add_columns_from_metrics(df_products_processed)
    # Добавляем столбцы из справочной таблицы
    df_products_with_catalog = add_columns_from_catalog(catalog, df_products_with_metrics)
    # Рассчитываем колонки участия в акциях
    df_products_actions_svod, action_cols, actions_with_products = calc_action_columns(df_products_actions_with_marketing_price, df_products_with_catalog)
    # Создаем таблицу для excel
    df_actions_svod_excel = create_svod_excel_actions(date_report_created, df_products_actions_svod, action_cols)
    # Сохранение результатов в excel
    save_excel(
        client_name,
        date_report_created,
        df_actions_svod_excel,
        actions_with_products
    )
    # Форматирование файла с акциями excel
    format_excel_actions(
        client_name,
        df_actions_svod_excel,
        date_report_created
    )
