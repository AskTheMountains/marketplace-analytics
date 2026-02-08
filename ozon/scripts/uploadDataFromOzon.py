# %% Определение всех функций
import argparse
import csv
import json
from multiprocessing import current_process
import os
import re
import time
from ast import literal_eval
from datetime import date, datetime, timedelta
from colorama import Fore, Back, Style, init
from pathlib import Path

import pandas as pd
from io import BytesIO
import requests
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

from wb.scripts.create_sales_and_orders_svod import get_sales

pd.set_option('future.no_silent_downcasting', True)

# Функция парсинга аргумента количества дней выгрузки
from ozon.scripts.command_args_parser import cli_days_ago
# Файл с некоторыми константами
from ozon.scripts.constants import (
    client_name,
    headers,
    marketplace_dir_path,
    ozon_seller_api_url,
    api_columns
)
# Функция обновления кластеров Озон
from ozon.scripts.update_clusters import update_cluster_list
# Некоторые доп. функции
from generic_functions import (
    custom_pd_explode
)

# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Upload data for Ozon marketplace",
    usage=(
        "python.exe -m ozon.scripts.uploadDataFromOzon [days_ago] [--days_ago N] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.uploadDataFromOzon 30\n\n"
        "  python.exe -m ozon.scripts.uploadDataFromOzon --days_ago 10 --client 2\n\n"
        "  python.exe -m ozon.scripts.uploadDataFromOzon 7 -c 3"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


# Определяем имя директории для загрузки файлов
uploaddir = (
    marketplace_dir_path
    / "Clients"
    / client_name
    / "UploadFiles"
)
uploaddir_today = uploaddir / f"UploadFiles_{str(date.today())}"


# Функция создания директорий для текущей выгрузки
def create_upload_dir():
    # # Определяем имя директории для загрузки файлов
    # uploaddir = (
    #     marketplace_dir_path
    #     / "Clients"
    #     / client_name
    #     / "UploadFiles"
    # )
    # uploaddir_today = uploaddir / f"UploadFiles_{str(date.today())}"
    # Если директории нет, то создаем её
    if not uploaddir_today.exists():
        uploaddir_today.mkdir(parents=True, exist_ok=True)

    return uploaddir_today


def generateDates(days_ago = 30):
    date_end = str(date.today()- timedelta(days=1)) + 'T23:59:59.000Z'
    date_start = str(date.today()- timedelta(days=1) - timedelta(days_ago)) + 'T00:00:00.000Z'
    logger.info(f"Generated dates: {date_start} - {date_end}")

    return date_start, date_end #,date_end_file,date_start_file

# GPT START ----
# Функция создания диапазона дат с указанной разницей
def generate_dates_new(days_ago=30, start_date_str=None, end_date_str=None):
    """
    Возвращает кортеж из строк начальной и конечной дат в формате ISO8601 с учётом приоритета параметров.

    Приоритет параметров:
    1. days_ago — если задан, вычисляются даты по количеству дней назад, другие параметры игнорируются.
    2. start_date_str и/или end_date_str — используются, если days_ago не задан.
       Если какая-то из дат не указана, подставляются значения по умолчанию:
           - для end_date_str: сегодня, время 23:59:59
           - для start_date_str: end_date минус 30 дней, время 00:00:00
    По умолчанию, если ничего не задано, возвращается диапазон "последние 30 дней".
    """
    # Если явно задан days_ago, используем его для расчёта диапазона дат
    if days_ago is not None:
        # Конечная дата: вчера, 23:59:59
        end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0) - timedelta(days=1)
        # Начальная дата: на days_ago дней раньше end_date, время 00:00:00
        start_date = end_date - timedelta(days=days_ago) + timedelta(days=1)
        start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        # Если хотя бы одна из дат передана явно, работаем с ними
        # Конечная дата: используем переданную или по умолчанию сегодня, 23:59:59
        if end_date_str:
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
            end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=0)
        else:
            end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
        # Начальная дата: используем переданную или по умолчанию минус 30 дней к end_date, 00:00:00
        if start_date_str:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            start_date = end_date - timedelta(days=30)
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Переводим к часовому поясу МСК
    # start_date_msk = msk_boundary_to_utc(
    #     start_date.strftime('%Y-%m-%dT%H:%M:%S')
    # )
    # end_date_msk = msk_boundary_to_utc(
    #     end_date.strftime('%Y-%m-%dT%H:%M:%S'),
    # )
    # Форматируем полученные даты в строковый вид ISO8601 (UTC с 'Z' на конце)
    start_date_str = start_date.strftime('%Y-%m-%dT%H:%M:%SZ')
    end_date_str = end_date.strftime('%Y-%m-%dT%H:%M:%SZ')

    return start_date_str, end_date_str
# GPT END ----


# Сохранение дат текущей выгрузки в csv
def save_dates_to_csv(date_start, date_end):
    # Формат дат для названий файла
    date_end_file = datetime.fromisoformat(date_end).strftime('%Y-%m-%d')
    date_start_file = datetime.fromisoformat(date_start).strftime('%Y-%m-%d')
    # Датафрейм с датами начала и окончания выгрузки данных
    df_dates = pd.DataFrame({'date_end': date_end,
                             'date_start': date_start,
                             'date_end_file': date_end_file,
                             'date_start_file': date_start_file},
                             index=[0])
    path_save = uploaddir_today / f"{str(date.today())}_dates_from_to.csv"
    df_dates.to_csv(path_save, sep=';', index=False, encoding='utf-8-sig')


# Отчет по товарам (список товаров)
def getOzonProduct(headers, to_save=True):
    # Параметры запроса на создание отчета по товарам
    params = json.dumps({
    "language": "DEFAULT",
    "offer_id": [],
    "search": "",
    "sku": [],
    "visibility": "ALL"
    })

    logger.info(f'Выгрузка списка товаров')

    # Отправляем запроса на создание отчета по товарам
    logger.info('Получаем ссылку на отчет по товарам')
    result = requests.post(f"{ozon_seller_api_url}/v1/report/products/create",headers=headers, data=params).json()
    code_id = json.dumps({
    "code": result["result"]["code"]
    })
    logger.info(f'Получен код отчета по товарам: {result["result"]["code"]}')

    # Ждем 5 секунд, пока отчет будет создаваться
    time.sleep(5)
    resp_data = {"result":{"file":""}}
    # Делаем запросы получения информации об отчете до тех пор, пока не получим ссылку на отчет
    while not resp_data["result"]["file"]:
        resp_data = requests.post(f"{ozon_seller_api_url}/v1/report/info", headers=headers,data=code_id).json()
        logger.info('Отчет по товарам еще не готов, ждем 5 секунд...')

    # Получаем ссылку на отчет по товарам
    path_csv = resp_data["result"]["file"]
    logger.info('Отчет по товарам готов, получаем файл')

    r = requests.get(path_csv, stream=True)
    if r.status_code == 200:
        # Считываем отчет по товарам (список товаров)
        df_products = pd.read_csv(path_csv, sep=';')
        # Убираем лишние символы из артикула
        df_products['Артикул'] = df_products['Артикул'].str.replace("'", "", regex=False)
        if to_save:
            file_name_products = uploaddir_today / f"{str(date.today())}_Товары.csv"
            df_products.to_csv(file_name_products, sep=';', index=False, encoding='utf-8-sig')
        else:
            return df_products
    else:
        logger.error(f'Ошибка при получении отчета по товарам, info: {r.text}')


# Информация о товарах
def get_products_info(headers, df_products=None):
    # Если список товаров не задан, считываем его из текущей выгрузки
    if df_products is None:
        logger.info("Получаем доп. информацию о товарах (задан локальный файл)")
        df_products_ = pd.read_csv(f"{uploaddir_today}/{str(date.today())}_Товары.csv", sep=';')
    # Если список товаров задан, берем его из переданного аргумента
    else:
        logger.info("Получаем доп. информацию о товарах (передан список товаров)")
        df_products_ = df_products.copy()

    # Разбиваем список товаров на диапазоны по 1000 товаров
    df_products_ = df_products_.reset_index(drop=True)
    step = 1000
    df_products_['chunks'] = df_products_.index // step + 1
    # Создаем переменные для прогресса выгрузки информации о товарах
    total_products = df_products_.shape[0]
    total_chunks = df_products_['chunks'].nunique()
    uploaded_products = 0

    # Убираем скрытые товары
    df_hidden_products = df_products_.loc[df_products_['SKU'] <= 0, :]
    df_products_ = df_products_.loc[df_products_['SKU'] > 0, :]
    if not df_hidden_products.empty:
        logger.info(f"Найдено {len(df_hidden_products)} скрытых товаров")

    # df для информации о товарах
    df_products_info_api = pd.DataFrame()

    # Цикл по каждому диапазону
    for chunk in df_products_['chunks'].unique():
        # Делаем выборку первой 1000 товаров
        product_ids = (
            df_products_
            .loc[df_products_['chunks'] == chunk, 'Артикул']
            .astype(str)
            .to_list()
        )
        # Формируем параметры запроса
        params = json.dumps({
            "offer_id": product_ids,
            # "visibility": "ALL"
        })

        current_progress = min(uploaded_products + step, total_products)

        logger.info(
            f"Выгружаем информацию о товарах: "
            f"{current_progress}/{total_products} товаров (чанк {chunk} из {total_chunks})"
        )

        # Делаем запрос
        resp_data_products_info = requests.post(f"{ozon_seller_api_url}/v3/product/info/list", headers=headers, data=params)
        # Если запрос прошел успешно, формируем df из ответа АПИ
        if resp_data_products_info.status_code == 200:
            tmp_df_products_info = pd.DataFrame(resp_data_products_info.json()['items'])
            # Объединяем с предыдущим чанком товаров
            df_products_info_api = pd.concat([df_products_info_api, tmp_df_products_info])
            # Обновляем прогресс выгрузки
            uploaded_products += step
        else:
            logger.error(
                f"Ошибка получения информации о товарах: "
                f"чанк {chunk} из {df_products_['chunks'].max()}, инфо: {resp_data_products_info.text}"
            )

    # Сбрасываем index после concat
    df_products_info_api = df_products_info_api.reset_index(drop=True)

    # Создаем колонку с признаком Super-товара
    df_products_info_api['Super-товар'] = (
        df_products_info_api['is_super']
        .apply(lambda x: 'Да' if x else 'Нет')
    )
    # Выбираем и переименовываем нужные колонки
    df_products_info = (
        df_products_info_api
        .loc[:, ['offer_id', 'Super-товар']]
        .rename(columns={
            'offer_id': 'Артикул',
            # 'is_super': 'Super-товар'
        })
    )

    return df_products_info

# Цены на товары
def get_products_prices(headers, df_products=None):
    # Если список товаров не задан, считываем его из текущей выгрузки
    if df_products is None:
        logger.info("Получаем цены на товары (задан локальный файл)")
        df_products_ = pd.read_csv(f"{uploaddir_today}/{str(date.today())}_Товары.csv", sep=';')
    # Если список товаров задан, берем его из переданного аргумента
    else:
        logger.info("Получаем цены на товары (передан список товаров)")
        df_products_ = df_products.copy()

    if 'РРЦ' in df_products_.columns:
        df_products_ = df_products_.drop(columns=['РРЦ'])
    # df_products['Артикул'] = df_products['Артикул'].str.replace("'", "", regex=False)
    # df для цен
    df_products_prices = pd.DataFrame()
    # Разбиваем df с товарами на равные части по 1000 товаров
    df_products_ = df_products_.reset_index(drop=True)
    step = 1000
    df_products_['chunks'] = df_products_.index // step + 1
    last_id = ''

    # Создаем переменные для отслеживания прогресса
    total_products = df_products_.shape[0]
    total_chunks = df_products_['chunks'].nunique()
    uploaded_products = 0


    for chunk in df_products_['chunks'].unique():
        # Выборка 1000 id товаров из df с товарами
        product_list = df_products_.loc[df_products_['chunks'] == chunk, 'Ozon Product ID'].to_list()
        product_list_string = [str(element) for element in product_list]
        # Передача списка товаров в параметры запроса
        params = json.dumps({
            "filter": {
                "offer_id": [],
                "product_id": product_list_string,
                "visibility": "ALL"
                },
            "cursor": last_id,
            "limit": 1000
        })

        current_progress = min(uploaded_products + step, total_products)

        logger.info(
            f"Выгружаем цены на товары: "
            f"{current_progress}/{total_products} товаров (чанк {chunk} из {total_chunks})"
        )

        resp_data = requests.post(f"{ozon_seller_api_url}/v5/product/info/prices", headers=headers,data=params)

        if resp_data.status_code != 200:
            logger.error(
                f"Ошибка получения цен на товары: "
                f"чанк {chunk} из {df_products_['chunks'].max()}, инфо: {resp_data.text}"
            )
            continue
        else:
            resp_data = resp_data.json()
            # по last_id будет выгружаться следующая страница в следующем проходе цикла
            last_id = resp_data['cursor']
            # В промежуточный df складываем данные текущей страницы
            tmp_df = pd.DataFrame(resp_data['items'])
            # Добавляем строки предыдущей страницы к текущей (по ходу цикла)
            df_products_prices = pd.concat([df_products_prices, tmp_df])

        # Обновляем прогресс выгрузки
        uploaded_products += step

    # Сбрасываем index после concat
    df_products_prices = df_products_prices.reset_index(drop=True)
    # Получаем список цен
    df_products_prices_all = pd.concat([
        df_products_prices.loc[:, ~df_products_prices.columns.isin(['price'])],
        pd.json_normalize(df_products_prices['price']),
    ], axis=1)
    # Переименовываем цены в соответствии с документацией
    df_products_prices_all = df_products_prices_all.rename(columns={
        'offer_id': 'Артикул',
        'product_id': 'Ozon Product ID',
        'marketing_price': 'Цена с учетом всех акций',
        'marketing_seller_price': 'Цена с учетом акций продавца',
        'min_price': 'Минимальная цена после применения всех скидок',
        'old_price': 'Цена до учета скидок (зачеркнутая)',
        'price': 'Цена с учетом скидок (на карточке товара)',
        'retail_price': 'Цена поставщика'
    })
    # Выбираем нужные колонки
    df_products_prices_all = df_products_prices_all.loc[:, df_products_prices_all.columns.isin([
        'Артикул',
        # 'Ozon Product ID',
        'Цена с учетом всех акций',
        'Цена с учетом акций продавца',
        'Минимальная цена после применения всех скидок',
        'Цена до учета скидок (зачеркнутая)',
        'Цена с учетом скидок (на карточке товара)',
        'Цена поставщика'
    ])]
    # Объединяем с df товаров
    df_products_with_prices = (
        df_products_
        .loc[:, df_products_.columns.isin(['Артикул'])]
        .merge(
            df_products_prices_all,
            how='left',
            on='Артикул'
        )
    )
    # Если стоит флаг сохранения, то сохраняем файл
    # if to_save:
    #     df_products_with_prices.to_csv(f"{uploaddir_today}/{str(date.today())}_Товары.csv", sep=';', index=False)
    # # иначе: возвращаем df с товарами и ценами
    # else:
    return df_products_with_prices


# Характеристики товаров
def get_products_attributes(headers, df_products=None):

    # Если список товаров не задан, считываем его из текущей выгрузки
    if df_products is None:
        logger.info("Получаем характеристики товаров (задан локальный файл)")
        df_products_ = pd.read_csv(f"{uploaddir_today}/{str(date.today())}_Товары.csv", sep=';')
    # Если список товаров задан, берем его из переданного аргумента
    else:
        logger.info("Получаем характеристики товаров (передан список товаров)")
        df_products_ = df_products.copy()

    # df, в который будем помещать результаты метода АПИ
    df_products_attributes = pd.DataFrame()

    # Разбиваем df с товарами на равные части по 100 товаров
    df_products_ = df_products_.reset_index(drop=True)
    step = 1000
    df_products_['chunks'] = df_products_.index // step + 1
    # Пагинация
    last_id = ''

    # Создаем переменные для отслеживания прогресса
    total_products = df_products_.shape[0]
    total_chunks = df_products_['chunks'].nunique()
    uploaded_products = 0

    # Цикл по каждому интервалу
    for chunk in df_products_['chunks'].unique():
        # Выборка 1000 id товаров из df с товарами
        product_list = df_products_.loc[df_products_['chunks'] == chunk, 'Ozon Product ID'].to_list()
        product_list_string = [str(element) for element in product_list]
        # Передача списка товаров в параметры запроса
        params = json.dumps({
            "filter": {
                "offer_id": [],
                "product_id": product_list_string,
                "sku": [],
                "visibility": "ALL"
                },
            "last_id": last_id,
            "sort_by": 'offer_id',
            "sort_id": 'asc',
            "limit": 1000
        })

        current_progress = min(uploaded_products + step, total_products)
        logger.info(
            f"Выгружаем характеристики товаров: "
            f"{current_progress}/{total_products} товаров (чанк {chunk} из {total_chunks})"
        )

        resp_data = requests.post(f"{ozon_seller_api_url}/v4/product/info/attributes", headers=headers,data=params)
        if resp_data.status_code != 200:
            logger.error(
                f"Ошибка получения характеристик товаров: "
                f"чанк {chunk} из {df_products_['chunks'].max()}, инфо: {resp_data.text}"
            )
            continue
        else:
            resp_data = resp_data.json()
            # по last_id будет выгружаться следующая страница в следующем проходе цикла
            last_id = resp_data['last_id']
            # В промежуточный df складываем данные текущей страницы
            tmp_df = pd.DataFrame(resp_data['result'])
            # Добавляем строки предыдущей страницы к текущей (по ходу цикла)
            df_products_attributes = pd.concat([df_products_attributes, tmp_df])

        # Обновляем прогресс выгрузки
        uploaded_products += step

    # Сбрасываем index после concat
    df_products_attributes = df_products_attributes.reset_index(drop=True)

    # Переименовываем колонки для удобства
    df_products_attributes = df_products_attributes.rename(columns={
        'name': 'Название товара',
        'offer_id': 'Артикул',
        'sku': 'SKU',
        'id': 'Ozon Product ID',
        'height': 'Высота',
        'depth': 'Глубина',
        'width': 'Ширина',
        'dimension_unit': 'Единицы измерения габаритов',
        'weight': 'Вес',
        'weight_unit': 'Единицы измерения веса',
    })
    # Получаем информацию о количестве моделей в артикуле
    df_products_model_info = (
        pd.concat([
            df_products_attributes,
            pd.json_normalize(df_products_attributes['model_info']), # Достаем характеристики из словаря в df
            ],
            axis=1
        )
        .rename(columns={
            'count': 'model_count'
        })
        .loc[:,[
                 'Артикул',
                 'model_id',
                 'model_count',
                 'Высота',
                 'Глубина',
                 'Ширина',
                 'Единицы измерения габаритов',
                 'Вес',
                 'Единицы измерения веса',
                 'type_id',
                 'description_category_id',
                 'attributes'
        ]]
    )

    # Получаем df с характеристиками товара
    df_attributes_all = (
            df_products_model_info
            # Выбираем нужные колонки
            .loc[:, ['Артикул', 'model_id', 'model_count', 'type_id', 'description_category_id', 'attributes']]
            # Распаковываем столбец со словарями характеристик товаров
            .explode(['attributes'])
            # Делаем сброс индекса после распаковки
            .reset_index(drop=True)
            # Объединяем df со списком товаров и df с характеристиками
            .pipe(lambda df:
                  pd.concat([
                        df, #.loc[:, ['Артикул']],
                        pd.json_normalize(df['attributes']), # Достаем характеристики из словаря в df
                    ],axis=1)
            )
        # Переименовываем столбец с id характеристики
        .rename(columns={
            'id': 'attribute_id'
        })
        # Распаковываем столбец со словарями значений характеристик товаров
        .explode(['values'])
        # Делаем сброс индекса после распаковки
        .reset_index(drop=True)
        # Объединяем df со списком товаров и df со значениями характеристик товаров
        .pipe(lambda df:
              pd.concat([
                    df, #.loc[:, ['Артикул']],
                    pd.json_normalize(df['values']), # Достаем характеристики из словаря в df
                ], axis=1)
        )
        # Переименовываем столбец со значением характеристики товара
        .rename(columns={
            'value': 'attribute_value'
        })
    )

    # Создаем словарь характеристик товаров
    product_attributes_dict = {
        'Цвет': 10096,
        'Размер': 9533,
        # 'Длина мм': 9802,
        # 'Высота мм': 6606,
        # 'Ширина мм': 6605,
        # 'Вес товара г': 4383,
    }

    # df, в который будем добавлять характеристики товаров
    df_products_attributes_final = df_products_model_info.copy()

    # Добавляем характеристики в df с товарами
    for attribute_name, attribute_id in product_attributes_dict.items():
        tmp_df_attributes = (
            df_attributes_all
            # Выбираем из df с характеристиками нужную характеристику товара
            .loc[df_attributes_all['attribute_id'] == attribute_id, ['Артикул', 'attribute_value']]
            # Переименовываем столбец со значением характеристики в соответствии со словарем выше
            .rename(columns={
                'attribute_value': attribute_name
            })
            # Делаем группировку по артикулу для избежания дубликатов
            .groupby(['Артикул'])
            # Если на один товар несколько характеристик, делаем их в одну строку,
            # разделяя точкой с запятой
            .agg({
                attribute_name: ';'.join
            })
            # Достаем столбец с артикулом из индекса после группировки
            .reset_index()
        )
        # Мерджим со списком товаров
        df_products_attributes_final = df_products_attributes_final.merge(
            tmp_df_attributes,
            how='left',
            on='Артикул'
        )
        # Заполняем пропуски в атрибутах
        df_products_attributes_final[attribute_name] = df_products_attributes_final[attribute_name].fillna(0)

    # Определяем список нужных колонок для итогового DataFrame
    required_columns = [
        'Артикул', 'model_id', 'model_count', 'Высота', 'Глубина', 'Ширина',
        'Единицы измерения габаритов', 'Вес', 'Единицы измерения веса'
    ] + list(product_attributes_dict.keys())

    # Выбираем нужные колонки
    df_products_attributes_final = (
        df_products_attributes_final
        .loc[:, required_columns]
    )

    # # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    # if to_save:
    #     df_products_with_attributes.to_csv(f"{uploaddir_today}/{str(date.today())}_Товары.csv", sep=';', index=False)
    # # иначе - возвращаем df с товарами
    # else:

    # return  df_products_model_info
    return df_products_attributes_final


# Описание характеристик товаров (не используется, вспомогательно для справки)
def get_attributes_description(headers, df_attributes_all):

    logger.info("Выгружаем описания характеристик товаров")

    # Создаем уникальные значение id описание и id типа товара
    df_attributes_descriptions_ids = (
        df_attributes_all
        .loc[:, ['type_id', 'description_category_id']]
        .copy()
        .drop_duplicates(keep='first')
        .reset_index(drop=True)
    )

    # df, куда будем помещать описания характеристик
    df_attributes_descriptions_api = pd.DataFrame()

    # Переменная для прогресса выгрузки
    total_descriptions = df_attributes_descriptions_ids.shape[0]
    uploaded_descriptions = 0

    # Цикл по каждой характеристике
    for idx, row in df_attributes_descriptions_ids.iterrows():
        # Выбираем id типа товара
        type_id = int(row['type_id'])
        # Выбираем id описания характеристики
        description_category_id = int(row['description_category_id'])
        # Формируем параметры запроса
        params_category_descriptions = json.dumps({
            'description_category_id':  description_category_id,
            'type_id': type_id
        })

        current_progress = uploaded_descriptions + 1

        logger.info(
            f"Выгружаем описание характеристики {current_progress} из {total_descriptions}"
        )

        # Отправляем запрос
        resp_data_description_category = requests.post(
            f'{ozon_seller_api_url}/v1/description-category/attribute',
            headers=headers,
            data=params_category_descriptions
        )
        # Если запрос прошел успешно, создаем df из результата запроса
        if resp_data_description_category.status_code == 200:
            tmp_df_category_description = pd.DataFrame(resp_data_description_category.json()['result'])
            # Объединяем с предыдущим проходом цикла
            df_attributes_descriptions_api = pd.concat([
                df_attributes_descriptions_api,
                tmp_df_category_description
            ], ignore_index=True)

        # Если запрос прошел с ошибкой, выводим текст ошибки:
        else:
            error_text = resp_data_description_category.text
            logger.error(f"Ошибка получения описания характеристики товаров, инфо: {error_text}")
            continue

        # Обновляем прогресс выгрузки
        uploaded_descriptions += 1



# Функция объединения различных данных по товарам
def merge_products_data(
        df_products,
        df_products_info,
        df_products_with_attributes,
        df_products_with_prices
        ):
    # Создаем список df для джойна
    dfs_to_join = [
        df_products_info,
        df_products_with_attributes,
        df_products_with_prices
    ]
    # df, в который будем помещать результаты джойна
    df_products_all_data = df_products.copy()
    # Джойним датафреймы в цикле
    for df in dfs_to_join:
        df_products_all_data = df_products_all_data.merge(df, how='left', on='Артикул')

    return df_products_all_data


# Функция выгрузки списка товаров и их характеристик
def get_ozon_product(headers, to_save=True):
    # Выгрузка списка товаров
    df_products = getOzonProduct(headers, to_save=False)
    # Получение информации о товарах
    df_products_info = get_products_info(headers, df_products)
    # Получение цен товаров
    df_products_with_prices = get_products_prices(headers, df_products)
    # Получение характеристик товаров
    df_products_with_attributes = get_products_attributes(headers, df_products)
    # Объединение данных по товарам в один df
    df_products_all_data = merge_products_data(
        df_products,
        df_products_info,
        df_products_with_attributes,
        df_products_with_prices,
    )
    # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    if to_save:
        file_name_products = uploaddir_today / f"{str(date.today())}_Товары.csv"
        df_products_all_data.to_csv(
            file_name_products,
            sep=';',
            encoding='utf-8-sig',
            index=False
        )
    # иначе - возвращаем df с товарами
    else:
        return df_products_all_data


# Отчет по отправлениям (заказы)
def getOrders(headers, date_start, date_end, delivery_schema="fbo", to_save=True):

    # забираем наименование отчета с заказами за установленный период
    params = json.dumps({
        "filter": {
            "processed_at_from": date_start,
            "processed_at_to": date_end,
            "delivery_schema": [
                delivery_schema

            ],
            "sku": [],
            "cancel_reason_id": [],
            "offer_id": "",
            "status_alias": [],
            "statuses": [],
            "title": ""
        },
        # "with": {
        #     "customer_data": True
        # },
        "language": "DEFAULT"
    })

    logger.info(f'Запрос отчета по заказам: Даты: {date_start} - {date_end}, Схема доставки: {delivery_schema.upper()}')

    result = requests.post(f"{ozon_seller_api_url}/v1/report/postings/create",headers=headers, data=params).json()
    code_id = json.dumps({
    "code": result["result"]["code"]
    })
    logger.info(f'Получен код отчета по заказам: {result["result"]["code"]}')
    time.sleep(5)

    # забираем ссылку на отчет с заказами
    resp_data = {"result":{"file":""}}
    while not resp_data["result"]["file"]:
        resp_data = requests.post(f"{ozon_seller_api_url}/v1/report/info", headers=headers,data=code_id).json()
        logger.info('Отчет по заказам еще не готов, ждем 5 секунд...')
        time.sleep(5)
    logger.info('Отчет по заказам готов, получаем ссылку на файл')
    path_csv = resp_data["result"]["file"]

    # Считываем отчет по заказам в датафрейм
    r = requests.get(path_csv, stream=True)
    if r.status_code == 200:
        df_orders = pd.read_csv(path_csv, sep=';')
    else:
        logger.error(f'Ошибка при получении отчета по заказам, info: {r.text}')

    # Добавляем схему доставки в датафрейм
    df_orders['Схема доставки'] = delivery_schema.upper()

    # Если стоит флаг сохранения, то сохраняем файл
    if to_save:
        # Задаем имя файла для заказов
        file_name_orders = (
            uploaddir_today /
            f"{str(date.today())}_Заказы_temp_{delivery_schema}.csv"
        )
        df_orders.to_csv(
            file_name_orders,
            sep=';',
            encoding='utf-8-sig',
            index=False
        )
            # with open(f"{uploaddir_today}/{str(date.today())}_Заказы_temp_{delivery_schema}.csv", 'wb') as f:
            #     r.raw.decode_content = True
            #     shutil.copyfileobj(r.raw, f)
            # f.close()
    # Если флаг сохранения False, возвращаем df с заказами
    else:
        return df_orders


# Дополнительная информация по отправлениям
def add_orders_info_fbo(
        headers,
        date_start,
        date_end,
        df_orders_fbo
    ):

    logger.info('Выгружаем доп. информацию по отправлениям FBO')

    # Выгружаем доп. информацию по отправлениям FBO
    df_all_orders_fbo = get_ozon_fbo_list(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    # Получаем товары в заказах FBO
    df_fbo_orders_processed = process_fbo_list(
        df_all_orders_fbo=df_all_orders_fbo,
    )
    # Удаляем дубликаты по номеру отправления
    df_fbo_orders_processed = df_fbo_orders_processed.drop_duplicates(
        subset=['Номер отправления'],
        keep='first'
    ).reset_index(drop=True)

    # Добавляем информацию об отправлениях к отчету по заказам
    columns_fbo_info = [
        'Номер отправления',
        'Кластер отправления',
        'Кластер доставки',
    ]
    df_orders_fbo_enriched = (
        df_orders_fbo
        .merge(
            df_fbo_orders_processed[columns_fbo_info],
            how='left',
            on='Номер отправления',
            indicator='Наличие в списке отправлений'
        )
    )
    # Находим отправления, по которым не удалось найти доп. информацию
    df_missing_fbo_postings = df_orders_fbo_enriched.loc[
        df_orders_fbo_enriched['Наличие в списке отправлений'] == 'left_only',
        :
    ]
    # Если такие отправления есть, выгружаем по ним информацию отдельно
    if not df_missing_fbo_postings.empty:
        missing_posting_numbers = df_missing_fbo_postings['Номер отправления'].tolist()
        logger.info(
            f"Не удалось найти доп. информацию по {len(missing_posting_numbers)} отправлениям FBO"
        )
        # Выгружаем доп. информацию по отсутствующим отправлениям
        df_missing_fbo_info = get_ozon_fbo_info(
            headers=headers,
            posting_numbers=missing_posting_numbers
        )
        # Обрабатываем полученный список отсутствующих отправлений
        df_missing_fbo_info_processed = process_fbo_list(
            df_all_orders_fbo=df_missing_fbo_info,
        )
        # Удаляем дубликаты по номеру отправления
        df_missing_fbo_info_processed = df_missing_fbo_info_processed.drop_duplicates(
            subset=['Номер отправления'],
            keep='first'
        ).reset_index(drop=True)

        # Добавляем отсутствующую информацию в df с пропущенными отправлениями
        df_missing_fbo_postings_enriched = (
            df_missing_fbo_postings
            .drop(columns=['Кластер отправления', 'Кластер доставки'])
            .merge(
                df_missing_fbo_info_processed[columns_fbo_info],
                how='left',
                on='Номер отправления'
            )
        )
        # Находим отправления, у которых изначально был кластер
        df_fbo_postings_not_missing = df_orders_fbo_enriched.loc[
            df_orders_fbo_enriched['Наличие в списке отправлений'] == 'both',
            :
        ]
        # Объединяем два df обратно
        df_orders_fbo_enriched = pd.concat([
            df_fbo_postings_not_missing,
            df_missing_fbo_postings_enriched
        ]).reset_index(drop=True)
    else:
        logger.success("Вся доп. информация по отправлениям FBO найдена в основном списке")

    return df_orders_fbo_enriched

# Функция выгрузки заказов FBO с кластерами
def getOrdersFBO(
        headers: dict[str, str],  # словарь заголовков HTTP
        date_start: str,          # ISO‑8601 начало периода
        date_end: str,            # ISO‑8601 конец периода
        add_clusters: bool = True,  # флаг добавления кластеров
        to_save: bool = True      # флаг сохранения файла
):

    # Выгружаем заказы FBO
    df_orders_fbo = getOrders(
        headers,
        date_start,
        date_end,
        delivery_schema="fbo",
        to_save=False
    )
    # Добавляем кластеры, если стоит флаг
    if add_clusters:
        df_orders_fbo = add_orders_info_fbo(
            headers,
            date_start,
            date_end,
            df_orders_fbo
        )
    # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    if to_save:
        file_name_orders_fbo = (
            uploaddir_today /
            f"{str(date.today())}_Заказы_fbo.csv"
        )
        df_orders_fbo.to_csv(
            file_name_orders_fbo,
            sep=';',
            encoding='utf-8-sig',
            index=False
        )
    # иначе - возвращаем df с заказами FBO
    else:
        return df_orders_fbo


# Дополнительная информация по отправлениям
def add_orders_info_fbs(
        headers,
        date_start,
        date_end,
        df_orders_fbs
    ):

    logger.info('Выгружаем доп. информацию по отправлениям FBS')

    # Выгружаем доп. информацию по отправлениям FBS
    df_all_orders_fbs = get_ozon_fbs_list(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    # Получаем товары в заказах FBS
    df_fbs_orders_processed = process_fbs_list(
        df_all_orders_fbs=df_all_orders_fbs,
    )
    # Удаляем дубликаты по номеру отправления
    df_fbs_orders_processed = df_fbs_orders_processed.drop_duplicates(
        subset=['Номер отправления'],
        keep='first'
    ).reset_index(drop=True)

    # Добавляем информацию об отправлениях к отчету по заказам
    columns_fbs_info = [
        'Номер отправления',
        'Кластер отправления',
        'Кластер доставки',
    ]
    df_orders_fbs_enriched = (
        df_orders_fbs
        .merge(
            df_fbs_orders_processed[columns_fbs_info],
            how='left',
            on='Номер отправления',
            indicator='Наличие в списке отправлений'
        )
    )
    # Находим отправления, по которым не удалось найти доп. информацию
    df_missing_fbs_postings = df_orders_fbs_enriched.loc[
        df_orders_fbs_enriched['Наличие в списке отправлений'] == 'left_only',
        :
    ]
    # Если такие отправления есть, выгружаем по ним информацию отдельно
    if not df_missing_fbs_postings.empty:
        missing_posting_numbers = df_missing_fbs_postings['Номер отправления'].tolist()
        logger.info(
            f"Не удалось найти доп. информацию по {len(missing_posting_numbers)} отправлениям FBS"
        )
        # Выгружаем доп. информацию по отсутствующим отправлениям
        df_missing_fbs_info = get_ozon_fbs_info(
            headers=headers,
            posting_numbers=missing_posting_numbers
        )
        # Обрабатываем полученный список отсутствующих отправлений
        df_missing_fbs_info_processed = process_fbs_list(
            df_all_orders_fbs=df_missing_fbs_info,
        )
        # Удаляем дубликаты по номеру отправления
        df_missing_fbs_info_processed = df_missing_fbs_info_processed.drop_duplicates(
            subset=['Номер отправления'],
            keep='first'
        ).reset_index(drop=True)

        # Добавляем отсутствующую информацию в df с пропущенными отправлениями
        df_missing_fbs_postings_enriched = (
            df_missing_fbs_postings
            .drop(columns=['Кластер отправления', 'Кластер доставки'])
            .merge(
                df_missing_fbs_info_processed[columns_fbs_info],
                how='left',
                on='Номер отправления'
            )
        )
        # Находим отправления, у которых изначально был кластер
        df_fbs_postings_not_missing = df_orders_fbs_enriched.loc[
            df_orders_fbs_enriched['Наличие в списке отправлений'] == 'both',
            :
        ]
        # Объединяем два df обратно
        df_orders_fbs_enriched = pd.concat([
            df_fbs_postings_not_missing,
            df_missing_fbs_postings_enriched
        ]).reset_index(drop=True)
    else:
        logger.success("Вся доп. информация по отправлениям FBS найдена в основном списке")

    return df_orders_fbs_enriched

# Функция выгрузки заказов FBS с кластерами
def getOrdersFBS(
        headers: dict[str, str],  # словарь заголовков HTTP
        date_start: str,          # ISO‑8601 начало периода
        date_end: str,            # ISO‑8601 конец периода
        add_clusters: bool = True,  # флаг добавления кластеров
        to_save: bool = True      # флаг сохранения файла
):

    # Выгружаем заказы FBS
    df_orders_fbs = getOrders(
        headers,
        date_start,
        date_end,
        delivery_schema="fbs",
        to_save=False
    )
    # Добавляем кластеры, если стоит флаг
    if add_clusters:
        df_orders_fbs = add_orders_info_fbs(
            headers,
            date_start,
            date_end,
            df_orders_fbs
        )
    # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    if to_save:
        file_name_orders_fbs = (
            uploaddir_today /
            f"{str(date.today())}_Заказы_fbs.csv"
        )
        df_orders_fbs.to_csv(
            file_name_orders_fbs,
            sep=';',
            encoding='utf-8-sig',
            index=False
        )
    # иначе - возвращаем df с заказами FBS
    else:
        return df_orders_fbs



# Кластеры отправлений (FBO)
def getFinalOrders_fbo(headers):
    # Имя файла с номерами заказов
    file_name_tmp_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_temp_fbo.csv"
    )
    # Загружаем данные по заказам
    df = pd.read_csv(file_name_tmp_orders, sep=";")
    # Имя файла для заказов с кластерами
    file_name_final_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_fbo.csv"
    )
    posting_numbers = df['Номер отправления'].to_list()
    # Если были заказы FBO, начинаем выгружать по ним кластеры
    if len(posting_numbers) > 0:
        df['Кластер отправления'] = ''
        df['Кластер доставки'] = ''
        df['Склад'] = ''
        total_posting_count = len(posting_numbers)
        loaded_posting_count = 0
        for number in posting_numbers:
            params = json.dumps({
                "posting_number": number,
                "translit": True,
                "with": {
                "analytics_data": True,
                "financial_data": True
                }
            })


            try:
            # записываем заказы и кластера в файл с заказами
                result = requests.post(f"{ozon_seller_api_url}/v2/posting/fbo/get",headers=headers, data=params).json()
                cluster_from = result['result']['financial_data']['cluster_from']
                cluster_to = result['result']['financial_data']['cluster_to']
                warehouse_name = result['result']['analytics_data']['warehouse_name']
                df.loc[df["Номер отправления"] == number, 'Склад'] = warehouse_name
                df.loc[df["Номер отправления"] == number, 'Кластер доставки'] = cluster_to
                df.loc[df["Номер отправления"] == number, 'Кластер отправления'] = cluster_from
                orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['created_at']))
                # Увеличиваем кол-во загруженных товаров на 1
                loaded_posting_count += 1
                # Сколько осталось выгрузить заказов в процентах
                loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
                logger.info(str(loaded_percent) + ' % ' + "Upload FBO orders on "+ orderDate.replace("['",'').replace("']",''))
                # time.sleep(1)
                df.to_csv(file_name_final_orders, sep=";", index = False, encoding='utf-8-sig')
            except Exception as e:
                print(f"{e}")
                logger.info("Connection is lost. Waiting  2 minutes and continue connection")
                time.sleep(120)
                result = requests.post(f"{ozon_seller_api_url}/v2/posting/fbo/get",headers=headers, data=params).json()
                cluster_from = result['result']['financial_data']['cluster_from']
                df.loc[df["Номер отправления"] == number, 'Кластер отправления'] = cluster_from
                cluster_to = result['result']['financial_data']['cluster_to']
                df.loc[df["Номер отправления"] == number, 'Кластер доставки'] = cluster_to
                warehouse_name = result['result']['analytics_data']['warehouse_name']
                df.loc[df["Номер отправления"] == number, 'Склад'] = warehouse_name
                orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['created_at']))
                # Увеличиваем кол-во загруженных товаров на 1
                loaded_posting_count += 1
                # Сколько осталось выгрузить заказов в процентах
                loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
                logger.info(str(loaded_percent) + ' % ' + "Upload FBO orders on "+ orderDate.replace("['",'').replace("']",''))
                # time.sleep(1)
                df.to_csv(file_name_final_orders, sep=";", index = False, encoding='utf-8-sig')


    # Если заказов FBO не было, сохраняем пустой df
    else:
        df_fbo_orders = pd.DataFrame(columns = ['Артикул', 'Кластер доставки', 'Статус', 'Сумма отправления', 'Объемный вес товаров, кг'])
        df_fbo_orders.to_csv(file_name_final_orders, sep=";", index=False, encoding='utf-8-sig')


# Кластеры отправлений (FBS)
def getFinalOrders_fbs(headers):
    # Имя файла с загруженными заказами
    file_name_tmp_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_temp_fbs.csv"
    )
    # Имя файла для заказов с кластерами
    file_name_final_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_fbs.csv"
    )
    # Загружаем данные по заказам
    df = pd.read_csv(file_name_tmp_orders, sep=";")
    # Создаем столбцы с информацией по кластерам
    df['Кластер отправления'] = ''
    df['Кластер доставки'] = ''
    df['Склад'] = ''
    # Если есть заказы по схеме fbs
    if df.shape[0]> 0:
        posting_numbers = df['Номер отправления'].to_list()
        total_posting_count = len(posting_numbers)
        loaded_posting_count = 0
        for number in posting_numbers:
            params = json.dumps({
            "posting_number": number,
            "with": {
                "analytics_data": True,
                "barcodes": False,
                "financial_data": True,
                "product_exemplars": False,
                "translit": False
            }
        })
            try:
            # записываем заказы и кластера в файл с заказами
                result = requests.post(f"{ozon_seller_api_url}/v3/posting/fbs/get",headers=headers, data=params).json()
                cluster_from = result['result']['financial_data']['cluster_from']
                df.loc[df["Номер отправления"] == number, 'Кластер отправления'] = cluster_from
                cluster_to = result['result']['financial_data']['cluster_to']
                df.loc[df["Номер отправления"] == number, 'Кластер доставки'] = cluster_to
                warehouse_name = result['result']['analytics_data']['warehouse']
                df.loc[df["Номер отправления"] == number, 'Склад'] = warehouse_name
                orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['in_process_at']))
                # Увеличиваем кол-во загруженных товаров на 1
                loaded_posting_count += 1
                # Сколько осталось выгрузить заказов в процентах
                loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
                logger.info(str(loaded_percent) + ' % ' + "Upload FBS orders on "+ orderDate.replace("['",'').replace("']",''))
                # time.sleep(1)
                df.to_csv(file_name_final_orders, sep=";", index = False, encoding='utf-8-sig')
            except Exception as e:
                print(f"{e}")
                logger.info("Connection is lost. Waiting  2 minites and continue connection")
                time.sleep(120)
                result = requests.post(f"{ozon_seller_api_url}/v3/posting/fbs/get",headers=headers, data=params).json()
                cluster_from = result['result']['financial_data']['cluster_from']
                df.loc[df["Номер отправления"] == number, 'Кластер отправления'] = cluster_from
                cluster_to = result['result']['financial_data']['cluster_to']
                df.loc[df["Номер отправления"] == number, 'Кластер доставки'] = cluster_to
                warehouse_name = result['result']['analytics_data']['warehouse']
                df.loc[df["Номер отправления"] == number, 'Склад'] = warehouse_name
                orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['created_at']))
                # Увеличиваем кол-во загруженных товаров на 1
                loaded_posting_count += 1
                # Сколько осталось выгрузить заказов в процентах
                loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
                logger.info(str(loaded_percent) + ' % ' + "Upload FBS orders on "+ orderDate.replace("['",'').replace("']",''))
                # time.sleep(1)
                df.to_csv(file_name_final_orders, sep=";", index = False, encoding='utf-8-sig')
    else:
        df.to_csv(file_name_final_orders, sep=";", index = False, encoding='utf-8-sig')

    #os.remove(f"{uploaddir_today}/{str(date.today())}_Заказы_temp_fbs.csv")






# Функция выгрузки списка отправлений FBO
def get_ozon_fbo_list(
    headers: dict[str, str],   # словарь заголовков HTTP
    date_start: str,           # ISO‑8601 начало периода
    date_end: str,             # ISO‑8601 конец периода
    limit: int = 1000,         # объём одной страницы
    retry_count: int = 5,      # сколько раз повторить запрос при ошибке
    offset_limit: int = 21000,  # число отправлений
) -> pd.DataFrame:
    """
    Загрузка всех заказов FBO из Ozon Seller API.

    Возвращает DataFrame: строка = товар + поля заказа.
    Поля заказа: order_id, order_number, posting_number,
                 status, substatus, cluster_from, cluster_to.

    При ошибке запроса выполняется до ``retry_count`` повторных попыток.
    Если запрос остаётся неудачным – переходим к следующему смещению/странице.
    """
    endpoint = f"{ozon_seller_api_url}/v2/posting/fbo/list"
    offset = 0
    all_orders_fbo = []

    logger.info(
        f"Начало выгрузки отправлений FBO с {date_start} по {date_end}"
    )

    # ------- Пагинация ----------
    while offset < offset_limit:
        payload = {
            "dir": "ASC",
            "filter": {
                "since": date_start,
                "status": "",
                "to": date_end,
            },
            "limit": limit,
            "offset": offset,
            # "translit": True,
            "with": {
                "analytics_data": True,
                "financial_data": True,
                "legal_info": False,
            },
        }

        logger.info(f"Запрос к Ozon: offset={offset}, limit={limit}")

        # -------- 5‑разовые попытки -----------
        for attempt in range(1, retry_count + 1):
            try:
                response = requests.post(
                    endpoint, headers=headers, json=payload, timeout=30
                )
                response.raise_for_status()
                break   # успешный запрос – выходим из цикла попыток
            except Exception as exc:
                logger.warning(
                    f"Запрос к Ozon не удался (попытка {attempt} из {retry_count}). "
                    f"Скоро будет выполнен повторный запрос."
                )
                time.sleep(2)  # небольшая пауза перед повтором
                if attempt == retry_count:
                    logger.error(
                        f"Ошибка запроса: {exc}. "
                        f"Пропускаем текущую страницу и переходим к следующей."
                    )
                    response = None

        # Если после всех попыток ответ не получен – переходим к следующему offset
        if response is None:
            offset += limit
            continue

        data = response.json()
        orders = data.get("result", [])
        if not orders:
            logger.info("Больше заказов не найдено.")
            break

        all_orders_fbo.extend(orders)
        logger.info(
            f"Получено {len(orders)} заказов; общее число: {len(all_orders_fbo)}"
        )

        offset += limit
        if len(orders) < limit:   # конец списка
            break

    logger.info(f"Собрано {len(all_orders_fbo)} заказов.")

    # Перевод в DataFrame
    try:
        df_all_orders_fbo = pd.DataFrame(all_orders_fbo)
    except Exception as exc:
        logger.error(f"Ошибка преобразования заказов в DataFrame: {exc}")
        df_all_orders_fbo = pd.DataFrame()

    if not df_all_orders_fbo.empty:
        # Получаем кластеры отправлений
        df_all_orders_fbo = custom_pd_explode(
            df_all_orders_fbo,
            column="financial_data",
        )
        # Добавляем схему доставки
        df_all_orders_fbo["delivery_schema"] = "FBO"
    else:
        df_all_orders_fbo = pd.DataFrame()

    return df_all_orders_fbo

# Функция выгрузки информации о конкретных отправлениях FBO
def get_ozon_fbo_info(
    headers: dict[str, str],   # словарь заголовков HTTP
    posting_numbers: list[str],  # список номеров отправлений
    retry_count: int = 5,      # сколько раз повторить запрос при ошибке
):
    """
    Загрузка информации по конкретным отправлениям FBO.
    В ответе теперь НЕ список – это один объект в поле ``result``.
    Функция всё равно собирает каждый объект в список, выводит прогресс
    в виде ``Progress: XX.XX% (N/M)`` и возвращает DataFrame.
    При возникновении ошибки запроса выполняется до 5 повторных попыток;
    если запрос остаётся неудачным, цикл продолжается с следующим
    отправлением, а в логгере печатается сообщение о повторе.
    """
    endpoint = f"{ozon_seller_api_url}/v2/posting/fbo/get"
    all_orders_fbo = []
    total_count = len(posting_numbers)
    logger.info(f"Начало выгрузки информации по {total_count} отправлениям FBO")
    for idx, posting_number in enumerate(posting_numbers, start=1):
        payload = {
            "posting_number": posting_number,
            # "translit": True,
            "with": {
                "analytics_data": True,
                "financial_data": True,
                "legal_info": False,
            },
        }
        # Прогресс в процентах
        percent = (idx / total_count) * 100
        logger.info(
            f"Запрос к Ozon: {percent:5.2f}% ({idx}/{total_count}) posting: {posting_number}"
        )
        # ----------- 5‑разовые попытки -----------
        for attempt in range(1, retry_count + 1):
            try:
                response = requests.post(
                    endpoint, headers=headers, json=payload, timeout=30
                )
                response.raise_for_status()
                break   # успешный запрос – выходим из цикла попыток
            except Exception as exc:
                # Если сервер вернул тело ошибки – логируем его
                logger.warning(
                    f"Запрос к Ozon не удался (попытка {attempt} из {retry_count}). "
                    f"Скоро будет выполнен повторный запрос."
                )
                if attempt == retry_count:
                    logger.error(f"Ошибка запроса: {exc}. Пропускаем это отправление.")
                    response = None
        # Если после {retry_count} попыток ответ не получен – переходим к следующему отправлению
        if response is None:
            continue
        data = response.json()
        # В новой версии API в ответе ровно один объект – ``result``.
        order = data.get("result")
        if not order:
            logger.warning("В ответе нет данных по этому отправлению.")
            continue
        all_orders_fbo.append(order)

    logger.info(f"Собрано {len(all_orders_fbo)} заказов.")
    # Превращаем список словарей в DataFrame
    try:
        df_all_orders_fbo = pd.DataFrame(all_orders_fbo)
    except Exception as exc:
        logger.error(f"Ошибка преобразования заказов в DataFrame: {exc}")
        df_all_orders_fbo = pd.DataFrame()
    # Если требуется разворачивать вложенную колонку «financial_data»
    df_all_orders_fbo = custom_pd_explode(
        df_all_orders_fbo, column="financial_data"
    )
    # Добавляем схему доставки
    df_all_orders_fbo["delivery_schema"] = "FBO"
    return df_all_orders_fbo

# Функция получения товаров в отправлениях FBO (заказы)
def process_fbo_list(
        df_all_orders_fbo: pd.DataFrame,
        additional_columns: list = None,
    ) -> pd.DataFrame:
    """
    Извлечение товаров из отправлений FBO из DataFrame заказов.

    Параметры:
        - df_all_orders_fbo: DataFrame с заказами FBO,
            полученный из функции get_ozon_fbo_list.
        - additional_columns: список дополнительных колонок для включения в итоговый DataFrame.
    """
    # Задаем имя файла для заказов
    file_name_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_fbo.csv"
    )
    if df_all_orders_fbo.empty:
        logger.warning("DataFrame заказов FBO пуст.")
        return pd.DataFrame(columns=api_columns.get('postings', []))

    # Базовые колонки для выбора
    order_columns = [
        'delivery_schema',
        'order_id',
        'order_number',
        'posting_number',
        'created_at',
        'in_process_at',
        'status',
        'financial_data.cluster_from',
        'financial_data.cluster_to',
        'financial_data.products',
        'products'
    ]
    # Если заданы дополнительные колонки, добавляем их в список
    if additional_columns:
        order_columns.extend(additional_columns)

    # Выбираем нужные колонки
    df_all_orders_fin_fbo = (
        df_all_orders_fbo
        .copy()
        .loc[:, order_columns]
    )
    # Распаковываем products
    df_products_fbo = custom_pd_explode(
        df_all_orders_fin_fbo,
        column='products'
    )

    # Переименовываем колонки
    df_products_fbo = df_products_fbo.rename(columns={
        'delivery_schema': 'Схема доставки',
        'posting_number': 'Номер отправления',
        'order_id': 'Идентификатор заказа',
        'order_number': 'Номер заказа',
        'status': 'Статус',
        'created_at': 'Дата создания',
        'in_process_at': 'Дата и время начала обработки отправления',
        'financial_data.cluster_from': 'Кластер отправления',
        'financial_data.cluster_to': 'Кластер доставки',
        'products.sku': 'SKU',
        'products.offer_id': 'Артикул',
        'products.name': 'Наименование товара',
        'products.quantity': 'Количество',
        'products.price': 'Цена товара',
    })
    # Переводим некоторые колонки в нужный тип
    column_types = {
        'Артикул': str,
        'Цена товара': float,
    }
    for col, col_type in column_types.items():
        if col in df_products_fbo.columns:
            df_products_fbo[col] = df_products_fbo[col].astype(col_type, errors='ignore')
    # Считаем заказы в рублях
    df_products_fbo['Сумма отправления'] = (
        df_products_fbo['Количество'] *
        df_products_fbo['Цена товара']
    )

    return df_products_fbo

# Функция выгрузки списка отправлений FBS
def get_ozon_fbs_list(
        headers : dict[str, str],   # словарь заголовков HTTP
        date_start : str,           # ISO‑8601 начало периода
        date_end : str,             # ISO‑8601 конец периода
        limit=1000 ,                # объём одной страницы
        retry_count=5,              # сколько раз повторить запрос при ошибке
        offset_limit=21000,         # максимальное число отправлений
    ) -> pd.DataFrame:
    """
    Загрузка всех заказов FBS (Fulfillment by Seller) из Ozon Seller API.

    Возвращает DataFrame: строка = товар + поля заказа.
    Поля заказа: posting_number, order_id, order_number,
                 status, substatus, cluster_from, cluster_to.

    При ошибке запроса выполняется до ``retry_count`` повторных попыток.
    Если запрос остаётся неудачным – переходим к следующему смещению/странице.
    """
    endpoint = f"{ozon_seller_api_url}/v3/posting/fbs/list"
    offset = 0
    all_orders_fbs = []

    logger.info(f"Начало выгрузки отправлений FBS с {date_start} по {date_end}")

    # ------- Пагинация ----------
    while offset < offset_limit:
        payload = {
            "dir": "ASC",
            "filter": {
                "since": date_start,
                "status": "",
                "to": date_end,
            },
            "limit": limit,
            "offset": offset,
            # "translit": True,
            "with": {
                "analytics_data": True,
                "financial_data": True,
                "legal_info": False,
            },
        }

        logger.info(f"Запрос списка отправлений FBS: offset={offset}, limit={limit}")

        # -------- 5‑разовые попытки ----------
        for attempt in range(1, retry_count + 1):
            try:
                response = requests.post(
                    endpoint,
                    headers=headers,
                    json=payload,
                    timeout=30
                )
                response.raise_for_status()
                break   # успешный запрос – выходим из цикла попыток
            except Exception as exc:
                logger.warning(
                    f"Запрос к Ozon не удался (попытка {attempt} из {retry_count}). "
                    f"Скоро будет выполнен повторный запрос."
                )
                if attempt == retry_count:
                    logger.error(
                        f"Ошибка запроса: {exc}. "
                        f"Пропускаем текущую страницу и переходим к следующей."
                    )
                    response = None

        # Если после всех попыток ответ не получен – переходим к следующему offset
        if response is None:
            offset += limit
            continue

        data = response.json()
        orders = data.get("result", {}).get("postings", [])
        has_next = data.get("result", {}).get("has_next", False)

        if not orders:
            logger.info("Больше заказов не найдено.")
            break

        all_orders_fbs.extend(orders)
        logger.info(
            f"Получено {len(orders)} заказов; общее число: {len(all_orders_fbs)}"
        )

        offset += limit
        if not has_next or len(orders) < limit:   # конец списка
            break

    logger.info(f"Собрано {len(all_orders_fbs)} заказов.")

    # Перевод в DataFrame
    try:
        df_all_orders_fbs = pd.DataFrame(all_orders_fbs)
    except Exception as exc:
        logger.error(f"Ошибка преобразования заказов в DataFrame: {exc}")
        df_all_orders_fbs = pd.DataFrame()

    if not df_all_orders_fbs.empty:
        # Распаковываем financial_data
        df_all_orders_fbs = custom_pd_explode(
            df_all_orders_fbs,
            column="financial_data",
        )
        # Добавляем схему доставки
        df_all_orders_fbs["delivery_schema"] = "FBS"
    else:
        df_all_orders_fbs = pd.DataFrame()

    return df_all_orders_fbs

# Функция выгрузки информации о конкретных отправлениях FBS
def get_ozon_fbs_info(
    headers : dict[str, str],   # словарь заголовков HTTP
    posting_numbers: list[str],  # список номеров отправлений
    retry_count: int = 5,          # сколько раз повторить запрос при ошибке
) -> pd.DataFrame:
    """
    Загрузка информации по конкретным отправлениям FBS.

    В ответе теперь НЕ список – это один объект в поле ``result``.
    Функция всё равно собирает каждый объект в список, выводит прогресс
    в виде ``Progress: XX.XX% (N/M)`` и возвращает DataFrame.

    При ошибке запроса выполняется до ``retry_count`` повторных попыток.
    Если запрос остаётся неудачным – переходим к следующему отправлению.
    """
    endpoint = f"{ozon_seller_api_url}/v3/posting/fbs/get"
    all_orders_fbs = []
    total_count = len(posting_numbers)

    for idx, posting_number in enumerate(posting_numbers, start=1):
        payload = {
            "posting_number": posting_number,
            # "translit": True,
            "with": {
                "analytics_data": True,
                "financial_data": True,
                "legal_info": False,
            },
        }
        # Прогресс в процентах
        percent = (idx / total_count) * 100
        logger.info(
            f"Запрос к Ozon: {percent:5.2f}% ({idx}/{total_count}) posting: {posting_number}"
        )

        # -------- 5‑разовые попытки ----------
        for attempt in range(1, retry_count + 1):
            try:
                response = requests.post(
                    endpoint, headers=headers, json=payload, timeout=30
                )
                response.raise_for_status()
                break   # успешный запрос – выходим из цикла попыток
            except Exception as exc:
                logger.warning(
                    f"Запрос к Ozon не удался (попытка {attempt} из {retry_count}). "
                    f"Скоро будет выполнен повторный запрос."
                )
                if attempt == retry_count:
                    logger.error(
                        f"Ошибка запроса для posting {posting_number}: {exc}. "
                        f"Пропускаем это отправление."
                    )
                    response = None

        # Если после всех попыток ответ не получен – переходим к следующему
        if response is None:
            continue

        # Пытаемся разобрать JSON
        try:
            data = response.json()
        except Exception as exc:
            logger.error(
                f"Не удалось разобрать JSON ответа для posting {posting_number}: {exc}"
            )
            continue

        # В новой версии API в ответе ровно один объект – ``result``.
        order = data.get("result")
        if not order:
            logger.warning(f"В ответе нет данных по отправлению {posting_number}.")
            continue

        all_orders_fbs.append(order)

    logger.info(f"Собрано {len(all_orders_fbs)} заказов.")

    # Превращаем список словарей в DataFrame
    try:
        df_all_orders_fbs = pd.DataFrame(all_orders_fbs)
    except Exception as exc:
        logger.error(f"Ошибка преобразования заказов в DataFrame: {exc}")
        df_all_orders_fbs = pd.DataFrame()

    # Если требуется разворачивать вложенную колонку «financial_data»
    df_all_orders_fbs = custom_pd_explode(
        df_all_orders_fbs,
        column="financial_data",
    )
    # Добавляем схему доставки
    df_all_orders_fbs["delivery_schema"] = "FBS"
    return df_all_orders_fbs


# Функция получения товаров в отправлениях fbs (заказы)
def process_fbs_list(
        df_all_orders_fbs: pd.DataFrame,
        additional_columns: list[str] = None,
    ) -> pd.DataFrame:
    """
    Извлечение товаров из отправлений fbs из DataFrame заказов.

    Параметры:
        - df_all_orders_fbs: DataFrame с заказами fbs,
            полученный из функции get_ozon_fbs_list.
        - additional_columns: список дополнительных колонок для включения в итоговый DataFrame.
    """
    # Задаем имя файла для заказов
    file_name_orders = (
        uploaddir_today /
        f"{str(date.today())}_Заказы_fbs.csv"
    )
    if df_all_orders_fbs.empty:
        logger.warning("DataFrame заказов FBS пуст.")
        return pd.DataFrame(columns=api_columns.get('postings', []))

    # Сравниваем длины products и financial_data.products
    # df_all_orders_fin_fbs['len_products'] = df_all_orders_fin_fbs['products'].apply(lambda x: len(x) if isinstance(x, list) else 0)
    # df_all_orders_fin_fbs['len_financial_products'] = df_all_orders_fin_fbs['financial_data.products'].apply(lambda x: len(x) if isinstance(x, list) else 0)

    # Выбираем нужные колонки
    order_columns = [
        'delivery_schema',
        'order_id',
        'order_number',
        'posting_number',
        # 'created_at',
        'in_process_at',
        'status',
        'financial_data.cluster_from',
        'financial_data.cluster_to',
        'financial_data.products',
        'products'
    ]
    # Если заданы дополнительные колонки, добавляем их в список
    if additional_columns:
        order_columns.extend(additional_columns)

    df_all_orders_fin_fbs = (
        df_all_orders_fbs
        .copy()
        .loc[:, order_columns]
    )

    # Распаковываем products
    df_products_fbs = custom_pd_explode(
        df_all_orders_fin_fbs,
        column='products'
    )

    # Переименовываем колонки
    df_products_fbs = df_products_fbs.rename(columns={
        'delivery_schema': 'Схема доставки',
        'posting_number': 'Номер отправления',
        'order_id': 'Идентификатор заказа',
        'order_number': 'Номер заказа',
        'status': 'Статус',
        'created_at': 'Дата создания',
        'in_process_at': 'Дата и время начала обработки отправления',
        'financial_data.cluster_from': 'Кластер отправления',
        'financial_data.cluster_to': 'Кластер доставки',
        'products.sku': 'SKU',
        'products.offer_id': 'Артикул',
        'products.name': 'Наименование товара',
        'products.quantity': 'Количество',
        'products.price': 'Цена товара',
    })
    # Переводим некоторые колонки в нужный тип
    column_types = {
        'Артикул': str,
        'Цена товара': float,
    }
    for col, col_type in column_types.items():
        if col in df_products_fbs.columns:
            df_products_fbs[col] = df_products_fbs[col].astype(col_type, errors='ignore')
    # Считаем заказы в рублях
    df_products_fbs['Сумма отправления'] = (
        df_products_fbs['Количество'] *
        df_products_fbs['Цена товара']
        # df_products_fbs['Цена товара']
    )

    return df_products_fbs

# Функция выгрузки заказов
def getOrders_v2(
        headers,
        date_start,
        date_end,
        to_save=True,
    ):
    """
    Выгружает заказы по API Ozon.

    Параметры:
        - headers: заголовки для запросов к API.
        - date_start: дата начала периода выгрузки (строка в формате ISO 8601).
        - date_end: дата окончания периода выгрузки (строка в формате ISO 8601).
        - to_save: флаг, указывающий, нужно ли сохранить файл с заказами.
    """
    logger.info(f"Выгрузка заказов с {date_start} по {date_end} начата.")

    # Выгружаем заказы FBO
    df_all_orders_fbo = get_ozon_fbo_list(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    # Получаем товары в заказах FBO
    df_fbo_orders_processed = process_fbo_list(
        df_all_orders_fbo=df_all_orders_fbo,
    )

    # Выгружаем заказы FBS
    df_all_orders_fbs = get_ozon_fbs_list(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    # Получаем товары в заказах FBS
    df_fbs_orders_processed = process_fbs_list(
        df_all_orders_fbs=df_all_orders_fbs,
    )
    # Если стоит флаг сохранения, сохраняем файлы с заказами
    if to_save:
        file_name_orders_fbo = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Заказы_fbo.csv"
        )

        file_name_orders_fbs = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Заказы_fbs.csv"
        )
        # Удаляем столбец с financial_data.products перед сохранением (для уменьшения размера файла)
        columns_to_remove = ['financial_data.products']
        for df in [df_fbo_orders_processed, df_fbs_orders_processed]:
            for col in columns_to_remove:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)

        df_fbo_orders_processed.to_csv(
            file_name_orders_fbo,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
        df_fbs_orders_processed.to_csv(
            file_name_orders_fbs,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    else:
        return df_fbo_orders_processed, df_fbs_orders_processed



    logger.info("Выгрузка заказов завершена.")

    # return


def generate_monthly_intervals(date_start: str, date_end: str):
    """Генерирует месячные интервалы, если период больше месяца"""
    dt_start = pd.to_datetime(date_start)
    dt_end = pd.to_datetime(date_end)
    intervals = []

    current_start = dt_start
    while current_start <= dt_end:
        next_month = (current_start + pd.DateOffset(months=1)).replace(day=1)
        current_end = min(next_month - pd.Timedelta(seconds=1), dt_end)
        intervals.append({
            'date_start': current_start.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'date_end': current_end.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'dt_start': current_start,
            'dt_end': current_end
        })
        current_start = next_month

    return pd.DataFrame(intervals)

# Функция выгрузки списка транзакций
def uploadTransactionReport(headers, date_start, date_end):
    """
    Выгружает транзакции по API Ozon с разбивкой по месяцам и логированием.
    """
    dt_start = pd.to_datetime(date_start)
    dt_end = pd.to_datetime(date_end)
    period_days = (dt_end - dt_start).days

    if period_days <= 31:
        # Один интервал на весь указанный диапазон
        date_range_df = pd.DataFrame([{
            'date_start': dt_start.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'date_end': dt_end.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'dt_start': dt_start,
            'dt_end': dt_end
        }])
    else:
        # Дробим на месяцы
        date_range_df = generate_monthly_intervals(date_start, date_end)

    logger.info(f"Выгрузка транзакций с {date_start} по {date_end} начата.")

    df_transactions_list = pd.DataFrame()

    for idx, interval in date_range_df.iterrows():
        logger.info(f"Обработка периода: {interval['dt_start']} - {interval['dt_end']}")
        page = 1
        total_pages = 1  # Пажинация: по умолчанию хотя бы одну страницу пробуем
        while page <= total_pages:
            payload = {
                "filter": {
                    "date": {
                        "from": interval['date_start'],
                        "to": interval['date_end']
                    },
                    "operation_type": [],
                    "posting_number": "",
                    "transaction_type": "all"
                },
                "page": page,
                "page_size": 1000
            }
            try:
                resp = requests.post(
                    f"{ozon_seller_api_url}/v3/finance/transaction/list",
                    headers=headers,
                    data=json.dumps(payload),
                    timeout=60
                )
                resp.raise_for_status()
                data = resp.json()
                # Получение страниц
                result = data.get('result', {})
                total_pages = result.get('page_count', 0)
                operations = result.get('operations', [])
                if not operations:
                    logger.info(f"Нет транзакций за этот период, страница {page}.")
                tmp_df = pd.DataFrame(operations)
                if not tmp_df.empty:
                    tmp_df = tmp_df.assign(
                        dt_start=interval['dt_start'],
                        dt_end=interval['dt_end']
                    )
                    df_transactions_list = pd.concat([df_transactions_list, tmp_df], ignore_index=True)
                logger.info(f"Страница {page} из {total_pages} обработана.")
            except requests.RequestException as ex:
                logger.error(f"Ошибка запроса API на странице {page}: {ex}")
            except Exception as ex:
                logger.error(f"Общая ошибка на странице {page}: {ex}")
            page += 1

    if not df_transactions_list.empty:
        logger.info(f"Выгрузка транзакций завершена. Найдено {df_transactions_list.shape[0]} транзакций.")
    else:
        logger.info("Транзакции не найдены по указанному периоду.")
        return pd.DataFrame(columns=api_columns.get("transactions", []))


    return df_transactions_list

# Функция обработки списка транзакций
def process_transaction_list(df_transactions_list):
    """
    Обрабатывает DataFrame со списком транзакций,
    извлекая информацию об отправлениях.
    """
    if not df_transactions_list.empty:
        # Создаем копию для избежания изменений оригинала
        df_transactions_list_processed = df_transactions_list.copy()
        # Распаковываем колонку информации об отправлениях
        df_fin_postings = custom_pd_explode(
            df_transactions_list_processed,
            column='posting'
        )

        return df_fin_postings
    else:
        return pd.DataFrame(columns=api_columns.get("transactions", []))

# Функция получения и обработки отчета по транзакциям
def get_transactions_report(headers, date_start, date_end, to_save=True):
    """
    Получает и обрабатывает отчет по транзакциям.

    Параметры:
        - headers: заголовки для API-запросов.
        - date_start: начало периода в формате ISO-8601.
        - date_end: конец периода в формате ISO-8601.
        - to_save: флаг сохранения результата в файл.
    """
    # Задаем имя файла для транзакций
    file_name_transactions = (
        uploaddir_today /
        f"{str(date.today())}_Транзакции.csv"
    )
    # Выгружаем список транзакций
    df_transactions_list = uploadTransactionReport(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    # Обрабатываем список транзакций
    df_transactions_list_processed = process_transaction_list(
        df_transactions_list=df_transactions_list
    )
    # Если стоит флаг сохранения, сохраняем файл
    if to_save:
        df_transactions_list_processed.to_csv(
            file_name_transactions,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    # Иначе возвращаем DataFrame с транзакциями
    else:
        return df_transactions_list_processed

# Функция получения продаж FBO (потоварно)
def get_sales_fbo_by_products(
        headers: dict[str, str],
        df_transactions_list_processed: pd.DataFrame,
):
    """
    Получает информацию о продажах FBO,
    сопоставляя транзакции с заказами FBO.

    Параметры:
        - df_transactions_list_processed: DataFrame со списком транзакций,
            обработанный функцией process_transaction_list.
    """
    # Фильтруем транзакции по fbo и типу операции "Доставка покупателю"
    df_sales_fbo_transactions = (
        df_transactions_list_processed
        .loc[
            (df_transactions_list_processed['posting.delivery_schema'] == 'FBO')
            & (df_transactions_list_processed['operation_type_name'] == 'Доставка покупателю'),
            :
        ]
    )
    # Если не нашли отправлений FBO, возвращаем пустой df
    if df_sales_fbo_transactions.empty:
        logger.info("Продажи FBO не найдены по указанному периоду.")
        return pd.DataFrame(columns=api_columns.get("postings", []))

    # Выбираем нужные колонки
    df_sales_fbo_transactions = df_sales_fbo_transactions[[
        'operation_id',
        'operation_date',
        'accruals_for_sale',
        'posting.posting_number',
        'posting.order_date',
        'posting.delivery_schema',
    ]]
    # Переименовываем колонки
    df_sales_fbo_transactions = df_sales_fbo_transactions.rename(columns={
        'posting.posting_number': 'posting_number',
        'posting.order_date': 'order_date',
        'posting.delivery_schema': 'delivery_schema',
    })
    # Удаляем дубликаты по номеру отправления
    df_sales_fbo_transactions = (
        df_sales_fbo_transactions
        .drop_duplicates(subset=['posting_number'], ignore_index=True)
    )
    # Получаем мин. и макс. даты заказов из транзакций
    min_date_sales_fbo = (
        pd.to_datetime(df_sales_fbo_transactions['order_date'].min())
        .tz_localize(None)
        .strftime('%Y-%m-%dT00:00:00Z')
    )
    max_date_sales_fbo = (
        pd.to_datetime(df_sales_fbo_transactions['order_date'].max())
        .tz_localize(None)
        .strftime('%Y-%m-%dT23:59:59Z')
    )
    logger.info(f"Даты заказов для продаж fbo: с {min_date_sales_fbo} по {max_date_sales_fbo}")

    # Выгружаем заказы fbo за нужный период
    df_fbo_list_for_sales = get_ozon_fbo_list(
        headers=headers,
        date_start=min_date_sales_fbo,
        date_end=max_date_sales_fbo
    )
    # Удаляем колонку со схемой доставки, т.к. она уже есть в списке транзакций
    df_fbo_list_for_sales = df_fbo_list_for_sales.drop(columns=['delivery_schema'])

    # Получаем товары в заказах fbo
    # df_fbo_orders_for_sales_processed = process_fbo_list(
    #     df_all_orders_fbo=df_fbo_list_for_sales,
    #     to_save=False
    # )

    # Логируем, если есть несопоставленные отправления
    unmatched_postings = (
        set(df_sales_fbo_transactions['posting_number'].unique()) -
        set(df_fbo_list_for_sales['posting_number'].unique())
    )
    unmatched_postings = list(unmatched_postings)
    # Если есть несопоставленные отправления, выгружаем по ним информацию
    if unmatched_postings:
        logger.warning(f"Не найдена информация по {len(unmatched_postings)} отправлениям")
        # Догружаем информацию по несопоставленным отправлениям
        df_unmatched_fbo_info = get_ozon_fbo_info(
            headers=headers,
            posting_numbers=unmatched_postings
        )
        # Удаляем колонку со схемой доставки, т.к. она уже есть в списке транзакций
        df_unmatched_fbo_info = df_unmatched_fbo_info.drop(columns=['delivery_schema'])
        # df_unmatched_fbo_processed = process_fbo_list(
        #     df_all_orders_fbo=df_unmatched_fbo_info,
        #     to_save=False
        # )

        # Объединяем с основными отправлениями
        df_all_postings_for_sales_fbo = pd.concat([df_fbo_list_for_sales, df_unmatched_fbo_info], ignore_index=True)
    else:
        df_all_postings_for_sales_fbo = df_fbo_list_for_sales.copy()

    # Удаляем дубликаты по номеру отправления (на всякий случай)
    df_all_postings_for_sales_fbo = (
        df_all_postings_for_sales_fbo
        .drop_duplicates(subset=['posting_number'], ignore_index=True)
    )
    # Мерджим с транзакциями
    df_sales_orders_fbo = df_sales_fbo_transactions.merge(
        df_all_postings_for_sales_fbo,
        how='left',
        on='posting_number',
    )

    # Обрабатываем полученный DataFrame для получения товаров
    df_sales_fbo_final = process_fbo_list(
        df_all_orders_fbo=df_sales_orders_fbo,
        additional_columns=['accruals_for_sale', 'operation_date'],
    )
    # У отправлений, где accruals_for_sale = 0, делаем стоимость отправления равной 0
    # df_sales_fbo_final.loc[
    #     df_sales_fbo_final['accruals_for_sale'] == 0,
    #     'Сумма отправления'
    # ] = 0

    return df_sales_fbo_final

# Функция получения продаж fbs (потоварно)
def get_sales_fbs_by_products(
        headers: dict[str, str],
        df_transactions_list_processed: pd.DataFrame,
):
    """
    Получает информацию о продажах fbs,
    сопоставляя транзакции с заказами fbs.

    Параметры:
        - df_transactions_list_processed: DataFrame со списком транзакций,
            обработанный функцией process_transaction_list.
        - df_fbs_orders_processed: DataFrame с товарами в заказах fbs,
            полученный из функции process_fbs_list.
    """
    # Задаем имя файла для продаж fbs
    file_name_sales_fbs = (
        uploaddir_today /
        f"{str(date.today())}_Продажи_fbs.csv"
    )
    # Фильтруем транзакции по fbs и типу операции "Доставка покупателю"
    df_sales_fbs_transactions = (
        df_transactions_list_processed
        .loc[
            (df_transactions_list_processed['posting.delivery_schema'].str.contains('FBS|RFBS', na=False))
            & (df_transactions_list_processed['operation_type_name'] == 'Доставка покупателю'),
            :
        ]
    )
    # Если не нашли отправлений FBS, возвращаем пустой df
    if df_sales_fbs_transactions.empty:
        logger.info("Продажи FBS не найдены по указанному периоду.")
        return pd.DataFrame(columns=api_columns.get("postings", []))

    # Выбираем нужные колонки
    df_sales_fbs_transactions = df_sales_fbs_transactions[[
                'operation_id',
                'operation_date',
                'accruals_for_sale',
                'posting.posting_number',
                'posting.order_date',
                'posting.delivery_schema',
    ]]
    # Переименовываем колонки
    df_sales_fbs_transactions = df_sales_fbs_transactions.rename(columns={
        'posting.posting_number': 'posting_number',
        'posting.order_date': 'order_date',
        'posting.delivery_schema': 'delivery_schema',
    })
    # Удаляем дубликаты по номеру отправления
    df_sales_fbs_transactions = (
        df_sales_fbs_transactions
        .drop_duplicates(subset=['posting_number'], ignore_index=True)
    )
    # Получаем мин. и макс. даты заказов из транзакций
    min_date_sales_fbs = (
        pd.to_datetime(df_sales_fbs_transactions['order_date'].min())
        .tz_localize(None)
        .strftime('%Y-%m-%dT00:00:00Z')
    )
    max_date_sales_fbs = (
        pd.to_datetime(df_sales_fbs_transactions['order_date'].max())
        .tz_localize(None)
        .strftime('%Y-%m-%dT23:59:59Z')
    )
    logger.info(f"Даты заказов для продаж fbs: с {min_date_sales_fbs} по {max_date_sales_fbs}")

    # Выгружаем заказы fbs за нужный период
    df_fbs_list_for_sales = get_ozon_fbs_list(
        headers=headers,
        date_start=min_date_sales_fbs,
        date_end=max_date_sales_fbs
    )
    # Удаляем колонку со схемой доставки, т.к. она уже есть в списке транзакций
    df_fbs_list_for_sales = df_fbs_list_for_sales.drop(columns=['delivery_schema'])
    # Получаем товары в заказах fbs
    # df_fbs_orders_for_sales_processed = process_fbs_list(
    #     df_all_orders_fbs=df_fbs_list_for_sales,
    #     to_save=False
    # )

    # Логируем, если есть несопоставленные отправления
    unmatched_postings = (
        set(df_sales_fbs_transactions['posting_number'].unique()) -
        set(df_fbs_list_for_sales['posting_number'].unique())
    )
    unmatched_postings = list(unmatched_postings)
    # Если есть несопоставленные отправления, выгружаем по ним информацию
    if unmatched_postings:
        logger.warning(f"Не найдена информация по {len(unmatched_postings)} отправлениям")
        # Догружаем информацию по несопоставленным отправлениям
        df_unmatched_fbs_info = get_ozon_fbs_info(
            headers=headers,
            posting_numbers=unmatched_postings
        )
        # Удаляем колонку со схемой доставки, т.к. она уже есть в списке транзакций
        df_unmatched_fbs_info = df_unmatched_fbs_info.drop(columns=['delivery_schema'])
        # df_unmatched_fbs_processed = process_fbs_list(
        #     df_all_orders_fbs=df_unmatched_fbs_info,
        #     to_save=False
        # )
        # Объединяем с основными отправлениями
        df_all_postings_for_sales_fbs = pd.concat([df_fbs_list_for_sales, df_unmatched_fbs_info], ignore_index=True)
    else:
        df_all_postings_for_sales_fbs = df_fbs_list_for_sales.copy()

    # Удаляем дубликаты по номеру отправления (на всякий случай)
    df_all_postings_for_sales_fbs = (
        df_all_postings_for_sales_fbs
        .drop_duplicates(subset=['posting_number'], ignore_index=True)
    )
    # Мерджим с транзакциями
    df_sales_orders_fbs = df_sales_fbs_transactions.merge(
        df_all_postings_for_sales_fbs,
        how='left',
        on='posting_number',
    )

    # Обрабатываем полученный DataFrame для получения товаров
    df_sales_fbs_final = process_fbs_list(
        df_all_orders_fbs=df_sales_orders_fbs,
        additional_columns=['accruals_for_sale', 'operation_date'],
    )
    # У отправлений, где accruals_for_sale = 0, делаем стоимость отправления равной 0
    # df_sales_fbs_final.loc[
    #     df_sales_fbs_final['accruals_for_sale'] == 0,
    #     'Сумма отправления'
    # ] = 0

    return df_sales_fbs_final

# Функция получения продаж FBO
def getSalesFBO_v2(
        headers,
        date_start,
        date_end,
        upload_transactions=True,
        to_save=True,

    ):
    """
    Выгружает продажи FBO по API Ozon.

    Параметры:
        - headers: заголовки для запросов к API.
        - date_start: дата начала периода выгрузки (строка в формате ISO 8601).
        - date_end: дата окончания периода выгрузки (строка в формате ISO 8601).
        - upload_transactions: флаг, указывающий, нужно ли выгружать список транзакций.
    """
    if upload_transactions:
        # Выгружаем список транзакций
        df_transactions_list_processed = get_transactions_report(
            headers=headers,
            date_start=date_start,
            date_end=date_end,
            to_save=False
        )
    else:
        file_name_transactions = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Транзакции.csv"
        )
        # Считываем файл со списком транзакций
        df_transactions_list_processed = pd.read_csv(file_name_transactions, sep=';')
        # Переводим некоторые колонки из строк в словари
        columns_to_eval = ['items', 'services']
        for col in columns_to_eval:
            if col in df_transactions_list_processed.columns:
                df_transactions_list_processed[col] = df_transactions_list_processed[col].apply(literal_eval)

    if not df_transactions_list_processed.empty:
        # Считаем продажи FBO из отправлений в списке транзакций
        df_sales_fbo_final = get_sales_fbo_by_products(
            headers,
            df_transactions_list_processed,
        )
        logger.success('Выгрузка продаж FBO завершена.')
    else:
        logger.info('Продажи FBO не найдены по указанному периоду.')
        df_sales_fbo_final = pd.DataFrame(columns=api_columns.get("postings", []))

    # Если стоит флаг сохранения, сохраняем файл
    if to_save:
        file_name_sales_fbo = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Продажи_fbo.csv"
        )
        # Удаляем некоторые колонки для уменьшения размера файла
        columns_to_remove = ['financial_data.products']
        for col in columns_to_remove:
            if col in df_sales_fbo_final.columns:
                df_sales_fbo_final.drop(columns=[col], inplace=True)
        # Сохраняем файл
        df_sales_fbo_final.to_csv(
            file_name_sales_fbo,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    else:
        return df_sales_fbo_final

# Функция получения продаж FBS
def getSalesFBS_v2(
        headers,
        date_start,
        date_end,
        upload_transactions=True,
        to_save=True,

    ):
    """
    Выгружает продажи FBS по API Ozon.

    Параметры:
        - headers: заголовки для запросов к API.
        - date_start: дата начала периода выгрузки (строка в формате ISO 8601).
        - date_end: дата окончания периода выгрузки (строка в формате ISO 8601).
        - upload_transactions: флаг, указывающий, нужно ли выгружать список транзакций.
    """
    if upload_transactions:
        # Выгружаем список транзакций
        df_transactions_list_processed = get_transactions_report(
            headers=headers,
            date_start=date_start,
            date_end=date_end,
            to_save=False
        )
    else:
        file_name_transactions = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Транзакции.csv"
        )
        # Считываем файл со списком транзакций
        df_transactions_list_processed = pd.read_csv(file_name_transactions, sep=';')
        # Переводим некоторые колонки из строк в словари
        columns_to_eval = ['items', 'services']
        for col in columns_to_eval:
            if col in df_transactions_list_processed.columns:
                df_transactions_list_processed[col] = df_transactions_list_processed[col].apply(literal_eval)

    if not df_transactions_list_processed.empty:
        # Считаем продажи FBS из отправлений в списке транзакций
        df_sales_fbs_final = get_sales_fbs_by_products(
            headers,
            df_transactions_list_processed,
        )
        logger.success('Выгрузка продаж FBS завершена.')
    else:
        logger.info('Продажи FBS не найдены по указанному периоду.')
        df_sales_fbs_final = pd.DataFrame(columns=api_columns.get("postings", []))

    # Если стоит флаг сохранения, сохраняем файл
    if to_save:
        file_name_sales_fbs = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Продажи_fbs.csv"
        )
        # Удаляем некоторые колонки для уменьшения размера файла
        columns_to_remove = ['financial_data.products']
        for col in columns_to_remove:
            if col in df_sales_fbs_final.columns:
                df_sales_fbs_final.drop(columns=[col], inplace=True)
        # Сохраняем файл
        df_sales_fbs_final.to_csv(
            file_name_sales_fbs,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    else:
        return df_sales_fbs_final


# Функция выгрузки списка возвратов
def get_ozon_returns_list(
    headers: dict[str, str],
    date_start: str,
    date_end: str,
    limit: int = 500,
) -> pd.DataFrame:
    """
    Загрузка всех возвратов (FBO/FBS) из Ozon Seller API.

    Параметры запроса:
        * `filter.logistic_return_date.time_from` / `time_to` – даты,
        * `limit` – сколько вернуть в каждом запросе,
        * `last_id` – курсор‑параметр для пагинации.

    Ответ содержит:
        * ``returns`` – список возвратов,
        * ``has_next`` – можно ли запрашивать дальше.

    Функция собирает все возвраты в один DataFrame

    """
    # ------------------------------------------------------------
    #  Константы
    # ------------------------------------------------------------
    endpoint = f"{ozon_seller_api_url}/v1/returns/list"
    sleep_time = 1  # пауза между запросами в секундах
    last_id = 0                      # курсор
    all_returns = []

    logger.info(f"Начало выгрузки возвратов с {date_start} по {date_end}")

    # ------------------------------------------------------------
    #  Пагинация
    # ------------------------------------------------------------
    while True:
        payload = {
            "filter": {
                "logistic_return_date": {
                    "time_from": date_start,
                    "time_to": date_end
                },
            },
            "limit": limit,
            "last_id": last_id
        }

        logger.info(f"Запрос к Ozon: last_id={last_id}, limit={limit}")

        try:
            response = requests.post(
                endpoint, headers=headers, json=payload, timeout=30
            )
            response.raise_for_status()
        except Exception as exc:
            logger.error(f"Ошибка запроса: {exc}")
            break

        data = response.json()

        batch = data.get("returns", [])
        if not batch:
            logger.info("Больше возвратов не найдено.")
            break

        all_returns.extend(batch)
        logger.info(
            f"Получено {len(batch)} возвратов; общее число: {len(all_returns)}"
        )

        if not data.get("has_next", False):
            logger.info("Последняя страница возвратов обработана.")
            break

        # Указываем курсор как id последнего возвращённого элемента
        # (у него поле ``id`` – строка)
        last_id = batch[-1]["id"]

        logger.info(f"Ждем {sleep_time} секунд перед следующим запросом...")
        # Пауза перед следующим запросом
        time.sleep(sleep_time)

    logger.info(f"Собрано {len(all_returns)} возвратов.")

    # ------------------------------------------------------------
    #  Перевод в DataFrame
    # ------------------------------------------------------------
    try:
        df_all_returns = pd.DataFrame(all_returns)
    except Exception as exc:
        logger.error(f"Ошибка преобразования возвратов в DataFrame: {exc}")
        df_all_returns = pd.DataFrame()

    # Переводим схему доставки в верхний регистр
    if not df_all_returns.empty and 'schema' in df_all_returns.columns:
        df_all_returns['schema'] = df_all_returns['schema'].str.upper()

    return df_all_returns

# Функция обработки списка возвратов
def process_returns_list(
        df_all_returns: pd.DataFrame,
    ) -> pd.DataFrame:
    """
    Обработка DataFrame со списком возвратов.

    Параметры:
        - df_all_returns: DataFrame с возвратами,
            полученный из функции get_ozon_returns_list.
    """
    if df_all_returns.empty:
        logger.warning("DataFrame возвратов пуст.")
        return pd.DataFrame(columns=api_columns.get('returns', []))

    # Выбираем нужные колонки
    return_columns = [
        'schema',  # FBO/FBS
        'return_reason_name',  # причина возврата
        'type',  # тип возврата
        'posting_number',  # номер отправления
        'place',  # склад, где находится возврат
        'product',  # информация о товаре
        'logistic',  # информация о логистике возврата

    ]

    df_returns_processed = (
        df_all_returns
        .copy()
        .loc[:, return_columns]
    )

    # Распаковываем product для получения информации о товаре
    df_returns_processed = custom_pd_explode(
        df_returns_processed,
        column='product'
    )

    # Распаковываем logistic для получения даты возврата
    df_returns_processed = custom_pd_explode(
        df_returns_processed,
        column='logistic'
    )

    # Переименовываем колонки
    df_returns_processed = df_returns_processed.rename(columns={
        'posting_number': 'Номер отправления',
        'schema': 'Схема',
        'return_reason_name': 'Причина возврата',
        'type': 'Тип возврата',
        'place': 'Склад находжения возврата',
        'product.sku': 'SKU',
        'product.offer_id': 'Артикул',
        'product.name': 'Наименование товара',
        'product.commission_percent': 'Процент комиссии',
        'product.quantity': 'Количество',
        'product.price.currency_code': 'Валюта цены',
        'product.price.price': 'Цена товара',
        'product.price_without_commission.currency_code': 'Валюта цены без комиссии',
        'product.price_without_commission.price': 'Цена товара без комиссии',
        'product.commission.currency_code': 'Валюта комиссии',
        'product.commission.price': 'Комиссия',
        'logistic.technical_return_moment': 'Дата технического возврата',
        'logistic.final_moment': 'Дата, когда возврат прибыл на фулфилмент или выдан продавцу',
        'logistic.cancelled_with_compensation_moment': 'Дата, когда продавцу компенсировали возврат',
        'logistic.return_date': 'Дата, когда покупатель вернул товар',
        'logistic.barcode': 'Штрихкод этикетки возврата',
    })
    # Считаем итоговую стоимость возврата
    df_returns_processed['Стоимость возврата'] = df_returns_processed['Количество'] * df_returns_processed['Цена товара']

    # Делаем фильтр по типу возврата (только возвраты от покупателей)
    df_returns_processed = (
        df_returns_processed
        .loc[df_returns_processed['Тип возврата'] == 'ClientReturn', :]
        .reset_index(drop=True)
    )

    return df_returns_processed

# Функция добавления кластеров к возвратам
def add_clusters_to_returns(
        headers: dict[str, str],
        date_start: str,
        date_end: str,
        df_returns_processed: pd.DataFrame,
        upload_orders: bool = True,
    ) -> pd.DataFrame:
    """
    Добавляет информацию о кластерах к возвратам.

    Параметры:
        - headers - заголовки для API-запросов.
        - date_start - начало периода в формате ISO-8601.
        - date_end - конец периода в формате ISO-8601.
        - df_returns_processed - DataFrame с возвратами,
        - upload_orders - флаг выгрузки заказов (если False, используется локальный файл).
    """

    if df_returns_processed.empty:
        logger.warning("DataFrame возвратов пуст.")
        return pd.DataFrame()

    # Создаем копию DataFrame для добавления кластеров
    df_returns_with_clusters = df_returns_processed.copy()

    # Выгружаем заказы (если нужно)
    if upload_orders:
        df_fbo_orders_processed, df_fbs_orders_processed = getOrders_v2(
            headers=headers,
            date_start=date_start,
            date_end=date_end,
            to_save=False
        )

    # Иначе читаем из локальных файлов
    else:
        file_name_orders_fbo = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Заказы_fbo.csv"
        )
        file_name_orders_fbs = (
            Path(uploaddir_today) /
            f"{str(date.today())}_Заказы_fbs.csv"
        )
        df_fbo_orders_processed = pd.read_csv(file_name_orders_fbo, sep=';')
        df_fbs_orders_processed = pd.read_csv(file_name_orders_fbs, sep=';')

    # Объединяем заказы FBO и FBS
    common_columns = [
        'Схема доставки',
        'Номер отправления',
        'Кластер доставки'
    ]
    df_orders = pd.concat([df_fbo_orders_processed, df_fbs_orders_processed], ignore_index=True)
    # Удаляем дубликаты по номеру отправления
    df_orders = df_orders.drop_duplicates(
        subset=['Номер отправления'],
        ignore_index=True
    )
    # Оставляем только нужные колонки
    df_orders = df_orders[common_columns]
    # Переименовываем схему доставки для мерджа
    df_orders = df_orders.rename(columns={
        'Схема доставки': 'Схема'
    })

    # Добавляем информацию о кластерах к возвратам
    df_returns_with_orders_info = df_returns_processed.merge(
        df_orders,
        on=['Схема', 'Номер отправления'],
        how='left',
        indicator='Наличие в заказах'
    )
    # Находим возвраты, у которых нет кластера
    df_returns_with_missing_clusters = (
        df_returns_with_orders_info
        .loc[df_returns_with_orders_info['Наличие в заказах'] == 'left_only', :]
    )
    if not df_returns_with_missing_clusters.empty:
        logger.info(f"Найдены возвраты без информации о кластерах: {df_returns_with_missing_clusters.shape[0]} шт.")
        # Разделяем возвраты без кластеров на FBO и FBS
        df_returns_with_missing_clusters_fbo = (
            df_returns_with_missing_clusters
            .loc[df_returns_with_missing_clusters['Схема'] == 'FBO', :]
            .reset_index(drop=True)
        )
        df_returns_with_missing_clusters_fbs = (
            df_returns_with_missing_clusters
            .loc[df_returns_with_missing_clusters['Схема'] == 'FBS', :]
            .reset_index(drop=True)
        )
        # Формируем список уникальных номеров отправлений для FBO и FBS
        fbo_returns_postings = df_returns_with_missing_clusters_fbo['Номер отправления'].unique().tolist()
        fbs_returns_postings = df_returns_with_missing_clusters_fbs['Номер отправления'].unique().tolist()

        # Выгружаем информацию по отправлениям FBO и FBS, если есть
        if fbo_returns_postings:
            df_fbo_missing_postings_info = get_ozon_fbo_info(
                headers=headers,
                posting_numbers=fbo_returns_postings
            )
        else:
            df_fbo_missing_postings_info = pd.DataFrame()

        if fbs_returns_postings:
            df_fbs_missing_postings_info = get_ozon_fbs_info(
                headers=headers,
                posting_numbers=fbs_returns_postings
            )
        else:
            df_fbs_missing_postings_info = pd.DataFrame()

        # Объединяем информацию по FBO и FBS
        df_missing_postings_info = pd.concat([df_fbo_missing_postings_info, df_fbs_missing_postings_info], ignore_index=True)
        # Переименовываем некоторые колонки
        df_missing_postings_info = df_missing_postings_info.rename(columns={
            'delivery_schema': 'Схема',
            'posting_number': 'Номер отправления',
            'financial_data.cluster_from': 'Кластер отправления',
            'financial_data.cluster_to': 'Кластер доставки',
        })

        # Добавляем информацию о кластерах к возвратам с пропусками
        df_returns_with_added_clusters = (
            df_returns_with_missing_clusters
            .drop(columns=['Кластер доставки'])
            .merge(
                df_missing_postings_info[['Схема', 'Номер отправления', 'Кластер доставки']],
                how='left',
                on=['Схема', 'Номер отправления']
            )
        )

        # Находим возвраты, у которых изначально был кластер
        df_returns_with_clusters = (
            df_returns_with_orders_info
            .loc[df_returns_with_orders_info['Наличие в заказах'] == 'both', :]
        )

        # Объединяем их с датафреймом добавленных кластеров
        df_returns_with_clusters = pd.concat([df_returns_with_clusters, df_returns_with_added_clusters], ignore_index=True)
    # Если информация по возвратам была найдена в заказах, просто копируем датафрейм
    else:
        logger.info("Информация по всем возвратам найдена в заказах.")
        df_returns_with_clusters = df_returns_with_orders_info.copy()

    logger.success("Добавление кластеров к возвратам завершено.")
    return df_returns_with_clusters


# Функция получения и обработки возвратов
def getReturns_v2(
        headers,
        date_start,
        date_end,
        add_clusters=True,
        to_save=True,
    ):
    """
    Получает и обрабатывает возвраты.

    Параметры:
        - headers: заголовки для API-запросов.
        - date_start: начало периода в формате ISO-8601.
        - date_end: конец периода в формате ISO-8601.
        - add_clusters: флаг добавления кластеров к возвратам.
        - to_save: флаг сохранения результата в файл.
    """
    # Задаем имя файла для возвратов
    file_name_returns = (
        uploaddir_today /
        f"{str(date.today())}_Возвраты.csv"
    )
    # Выгружаем список возвратов
    df_all_returns = get_ozon_returns_list(
        headers=headers,
        date_start=date_start,
        date_end=date_end
    )
    if not df_all_returns.empty:
        # Обрабатываем список возвратов
        df_returns_processed = process_returns_list(
            df_all_returns=df_all_returns
        )
        if add_clusters:
            df_returns_processed = add_clusters_to_returns(
                headers=headers,
                date_start=date_start,
                date_end=date_end,
                df_returns_processed=df_returns_processed,
                upload_orders=False
            )
    else:
        logger.info("Возвраты за указанный период не найдены.")
        df_returns_processed = pd.DataFrame(columns=api_columns.get('returns', []))

    logger.success('Выгрузка возвратов завершена.')

    # Если стоит флаг сохранения, сохраняем файл
    if to_save:
        df_returns_processed.to_csv(
            file_name_returns,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    # Иначе возвращаем DataFrame с возвратами
    else:
        return df_returns_processed


# Продажи FBO
def getSalesFBO(headers):
    # Задаем имя файла для списка транзакций
    file_name_transactions = (
        uploaddir_today /
        f"{str(date.today())}_Транзакции.csv"
    )
    # Задаем имя файла для продаж FBO
    file_name_sales_fbo = (
        uploaddir_today /
        f"{str(date.today())}_Продажи_fbo.csv"
    )
    # Считываем файл со списком транзакций
    # BUG: при чтении excel все словари в колонках переводятся в строки. Приходится их переводить обратно при помощи eval()
    df_transactions_list = pd.read_csv(file_name_transactions, sep=';')
    df_transactions_list['posting'] = df_transactions_list['posting'].apply(literal_eval)
    # Достаем информацию об отправлениях из списка транзакций
    postings = pd.json_normalize(df_transactions_list['posting'])
    # Добавляем тип (заказ\возврат), стоимость и количество товаров в отправлении
    postings = pd.concat([postings, df_transactions_list[['type', 'operation_date', 'operation_type', 'operation_type_name', 'amount', 'accruals_for_sale', 'items_amount']]], axis=1)

    # Делаем выборку только по FBO и считаем итоговую цену для каждого отправления
    postings_fbo = (
        postings
        # Продажей считаем тип начисления - Доставка покупателю
        .loc[(postings['delivery_schema'] == 'FBO')& (postings['operation_type_name'].isin(['Доставка покупателю'])), :]
        .groupby(['posting_number', 'type', 'order_date'])
        .agg(amount=('amount', 'sum'))
        .reset_index()
    )
    # Делаем фильтр по типу отправления, нужны только заказы (orders) без возвратов
    postings_fbo = (
        postings_fbo
        .loc[postings_fbo['type'] == 'orders', :]
        .drop_duplicates(subset=['posting_number'])
        .sort_values(['order_date'])
    )

    # Если есть продажи по FBO
    if postings_fbo.shape[0] > 0:
        # Выгружаем данные по отправлениям FBO
        sales_fbo = pd.DataFrame()
        loaded_posting_count = 0
        total_posting_count = len(postings_fbo['posting_number'])
        for posting_number in postings_fbo['posting_number']:
            params_fbo = json.dumps({
                "posting_number": posting_number,
                "translit": True,
                "with": {
                    "analytics_data": True,
                    "financial_data": True
                }
            })
            result = requests.post(f"{ozon_seller_api_url}/v2/posting/fbo/get",headers=headers, data=params_fbo).json()
            # Получаем дату заказа (для логов)
            orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['created_at']))
            # Достаем из результата данные о товаре
            products_fbo = pd.DataFrame(result['result']['products'])
            products_fbo = products_fbo.assign(
                Номер_заказа=result['result']['order_number'],
                Номер_отправления=result['result']['posting_number'],
                Склад = result['result']['analytics_data']['warehouse_name'],
                Кластер_отправления=result['result']['financial_data']['cluster_from'],
                Кластер_доставки=result['result']['financial_data']['cluster_to'],
                Сумма_отправления=postings_fbo.loc[postings_fbo['posting_number'] == posting_number, 'amount'].values[0],
                Статус=result['result']['status']
            )
            # Объединяем данные с предыдущим заказом
            sales_fbo = pd.concat([sales_fbo, products_fbo])
            # Сколько уже загружено заказов
            loaded_posting_count += 1
            # Сколько осталось выгрузить заказов в процентах
            loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
            logger.info(str(loaded_percent) + ' % ' + "Upload FBO sales from "+ orderDate.replace("['",'').replace("']",''))
            # Делаем паузу 3 с., чтоб ы часто не вызывать метод апи
            # time.sleep(3)

        # Переименовываем некоторые колонки для удобства и убираем лишние
        sales_fbo = sales_fbo.rename(columns={
            'offer_id': 'Артикул',
            'name': 'Название товара',
            'quantity': 'Количество',
            'currency_code': 'Валюта отправления'
        })
        # Указываем, что схема доставки FBO
        sales_fbo['Схема доставки'] = 'FBO'
        sales_fbo = sales_fbo.loc[:, ~sales_fbo.columns.isin(['price', 'digital_codes'])]
    # Если продаж не было, сохраняем пустой df с продажами
    else:
        colnames_sales_fbo = ['sku', 'Название товара', 'Количество', 'Артикул',
       'Валюта отправления', 'Номер_заказа', 'Номер_отправления', 'Склад',
       'Кластер_отправления', 'Кластер_доставки', 'Сумма_отправления', 'Схема доставки', 'Статус']
        sales_fbo = pd.DataFrame(columns=colnames_sales_fbo)
    sales_fbo.to_csv(file_name_sales_fbo, sep=';', index=False, encoding='utf-8-sig')


# Продажи FBS
def getSalesFBS(headers):
    # Задаем имя файла для списка транзакций
    file_name_transactions = (
        uploaddir_today /
        f"{str(date.today())}_Транзакции.csv"
    )
    # Задаем имя файла для продаж FBS
    file_name_sales_fbs = (
        uploaddir_today /
        f"{str(date.today())}_Продажи_fbs.csv"
    )
    # Считываем файл со списком транзакций
    # BUG: при чтении excel все словари в колонках переводятся в строки. Приходится их переводить обратно при помощи eval()
    df_transactions_list = pd.read_csv(file_name_transactions, sep=';')
    df_transactions_list['posting'] = df_transactions_list['posting'].apply(literal_eval)
    # Достаем информацию об отправлениях из списка транзакций
    postings = pd.json_normalize(df_transactions_list['posting'])
    # Добавляем тип (заказ\возврат), стоимость и количество товаров в отправлении
    postings = pd.concat([postings, df_transactions_list[['type', 'operation_date', 'operation_type', 'operation_type_name', 'amount', 'accruals_for_sale', 'items_amount']]], axis=1)

    # Делаем выборку только по FBO и считаем итоговую цену для каждого отправления
    postings_fbs = (
        postings
        # Продажей считаем тип начисления - Доставка покупателю
        .loc[(postings['delivery_schema'] == 'FBS')& (postings['operation_type_name'].isin(['Доставка покупателю'])), :]
        .groupby(['posting_number', 'type', 'order_date'])
        .agg(amount=('amount', 'sum'))
        .reset_index()
    )
    # Делаем фильтр по типу отправления, нужны только заказы (orders) без возвратов
    postings_fbs = (
        postings_fbs
        .loc[postings_fbs['type'] == 'orders', :]
        .drop_duplicates(subset=['posting_number'])
        .sort_values(['order_date'])
    )

    # Если есть продажи по FBS
    if postings_fbs.shape[0] > 0:
        # Выгружаем данные по отправлениям FBO
        sales_fbs = pd.DataFrame()
        loaded_posting_count = 0
        total_posting_count = postings_fbs['posting_number'].shape[0]
        for posting_number in postings_fbs['posting_number']:
            params_fbs = json.dumps({
                "posting_number": posting_number,
                "translit": True,
                "with": {
                    "analytics_data": True,
                    "financial_data": True
                }
            })
            result = requests.post(f"{ozon_seller_api_url}/v3/posting/fbs/get",headers=headers, data=params_fbs).json()
            # Получаем дату заказа (для логов)
            orderDate = str(re.findall(r'[0-9]*-[0-9]*-[0-9][0-9]',result['result']['in_process_at']))
            # Достаем из результата данные о товаре
            products_fbs = pd.DataFrame(result['result']['products'])
            products_fbs = products_fbs.assign(
                Номер_заказа=result['result']['order_number'],
                Номер_отправления=result['result']['posting_number'],
                Склад = result['result']['analytics_data']['warehouse'],
                Кластер_отправления=result['result']['financial_data']['cluster_from'],
                Кластер_доставки=result['result']['financial_data']['cluster_to'],
                Сумма_отправления=postings_fbs.loc[postings_fbs['posting_number'] == posting_number, 'amount'].values[0],
                Статус=result['result']['status']
            )
            # Объединяем данные с предыдущим заказом
            sales_fbs = pd.concat([sales_fbs, products_fbs])
            # Сколько уже загружено заказов
            loaded_posting_count += 1
            # Сколько осталось выгрузить заказов в процентах
            loaded_percent = round(loaded_posting_count / total_posting_count * 100, 2)
            logger.info(str(loaded_percent) + ' % ' + "Upload FBS sales from "+ orderDate.replace("['",'').replace("']",''))
            # Делаем паузу 3 с., чтоб ы часто не вызывать метод апи
            # time.sleep(3)

        # Переименовываем некоторые колонки для удобства и убираем лишние
        sales_fbs = sales_fbs.rename(columns={
            'offer_id': 'Артикул',
            'name': 'Название товара',
            'quantity': 'Количество',
            'currency_code': 'Валюта отправления'
        })
        # Указываем, что схема доставки FBS
        sales_fbs['Схема доставки'] = 'FBS'
        sales_fbs = sales_fbs.loc[:, ~sales_fbs.columns.isin(['price', 'digital_codes', 'mandatory_mark', 'dimensions'])]
    # Если продаж не было, сохраняем пустой df с продажами
    else:
        colnames_sales_fbs = ['Артикул', 'Название товара', 'sku', 'Количество', 'Валюта отправления',
       'Номер_заказа', 'Номер_отправления', 'Склад', 'Кластер_отправления',
       'Кластер_доставки', 'Сумма_отправления', 'Статус', 'Схема доставки']
        sales_fbs = pd.DataFrame(columns=colnames_sales_fbs)
    sales_fbs.to_csv(file_name_sales_fbs, sep=';', index=False, encoding='utf-8-sig')

# Остатки fbo (старый метод)
def getStockRemainders_fbo(headers):

    # Задаем имя файла для сохранения
    file_name_stock_rem_fbo = (
        uploaddir_today /
        f"{str(date.today())}_Остатки_fbo.csv"
    )
    # Задаем имя файла для временных остатков
    file_name_stock_rem_fbo_temp = (
        uploaddir_today /
        f"{str(date.today())}_Остатки_fbo_temp.csv"
    )
    # Задаем имя файла для обработанных данных по остаткам
    file_name_stock_rem_fbo_processed = (
        uploaddir_today /
        f"{str(date.today())}_Остатки_fbo_all.csv"
    )


    logger.info("Upload stock remainders on  "+ str(date.today()))
    # df, куда будет помещаться итоговый результат
    df_stock_rem_fbo = pd.DataFrame()
    # забираем остатки во временный файл
    for i in range(0,10000,1000):
        params = json.dumps({
        "limit": 1000,
        "offset": i,
        "warehouse_type": "ALL"
        })

        result = requests.post(f"{ozon_seller_api_url}/v2/analytics/stock_on_warehouses",headers=headers, data=params)
        # Если есть остатки, то сохраняем результат запроса в файл и начинаем обработку файла
        if result.status_code == 200 and result.json()['result']['rows']:
            remind = result.json()
            with open(file_name_stock_rem_fbo_temp, "a", newline="",encoding="utf-8-sig") as f:
            # with open(f"UploadFiles/UploadFiles_"+str(date.today())+'/'+str(date.today())+'_Остатки_temp'+'_from_'+date_start_file+'_to_'+date_end_file+'.csv', "a", newline="",encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, remind['result']['rows'][0].keys(), delimiter =',')
                w.writeheader()
                for row in remind['result']['rows']:
                    w.writerow(row)
                f.close()
        # Если остатков нет, сохраняем пустой файл
        else:
            df_stock_rem_fbo = pd.DataFrame(columns=['sku','Кластер','Название склада','Артикул','Название товара','Товары в пути','Доступный к продаже товар','Резерв'])
            df_stock_rem_fbo.to_csv(file_name_stock_rem_fbo, sep=";", index=False, encoding='utf-8-sig')

    # Если есть остатки, начинаем обработку файла с остатками
    if os.path.exists(file_name_stock_rem_fbo_temp):
        headerList = ['sku','Название склада','Артикул','Название товара','Товары в пути','Доступный к продаже товар','Резерв']
        df = pd.read_csv(file_name_stock_rem_fbo_temp, delimiter=",")
        if 'idc' in df.columns:
            df.drop(columns=['idc'], inplace=True)
        df = df[df['sku'] != 'sku']
        df.to_csv(file_name_stock_rem_fbo_processed, sep=";",header=headerList)
        # джойним остатки с кластерами и записываем в файл с остатками
        colnames=['sku','warehouse_name','item_code','item_name','promised_amount','free_to_sell_amount','reserved_amount']
        df_stok_rem = pd.read_csv(file_name_stock_rem_fbo_temp, delimiter=";",names=colnames,skiprows=1)
        colnames_df2 = ['warehouse_name', 'cluster']
        df_cluster = pd.read_csv('clasters_warehouse.csv',delimiter=";", names=colnames_df2,skiprows=1)
        data_with_cluster = df_stok_rem.merge(df_cluster, left_on='warehouse_name',right_on='warehouse_name',how='left')
        new_df = data_with_cluster[['sku','cluster','warehouse_name','item_code','item_name','promised_amount','free_to_sell_amount','reserved_amount']]
        headerList = ['sku','Кластер','Название склада','Артикул','Название товара','Товары в пути','Доступный к продаже товар','Резерв']
        # Кластеры, которых нет в списке, помечаем как "Неизвестный кластер"
        new_df['cluster'] = new_df['cluster'].fillna('Неизвестный кластер')
        new_df.to_csv(file_name_stock_rem_fbo,sep=";",header=headerList, index=False, encoding='utf-8-sig')
        os.remove(file_name_stock_rem_fbo_temp)
        os.remove(file_name_stock_rem_fbo_processed)
        # Выводим предупреждение, если есть кластеры, отсутствующие в файле
        warehouse_list_stock_rem = pd.DataFrame({'warehouse_name':df_stok_rem['warehouse_name'].unique()})
        df_all_clusters = pd.merge(df_cluster, warehouse_list_stock_rem,
                                on='warehouse_name',
                                how='outer',
                                indicator=True)
        missing_clusters = df_all_clusters.loc[df_all_clusters['_merge'] == 'right_only', 'warehouse_name'].to_list()
        if len(missing_clusters) > 0:
            logger.warning(f"No cluster found for warehouses {missing_clusters}")


# Остатки FBO
def getStockReminders_fbo_v2(headers, df_products=None, to_save=True):

    logger.info(f"Выгрузка остатков FBO на {str(date.today())}")

    # Интервал между запросами (в секундах)
    sleep_interval = 5
    # Задаем имя файла для товаров
    file_name_products = (
        uploaddir_today /
        f"{str(date.today())}_Товары.csv"
    )
    # Задаем имя файла для сохранения остатков
    file_name_stock_rem_fbo = (
        uploaddir_today /
        f"{str(date.today())}_Остатки_fbo.csv"
    )
    # Если список товаров не задан, считываем его из текущей выгрузки
    if df_products is None:
        df_products_ = pd.read_csv(file_name_products, sep=';')
    # Если список товаров задан, берем его из переданного аргумента
    else:
        df_products_ = df_products.copy()
    # Удаляем на всякий случай дубликаты
    df_products_ = df_products_.drop_duplicates(subset=['SKU'])
    # Удаляем товары с SKU = 0 (скрытые товары)
    df_products_ = df_products_.loc[~df_products_['SKU'].isin([0, '0']), :]
    # Сбрасываем индекс
    df_products_ = df_products_.reset_index(drop=True)
    # Разбиваем список товаров на диапазоны по 100 шт
    step = 100
    df_products_['chunks'] = (df_products_.index // step) + 1

    # df, куда будем помещать результаты выгрузки АПИ
    df_fbo_reminders_api = pd.DataFrame()

    # Переменная для отслеживания количества товаров
    uploaded_products_count = 0
    total_products_count = df_products_.shape[0]

    # Цикл по каждому интервалу
    for chunk in df_products_['chunks'].unique():
        # Делаем выборку одного интервала
        df_products_chunk = df_products_.loc[df_products_['chunks'] == chunk, :]
        # Считаем, сколько товаров в данном интервале
        products_count_in_chunk = df_products_chunk.shape[0]
        # Формируем лист с SKU
        sku_list = df_products_chunk['SKU'].to_list()
        # Переводим SKU в строку
        sku_list = [str(sku) for sku in sku_list]

        # Параметры запроса
        params_fbo_reminders = json.dumps({
            'skus': sku_list
        })

        # Запрос к АПИ
        resp_data_reminders_fbo = requests.post(
            f"{ozon_seller_api_url}//v1/analytics/stocks",
            headers=headers,
            data=params_fbo_reminders
        )
        if resp_data_reminders_fbo.status_code == 200:
            resp_data_reminders_fbo = resp_data_reminders_fbo.json()
            # Переводим в df
            tmp_df_fbo_reminders = pd.DataFrame(resp_data_reminders_fbo['items'])

            # Объединяем с предыдущим проходом цикла
            df_fbo_reminders_api = pd.concat([df_fbo_reminders_api, tmp_df_fbo_reminders])

            # Обновляем количество загруженных товаров
            uploaded_products_count += products_count_in_chunk

            logger.info(
                f"Выгружены остатки для {uploaded_products_count} из {total_products_count} товаров "
                f"(интервал {chunk}/{df_products_['chunks'].max()})"
            )

            # Делаем паузу между запросами
            logger.info(f"Пауза {sleep_interval} секунд перед следующим запросом...")
            time.sleep(sleep_interval)

        else:
            logger.error(f"Ошибка получения остатков FBO: {resp_data_reminders_fbo.text}")

    # Переименовываем и выбираем нужные колонки
    rename_dict = {
        'sku': 'SKU',
        'offer_id': 'Артикул',
        'name': 'Наименование товара',
        'cluster_name': 'Кластер',
        'warehouse_name': 'Склад',
        'return_from_customer_stock_count': 'Товары в процессе возврата',
        'requested_stock_count': 'Товары в заявках на поставку',
        'transit_stock_count': 'Товары в пути',
        'available_stock_count': 'Доступный к продаже товар'
    }
    df_fbo_reminders_api = df_fbo_reminders_api.rename(columns=rename_dict)

    df_fbo_reminders = (
        df_fbo_reminders_api
        .loc[:, df_fbo_reminders_api.columns.isin([
            'SKU',
            'Артикул',
            'Наименование товара',
            'Кластер',
            'Склад',
            'Товары в процессе возврата',
            'Товары в заявках на поставку',
            'Товары в пути',
            'Доступный к продаже товар'
        ])]
    )
    # Если пришел пустой df, то сохраняем df с колонками, которые приходят по апи
    if df_fbo_reminders.empty:
        df_fbo_reminders = pd.DataFrame(columns=list(rename_dict.values()))
    # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    if to_save:
        df_fbo_reminders.to_csv(file_name_stock_rem_fbo, sep=";", index=False, encoding='utf-8-sig')

    # иначе - возвращаем df с товарами
    else:
        return df_fbo_reminders


# Остатки FBS
def getStockReminders_fbs(headers, to_save=True):

    # Задаем имя файла для остатков FBS
    file_name_stock_rem_fbs = (
        uploaddir_today /
        f"{str(date.today())}_Остатки_fbs.csv"
    )

    # Колонки, которые всегда возвращаются по апи в таком порядке
    fbs_reminders_columns = ['Идентификатор склада', 'Название склада', 'Артикул', 'Наименование товара', 'Доступно на моем складе, шт', 'Зарезервировано на моем складе, шт']
    # Получаем список складов FBS
    resp_data_fbs_warehouses= requests.post(f"{ozon_seller_api_url}/v1/warehouse/list", headers=headers).json()
    fbs_warehouses = pd.DataFrame(resp_data_fbs_warehouses['result'])
    # Если нет складов fbs, то сохраняем пустой df
    if fbs_warehouses.shape[0] == 0:
        df_fbs_reminders = pd.DataFrame(columns=fbs_reminders_columns)
    # Если есть склады FBS, забираем по ним остатки
    else:
        params_warehouses = json.dumps({
            "language": "DEFAULT",
            "warehouseId": fbs_warehouses['warehouse_id'].to_list()
        })
        result = requests.post(f"{ozon_seller_api_url}/v1/report/warehouse/stock", headers=headers, data=params_warehouses).json()
        code_id = json.dumps({
        "code": result["result"]["code"]
        })
        logger.info('Get name of fbs warehouses report')
        logger.info(result["result"]["code"])
        time.sleep(5)

        # забираем ссылку на отчет с остатками на складе fbs
        resp_data = {"result":{"file":""}}
        while not resp_data["result"]["file"]:
            resp_data = requests.post(f"{ozon_seller_api_url}/v1/report/info", headers=headers, data=code_id).json()
            time.sleep(2)  # добавьте небольшую задержку, чтобы избежать лишних запросов

        time.sleep(5)  # дайте пару секунд на саму генерацию отчета
        path_csv = resp_data["result"]["file"]
        logger.info('Get link on FBS warehouses report')
        # logger.info(f"FBS stock reminders report link: {path_csv}")

        # Качаем файл
        MAX_TRIES = 20
        DELAY = 3  # секунд между попытками

        for i in range(MAX_TRIES):
            resp = requests.get(path_csv, stream=True)
            if resp.status_code == 200:
            # and int(resp.headers.get("Content-Length", 0)) > 5000:
                break
            time.sleep(DELAY)
        else:
            raise Exception("Файл так и не стал доступен для скачивания")

        # Читаем в pandas из памяти
        df_fbs_reminders = pd.read_excel(BytesIO(resp.content))


        # Если на складах есть остатки, то переименовываем колонки в нормальные и сохраняем в csv
        if df_fbs_reminders.shape[0] > 0:
            df_fbs_reminders.columns = fbs_reminders_columns
        # Если нет, то создаём пустой DataFrame с нужными колонками
        else:
            df_fbs_reminders = pd.DataFrame(columns=fbs_reminders_columns)
    # Если стоит флаг сохранения, то сохраняем файл в папку текущей выгрузки
    if to_save:
        df_fbs_reminders.to_csv(file_name_stock_rem_fbs, sep=";", encoding='utf-8-sig', index=False)
    # иначе - возвращаем df с товарами
    else:
        return df_fbs_reminders



# Функция выгрузки списка заявок на поставку
def get_supply_order_list(
    headers: dict,
    dropoff_warehouse_ids: list = None,
    order_number_search: list = None,
    states: list = [
        "READY_TO_SUPPLY",
        "IN_TRANSIT",
        "ACCEPTED_AT_SUPPLY_WAREHOUSE",
        "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
        "DATA_FILLING"
    ],
    timeslot_from: str = None,
    timeslot_to: str = None,
    timeslot_filter_type: str = "UNSPECIFIED",
    limit: int = 100,
    sort_by: str = "ORDER_CREATION",
    sort_dir: str = "DESC",
    uploaddir_today: Path = Path('.'),
    marketplace_dir_path: Path = Path('.'),
    method_url: str = "/v3/supply-order/list"
):
    """
    Выгружает заявки на поставку через Ozon Seller API с поддержкой пагинации и фильтрации по списку номеров заявок.

    Параметры:
    ----------
    headers : dict
        Заголовки запроса с авторизационным токеном Ozon API.

    dropoff_warehouse_ids : list of str or int, optional
        Список идентификаторов пунктов отгрузки для фильтрации.

    order_number_search : list of str, optional
        Список номеров заявок на поставку. Если указан и не пустой —
        выгружаются все заявки с этими номерами (без фильтрации по статусу).

    states : list of str, optional
        Список статусов заявок на поставку для фильтрации.
        Игнорируется, если указан параметр order_number_search.

    timeslot_from : str, optional
        Дата начала периода для фильтрации по таймслоту (в формате ISO8601: 'YYYY-MM-DDTHH:MM:SSZ').

    timeslot_to : str, optional
        Дата окончания периода для фильтрации по таймслоту (в формате ISO8601).

    timeslot_filter_type : str, default="UNSPECIFIED"
        Тип фильтрации по времени таймслота: "UNSPECIFIED", "BY_LOCAL_TIME" или "BY_UTC_TIME".

    limit : int, default=100
        Количество заявок на одной странице (максимум 100).

    sort_by : str, default="ORDER_CREATION"
        Критерий сортировки заявок: "UNSPECIFIED", "ORDER_CREATION", "ORDER_STATE_UPDATED_AT", "TIMESLOT_FROM_UTC", "TIMESLOT_FROM_LOCAL".

    sort_dir : str, default="DESC"
        Направление сортировки: "UNSPECIFIED", "ASC", "DESC".

    uploaddir_today : pathlib.Path, default=Path('.')
        Путь к директории для сохранения результата выгрузки.

    marketplace_dir_path : pathlib.Path, default=Path('.')
        Базовый путь для внутренних служебных скриптов (необязательный параметр).

    method_url : str, default="/v3/supply-order/list"
        URL метода Seller API Ozon для получения списка заявок на поставку.

    Возвращает:
    -----------
    pd.DataFrame
        DataFrame с идентификаторами всех найденных заявок (order_id).
    """
    supply_api_url = f'{ozon_seller_api_url}/{method_url}'
    file_name_supply_orders = uploaddir_today / f"{str(date.today())}_Поставки.csv"
    all_order_ids = []

    use_order_number_search = order_number_search is not None and len(order_number_search) > 0

    # Сценарий 1: по списку номеров заявок
    if use_order_number_search:
        logger.info("Выгрузка по списку номеров заявок")
        # При поиске по номеру статусы игнорируем — ищем любые (или UNSPECIFIED)
        for order_num in order_number_search:
            last_id = None
            while True:
                filter_block = {
                    "order_number_search": order_num,
                    # По ТЗ — если по номеру, статус не фильтруем
                    "states": [
                        "UNSPECIFIED",
                        "DATA_FILLING",
                        "READY_TO_SUPPLY",
                        "ACCEPTED_AT_SUPPLY_WAREHOUSE",
                        "IN_TRANSIT",
                        "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
                        "REPORTS_CONFIRMATION_AWAITING",
                        "REPORT_REJECTED",
                        "COMPLETED",
                        "REJECTED_AT_SUPPLY_WAREHOUSE",
                        "CANCELLED",
                        "OVERDUE"
                    ]
                }
                if dropoff_warehouse_ids:
                    filter_block["dropoff_warehouse_ids"] = dropoff_warehouse_ids
                if timeslot_from or timeslot_to:
                    filter_block["timeslot_from_range"] = {
                        "from": timeslot_from,
                        "to": timeslot_to,
                        "timeslot_filter_type": timeslot_filter_type
                    }
                else:
                    filter_block["timeslot_from_range"] = None

                payload = {
                    "filter": {k: v for k, v in filter_block.items() if v not in [None, [], {}]},
                    "limit": limit,
                    "last_id": last_id,
                    "sort_by": sort_by,
                    "sort_dir": sort_dir
                }
                payload = {k: v for k, v in payload.items() if v not in [None, [], {}]}

                logger.debug(f"Поиск заявки с номером {order_num}, last_id={last_id}")
                response = requests.post(supply_api_url, headers=headers, json=payload)
                if response.status_code != 200:
                    logger.error(f"Ошибка ответа API: {response.status_code} — {response.text}")
                    break

                result = response.json()
                order_ids = result.get("order_ids", [])
                last_id = result.get("last_id", None)

                logger.info(f"Найдено заявок с номером {order_num}: {len(order_ids)}")
                all_order_ids.extend(order_ids)

                if not order_ids or not last_id:
                    break

    # Сценарий 2: стандартная выгрузка по статусам и остальным фильтрам
    else:
        logger.info('Выгрузка по статусам и другим фильтрам')
        last_id = None
        while True:
            filter_block = {
                "states": states
            }
            if dropoff_warehouse_ids:
                filter_block["dropoff_warehouse_ids"] = dropoff_warehouse_ids
            if timeslot_from or timeslot_to:
                filter_block["timeslot_from_range"] = {
                    "from": timeslot_from,
                    "to": timeslot_to,
                    "timeslot_filter_type": timeslot_filter_type
                }
            else:
                filter_block["timeslot_from_range"] = None

            payload = {
                "filter": {k: v for k, v in filter_block.items() if v not in [None, [], {}]},
                "limit": limit,
                "last_id": last_id,
                "sort_by": sort_by,
                "sort_dir": sort_dir
            }
            payload = {k: v for k, v in payload.items() if v not in [None, [], {}]}

            logger.debug(f"Запрос данных с last_id={last_id}")
            response = requests.post(supply_api_url, headers=headers, json=payload)
            if response.status_code != 200:
                logger.error(f"Ошибка ответа API: {response.status_code} — {response.text}")
                break

            result = response.json()
            order_ids = result.get("order_ids", [])
            last_id = result.get("last_id", None)

            logger.info(f"Загружено заявок: {len(order_ids)}")
            all_order_ids.extend(order_ids)

            if not order_ids or not last_id:
                break

    df_supply_order_list = pd.DataFrame({"order_id": all_order_ids})

    logger.success(f"Список заявок выгружен: выгружено {len(all_order_ids)} заявок")

    return df_supply_order_list

# Функция получения деталей по заявкам на поставку
def get_supply_orders_details(
    df_supply_order_list: pd.DataFrame,
    headers: dict,
    batch_size: int = 50,
    method_url: str = "/v3/supply-order/get"
):
    """
    Получает DataFrame с деталями по заявкам на поставку через Seller API Ozon.

    Параметры:
    ----------
    df : pd.DataFrame
        DataFrame с колонкой 'order_id'.

    headers : dict
        Заголовки для API.

    batch_size : int, default=100
        Размер блока запросов к API.

    method_url : str, default="/v3/supply-order/get"
        URL метода Ozon Seller API.

    Возвращает:
    -----------
    pd.DataFrame
        DataFrame с деталями (order_id, order_number, dropoff_warehouse_id, dropoff_warehouse_name,
        bundle_id, storage_warehouse_id, storage_warehouse_name, supply_id)
    """
    supply_info_url = f'{ozon_seller_api_url}/{method_url}'
    order_ids = df_supply_order_list["order_id"].tolist()
    orders_details = []

    logger.info(f"Запросим детали по {len(order_ids)} заявкам (batch_size={batch_size})")

    for i in range(0, len(order_ids), batch_size):
        batch_ids = order_ids[i:i+batch_size]
        payload = {"order_ids": batch_ids}
        response = requests.post(supply_info_url, headers=headers, json=payload)
        if response.status_code != 200:
            logger.error(f"Ошибка API {supply_info_url}: {response.status_code} — {response.text}")
            continue
        result = response.json()
        orders = result.get("orders", [])
        if not orders:
            logger.warning(f"В ответе пустой список 'orders'.")
        orders_details.extend(orders)

    # Обработка и вытаскивание нужных полей в итоговый список
    extracted_rows = []
    for order in orders_details:
        order_id = order.get("order_id")
        order_number = order.get("order_number")
        drop = order.get("drop_off_warehouse", {})
        drop_id = drop.get("warehouse_id")
        drop_name = drop.get("name")
        supplies = order.get("supplies", [])
        if supplies:
            for supply in supplies:
                bundle_id = supply.get("bundle_id")
                supply_id = supply.get("supply_id")
                storage = supply.get("storage_warehouse", {})
                storage_id = storage.get("warehouse_id")
                storage_name = storage.get("name")
                extracted_rows.append({
                    "order_id": order_id,
                    "order_number": order_number,
                    "dropoff_warehouse_id": drop_id,
                    "dropoff_warehouse_name": drop_name,
                    "bundle_id": bundle_id,
                    "storage_warehouse_id": storage_id,
                    "storage_warehouse_name": storage_name,
                    "supply_id": supply_id
                })
        else:
            # На случай если supplies пустой, просто добавим по заявке (bundle, storage — None)
            extracted_rows.append({
                "order_id": order_id,
                "order_number": order_number,
                "dropoff_warehouse_id": drop_id,
                "dropoff_warehouse_name": drop_name,
                "bundle_id": None,
                "storage_warehouse_id": None,
                "storage_warehouse_name": None,
                "supply_id": None
            })

    df_supply_details = pd.DataFrame(extracted_rows)

    # Проверка — все ли order_ids были в ответе
    received_orders = set(df_supply_details["order_id"].unique())
    requested_orders = set(order_ids)
    missing = requested_orders - received_orders
    if missing:
        logger.warning(f"Не получены детали для order_ids: {missing}")
    else:
        logger.success("Все order_ids найдены в ответе!")

    return df_supply_details

# Функция получения товаров в заявках на поставку
def get_supply_bundle_items(
    df_supply_details: pd.DataFrame,
    headers: dict,
    batch_size: int = 100,  # Это про limit для API, не кол-во bundle_id за цикл
    method_url: str = "/v1/supply-order/bundle"
):
    """
    Получает DataFrame с товарами для каждого bundle_id, с привязкой к заявке на поставку.

    Параметры:
    ----------
    df_details : pd.DataFrame
        DataFrame c результатами get_supply_orders_details_df (bundle_id + мета-инфо по поставкам).
    headers : dict
        Заголовки для API.
    batch_size : int, default=100
        Сколько товаров забирать за один запрос (limit).
    api_url : str
        URL метода Ozon Seller API.

    Возвращает:
    -----------
    pd.DataFrame
        DataFrame со всеми нужными полями:
        SKU, Название товара, product_id, количество, order_id, order_number, dropoff_warehouse_id,
        dropoff_warehouse_name, bundle_id, storage_warehouse_id, storage_warehouse_name, supply_id
    """
    # Собираем только уникальные bundle_id с метаданными к ним (остальные поля будут подмешаны)
    bundle_meta = df_supply_details.drop_duplicates('bundle_id').set_index('bundle_id')
    all_rows = []
    bundle_info_url = f'{ozon_seller_api_url}/{method_url}'

    unique_bundle_ids = df_supply_details["bundle_id"].dropna().unique().tolist()
    logger.info(f"Всего уникальных bundle_id для запроса: {len(unique_bundle_ids)}")

    for idx, bundle_id in enumerate(unique_bundle_ids, start=1):
        last_id = None
        has_next = True
        page_num = 1
        bundle_info = bundle_meta.loc[bundle_id].to_dict()
        order_id = bundle_info.get('order_id')
        order_number = bundle_info.get('order_number')

        logger.info(f"({idx}/{len(unique_bundle_ids)}) Обработка bundle_id={bundle_id} — заявка order_number={order_number}, order_id={order_id}")

        items_collected = 0

        while has_next:
            payload = {
                "bundle_ids": [bundle_id],
                "is_asc": True,
                "limit": batch_size,
                "last_id": last_id,
                "sort_field": "UNSPECIFIED"
            }
            response = requests.post(bundle_info_url, headers=headers, json=payload)
            if response.status_code != 200:
                logger.error(f"Ошибка API: {response.status_code} for bundle_id={bundle_id} — {response.text}")
                break
            result = response.json()
            items = result.get("items", [])
            items_collected += len(items)
            # logger.info(f"    bundle_id={bundle_id}, страница {page_num}, получено товаров: {len(items)}, всего собрано: {items_collected}")

            for item in items:
                all_rows.append({
                    "offer_id": item.get("offer_id"),
                    "sku": item.get("sku"),
                    "name": item.get("name"),
                    "product_id": item.get("product_id"),
                    "quantity": item.get("quantity"),
                    **bundle_info
                })
            has_next = result.get("has_next", False)
            last_id = result.get("last_id", None)
            page_num += 1
            time.sleep(1)

        logger.info(f"Завершена обработка bundle_id, всего товаров собрано: {items_collected}")

    columns_out = [
        'offer_id', 'sku', 'name', 'product_id', 'quantity', 'order_id', 'order_number',
        'dropoff_warehouse_id', 'dropoff_warehouse_name', 'bundle_id',
        'storage_warehouse_id', 'storage_warehouse_name', 'supply_id'
    ]
    df_supply_items = pd.DataFrame(all_rows)

    df_supply_items = df_supply_items[[c for c in columns_out if c in df_supply_items.columns]]

    # Переименовываем колонки
    df_supply_items = df_supply_items.rename(columns={
        'offer_id': 'Артикул',
        'sku': 'SKU',
        'product_id': 'Ozon Product ID',
        'name': 'Наименование товара',
        'barcode': 'Штрихкод',
        'quantity': 'Количество',
    })
    # Считаем количество товаров по собранным заявкам
    total_items = df_supply_items['Количество'].sum()
    logger.success(f"Итого товаров в полученных заявках: {total_items}")

    return df_supply_items

# Функция добавления кластеров к товарам в заявках на поставку
def add_clusters_to_supply_order_items(
        df_supply_order_items: pd.DataFrame
) -> pd.DataFrame:
    """
    Добавляет информацию о кластерах к товарам в заявках на поставку.

    Параметры:
    -----------
    df_supply_items : pd.DataFrame
        DataFrame с товарами в заявках на поставку.

    Возвращает:
    -----------
    pd.DataFrame
        DataFrame с добавленной информацией о кластерах.
    """
    # Считываем файл с маппингом кластеров
    df_clusters = pd.read_csv(
        marketplace_dir_path
        / 'scripts' / 'clusters_and_warehouses.csv',
        sep=';'
    )
    # Задаем колонку, по которой мерджим
    supply_orders_col = 'storage_warehouse_id'
    # Мерджим со списком кластеров
    df_supply_items_with_clusters = df_supply_order_items.merge(
        df_clusters[['warehouse_id', 'cluster_name']],
        left_on=supply_orders_col,
        right_on='warehouse_id',
        how='left',
        indicator=True,
    )
    # Выводим сообщение, если какие-то кластеры не подтянулись
    missing_clusters = df_supply_items_with_clusters[df_supply_items_with_clusters['_merge'] == 'left_only']
    if not missing_clusters.empty:
        logger.warning(f"Не найдены кластеры для следующих складов: {missing_clusters[supply_orders_col].unique()}")
    # Удаляем временную колонку
    df_supply_items_with_clusters = df_supply_items_with_clusters.drop(columns=['_merge'])
    # Переименовываем колонки
    df_supply_items_with_clusters = df_supply_items_with_clusters.rename(columns={
        'cluster_name': 'Кластер',
    })

    return df_supply_items_with_clusters

# Функция получения и обработки товаров в заявках на поставку
def get_supply_orders_full(
    headers: dict,
    dropoff_warehouse_ids: list = None,
    order_number_search: list = None,
    states: list = [
        # "READY_TO_SUPPLY",
        "IN_TRANSIT",
        "ACCEPTED_AT_SUPPLY_WAREHOUSE",
        "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
        "DATA_FILLING"
    ],
    timeslot_from: str = None,
    timeslot_to: str = None,
    timeslot_filter_type: str = "UNSPECIFIED",
    limit: int = 100,
    sort_by: str = "ORDER_CREATION",
    sort_dir: str = "DESC",
    uploaddir_today: Path = Path('.'),
    marketplace_dir_path: Path = Path('.'),
    batch_size: int = 100,
    order_list_url: str = "/v3/supply-order/list",
    order_details_url: str = "/v3/supply-order/get",
    bundle_items_url: str = "/v1/supply-order/bundle",
    to_save: bool = True,
):
    """
    Полный цикл: список заявок -> детали заявок -> товары в заявках.
    Возвращает итоговый DataFrame с деталями по товарам.
    """
    # Колонки на случай отсутствия заявок
    supply_order_columns = [
        'Артикул',
        'SKU',
        'Ozon Product ID',
        'Наименование товара',
        'Количество',
        'Кластер',
    ]
    df_empty_supply_orders = pd.DataFrame(columns=supply_order_columns)

    # ===== 1. Получаем список заявок =====
    df_supply_order_list = get_supply_order_list(
        headers=headers,
        dropoff_warehouse_ids=dropoff_warehouse_ids,
        order_number_search=order_number_search,
        states=states,
        timeslot_from=timeslot_from,
        timeslot_to=timeslot_to,
        timeslot_filter_type=timeslot_filter_type,
        limit=limit,
        sort_by=sort_by,
        sort_dir=sort_dir,
        uploaddir_today=uploaddir_today,
        marketplace_dir_path=marketplace_dir_path,
        method_url=order_list_url
    )
    # Если были найдены заявки на поставку, получаем по ним инфу
    if not df_supply_order_list.empty:
        # ===== 2. Получаем детали заявок =====
        df_supply_details = get_supply_orders_details(
            df_supply_order_list=df_supply_order_list,
            headers=headers,
            batch_size=50,
            method_url=order_details_url
        )
        if df_supply_details.empty:
            logger.warning("Не удалось получить детали заявок.")


        # ===== 3. Получаем товары по поставкам =====
        df_supply_order_items = get_supply_bundle_items(
            df_supply_details=df_supply_details,
            headers=headers,
            batch_size=batch_size,
            method_url=bundle_items_url
        )
        if df_supply_order_items.empty:
            logger.warning("Не удалось получить товары в заявках.")

        # ===== 4. Добавляем кластеры к товарам в заявках на поставку =====
        df_supply_order_items_with_clusters = add_clusters_to_supply_order_items(
            df_supply_order_items=df_supply_order_items,

        )

        logger.success("Выгрузка всех этапов заявок на поставку завершена успешно.")
    else:
        df_supply_order_items_with_clusters = df_empty_supply_orders.copy()
        logger.info("Не найдено ни одной заявки на поставку.")

    if to_save:
        # Задаем имя файла для сохранения
        file_name_supply_orders = (
            uploaddir_today / f"{str(date.today())}_Поставки.csv"
        )
        df_supply_order_items_with_clusters.to_csv(
            file_name_supply_orders,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
    else:
        return df_supply_order_items_with_clusters


# %% Вызов всех функций
if __name__ == '__main__':

    # Создаем директорию для текущей выгрузки
    uploaddir_today = create_upload_dir()

    # Парсим аргументы командной строки
    days_ago, cli_start_date, cli_end_date = cli_days_ago(parser)

    date_start, date_end = generate_dates_new(
        days_ago=days_ago,
        start_date_str=cli_start_date,
        end_date_str=cli_end_date,
        # days_ago=None,
        # start_date_str = '2025-11-30',
        # end_date_str = '2025-11-30',
    )
    # date_end = '2025-07-12T23:59:59.000Z'
    # date_start = '2025-06-13T00:00:00.000Z'
    logger.info(f"Upload Files from Ozon for client {client_name} for dates {date_start} - {date_end}")

    # Сохранение периода выгрузки данных в csv
    save_dates_to_csv(date_start, date_end)

    # Обновление кластеров Озон
    update_cluster_list(headers)

    # Список товаров
    get_ozon_product(headers)

    # Заказы FBO
    getOrdersFBO(headers, date_start, date_end)
    # Заказы FBS
    getOrdersFBS(headers, date_start, date_end)

    # # Заказы FBO
    # getOrders(headers, date_start, date_end, delivery_schema="fbo")
    # # Заказы FBS
    # getOrders(headers, date_start, date_end, delivery_schema="fbs")
    # # Кластер доставки заказов FBO
    # getFinalOrders_fbo(headers)
    # # Кластер доставки заказов FBS
    # getFinalOrders_fbs(headers)


    # Заказы
    # getOrders_v2(headers, date_start, date_end)

    # Список транзакций
    get_transactions_report(headers, date_start, date_end)
    # Продажи FBO
    # getSalesFBO(headers)
    getSalesFBO_v2(headers, date_start, date_end, upload_transactions=False)
    # Продажи FBS
    # getSalesFBS(headers)
    getSalesFBS_v2(headers, date_start, date_end, upload_transactions=False)

    # Возвраты (FBO + FBS)
    getReturns_v2(headers, date_start, date_end)

    # Остатки FBS
    getStockReminders_fbs(headers)
    # Остатки FBO
    getStockReminders_fbo_v2(headers)

    # Поставки FBO
    get_supply_orders_full(headers, uploaddir_today=uploaddir_today)

    logger.info(f"DONE UPLOADING FILES FOR CLIENT {client_name}")
    logger.info(f"Uploaded files saved in: {uploaddir_today}")
    print(f"{Back.BLUE}{Fore.WHITE}DONE UPLOADING FILES FOR CLIENT {client_name}{Style.RESET_ALL}")

# %%
