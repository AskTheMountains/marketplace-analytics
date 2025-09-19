
# %% Определение функций
import requests
import json
import time
from datetime import date,datetime,timedelta
import pandas as pd
import shutil
import os
import csv
import numpy as np
from openpyxl import Workbook
import re
from loguru import logger
import getopt
import sys
from pathlib import Path
pd.set_option('chained_assignment',None)

# Файл с некоторыми константами
from wb.scripts.constants import (
    headers,
    client_name,
    marketplace_dir_name,
)


# Имя директории для текущей выгрузки
uploaddir = marketplace_dir_name / "Clients" / client_name / "UploadFiles"
uploaddir_today = uploaddir / f"UploadFiles_{str(date.today())}"


# Функция генерации дат периода выгрузки данных
def generateDates(days_ago = 30):
    date_end = str(date.today()- timedelta(days=1)) + 'T23:59:59.000Z'
    date_start = str(date.today()- timedelta(days=1) - timedelta(days_ago)) + 'T00:00:00.000Z'
    logger.info(f"Uploading orders from "+ date_start + " to " + date_end)
    return date_end,date_start #,date_end_file,date_start_file

# GPT START ----
# Функция создания диапазона дат с указанной разницей
def generate_dates_new(days_ago=30):
    # Получаем сегодняшнюю дату и округляем к 23:59:59 предыдущих суток
    end_date = datetime.now()
    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=0) - timedelta(days=1)

    # Вычисляем начальную дату на 00:00:00
    start_date = end_date - timedelta(days=days_ago) + timedelta(days=1)
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Устанавливаем конечную дату на 23:59:59
    # end_date = start_date + timedelta(days=days_ago) - timedelta(seconds=1)

    # Преобразуем даты в нужный формат
    start_date_str = start_date.strftime('%Y-%m-%dT%H:%M:%S')
    end_date_str = end_date.strftime('%Y-%m-%dT%H:%M:%S')

    return start_date_str, end_date_str
# GPT END ----

# Функция сохранения дат в отдельный файл csv, чтобы было понятно, за какой период был сформирован отчет
def save_dates_to_csv(date_start, date_end):
    # Формат дат для названий файла
    date_end_file = datetime.fromisoformat(date_end).strftime('%Y-%m-%d')
    date_start_file = datetime.fromisoformat(date_start).strftime('%Y-%m-%d')
    # Датафрейм с датами начала и окончания выгрузки данных
    df_dates = pd.DataFrame({
        'date_end': date_end,
        'date_start': date_start,
        'date_end_file': date_end_file,
        'date_start_file': date_start_file
    }, index=[0])
    # Сохраняем файл с датами через pathlib
    dates_file = uploaddir_today / f"{str(date.today())}_dates_from_to.csv"
    df_dates.to_csv(dates_file, sep=';', index=False)



# Список товаров
def getWBProduct(headers, type_products='not_from_recycle', to_save=True):
    # Начальные значения для цикла
    # limit = 100
    # products_loaded = 0
    # total = 101
    # df_products = pd.DataFrame()
    # cursor = {"limit" : limit}
    # logger.info("Uploading products list")
    # # Выгружаем товары, пока не выгрузим все товары, которые содержатся в переменной total
    # while products_loaded < total:
    #     params = json.dumps({
    #     "settings": {
    #         "cursor": cursor,
    #         "filter": {
    #             "withPhoto": -1
    #         }
    #     }
    #     })
    #     resp_data = requests.post("https://content-api.wildberries.ru/content/v2/get/cards/list", headers=headers, data=params).json()
    #     # Если товаров больше 100, передаем данным переменным значения из предыдущей итерации цикла
    #     cursor = {"limit" : limit,
    #               # "updatedAt": resp_data['cursor']['updatedAt'],
    #               "nmID": resp_data['cursor']['nmID']
    #               }
    #     # Добавляем товары в df на каждом проходе цикла
    #     tmp_df = pd.DataFrame(resp_data['cards'])
    #     df_products = pd.concat([tmp_df, df_products])
    #     # Сколько было выгружено товаров
    #     products_loaded = len(df_products)
    #     # Сколько осталось выгрузить товаров
    #     total = resp_data['cursor']['total']

    # Настройки параметров запроса отдельно для товаров из корзины и не из корзины
    if type_products == 'not_from_recycle':
        products_url = "https://content-api.wildberries.ru/content/v2/get/cards/list"
        updated_name = 'updatedAt'
        start_params = {
            "settings": {
                "cursor": {},
                "filter": {
                    "withPhoto": -1
                }
            }
        }
        file_name_products = uploaddir_today / f"{str(date.today())}_Товары.csv"
    else:
        products_url = "https://content-api.wildberries.ru/content/v2/get/cards/trash"
        updated_name = 'trashedAt'
        start_params = {
            "settings": {
                "cursor": {}
            }
        }
        file_name_products = uploaddir_today / f"{str(date.today())}_Товары_Корзина.csv"
    # Начальные значения для цикла
    page_number = 1
    limit = 100
    # products_loaded = 0
    total = 101
    df_products = pd.DataFrame()
    cursor = {"limit" : limit}
    logger.info("Uploading products list")
    # Выгружаем товары, пока не выгрузим все товары, которые содержатся в переменной total
    while total >= limit:
        logger.info(f"Uploading page {page_number}")
        params = start_params
        params['settings']['cursor'] = cursor
        params = json.dumps(params)
        resp_data = requests.post(products_url, headers=headers, data=params).json()
        # Если есть товары
        if len(resp_data['cards']) > 0:
            # Если товаров больше 100, передаем данным переменным значения из предыдущей итерации цикла
            cursor = {"limit" : limit,
                    updated_name: resp_data['cursor'][updated_name],
                    "nmID": resp_data['cursor']['nmID']
                    }
            # Добавляем товары в df на каждом проходе цикла
            tmp_df = pd.DataFrame(resp_data['cards'])
            df_products = pd.concat([tmp_df, df_products])
            # Сколько было выгружено товаров
            # products_loaded = len(df_products)
            # Сколько осталось выгрузить товаров
            total = resp_data['cursor']['total']
            logger.info(f"Uploaded page {page_number}, products count: {total}")
            page_number = page_number + 1
        # Если нет товаров, то пустой df
        else:
            df_products = pd.DataFrame(columns=['cards'])
    logger.info("Done uploading product list")

    # Если есть товары, начинаем получение нужных атрибутов товаров
    if df_products.shape[0] > 0:
        df_products = df_products.reset_index(drop=True)
        # Распаковка размеров для одного nmID
        df_products_unpacked = df_products.explode(column=['sizes'], ignore_index=True)
        # Из словарей размеров достаем размер и sku
        df_products_unpacked = df_products_unpacked.assign(
            Размер=[d.get('techSize') for d in df_products_unpacked.sizes],
            sku_list=[d.get('skus') for d in df_products_unpacked.sizes],
            chrtID=[d.get('chrtID') for d in df_products_unpacked.sizes],
        )
        # Получаем массив баркодов
        df_skus = df_products_unpacked['sku_list'].explode()
        # Удаляем дубликаты из баркодов, выбирая последний
        df_skus = df_skus[~df_skus.index.duplicated(keep='last')]
        # Переименовываем Series с баркодами
        df_skus = df_skus.rename('last_barcode')
        # Добавляем их к списку товаров
        df_products_unpacked = pd.concat([
            df_products_unpacked,
            df_skus,
        ],
        axis=1)
        # Переводим из list в str
        df_products_unpacked['sku'] = [';'.join(map(str, l)) for l in df_products_unpacked['sku_list']]

        # Получение некоторых характеристик товара
        # Распаковка столбцов со словарями
        t = df_products_unpacked.explode("characteristics")
        products_characteristics = pd.concat([
                t[["sku", "vendorCode", "Размер"]].reset_index(drop=True),
                pd.json_normalize(t["characteristics"])],
            axis=1
        )
        # Цвет
        colors = products_characteristics.loc[products_characteristics['id'] == 14177449, :]
        df_products_unpacked = (
            df_products_unpacked
            .merge(colors[['vendorCode', 'Размер', 'value']],
                   how='left',
                   on=['vendorCode', 'Размер'])
            .rename(columns={'value':'Цвет'})
        )
        # df_products_unpacked['Цвет'] = ''
        # for i in range(len(df_products_unpacked)):
        #     # Скобки нужны, чтобы избавиться от двойного list
        #     df_products_unpacked['Цвет'][i] = [d.get('value') for d in df_products_unpacked['characteristics'][i] if d['id']==14177449]
        # #     # df_products_unpacked.loc[i, 'Цвет'] = [element for element in [d.get('value') for d in df_products_unpacked['characteristics'][i] if d['id']==14177449]]
        # #     #df_products_unpacked['Цвет'] = df_products_unpacked['characteristics'].apply(lambda x: [d.get('value') for d in x if d['id']==14177449])

        # df_products_unpacked['Цвет'] = [';'.join(map(str, l)) for l in df_products_unpacked['Цвет'] if not pd.isnull(l)]
        df_products_unpacked['Цвет'] = df_products_unpacked['Цвет'].str.join(';')

        # Переименование некоторых колонок для удобства
        df_products_unpacked = df_products_unpacked.rename(columns={
            "subjectName": "Предмет",
            "title": "Наименование товара",
            "vendorCode": "Артикул продавца",
            "sku": "barcode",
            })
        # На всякий случай удаление дубликатов
        df_products_unpacked = df_products_unpacked.drop_duplicates(subset=['Артикул продавца', 'Размер', 'barcode'], keep='first')
        # На всякий случай заполнение nan в размере
        df_products_unpacked['Размер'] = df_products_unpacked['Размер'].fillna(0)
        # Сохранение в csv нужных колонок
        df_products_unpacked_csv = df_products_unpacked[["nmID", "imtID", "chrtID", "Артикул продавца", "barcode", "last_barcode", "Наименование товара", "Предмет", "Размер", "Цвет"]]
        return df_products_unpacked_csv

    else:
        df_products.to_csv(file_name_products, sep=';', index=False)


# Цены товаров
def get_prices_WB(headers, df_products, to_save=True):
    df_prices = pd.DataFrame()
    # Диапазоны выгрузок (с шагом 1000 до кол-ва товаров)
    l = list(range(0, len(df_products), 1000)) + [len(df_products)]
    offset = 0
    logger.info("Uploading products prices")
    for z in zip(l[:-1], l[1:]):
        params = {
            "limit": 1000,
            "offset": z[0],
            # "filterNmID": ""
        }
        resp_data = requests.get("https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter",
                                 headers=headers, params=params).json()
        tmp_df = pd.DataFrame(resp_data['data']['listGoods'])
        df_prices = pd.concat([df_prices, tmp_df])
        time.sleep(5)
    # Распаковываем данные по каждому размеру
    df_prices = df_prices.explode(column=['sizes'])
    # Получаем id размера и цену
    df_prices = df_prices.assign(
        # ID размера
        chrtID=[d.get('sizeID') for d in df_prices.sizes],
        # Базовая цена
        price=[d.get('price') for d in df_prices.sizes],
        # Цена со скидкой
        discount_price=[d.get('discountedPrice') for d in df_prices.sizes],
        # Цена со скидкой клуба
        club_discount_price=[d.get('clubDiscountedPrice') for d in df_prices.sizes],
    )
    # Объединяем с df по товарам
    df_products_prices = df_products.merge(
        df_prices[['chrtID', 'price', 'discount', 'discount_price', 'club_discount_price']],
        on='chrtID',
        how='left'
    )
    # На всякий случай удаляем дубликаты
    df_products_prices = df_products_prices.drop_duplicates(subset=['nmID', 'chrtID'], keep='first')
    # На всякий случай заполнение nan в размере
    df_products_prices['Размер'] = df_products_prices['Размер'].fillna(0)
    if to_save:
        # Сохраняем df с ценами через pathlib
        products_prices_file = uploaddir_today / f"{str(date.today())}_Товары.csv"
        df_products_prices.to_csv(products_prices_file, sep=';', index=False)
    else:  return df_products_prices


# Функция объединения списка товаров
# def merge_products_data():

# Функция выгрузки списка товаров WB
def get_wb_product(headers, to_save=True):
    # Получаем список товаров
    df_products = getWBProduct(headers, to_save=False)
    # Получаем цены товаров
    df_products_prices = get_prices_WB(headers, df_products, to_save=False)
    # Если стоит флаг сохранения, то сохраняем df через pathlib
    if to_save:
        products_file = uploaddir_today / f"{str(date.today())}_Товары.csv"
        df_products_prices.to_csv(products_file, sep=';', index=False)

    return df_products_prices


# Заказы
def getOrdersWB(headers, date_start, date_end, to_save=True):
    params = {
        "dateFrom": date_start,
        "dateTo": date_end,
        # "flag": 0
    }
    logger.info("Uploading orders")
    try:
        resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/orders",
                                headers=headers, params=params)

        # Если ошибка, что много запросов, ждем 1 минуту, потом пробуем снова
        if resp_data.status_code == 429:
            logger.info("Waiting 1 minute for next attempt")
            time.sleep(65)
            resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/orders",
                                    headers=headers, params=params)
        # Если запрос выполнен успешно, забираем отчет из ответа
        if resp_data.status_code == 200:
            result = resp_data.json()
            df_orders = pd.DataFrame(result)
            df_orders = df_orders[(df_orders['date'] >= date_start) & (df_orders['date'] <= date_end)]
            df_orders = df_orders.rename(columns={
                "techSize": "Размер",
                "supplierArticle": "Артикул продавца",
                "nmId": "Артикул WB"
            })
            # Если стоит флаг сохранения, сохраняем отчет в csv, иначе - возвращаем df
            if to_save:
                orders_file = uploaddir_today / f"{str(date.today())}_Заказы.csv"
                df_orders.to_csv(orders_file, sep=';', index=False)
            else:
                return df_orders
    except Exception as e:
        logger.error(f"Ошибка получения заказов: {e}")


# Продажи
def getSalesWB(headers, date_start, date_end, to_save=True):
    params = {
        "dateFrom": date_start,
        "dateTo": date_end,
        # "flag": 0
    }
    logger.info("Uploading sales")
    try:
        resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/sales",
                                headers=headers, params=params)
        # Если ошибка, что много запросов, ждем 1 минуту, потом пробуем снова
        if resp_data.status_code == 429:
            logger.info("Waiting 1 minute for next attempt")
            time.sleep(65)
            resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/sales",
                                    headers=headers, params=params)
        # Если запрос выполнен успешно, забираем отчет из ответа
        if resp_data.status_code == 200:
            result = resp_data.json()
            df_sales = pd.DataFrame(result)
            df_sales = df_sales[(df_sales['date'] >= date_start) & (df_sales['date'] <= date_end)]
            df_sales = df_sales.rename(columns={
                            "techSize": "Размер",
                            "supplierArticle": "Артикул продавца",
                            "nmId": "Артикул WB"
                        })
            # Если стоит флаг сохранения, сохраняем отчет в csv, иначе - возвращаем df
            if to_save:
                sales_file = uploaddir_today / f"{str(date.today())}_Продажи.csv"
                df_sales.to_csv(sales_file, sep=';', index=False)
            else:
                return df_sales
    except Exception as e:
        logger.error(f"Ошибка получения продаж {e}")

# Остатки
def getStockRemindersWB(headers, date_start, date_end):
    params = {
        "dateFrom": "2010-01-01T00:00:00Z",
        # "flag": 0
    }
    logger.info("Uploading stock reminders")
    try:
        resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/stocks",
                                headers=headers, params=params)
        # Если ошибка, что много запросов, ждем 1 минуту, потом пробуем снова
        if resp_data.status_code == 429:
            logger.info("Waiting 1 minute for next attempt")
            time.sleep(65)
            resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/stocks",
                                    headers=headers, params=params)
        if resp_data.status_code == 200:
            result = resp_data.json()
            df_reminders = pd.DataFrame(result)
            df_reminders = df_reminders.rename(columns={
                "techSize": "Размер",
                "supplierArticle": "Артикул продавца",
                "nmId": "Артикул WB"
                })
            df_reminders.to_csv(uploaddir_today / f"{str(date.today())}_Остатки.csv", sep=';', index=False)
    except Exception as e:
        logger.error(f"Ошибка получения Остатков:{e}")


# Остатки fbs
def getStockReminders_fbs(headers):
    # Считываем df со списком товаров
    df_products = pd.read_csv(uploaddir_today / f"{str(date.today())}_Товары.csv", sep=';')
    # Получаем склады продавца
    resp_data_seller_warehouses = requests.get("https://marketplace-api.wildberries.ru/api/v3/warehouses", headers=headers).json()
    df_seller_warehouses = pd.DataFrame(resp_data_seller_warehouses)
    # Пустой df, в который будем добавлять остатки
    df_reminders_fbs = pd.DataFrame(columns=['sku', 'amount'])
    # Переводим баркоды в список
    df_products['barcode_list'] = df_products['barcode'].apply(
        lambda x: [str(item.strip()) for item in x.split(';')] if isinstance(x, str) else x
    )
    # Диапазоны выгрузок (с шагом 1000 до кол-ва товаров)
    step = 1000
    df_products['chunks'] = df_products.index.map(lambda x: int(x/step) + 1)
    logger.info("Uploading fbs reminders")
    for i in range(df_seller_warehouses.shape[0]):
        for chunk in df_products['chunks'].unique():
            sku_list = df_products.loc[df_products['chunks'] == chunk, 'last_barcode'].to_list()
            sku_list_string = [str(elem) for elem in sku_list]
            params_fbs_stocks = {# "warehouse_id": warehouse_id,
                                "skus": sku_list_string
                                }
            resp_data_fbs_stocks = requests.post(f"https://marketplace-api.wildberries.ru/api/v3/stocks/{df_seller_warehouses['id'][i]}", headers=headers, json=params_fbs_stocks).json()
            tmp_df = pd.DataFrame(resp_data_fbs_stocks['stocks'])
            df_reminders_fbs = pd.concat([df_reminders_fbs, tmp_df])
            df_reminders_fbs['warehouseName'] = df_seller_warehouses['name'][i]
    # Пересоздаем индекс после concat
    df_reminders_fbs = df_reminders_fbs.reset_index(drop=True)
    # Переименовываем колонку с баркодом
    df_reminders_fbs = df_reminders_fbs.rename(columns={'sku': 'last_barcode'})
    # Переводим остатки и баркод в число
    for col in ['amount', 'last_barcode']:
        df_reminders_fbs[col] = df_reminders_fbs[col].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_reminders_fbs[col] = df_reminders_fbs[col].fillna(0)
    # По баркоду получаем имя и размер товара
    df_reminders_fbs = df_reminders_fbs.merge(df_products[['Артикул продавца', 'Размер', 'last_barcode']],
                                     how='left',
                                     on='last_barcode')
    # Сохраняем файл
    df_reminders_fbs.to_csv(uploaddir_today / f"{str(date.today())}_Остатки_fbs.csv", sep=';', index=False)

# Воронка продаж
def getSalesCrater(headers, date_start, date_end):
    page_number = 1
    is_next_page = True
    while is_next_page:
        params = {
                "brandNames": [],
                "objectIDs": [],
                "tagIDs": [],
                "nmIDs": [],
                "timezone": "Europe/Moscow",
                "period": {
                    "begin": "2024-07-30 00:00:00",
                    "end": "2024-07-31 00:00:00"
                },
                "orderBy": {
                    "field": "openCard",
                    "mode": "asc"
                },
                "page": page_number
        }
        resp_data = requests.post("https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail", headers=headers, json=params).json()
        tmp_df = pd.DataFrame(resp_data['data']['cards'])


# Поставки
def get_supply_orders_WB(headers, date_start, date_end):
    params = {
        "dateFrom": date_start,
        # "flag": 0
    }
    logger.info("Uploading supply orders")
    resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/incomes",
                             headers=headers, params=params)
    # Если ошибка, что много запросов, ждем 1 минуту, потом пробуем снова
    if resp_data.status_code == 429:
        logger.info("Waiting 1 minute for next attempt")
        time.sleep(65)
        resp_data = requests.get("https://statistics-api.wildberries.ru/api/v1/supplier/incomes",
                                 headers=headers, params=params)
    if resp_data.status_code == 200:
        result = resp_data.json()
        # Если были поставки
        if len(result) > 0:
            df_supplies = pd.DataFrame(result)
            df_supplies = df_supplies.rename(columns={
                "techSize": "Размер",
                "supplierArticle": "Артикул продавца",
                "nmId": "Артикул WB"
            })
            df_supplies = df_supplies[(df_supplies['date'] >= date_start) & (df_supplies['date'] <= date_end)]
        # Если нет, сохраняем пустой df
        else:
            df_supplies = pd.DataFrame(columns=[
                'Артикул продавца',
                'Размер',
                'supplies'
            ])
        df_supplies.to_csv(uploaddir_today / f"{str(date.today())}_Поставки.csv", sep=';', index=False)
    else:
        logger.error(f"Ошибка получения поставок")

def get_supply_orders_WB_v2(headers, date_start, date_end, status_ids=[3, 4,]):

    base_url = "https://supplies-api.wildberries.ru/api/v1/supplies"
    limit = 1000
    offset = 0
    all_supplies = []

    payload_template = {
        # "dates": [
        #     {
        #         "from": from_date,
        #         "till": till_date,
        #         "type": "factDate"
        #     }
        # ],
        "statusIDs": status_ids,
        "limit": limit,
        "offset": offset
    }

    logger.info(f"Загружаем заявки поставок по статусам {status_ids} и датам {date_start} - {date_end}")

    while True:
        payload = payload_template.copy()
        payload["offset"] = offset
        try:
            response = requests.post(base_url, headers=headers, json=payload)
        except Exception as e:
            logger.error(f"Ошибка соединения: {e}")
            break

        if response.status_code == 200:
            supplies = response.json()
            if not supplies:
                break
            all_supplies.extend(supplies)
            logger.info(f"Получено {len(supplies)} заявок (offset={offset})")
            if len(supplies) < limit:
                break
            offset += limit
            time.sleep(1)  # Лимит: 1000 запросов в минуту — 1 запрос в секунду (на практике обычно достаточно)
        elif response.status_code == 429:
            logger.warning("429 Too Many Requests на /supplies. Ждём 65 секунд...")
            time.sleep(65)
        else:
            logger.error(f"Ошибка получения списка заявок: {response.status_code} {response.text}")
            break

    df_supplies = pd.DataFrame(all_supplies)
    if df_supplies.empty:
        logger.info("Нет заявок с такими статусами и датами.")
        # return pd.DataFrame()

    # --- Получаем склад и товары в поставках ---
    all_items = []
    for idx, row in df_supplies.iterrows():
        supply_id = row["supplyID"]
        status_name = row.get("statusName", "")  # или другое поле со статусом
        logger.info(f"Получаем склад и товары для заявки {supply_id}, статус заявки: {status_name}")
        # --- Получаем склад ---
        supply_url = f"{base_url}/{supply_id}"
        attempts = 0
        sc_name = ""
        while attempts < 3:
            r = requests.get(supply_url, headers=headers)
            if r.status_code == 200:
                sc_name = r.json().get('warehouseName', "")
                break
            elif r.status_code == 429:
                logger.warning(f"429 на склад {supply_id}, ждем 61 сек")
                time.sleep(61)
                attempts += 1
            else:
                logger.error(f"Ошибка {r.status_code} на склад {supply_id}: {r.text}")
                break
        else:
            logger.error(f"Не получилось получить склад для {supply_id}")
            continue
        time.sleep(2.1)

        # --- Получаем товары ---
        goods_url = f"{base_url}/{supply_id}/goods"
        attempts = 0
        while attempts < 3:
            r = requests.get(goods_url, headers=headers)
            if r.status_code == 200:
                try:
                    goods = r.json()
                    if isinstance(goods, list) and goods:
                        temp_df = pd.DataFrame(goods)
                        temp_df["Склад"] = sc_name
                        temp_df["Номер поставки"] = supply_id
                        temp_df["Статус"] = status_name
                        all_items.append(temp_df[[
                            "nmID",
                            "vendorCode",
                            "techSize",
                            "Склад",
                            "quantity",
                            "Номер поставки",
                            "Статус"
                            ]]
                        )
                except Exception as e:
                    logger.error(f"Ошибка парсинга JSON на товарах {supply_id}: {e}")
                break
            elif r.status_code == 429:
                logger.warning(f"429 на товары {supply_id}, ждем 61 сек")
                time.sleep(61)
                attempts += 1
            else:
                logger.error(f"Ошибка {r.status_code} на товары {supply_id}: {r.text}")
                break
        time.sleep(2.1)

    # Конкатенация всех датафреймов по товарам в один
    if all_items:
        df_items = pd.concat(all_items, ignore_index=True)
        df_items = df_items.rename(columns={
            "nmId": "Артикул WB",
            "vendorCode": "Артикул продавца",
            "techSize": "Размер",
            "quantity": "Количество товара"
        })
    else:
        df_items = pd.DataFrame(columns=[
            "Артикул WB",
            "Артикул продавца",
            "Размер",
            "Склад",
            "Количество товара",
            "Номер поставки",
            "Статус"
        ])

    # Сохраняем df с поставками в csv
    df_items.to_csv(
        uploaddir_today / f"{str(date.today())}_Поставки.csv",
        sep=";", index=False, encoding="utf-8-sig"
    )

    return df_items


# Поставки fbs
def get_supply_orders_WB_fbs(headers):
    # Получаем склады продавца
    resp_data_seller_warehouses = requests.get("https://marketplace-api.wildberries.ru/api/v3/warehouses", headers=headers).json()
    df_seller_warehouses = pd.DataFrame(resp_data_seller_warehouses)

    # Получаем список сборочных заданий
    resp_data_supplies_fbs = requests.get("https://marketplace-api.wildberries.ru/api/v3/orders/new", headers=headers).json()
    df_supplies_list_fbs = pd.DataFrame(resp_data_supplies_fbs['orders'])

    # Получаем статусы сборочны заданий
    # Если были поставки
    if df_supplies_list_fbs.shape[0] > 0:
        # Диапазон выгрузок
        l = list(range(0, len(df_supplies_list_fbs), 1000)) + [len(df_supplies_list_fbs)]
        df_supply_statuses = pd.DataFrame()
        for z in zip(l[:-1], l[1:]):
            params_supply_statuses = {"orders": df_supplies_list_fbs.loc[z[0]:z[1], 'id'].to_list()}
            resp_data_supply_statuses = requests.post("https://marketplace-api.wildberries.ru/api/v3/orders/status", headers=headers, json=params_supply_statuses).json()
            tmp_df = pd.DataFrame(resp_data_supply_statuses['orders'])
            df_supply_statuses = pd.concat([tmp_df, df_supply_statuses])

        # Добавляем к списку сборочных заданий их статусы
        df_supplies_list_fbs = df_supplies_list_fbs.merge(df_supply_statuses, how='left', on='id')

        # Убираем list из некоторых колонок
        df_supplies_list_fbs['sku'] = [';'.join(map(str, l)) for l in df_supplies_list_fbs['skus']]
        df_supplies_list_fbs['offices'] = [';'.join(map(str, l)) for l in df_supplies_list_fbs['offices']]
    # Если не было поставок, то сохраняем пустой df
    else:
        df_supplies_list_fbs = pd.DataFrame(columns=['supplies_fbs'])
    # Сохраняем файл
    df_supplies_list_fbs.to_csv(uploaddir_today / f"{str(date.today())}_Поставки_fbs.csv", sep=';', index=False)


# %% Вызов всех функций
if __name__ == '__main__':
    # Создаем директории выгрузки, если их нет
    uploaddir_today.mkdir(parents=True, exist_ok=True)
    # Генерируем даты (по умолчанию - 30 дней от текущей даты)
    date_start, date_end = generate_dates_new()
    # date_end = '2025-07-09T23:59:59'
    # date_start = '2025-06-10T00:00:00'
    logger.info(f"Uploading files for client {client_name} for dates {date_start} - {date_end}")
    save_dates_to_csv(date_start, date_end)
    # Товары, не находящиеся в корзине
    # df_products = getWBProduct(headers, type_products='not_from_recycle')
    # Товары из корзины
    # getWBProduct(headers, type_products='from_recycle')
    # Цены
    # get_prices_WB(headers, df_products)
    df_products = get_wb_product(headers, to_save=True)
    # Заказы
    getOrdersWB(headers, date_start, date_end)
    # Продажи и возвраты
    getSalesWB(headers, date_start, date_end)
    # Остатки
    getStockRemindersWB(headers, date_start, date_end)
    # Остатки fbs
    getStockReminders_fbs(headers)
    # Поставки
    get_supply_orders_WB(headers, date_start, date_end)
    # Поставки fbs
    get_supply_orders_WB_fbs(headers)
    logger.info(f"DONE UPLOADING FILES FOR CLIENT {client_name}")
    logger.info(f"Uploaded files saved in: {uploaddir_today}")

    print(f"\033[45m\033[37mDONE UPLOADING FILES FOR CLIENT {client_name}")

# %%
