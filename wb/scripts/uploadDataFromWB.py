
# %% Определение функций
import argparse
import json
import time
from datetime import date, datetime, timedelta
from colorama import Fore, Back, Style, init

import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.set_option('chained_assignment',None)

# Файл с некоторыми константами
from wb.scripts.constants import client_name, headers, marketplace_dir_path

# Имя директории для текущей выгрузки
uploaddir = marketplace_dir_path / "Clients" / client_name / "UploadFiles"
uploaddir_today = uploaddir / f"UploadFiles_{str(date.today())}"

# Функция парсинга аргумента количества дней выгрузки
from wb.scripts.command_args_parser import cli_days_ago

# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Upload data for WB marketplace",
    usage=(
        "python.exe -m wb.scripts.uploadDataFromWB [days_ago] [--days_ago N] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.uploadDataFromWB 30\n\n"
        "  python.exe -m wb.scripts.uploadDataFromWB --days_ago 10 --client 2\n\n"
        "  python.exe -m wb.scripts.uploadDataFromWB 7 -c 3"
    ),
    formatter_class=ParagraphRichHelpFormatter
)

# Функция генерации дат периода выгрузки данных
def generateDates(days_ago = 30):
    date_end = str(date.today()- timedelta(days=1)) + 'T23:59:59.000Z'
    date_start = str(date.today()- timedelta(days=1) - timedelta(days_ago)) + 'T00:00:00.000Z'
    logger.info(f"Uploading orders from "+ date_start + " to " + date_end)
    return date_end,date_start #,date_end_file,date_start_file

# GPT START ----
# Функция создания диапазона дат с указанной разницей
def generate_dates_new(days_ago=30, start_date_str=None, end_date_str=None):
    """
    Формирует диапазон дат
    Приоритет параметров:
    1. days_ago — если указан, вычисляем диапазон дней (от end_date минус days_ago к end_date), игнорируя start_date_str и end_date_str.
    2. start_date_str и/или end_date_str — используются, если days_ago не задан (оба параметра могут быть заданы частично).
       По умолчанию:
           - end_date_str: сегодняшняя дата, 23:59:59
           - start_date_str: end_date минус 30 дней, 00:00:00
    Возвращает кортеж из строк начальной и конечной дат в формате '%Y-%m-%dT%H:%M:%S'.
    """
    # 1. Приоритет days_ago: вычисляем оконные даты на основе количества дней назад
    if days_ago is not None:
        # Конечная дата: вчера, 23:59:59
        end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0) - timedelta(days=1)
        # Начальная дата: end_date минус days_ago дней + 1, время 00:00:00
        start_date = end_date - timedelta(days=days_ago) + timedelta(days=1)
        start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        # 2. Если хотя бы одна дата задана явно
        # Конечная дата: end_date_str или по умолчанию сегодня, 23:59:59
        if end_date_str:
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
            end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=0)
        else:
            end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
        # Начальная дата: start_date_str или по умолчанию end_date минус 30 дней, время 00:00:00
        if start_date_str:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            start_date = end_date - timedelta(days=30)
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Преобразуем даты в нужный формат для WB API ('%Y-%m-%dT%H:%M:%S', без 'Z')
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
    """
    Выгрузка карточек товаров через API Wildberries с поддержкой пагинации.
    Возвращает результат в виде pandas DataFrame.
    """
    file_name_products = uploaddir_today / f"{str(date.today())}_Товары.csv"
    url = "https://content-api.wildberries.ru/content/v2/get/cards/list"

    # Пустой список для сохранения всех карточек
    all_cards = []

    # Начальные параметры пагинации
    cursor = {
        "limit": 100  # Запрашиваем по 100 карточек за раз
    }

    logger.info("Начало выгрузки карточек товаров")

    while True:
        # Подготовка тела запроса
        payload = {
            "settings": {
                "sort": {
                    "ascending": True  # Сортировка по возрастанию
                },
                "cursor": cursor,
                "filter": {
                    "withPhoto": -1  # Фильтр по карточкам с фото
                }
            }
        }

        logger.info(f"Выполняем запрос к API с параметрами курсора: {cursor}")

        # Выполняем POST запрос к API Wildberries
        response = requests.post(url, json=payload, headers=headers)

        if response.status_code == 200:
            data = response.json()
            logger.info(f"Запрос выполнен успешно. Количество карточек в ответе: {len(data.get('cards', []))}")

            # Добавляем карточки в общий список
            all_cards.extend(data.get('cards', []))

            # Проверяем, достигли ли мы конца данных
            if len(data['cards']) < cursor['limit']:
                logger.info("Все карточки выгружены.")
                break

            # Обновляем курсор для следующего запроса
            cursor.update({
                "updatedAt": data['cursor']['updatedAt'],
                "nmID": data['cursor']['nmID']
            })
        else:
            logger.error(f"Ошибка при выполнении запроса: {response.status_code}, {response.text}")
            raise Exception(f"Ошибка: {response.status_code}, {response.text}")

        # Пауза между запросами, чтобы не перегружать API
        pause_interval = 0.7
        logger.info(f"Пауза {pause_interval} секунды перед следующим запросом")
        time.sleep(pause_interval)

    # Конвертируем данные в DataFrame
    logger.info("Конвертация данных в pandas DataFrame")
    df_products = pd.DataFrame(all_cards)
    logger.info("Выгрузка успешно завершена.")

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

        if to_save:
            # Сохраняем df с товарами через pathlib
            df_products_unpacked_csv.to_csv(file_name_products, sep=';', index=False)
        else:
            return df_products_unpacked


def get_prices_WB(headers: dict, df_products: pd.DataFrame, to_save: bool = True) -> pd.DataFrame:
    """
    Выгрузка цен товаров с Wildberries API с постраничной обработкой данных.

    :param headers: Заголовки для API запросов.
    :param df_products: DataFrame с данными по товарам.
    :param to_save: Если True, данные сохраняются в файл CSV, иначе возвращаются DataFrame.
    :return: DataFrame с объединенной информацией о товарах и ценах (если to_save=False).
    """
    logger.info("Начало выгрузки цен товаров с Wildberries API.")

    df_prices = pd.DataFrame()
    try:
        # Создаем диапазоны индексов для пагинации
        index_ranges = list(range(0, len(df_products), 1000)) + [len(df_products)]
        logger.info(f"Всего диапазонов для обработки: {len(index_ranges) - 1}")

        # Обрабатываем каждый диапазон в цикле
        for start, end in zip(index_ranges[:-1], index_ranges[1:]):
            params = {
                "limit": 1000,
                "offset": start,
            }
            logger.info(f"Запрос цен товаров с offset={start}, limit=1000")

            # Выполняем запрос
            response = requests.get(
                "https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter",
                headers=headers,
                params=params,
            )

            if response.status_code == 200:
                logger.info(f"Данные успешно получены для диапазона {start}-{end}")
                resp_data = response.json()

                # Создаем временный DataFrame
                tmp_df = pd.DataFrame(resp_data.get('data', {}).get('listGoods', []))
                df_prices = pd.concat([df_prices, tmp_df], ignore_index=True)
            else:
                logger.error(f"Ошибка при запросе: {response.status_code} - {response.text}")
                break  # Прерываем цикл при ошибке

            time.sleep(5)  # Делаем паузу между запросами

        # Проверка, если нет данных
        if df_prices.empty:
            logger.warning("Цены не выгружены. Результат пустой.")
            return None

        # Распаковываем данные по колонке sizes
        logger.info("Раскрытие данных по размерам товаров.")
        df_prices = df_prices.explode('sizes')

        # Убираем строки с пустыми размерами
        df_prices.dropna(subset=['sizes'], inplace=True)

        # Извлекаем данные о ценах
        price_data = pd.json_normalize(df_prices['sizes'])
        df_prices = pd.concat([df_prices.drop(columns=['sizes']), price_data], axis=1)

        df_prices.rename(
            columns={
                'sizeID': 'chrtID',
                'price': 'price',
                'discountedPrice': 'discount_price',
                'clubDiscountedPrice': 'club_discount_price',
            },
            inplace=True,
        )

        # Объединяем выгрузку с основным DataFrame
        logger.info("Объединение данных о товарах и ценах.")
        df_products_prices = df_products.merge(
            df_prices[['chrtID', 'price', 'discount_price', 'club_discount_price']],
            on='chrtID',
            how='left'
        )

        # Удаляем дубликаты и заполняем пропущенные данные
        df_products_prices.drop_duplicates(subset=['nmID', 'chrtID'], keep='first', inplace=True)
        df_products_prices['Размер'] = df_products_prices['Размер'].fillna(0)

        # Сохраняем или возвращаем результат
        if to_save:
            products_prices_file = uploaddir_today / f"{str(date.today())}_Товары.csv"
            logger.info(f"Сохраняем результаты в файл: {products_prices_file}")
            df_products_prices.to_csv(products_prices_file, sep=';', index=False)
        else:
            logger.info("Возврат результирующего DataFrame.")
            return df_products_prices

    except Exception as e:
        logger.error(f"Произошла ошибка при выгрузке цен: {e}")
        return None


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

def get_supply_orders_WB_v2(
    headers,
    status_ids=[2, 3, 4],
    filter_by='date',  # 'date' (по умолчанию) или 'status'
):
    """
    Получить поставки Wildberries с фильтрацией по дате или статусу.

    :param headers: dict — заголовки для API Wildberries
    :param date_start: str|None — дата начала фильтра (ГГГГ-ММ-ДД)
    :param date_end: str|None — дата конца фильтра (ГГГГ-ММ-ДД)
    :param status_ids: list — список ID статусов для фильтрации
    :param filter_by: str — 'date' или 'status', определяет тип фильтра
    :param uploaddir_today: str — каталог для сохранения файла
    :return: pd.DataFrame — данные по товарам в поставках
    """
    base_url = "https://supplies-api.wildberries.ru/api/v1/supplies"
    limit = 1000
    offset = 0
    all_supplies = []

    # df для сохранения, если не было заявок на поставку
    df_supplies_template = pd.DataFrame(columns=[
        "Артикул WB", "Артикул продавца", "Размер", "Склад",
        "Количество товара", "Номер поставки", "Статус"
    ])
    output_path = uploaddir_today / f"{str(date.today())}_Поставки.csv"


    # --- Установка дат по умолчанию ---
    from_date = str(date.today())
    logger.info(f"Фильтрация заявок по: {filter_by}")

    # --- Формирование payload в зависимости от фильтрации ---
    if filter_by == 'date':
        payload_template = {
            "dates": [{
                "from": from_date,
                # "till": till_date,
                "type": "supplyDate"
            }],
            "limit": limit,
            "offset": offset
        }
    elif filter_by == 'status':
        payload_template = {
            "statusIDs": status_ids,
            "limit": limit,
            "offset": offset
        }
    else:
        logger.error("Некорректное значение filter_by! Используйте 'date' или 'status'")
        return pd.DataFrame()

    logger.info(
        f"Начинаем загрузку заявок: фильтр по {filter_by}."
        # f"Статусы: {status_ids}"
    )

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
            time.sleep(1)
        elif response.status_code == 429:
            logger.warning("429 Too Many Requests на /supplies. Ждём 65 секунд...")
            time.sleep(65)
        else:
            logger.error(f"Ошибка получения списка заявок: {response.status_code} {response.text}")
            break

    df_supplies = pd.DataFrame(all_supplies)

    # Если нашлись заявки на поставку, начинаем выгрузку инфы по ним
    if not df_supplies.empty:
        # --- Получаем сведения о складах и товарах в заявках ---
        all_items = []
        for idx, row in df_supplies.iterrows():
            supply_id = row["supplyID"]
            status_name = row.get("statusName", "")
            logger.info(f"Выгрузка данных по заявке {supply_id}, статус заявки: {status_name}")

            # --- Получаем название склада по supplyID ---
            supply_url = f"{base_url}/{supply_id}"
            attempts = 0
            sc_name = ""
            while attempts < 3:
                r = requests.get(supply_url, headers=headers)
                if r.status_code == 200:
                    sc_name = r.json().get('warehouseName', "")
                    break
                elif r.status_code == 429:
                    logger.warning(f"429 на склад поставки {supply_id}, ждем 61 сек")
                    time.sleep(61)
                    attempts += 1
                else:
                    logger.error(f"Ошибка {r.status_code} на склад поставки {supply_id}: {r.text}")
                    break
            else:
                logger.error(f"Не получилось получить склад для {supply_id}")
                continue
            time.sleep(2.1)

            # --- Получаем товары по заявке ---
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
                                "nmID", "vendorCode", "techSize", "Склад",
                                "quantity", "Номер поставки", "Статус"
                            ]])
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

        # --- Объединение результатов в единый датафрейм ---
        if all_items:
            df_items = pd.concat(all_items, ignore_index=True)
            df_items = df_items.rename(columns={
                "nmID": "Артикул WB",
                "vendorCode": "Артикул продавца",
                "techSize": "Размер",
                "quantity": "Количество товара"
            })
        else:
            df_items = df_supplies_template.copy()

        df_items.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")
    # Если заявки на поставку не нашлись, сохраняем пустой df
    else:
        logger.info("Нет заявок по заданным параметрам.")
        df_supplies_template.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")


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

    # Парсим аргументы командной строки
    days_ago, cli_start_date, cli_end_date = cli_days_ago(parser)

    # Создаем директории выгрузки, если их нет
    uploaddir_today.mkdir(parents=True, exist_ok=True)

    # Генерируем даты (по умолчанию - 30 дней от текущей даты)
    date_start, date_end = generate_dates_new(
        days_ago=days_ago,
        start_date_str=cli_start_date,
        end_date_str=cli_end_date,
        # start_date_str = '2025-07-12T23:59:59'
        # end_date_str = '2025-06-13T00:00:00'
    )
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
    get_supply_orders_WB_v2(headers)
    # Поставки fbs
    get_supply_orders_WB_fbs(headers)
    logger.info(f"DONE UPLOADING FILES FOR CLIENT {client_name}")
    logger.info(f"Uploaded files saved in: {uploaddir_today}")

    print(f"{Back.MAGENTA}{Fore.WHITE}DONE UPLOADING FILES FOR CLIENT {client_name}{Style.RESET_ALL}")
# %%
