
# %% Определение функций
import argparse
import json
import time
from datetime import date, datetime, timedelta
from colorama import Fore, Back, Style, init

import pandas as pd
import numpy as np
import requests
from requests import RequestException
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.set_option('chained_assignment',None)

# Файл с некоторыми константами
from wb.scripts.constants import client_name, headers, session, marketplace_dir_path
from wb.scripts.wb_generic_functions import WbRateLimiter

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

        # --- ДОБАВЛЕНО: распаковка габаритов из dimensions ---
        # dimensions – это словарь с полями length, width, height, weightBrutto, isValid
        # Сохраняем их как отдельные колонки
        df_products = df_products.assign(
            **{
                "Длина, см": df_products["dimensions"].apply(
                    lambda d: d.get("length") if isinstance(d, dict) else None
                ),
                "Ширина, см": df_products["dimensions"].apply(
                    lambda d: d.get("width") if isinstance(d, dict) else None
                ),
                "Высота, см": df_products["dimensions"].apply(
                    lambda d: d.get("height") if isinstance(d, dict) else None
                ),
                "Вес, кг": df_products["dimensions"].apply(
                    lambda d: d.get("weightBrutto") if isinstance(d, dict) else None
                ),
                "Корректность габаритов": df_products["dimensions"].apply(
                    lambda d: d.get("isValid") if isinstance(d, dict) else None
                ),
            }
        )
        # -----------------------------------------------

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

        # --- ДОБАВЛЕНО: включаем габариты в CSV-выгрузку ---
        df_products_unpacked_csv = df_products_unpacked[
            [
                "nmID",
                "imtID",
                "chrtID",
                "Артикул продавца",
                "barcode",
                "last_barcode",
                "Наименование товара",
                "Предмет",
                "Размер",
                "Цвет",
                "Длина, см",
                "Ширина, см",
                "Высота, см",
                "Вес, кг",
                "Корректность габаритов",
            ]
        ]
        # -------------------------------------------------

        if to_save:
            file_name_products = uploaddir_today / f"{str(date.today())}_Товары.csv"
            # Сохраняем df с товарами через pathlib
            df_products_unpacked_csv.to_csv(file_name_products, sep=';', encoding='utf-8-sig', index=False)
        logger.success(f"Выгрузка карточек товаров успешно завершена, всего товаров: {len(df_products_unpacked)}.")
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
        df_prices = df_prices.dropna(subset=['sizes'], ignore_index=True)

        # Извлекаем данные о ценах
        price_data = pd.json_normalize(df_prices['sizes'])
        df_prices = pd.concat([df_prices.drop(columns=['sizes']), price_data], axis=1)

        df_prices = df_prices.rename(
            columns={
                'sizeID': 'chrtID',
                'price': 'price',
                'discountedPrice': 'discount_price',
                'clubDiscountedPrice': 'club_discount_price',
            },
        )

        # Объединяем выгрузку с основным DataFrame
        logger.info("Объединение данных о товарах и ценах.")
        df_products_prices = df_products.merge(
            df_prices[['chrtID', 'price', 'discount_price', 'club_discount_price']],
            on='chrtID',
            how='left'
        )

        # Удаляем дубликаты и заполняем пропущенные данные
        df_products_prices = df_products_prices.drop_duplicates(subset=['nmID', 'chrtID'], keep='first', ignore_index=True)
        df_products_prices['Размер'] = df_products_prices['Размер'].fillna(0)

        # Сохраняем или возвращаем результат
        if to_save:
            products_prices_file = uploaddir_today / f"{str(date.today())}_Товары.csv"
            logger.info(f"Сохраняем результаты в файл: {products_prices_file}")
            df_products_prices.to_csv(products_prices_file, sep=';', encoding='utf-8-sig', index=False)

        logger.success("Выгрузка цен товаров успешно завершена.")
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
        df_products_prices.to_csv(products_file, sep=';', encoding='utf-8-sig', index=False)

    return df_products_prices


# Заказы
def getOrdersWB(
    headers: dict,
    date_start: str,
    date_end: str,
    to_save: bool = True,
    max_retries: int = 2,
    retry_delay: int = 65,
):
    """
    Загрузка заказов Wildberries за указанный период.

    :param headers: Заголовки запроса (должен содержать токен авторизации).
    :param date_start: Начальная дата периода (строка, как в API WB).
    :param date_end: Конечная дата периода (строка, как в API WB).
    :param to_save: Сохранять ли результат в CSV.
    :param max_retries: Кол-во попыток при ошибке 429 или сетевых ошибках.
    :param retry_delay: Пауза между попытками (сек).
    :return: DataFrame с заказами.
    :raises: RequestException, ValueError, RuntimeError при неуспехе.
    """
    orders_endpoint = "https://statistics-api.wildberries.ru/api/v1/supplier/orders"
    params = {
        "dateFrom": date_start,
        "dateTo": date_end,
    }

    logger.info(
        "Загрузка заказов WB с {} по {} (max_retries={}, retry_delay={})".format(
            date_start, date_end, max_retries, retry_delay
        )
    )

    attempt = 0
    while attempt < max_retries:
        attempt += 1
        try:
            logger.debug("Отправка запроса к WB (попытка {} из {})", attempt, max_retries)
            response = requests.get(
                orders_endpoint,
                headers=headers,
                params=params,
                timeout=30,
            )
        except RequestException as exc:
            logger.error(
                "Сетевая ошибка при запросе заказов WB (попытка {} из {}): {}",
                attempt, max_retries, exc
            )
            if attempt >= max_retries:
                raise
            logger.info("Ожидание {} секунд перед повторной попыткой...", retry_delay)
            time.sleep(retry_delay)
            continue

        # Обработка статусов
        if response.status_code == 200:
            logger.debug("Успешный ответ WB (статус 200)")
            break

        if response.status_code == 429 and attempt < max_retries:
            logger.warning(
                "Получен статус 429 Too Many Requests (попытка {} из {}). "
                "Ожидание {} секунд перед повтором...",
                attempt, max_retries, retry_delay
            )
            time.sleep(retry_delay)
            continue

        # Любой другой неожиданный статус
        logger.error(
            "Неожиданный статус ответа WB: {}. Тело ответа: {}",
            response.status_code,
            response.text[:1000],
        )
        response.raise_for_status()

    # Если после цикла статус всё ещё не 200 — это ошибка логики
    if response.status_code != 200:
        raise RuntimeError(
            f"Не удалось получить заказы WB после {max_retries} попыток. "
            f"Последний статус: {response.status_code}"
        )

    # Разбор JSON
    try:
        data = response.json()
    except ValueError as exc:
        logger.error(
            "Ошибка парсинга JSON от WB: {}. Тело ответа: {}",
            exc,
            response.text[:1000],
        )
        raise

    # Формирование DataFrame
    df_orders = pd.DataFrame(data)
    if df_orders.empty:
        logger.warning("WB вернул пустой список заказов за указанный период.")
    else:
        # Фильтрация по диапазону дат (если поле 'date' есть)
        if "date" in df_orders.columns:
            df_orders = df_orders[
                (df_orders["date"] >= date_start) &
                (df_orders["date"] <= date_end)
            ]
        else:
            logger.warning(
                "В ответе WB отсутствует колонка 'date'. Фильтрация по датам не выполнена."
            )

        # Переименование колонок
        rename_map = {
            "techSize": "Размер",
            "supplierArticle": "Артикул продавца",
            "nmId": "Артикул WB",
        }
        df_orders = df_orders.rename(columns=rename_map)

    # Сохранение в файл
    if to_save:
        orders_file = uploaddir_today / f"{date.today()}_Заказы.csv"
        try:
            df_orders.to_csv(orders_file, sep=";", encoding="utf-8-sig", index=False)
            logger.info("Файл с заказами сохранён: {}", orders_file)
        except Exception as exc:
            logger.error("Не удалось сохранить файл с заказами {}: {}", orders_file, exc)
            raise

    return df_orders

def getSalesWB(
    headers: dict,
    date_start: str,
    date_end: str,
    to_save: bool = True,
    max_retries: int = 2,
    retry_delay: int = 65,
    min_days: int | None = None,
):
    """
    Загрузка продаж Wildberries за указанный период.

    :param headers: Заголовки запроса (должен содержать токен авторизации).
    :param date_start: Начальная дата периода (строка, как в API WB).
    :param date_end: Конечная дата периода (строка, как в API WB).
    :param to_save: Сохранять ли результат в CSV.
    :param max_retries: Кол-во попыток при ошибке 429 или сетевых ошибках.
    :param retry_delay: Пауза между попытками (сек).
    :param min_days: минимальная длина периода в днях (включительно).
                     Если задано и исходный период КОРОЧЕ, то date_start
                     сдвигается назад так, чтобы длина была ровно min_days.
                     Если исходный период уже длиннее или равен min_days – не меняем.
    :return: DataFrame с продажами.
    :raises: RequestException, ValueError, RuntimeError при неуспехе.
    """
    # --- 0. Корректировка периода, если задан min_days ---
    if min_days is not None:
        # Берём только даты, без времени (WB обычно принимает YYYY-MM-DD)
        end_date = pd.to_datetime(date_end).date()
        start_date = pd.to_datetime(date_start).date()
        orig_days = (end_date - start_date).days  # длина исходного периода

        if orig_days < min_days:
            # Сдвигаем начало так, чтобы длина была min_days
            new_start_date = end_date - pd.Timedelta(days=min_days)
            logger.info(
                "Период {}–{} короче min_days={} ({} дней). "
                "Смещаем date_start на {}.",
                start_date, end_date, min_days, orig_days, new_start_date,
            )
            # Собираем обратно строки в формате YYYY-MM-DD
            date_start = new_start_date.strftime("%Y-%m-%d")
            date_end = end_date.strftime("%Y-%m-%d")
        else:
            logger.info(
                "Исходный период {}–{} ({} дней) не короче min_days={}, "
                "оставляем как есть.",
                start_date, end_date, orig_days, min_days,
            )

    sales_endpoint = "https://statistics-api.wildberries.ru/api/v1/supplier/sales"
    params = {
        "dateFrom": date_start,
        "dateTo": date_end,
    }
    logger.info(
        "Загрузка продаж WB с {} по {} (max_retries={}, retry_delay={})",
        date_start, date_end, max_retries, retry_delay
    )

    attempt = 0
    response = None
    while attempt < max_retries:
        attempt += 1
        try:
            logger.debug(
                "Отправка запроса к WB (продажи), попытка {} из {}",
                attempt, max_retries
            )
            response = requests.get(
                sales_endpoint,
                headers=headers,
                params=params,
                timeout=30,
            )
        except RequestException as exc:
            logger.error(
                "Сетевая ошибка при запросе продаж WB (попытка {} из {}): {}",
                attempt, max_retries, exc
            )
            if attempt >= max_retries:
                raise
            logger.info(
                "Ожидание {} секунд перед повторной попыткой...",
                retry_delay
            )
            time.sleep(retry_delay)
            continue

        # Обработка статусов
        if response.status_code == 200:
            logger.debug("Успешный ответ WB по продажам (статус 200)")
            break

        if response.status_code == 429 and attempt < max_retries:
            logger.warning(
                "Получен статус 429 Too Many Requests при запросе продаж "
                "(попытка {} из {}). Ожидание {} секунд перед повтором...",
                attempt, max_retries, retry_delay
            )
            time.sleep(retry_delay)
            continue

        logger.error(
            "Неожиданный статус ответа WB по продажам: {}. Тело ответа: {}",
            response.status_code,
            response.text[:1000],
        )
        response.raise_for_status()

    if response is None or response.status_code != 200:
        raise RuntimeError(
            f"Не удалось получить продажи WB после {max_retries} попыток. "
            f"Последний статус: {getattr(response, 'status_code', 'нет ответа')}"
        )

    # Разбор JSON
    try:
        data = response.json()
    except ValueError as exc:
        logger.error(
            "Ошибка парсинга JSON от WB (продажи): {}. Тело ответа: {}",
            exc,
            response.text[:1000],
        )
        raise

    # Формирование DataFrame
    df_sales = pd.DataFrame(data)

    if df_sales.empty:
        logger.warning("WB вернул пустой список продаж за указанный период.")
    else:
        # Фильтрация по датам, если колонка есть
        if "date" in df_sales.columns:
            df_sales = df_sales[
                (df_sales["date"] >= date_start) &
                (df_sales["date"] <= date_end)
            ]
        else:
            logger.warning(
                "В ответе WB по продажам отсутствует колонка 'date'. "
                "Фильтрация по датам не выполнена."
            )

        rename_map = {
            "techSize": "Размер",
            "supplierArticle": "Артикул продавца",
            "nmId": "Артикул WB",
        }
        df_sales = df_sales.rename(columns=rename_map)

    # Сохранение в файл
    if to_save:
        sales_file = uploaddir_today / f"{date.today()}_Продажи.csv"
        try:
            df_sales.to_csv(
                sales_file,
                sep=";",
                encoding="utf-8-sig",
                index=False,
            )
            logger.info("Файл с продажами сохранён: {}", sales_file)
        except Exception as exc:
            logger.error(
                "Не удалось сохранить файл с продажами {}: {}",
                sales_file,
                exc,
            )
            raise

    return df_sales

# Остатки FBO
def getStockRemindersWB(
    headers: dict,
    date_start: str,  # для совместимости, не используется
    date_end: str,    # для совместимости, не используется
    to_save: bool = True,
    max_retries: int = 3,
    limit: int = 250000,
):
    """
    Загрузка остатков Wildberries через POST с пагинацией по offset, начиная с 0.
    nmIds/chrtIds не передаём — получаем остатки по всем товарам.
    """

    stock_rem_endpoint = "https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses"

    logger.info(
        f"Загрузка остатков WB (POST, пагинация от offset=0), limit={limit}, max_retries={max_retries}"
    )

    session = requests.Session()
    session.headers.update(headers)

    # Лимиты: 1 мин / 3 запроса / 20 сек / burst=1
    limiter = WbRateLimiter(
        period_seconds=60,
        period_limit=3,
        max_burst=1,
    )

    all_items = []

    # offset задаём явно с 0 и инкрементируем каждым шагом
    offset = 0
    page_idx = 1

    while True:
        payload = {
            "nmIds": [],
            "chrtIds": [],
            "limit": limit,
            "offset": offset,
        }

        logger.info(
            f"Запрос остатков WB: страница {page_idx}, offset={offset}, limit={limit}"
        )

        resp, ok = limiter.call(
            session.post,
            url=stock_rem_endpoint,
            json=payload,
            max_retries=max_retries,
        )
        if not ok or resp is None:
            logger.error(
                f"Не удалось получить остатки WB на странице {page_idx} "
                f"(offset={offset}) после {max_retries} попыток. Прерываем пагинацию."
            )
            break

        if resp.status_code != 200:
            try:
                err_json = resp.json()
            except Exception:
                err_json = resp.text
            logger.error(
                f"Неожиданный статус ответа WB по остаткам (POST): {resp.status_code}. "
                f"Тело ответа: {err_json}"
            )
            break

        try:
            data = resp.json()
        except Exception as exc:
            logger.error(
                f"Ошибка парсинга JSON от WB (остатки, POST), страница {page_idx}, offset={offset}: {exc}"
            )
            break

        # Формат {"data": {"items": [...]}}
        if (
            isinstance(data, dict)
            and "data" in data
            and isinstance(data["data"], dict)
            and "items" in data["data"]
        ):
            items = data["data"]["items"] or []
        else:
            items = data if isinstance(data, list) else []

        logger.info(
            f"Страница {page_idx}: получено {len(items)} строк остатков."
        )

        if not items:
            # Пустая страница — дальше данных нет
            logger.info(
                f"Пустой результат на странице {page_idx} (offset={offset}). "
                f"Пагинация завершена."
            )
            break

        all_items.extend(items)

        # Если вернулось меньше, чем limit — это последняя страница
        if len(items) < limit:
            logger.info(
                f"Страница {page_idx}: получено {len(items)} < limit={limit}. "
                f"Пагинация завершена."
            )
            break

        # Следующая страница: смещаем offset ровно на limit
        offset += limit
        page_idx += 1

    if not all_items:
        logger.warning("WB вернул пустой список остатков (по всем страницам).")
        df_reminders = pd.DataFrame()
    else:
        df_reminders = pd.DataFrame(all_items)
        rename_map = {
            "techSize": "Размер",
            "supplierArticle": "Артикул продавца",
            "nmId": "Артикул WB",
        }
        df_reminders = df_reminders.rename(columns=rename_map)

    if to_save:
        reminders_file = uploaddir_today / f"{date.today()}_Остатки.csv"
        try:
            df_reminders.to_csv(reminders_file, sep=";", encoding="utf-8-sig", index=False)
            logger.info(
                f"Файл с остатками сохранён: {reminders_file}. "
                f"Всего строк: {df_reminders.shape[0]}"
            )
        except Exception as exc:
            logger.error(f"Не удалось сохранить файл с остатками {reminders_file}: {exc}")
            raise

    return df_reminders

# Остатки FBS
def getStockReminders_fbs(
    headers: dict,
    to_save: bool = True,
    chunk_size: int = 1000,
    max_retries: int = 3,
    retry_delay: int = 65,
    df_products: pd.DataFrame | None = None,
):
    """
    Остатки FBS по складам WB.

    ОСОБЕННОСТИ:
    1. В запрос передаётся список chrtID (а не баркодов).
    2. Объединение остатков с товарами происходит по колонке `chrtId`
       (внутри функции все варианты chrtID/chrtId/chrtid приводятся к `chrtId`).

    Параметры
    ---------
    headers : dict
        Заголовки для запросов к API WB.
    to_save : bool, по умолчанию True
        Сохранять ли результат в CSV.
    chunk_size : int
        Размер чанка chrtId для одного запроса.
    max_retries : int
        Максимальное число повторов запроса при ошибках/429.
    retry_delay : int
        Пауза между повторами.
    df_products : pandas.DataFrame | None
        Если передан, используется как источник товаров и chrtId.
        Ожидается колонка 'chrtId' или 'chrtID'.
        Если None — список товаров читается из файла
        `{uploaddir_today}/{date.today()}_Товары.csv`.
    """
    def _request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
        """Вспомогательная функция для запросов с ретраями."""
        attempt = 0
        while attempt < max_retries:
            attempt += 1
            try:
                logger.debug(
                    "HTTP {} {} (попытка {}/{})",
                    method.upper(),
                    url,
                    attempt,
                    max_retries,
                )
                resp = requests.request(method=method, url=url, timeout=30, **kwargs)
            except RequestException as exc:
                logger.error(
                    "Сетевая ошибка при запросе {} {} (попытка {}/{}) : {}",
                    method.upper(),
                    url,
                    attempt,
                    max_retries,
                    exc,
                )
                if attempt >= max_retries:
                    raise
                logger.info(
                    "Ожидание {} секунд перед повторной попыткой...",
                    retry_delay,
                )
                time.sleep(retry_delay)
                continue
            if resp.status_code == 200:
                return resp
            if resp.status_code == 429 and attempt < max_retries:
                logger.warning(
                    "Получен статус 429 Too Many Requests при запросе {} {} "
                    "(попытка {}/{}). Ожидание {} секунд перед повтором...",
                    method.upper(),
                    url,
                    attempt,
                    max_retries,
                    retry_delay,
                )
                time.sleep(retry_delay)
                continue
            logger.error(
                "Неожиданный статус ответа {} {}: {}. Тело ответа: {}",
                method.upper(),
                url,
                resp.status_code,
                resp.text[:1000],
            )
            resp.raise_for_status()
        raise RuntimeError(
            "Не удалось выполнить запрос {} {} после {} попыток.".format(
                method.upper(), url, max_retries
            )
        )

    # URL для запросов
    WB_WAREHOUSES_URL = "https://marketplace-api.wildberries.ru/api/v3/warehouses"
    WB_STOCKS_FBS_URL = "https://marketplace-api.wildberries.ru/api/v3/stocks"

    # 1. Получаем список товаров
    if df_products is None:
        products_file = uploaddir_today / f"{date.today()}_Товары.csv"
        logger.info("Загрузка списка товаров из файла: {}", products_file)
        try:
            df_products = pd.read_csv(products_file, sep=";")
        except Exception as exc:
            logger.error("Не удалось прочитать файл товаров {}: {}", products_file, exc)
            raise
    else:
        logger.info(
            "Используется переданный df_products для запроса остатков FBS. Строк: {}",
            len(df_products),
        )

    # Приводим chrtID/chrtId/chrtid к единому названию chrtId
    rename_wb_ids = {
        "chrtID": "chrtId",
        "chrtid": "chrtId",
    }
    df_products = df_products.rename(
        columns={c: rename_wb_ids[c] for c in df_products.columns if c in rename_wb_ids}
    )

    if "chrtId" not in df_products.columns:
        raise ValueError("В df_products отсутствует колонка 'chrtId' (или 'chrtID').")

    df_products = df_products[df_products["chrtId"].notna()].copy()
    if df_products.empty:
        logger.warning(
            "Список товаров пуст или не содержит chrtId. Возвращаю пустой DataFrame."
        )
        return pd.DataFrame(columns=["chrtId", "amount", "warehouseId", "warehouseName"])

    df_products["chrtId"] = df_products["chrtId"].astype(str)
    df_products = df_products.reset_index(drop=True)
    df_products["chunk"] = df_products.index // chunk_size

    logger.info(
        "Всего товаров: {}, размер чанка: {}, всего чанков: {}",
        len(df_products),
        chunk_size,
        df_products["chunk"].nunique(),
    )

    # 2. Получаем склады продавца
    logger.info("Загрузка складов продавца (FBS)...")
    resp_warehouses = _request_with_retry("get", WB_WAREHOUSES_URL, headers=headers)
    try:
        warehouses_data = resp_warehouses.json()
    except ValueError as exc:
        logger.error(
            "Ошибка парсинга JSON по складам FBS: {}. Ответ: {}",
            exc,
            resp_warehouses.text[:1000],
        )
        raise

    df_seller_warehouses = pd.DataFrame(warehouses_data)
    if df_seller_warehouses.empty:
        logger.warning("Список складов продавца пуст. Возвращаю пустой DataFrame.")
        return pd.DataFrame(columns=["chrtId", "amount", "warehouseId", "warehouseName"])

    if "id" not in df_seller_warehouses.columns or "name" not in df_seller_warehouses.columns:
        raise ValueError(
            "В ответе по складам нет ожидаемых колонок 'id' и 'name'. "
            "Колонки: {}".format(list(df_seller_warehouses.columns))
        )

    logger.info(
        "Найдено складов FBS: {} ({})",
        df_seller_warehouses.shape[0],
        ", ".join(map(str, df_seller_warehouses["name"].tolist())),
    )

    # 3. Запрашиваем остатки по всем складам и чанкам chrtId
    logger.info("Загрузка FBS-остатков по chrtId...")
    all_stocks = []
    for _, wh_row in df_seller_warehouses.iterrows():
        warehouse_id = wh_row["id"]
        warehouse_name = wh_row["name"]
        logger.info("Обработка склада '{}' (id={})...", warehouse_name, warehouse_id)

        for chunk_id, chunk_df in df_products.groupby("chunk"):
            chrt_list = chunk_df["chrtId"].tolist()
            payload = {"skus": chrt_list}
            url = f"{WB_STOCKS_FBS_URL}/{warehouse_id}"

            logger.debug(
                "Запрос остатков FBS: склад={}, chunk={}, кол-во chrtId={}",
                warehouse_name,
                chunk_id,
                len(chrt_list),
            )

            resp_stocks = _request_with_retry("post", url, headers=headers, json=payload)
            try:
                stocks_data = resp_stocks.json()
            except ValueError as exc:
                logger.error(
                    "Ошибка парсинга JSON по остаткам FBS (склад {}): {}. Ответ: {}",
                    warehouse_name,
                    exc,
                    resp_stocks.text[:1000],
                )
                raise

            stocks_list = stocks_data.get("stocks", [])
            if not stocks_list:
                logger.debug(
                    "Пустые остатки для склада {}, chunk={}",
                    warehouse_name,
                    chunk_id,
                )
                continue

            tmp_df = pd.DataFrame(stocks_list)
            if tmp_df.empty:
                continue

            tmp_df["warehouseId"] = warehouse_id
            tmp_df["warehouseName"] = warehouse_name
            all_stocks.append(tmp_df)

    if not all_stocks:
        logger.warning("WB вернул пустые остатки FBS по всем складам.")
        df_reminders_fbs = pd.DataFrame(columns=["chrtId", "amount", "warehouseId", "warehouseName"])
    else:
        df_reminders_fbs = pd.concat(all_stocks, ignore_index=True)

    # 4. Приведение типов и объединение с товарами по chrtId
    if "sku" not in df_reminders_fbs.columns:
        logger.warning(
            "В ответе по остаткам FBS нет колонки 'sku'. Колонки: {}",
            list(df_reminders_fbs.columns),
        )
    else:
        df_reminders_fbs = df_reminders_fbs.rename(columns={"sku": "chrtId"})

    # Приводим chrtId и amount к числам/строкам
    if "chrtId" in df_reminders_fbs.columns:
        df_reminders_fbs["chrtId"] = df_reminders_fbs["chrtId"].astype(str)

    if "amount" in df_reminders_fbs.columns:
        df_reminders_fbs["amount"] = pd.to_numeric(
            df_reminders_fbs["amount"],
            errors="coerce",
        ).fillna(0)

    if "chrtId" in df_reminders_fbs.columns:
        df_reminders_fbs = df_reminders_fbs.merge(
            df_products[["Артикул продавца", "Размер", "chrtId"]],
            how="left",
            on="chrtId",
        )
    else:
        logger.warning(
            "В остатках отсутствует колонка 'chrtId', объединение с товарами не выполнено."
        )

    # 5. Сохранение
    reminders_fbs_file = uploaddir_today / f"{date.today()}_Остатки_fbs.csv"
    if to_save:
        try:
            df_reminders_fbs.to_csv(
                reminders_fbs_file,
                sep=";",
                encoding="utf-8-sig",
                index=False,
            )
            logger.info("Файл с остатками FBS сохранён: {}", reminders_fbs_file)
        except Exception as exc:
            logger.error(
                "Не удалось сохранить файл с остатками FBS {}: {}",
                reminders_fbs_file,
                exc,
            )
            raise

    return df_reminders_fbs

# 1. Чистая выгрузка данных из API
def get_warehouses_load_raw(
    headers=None,
    warehouse_ids=None,
    base_url: str = "https://common-api.wildberries.ru/api/tariffs/v1/acceptance/coefficients",
    max_retries: int = 3,
) -> pd.DataFrame:
    """
    Получает СЫРЫЕ данные загруженности складов WB из API.
    Никакой обработки, только запрос и DataFrame по ответу.
    """

    if headers is not None:
        session.headers.update(headers)

    # Лимиты: 6 запросов / 60 сек, интервал ~10 сек, всплеск 6
    limiter = WbRateLimiter(
        period_seconds=60,
        period_limit=6,
        max_burst=6,
    )

    params = {}
    if warehouse_ids:
        ids_str = ",".join(str(wid) for wid in warehouse_ids)
        params["warehouseIDs"] = ids_str
        logger.info(f"Запрос загруженности по складам: {ids_str}")
    else:
        logger.info("Запрос загруженности по всем складам (без warehouseIDs).")

    resp, ok = limiter.call(
        session.get,
        url=base_url,
        params=params,
        max_retries=max_retries,
    )
    if not ok or resp is None:
        logger.error("Не удалось получить загруженность складов WB (ошибка соединения/лимита).")
        return pd.DataFrame()

    if resp.status_code != 200:
        logger.error(
            f"Ошибка ответа при получении загруженности складов: "
            f"HTTP {resp.status_code} | {resp.text}"
        )
        return pd.DataFrame()

    try:
        data = resp.json()
    except Exception as e:
        logger.error(f"Ошибка парсинга JSON загруженности складов: {e}")
        return pd.DataFrame()

    if not isinstance(data, list) or not data:
        logger.warning("Ответ по загруженности складов пуст или не является списком.")
        return pd.DataFrame()

    df = pd.DataFrame(data)
    logger.info(f"Сырые данные по загруженности: {len(df)} записей.")
    return df

# 2. Обработка/расшифровка полей + столбец boxAcceptanceType
def process_warehouses_load(df: pd.DataFrame) -> pd.DataFrame:
    """
    Обрабатывает сырой DataFrame загруженности складов:
      - приводит типы;
      - добавляет расшифровки coefficient, boxTypeID, allowUnload, isSortingCenter;
      - добавляет столбец boxAcceptanceType (какие типы поставок принимает склад).
    """

    if df.empty:
        return df

    # --- Приведение типов ---
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # coefficient_explained (векторно)
    if "coefficient" in df.columns:
        coef = pd.to_numeric(df["coefficient"], errors="coerce")
        coef_expl = pd.Series(index=df.index, dtype=object)

        mask_minus1 = coef == -1
        mask_zero = coef == 0
        mask_ge1 = coef >= 1

        coef_expl.loc[mask_minus1] = "Приёмка недоступна (игнорировать allowUnload)"
        coef_expl.loc[mask_zero] = "Бесплатная приёмка"
        coef_expl.loc[mask_ge1] = "Множитель стоимости приёмки x" + coef.loc[mask_ge1].astype(str)

        df["coefficient_explained"] = coef_expl

    # boxTypeName (векторно)
    if "boxTypeID" in df.columns:
        box = pd.to_numeric(df["boxTypeID"], errors="coerce")
        box_name = pd.Series(index=df.index, dtype=object)

        mask_nan = box.isna()
        mask_2 = box == 2
        mask_5 = box == 5
        mask_6 = box == 6
        mask_other = ~(mask_nan | mask_2 | mask_5 | mask_6)

        box_name.loc[mask_nan] = "QR-поставка с коробами (boxTypeID не возвращается)"
        box_name.loc[mask_2] = "Короба"
        box_name.loc[mask_5] = "Монопаллеты"
        box_name.loc[mask_6] = "Суперсейф"
        box_name.loc[mask_other] = (
            "Неизвестный тип поставки (" + box.loc[mask_other].astype("Int64").astype(str) + ")"
        )

        df["boxTypeName"] = box_name

    # allowUnload_explained (векторно)
    if "allowUnload" in df.columns:
        allow = df["allowUnload"]
        allow_exp = pd.Series(index=df.index, dtype=object)

        mask_true = allow == True
        mask_false = allow == False

        allow_exp.loc[mask_true] = "Приёмка доступна для данного типа поставки"
        allow_exp.loc[mask_false] = "Приёмка недоступна для данного типа поставки"

        df["allowUnload_explained"] = allow_exp

    # warehouseType (векторно)
    if "isSortingCenter" in df.columns:
        isc = df["isSortingCenter"]
        wh_type = pd.Series(index=df.index, dtype=object)

        mask_true = isc == True
        mask_false = isc == False

        wh_type.loc[mask_true] = "Сортировочный центр (СЦ)"
        wh_type.loc[mask_false] = "Обычный склад"

        df["warehouseType"] = wh_type

    # --- boxAcceptanceType: какие типы поставок принимает склад ---
    required_cols = {"warehouseID", "boxTypeID", "allowUnload", "coefficient"}
    if required_cols.issubset(df.columns):

        active_mask = (
            (df["coefficient"] != -1) &
            (df["allowUnload"] == True)
        )

        mask_boxes_active = active_mask & (df["boxTypeID"] == 2)
        mask_pallets_active = active_mask & (df["boxTypeID"] == 5)

        # считаем по складам, где есть хотя бы одна активная запись
        accepts_boxes = (
            df[mask_boxes_active]
            .groupby("warehouseID")["boxTypeID"]
            .size() > 0
        )
        accepts_pallets = (
            df[mask_pallets_active]
            .groupby("warehouseID")["boxTypeID"]
            .size() > 0
        )

        # создаём availability и явно приводим к bool без fillna
        availability = pd.DataFrame(index=df["warehouseID"].unique())
        availability.index.name = "warehouseID"

        availability["accepts_boxes"] = accepts_boxes.reindex(availability.index, fill_value=False).astype(bool)
        availability["accepts_pallets"] = accepts_pallets.reindex(availability.index, fill_value=False).astype(bool)

        conds = [
            availability["accepts_boxes"] & availability["accepts_pallets"],
            availability["accepts_boxes"] & ~availability["accepts_pallets"],
            ~availability["accepts_boxes"] & availability["accepts_pallets"],
        ]
        choices = [
            "Короба и монопаллеты",
            "Короба",
            "Монопаллеты",
        ]
        box_type = np.select(
            conds,
            choices,
            default="Не принимает короба и монопаллеты"
        )

        availability["boxAcceptanceType"] = box_type

        df = df.merge(
            availability[["boxAcceptanceType"]],
            left_on="warehouseID",
            right_index=True,
            how="left",
        )

    logger.info(f"После обработки: {len(df)} записей по загруженности складов.")
    return df


def get_warehouses_load_full(
    headers=None,
    warehouse_ids=None,
    base_url: str = "https://common-api.wildberries.ru/api/tariffs/v1/acceptance/coefficients",
    max_retries: int = 3,
    to_save: bool = True,
):
    """
    Генеральная функция: получает загруженность складов WB и сразу обрабатывает её.
    Использует:
      - get_warehouses_load_raw  — запрос к API
      - process_warehouses_load  — добавление расшифровок и boxAcceptanceType
    """

    # 1. Сырые данные из API
    df_raw = get_warehouses_load_raw(
        headers=headers,
        warehouse_ids=warehouse_ids,
        base_url=base_url,
        max_retries=max_retries,
    )

    # 2. Обработка / расшифровка
    df_processed = process_warehouses_load(df_raw)

    # 3. Сохранение в CSV
    if to_save:
        output_file = uploaddir_today / f"{date.today()}_Склады_WB.csv"
        try:
            df_processed.to_csv(output_file, sep=";", encoding="utf-8-sig", index=False)
            logger.info("Файл с загруженностью складов сохранён: {}", output_file)
        except Exception as exc:
            logger.error("Не удалось сохранить файл с загруженностью складов {}: {}", output_file, exc)
            raise

    return df_processed

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


# Поставки v2
def get_supply_orders_WB_v2(
    headers,
    status_ids=[2, 3, 4],
    filter_by='date',  # 'date' (по умолчанию) или 'status'
    status_ids_map = {
        1: "Не запланировано",
        2: "Запланировано",
        3: "Отгрузка разрешена",
        4: "Идёт приёмка",
        5: "Принято",
        6: "Отгружено на воротах"
    },
    to_save=True,
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
            status_id = row.get("statusID", "")
            status_name = status_ids_map.get(status_id, f"Неизвестный статус {status_id}")
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

        # Сохранение и возврат результата
        if to_save:
            df_items.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")
        return df_items

    # Если заявки на поставку не нашлись, сохраняем/возвращаем пустой df
    else:
        logger.info("Нет заявок по заданным параметрам.")
        if to_save:
            df_supplies_template.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")
        return df_supplies_template

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

    logger.info(f"Uploading files for client {client_name} for dates {date_start} - {date_end}")
    save_dates_to_csv(date_start, date_end)


    df_products = get_wb_product(headers, to_save=True)

    # Заказы
    df_orders = getOrdersWB(headers, date_start, date_end)
    # Продажи и возвраты
    df_sales = getSalesWB(headers, date_start, date_end, min_days=61)
    # Остатки
    df_reminders = getStockRemindersWB(headers, date_start, date_end)
    # Остатки fbs
    df_reminders_fbs = getStockReminders_fbs(headers)

    # Поставки
    df_supply_items = get_supply_orders_WB_v2(headers)

    # Поставки fbs
    get_supply_orders_WB_fbs(headers)

    # Загруженность складов
    df_warehouses_load = get_warehouses_load_full(headers=headers)
    logger.info(f"DONE UPLOADING FILES FOR CLIENT {client_name}")
    logger.info(f"Uploaded files saved in: {uploaddir_today}")

    print(f"{Back.MAGENTA}{Fore.WHITE}DONE UPLOADING FILES FOR CLIENT {client_name}{Style.RESET_ALL}")
# %%
