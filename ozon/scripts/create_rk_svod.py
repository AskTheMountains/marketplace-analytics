
# %% Определение всех функций
import argparse
import copy
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest import result


import numpy as np
import pandas as pd
import re
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.set_option('future.no_silent_downcasting', True)

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent


# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_companies
# Файл с настройками и номером клиента
# from options import settings, client_number, headers
# Некоторые константы
from ozon.scripts.constants import (
    client_id_performance, client_name,
    client_secret_performance, headers,
    PROMOTION_COMPANIES, shop_type,
    CLIENT_SUMMARY_COL_FORMAT,
    CAMPAIGNS_REPORT_COL_FORMAT
)
# Функции выгрузки данных по API Performane
# Функция формирования статистики РК
from ozon.scripts.create_campaigns_report import \
    create_campaigns_report as upload_campaigns_report
# Функции выгрузки данных по API Seller (доп.)
from ozon.scripts.ozon_generic_functions import (
    get_orders,
    get_reminders_by_product,
)
# Функции выгрузки данных по API Seller
from ozon.scripts.uploadDataFromOzon import get_ozon_product

# Некоторые вспомогательные функции


parser = argparse.ArgumentParser(
    description="Create promotion report (Ozon)",
    usage=(
        "python.exe -m ozon.scripts.create_rk_svod [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD --reference_date YYYY-MM-DD] "
        "[-c N | --client N] "
        "[--promotion_companies \"ID1, ID2, ...\"]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod 2024-11-13 2024-11-20\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --reference_date 2024-11-21 -c 3\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --promotion_companies \"13687629, 14962180, 16997671\"\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --promotion_companies \"13687629 14962180 16997671\"\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --promotion_companies \"13687629;14962180;16997671\"\n\n"
        "  python.exe -m ozon.scripts.create_rk_svod --promotion_companies \"13687629\\n14962180\\n16997671\"\n\n"
        "Можно комбинировать любые разделители — запятая, пробел, точка с запятой, перевод строки.\n"
    ),
    formatter_class=ParagraphRichHelpFormatter
)

# Функция парсинга списка компаний из аргументов командной строки
def parse_cli_companies(cli_companies):
    """
    Принимает строку из аргумента командной строки.
    Возвращает список int ID кампаний.
    Разделители: запятая, точка с запятой, пробел, перевод строки и любые их комбинации.
    """
    if not cli_companies:
        return []
    # Делим по всем: запятая, точка с запятой, пробел, \n, \r
    parts = re.split(r"[,\s;]+", cli_companies.strip())
    cli_companies_parsed = [int(s) for s in parts if s.strip().isdigit()]

    return cli_companies_parsed


# Функция создания директории, куда будет помещен итоговый отчет
def create_client_rk_svod_dir(client_name, date_report=str(date.today())):
    client_rk_path = (
        BASE_DIR
        / "Clients"
        / client_name
        / "CampaignsReport"
    )

    return client_rk_path

# Функция создания директории, куда будет помещаться статистика кампаний
def create_upload_statistic_dir(client_name, date_report=str(date.today())):
    # Задаем путь к директории
    upload_path_statistic = (
        BASE_DIR
        / "Clients"
        / client_name
        / "CampaignsReport"
        / "Данные по РК"
        / date_report
        / f"{date_report}_Кампании_API"
    )

    return upload_path_statistic


# Функция создания диапазона дат
# GPT START ----
def generate_dates(reference_date: str = None, start_date: str = None, end_date: str = None) -> pd.DataFrame:
    # Проверка на несовместимость параметров
    if reference_date and (start_date or end_date):
        raise ValueError("Нельзя одновременно задавать reference_date и start_date/end_date.")

    # Если указан диапазон — используем его
    if start_date and end_date:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    elif start_date or end_date:
        raise ValueError("Нужно задать одновременно start_date и end_date.")
    else:
        # Старая логика по reference_date
        if reference_date is None:
            today = datetime.now()
        else:
            today = datetime.strptime(reference_date, "%Y-%m-%d")

        # Логика в зависимости от дня недели
        # if today.weekday() == 0:
        # Если сегодня понедельник (weekday() == 0)
        # Понедельник предыдущей недели:
        start_dt = today - timedelta(days=7)
        start_dt = start_dt - timedelta(days=start_dt.weekday())
        # Воскресенье предыдущей недели:
        end_dt = start_dt + timedelta(days=6)
        # else:
        # # Если не понедельник:
        # # Понедельник текущей недели:
        # start_dt = today - timedelta(days=today.weekday())
        # # Вчерашний день:
        # end_dt = today - timedelta(days=1)

    start_date_str = start_dt.strftime('%Y-%m-%d')
    end_date_str = end_dt.strftime('%Y-%m-%d')
    start_date_iso = start_dt.strftime('%Y-%m-%dT00:00:00Z')
    end_date_iso = end_dt.strftime('%Y-%m-%dT23:59:59Z')

    df_dates = pd.DataFrame([{
        'date_start': start_date_str,
        'date_end': end_date_str,
        'datetime_start': start_date_iso,
        'datetime_end': end_date_iso
    }])

    return df_dates
# GPT END ----

# Функция получения начальной и конечной даты в формате строки
def get_start_end_date(df_dates, type_dates='companies'):
    # Если задан тип дат - даты для заказов, то берем даты со временм
    if type_dates == 'orders':
        date_start, date_end = df_dates.loc[0, ['datetime_start', 'datetime_end']]
    # В остальных случаях берем даты без времени
    else:
        date_start, date_end = df_dates.loc[0, ['date_start', 'date_end']]

    return date_start, date_end


# Функция создания списка РК с датами
def create_input_companies(
        client_name,
        df_date_range,
        cli_companies_parsed=[]
    ):
    # Если передан список из CLI, используем его
    if cli_companies_parsed:
        input_companies = [[str(company)] for company in cli_companies_parsed]
    else:
        # Берём по client_name из словаря
        companies = PROMOTION_COMPANIES.get(client_name, [])
        if not companies:
            raise ValueError(f"Список кампаний пуст для клиента {client_name}!")
        input_companies = [[str(company)] for company in companies]

    # Получаем даты для отчёта РК
    date_start_performance, date_end_performance = get_start_end_date(df_date_range, type_dates='companies')
    # Добавляем даты к каждой кампании
    for company in input_companies:
        company.extend([date_start_performance, date_end_performance])

    return input_companies


# Функция проверки наличия отчета по РК
def check_rk_report_exist(client_name, date_report=str(date.today())):
    # Директория, где хранятся отчеты по РК
    rk_report_dir = BASE_DIR / "Clients" / client_name / "CampaignsReport" / date_report
    # Имя файла с отчетом РК
    rk_file_name = f"{date_report}_Отчет_РК_{client_name}_Ozon.xlsx"
    # Полный путь к файлу
    filepath_rk_file = rk_report_dir / rk_file_name

    # Проверяем наличие файла в директории
    rk_report_exists = filepath_rk_file.exists()
    # Если файл существует, то возвращаем путь к файлу
    if rk_report_exists:
        return filepath_rk_file
    # Если файл не существует, то возвращаем False
    else:
        return False


# Функция выгрузки отчета по РК
def create_campaigns_report(
        client_name,
        client_id_performance,
        client_secret_performance,
        headers,
        df_dates,
        input_companies,
        upload_path_statistic,
        date_campaigns_report=str(date.today()),
        task_id=None,
        to_save=True,
        delete_files=False

    ):

    # Формируем отчет РК
    df_campaigns_report = upload_campaigns_report(
        client_name,
        client_id_performance,
        client_secret_performance,
        headers,
        df_dates,
        input_companies,
        upload_path_statistic,
        date_campaigns_report,
        task_id=task_id,
        to_save=to_save,
        delete_files=delete_files
    )

    return df_campaigns_report

# Функция обработки отчета РК
def process_campaigns_report(df_campaigns_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_campaigns_report_processed = df_campaigns_report.copy()
    # Убираем сшитые товары
    # df_campaigns_report_processed = (
    #     df_campaigns_report_processed
    #     .loc[df_campaigns_report_processed['Тип товара'].isin(['Товар из РК', 'Товар из заказов']), :]
    #     .reset_index(drop=True)
    # )
    # Переименовываем некоторые колонки
    df_campaigns_report_processed = df_campaigns_report_processed.rename(columns={
        'Расход по кампаниям с НДС, руб.': 'Расход РК, руб.',
        'ДРР по РК': 'ДРР',
        'Количество размеров в артикуле': 'Кол-во размеров в артикуле',
        'Заказы ОЗЗ 5% по размерам, шт.': 'Заказы по размерам ОЗЗ 5%, шт.',
        'Расход ОЗЗ 5% по размерам, руб.': 'Расход по размерам ОЗЗ 5%, руб.',
    })
    # Делаем порядок колонок
    column_order = [
        'Номер кампании',
        'SKU',
        'Артикул',
        'Название товара',
        'Тип товара',
        'Цена товара, руб.',
        'Показы',
        'Клики',
        'CTR (%)',
        'В корзину',
        'Ср. цена клика, руб.',
        'Заказы, шт. (РК)',
        'Заказы, руб. (РК)',
        'Заказы, шт. (Заказы)',
        'Заказы, руб. (Заказы)',
        'Заказы модели, шт. (РК)',
        'Заказы модели, руб. (РК)',
        'Заказы модели, шт. (Заказы)',
        'Заказы модели, руб. (Заказы)',
        'Кол-во размеров в артикуле',
        'Заказы размеров, шт. (Заказы)',
        'Заказы размеров, руб. (Заказы)',
        'Заказы всего, шт.',
        'Заказы всего, руб.',
        'Заказы ОЗЗ 5%, шт.',
        'Расход ОЗЗ 5%, руб.',
        'Заказы по размерам ОЗЗ 5%, шт.',
        'Расход по размерам ОЗЗ 5%, руб.',
        'Расход РК, руб.',
        'Расход РК и ОЗЗ, руб.',
        'Расход РК по размерам, руб.',
        'Расход РК и ОЗЗ по размерам, руб',
        'ДРР',
        'ДРР артикула',
        'Остаток',
        'Дата добавления',
    ]
    existing_column_order = [col for col in column_order if col in df_campaigns_report_processed.columns]
    df_campaigns_report_processed = df_campaigns_report_processed[existing_column_order]

    # У магазинов с товарами без размеров удаляем некоторые колонки, если они есть
    columns_to_drop_no_sizes = [
        'Кол-во размеров в артикуле',
        'Заказы размеров, шт. (Заказы)',
        'Заказы размеров, руб. (Заказы)',
        'Заказы по размерам ОЗЗ 5%, шт.',
        'Расход по размерам ОЗЗ 5%, руб.',
        'Расход РК по размерам, руб.',
        'Расход РК и ОЗЗ по размерам, руб.'
    ]
    if shop_type != 'clothes':
        logger.info("Магазин не одежды, удаляем колонки с размерами, если они есть")
        existing_columns_to_drop = [col for col in columns_to_drop_no_sizes if col in df_campaigns_report_processed.columns]
        df_campaigns_report_processed = df_campaigns_report_processed.drop(columns=existing_columns_to_drop)

    # Переводим Артикул в строку
    df_campaigns_report_processed['Артикул'] = df_campaigns_report_processed['Артикул'].astype(str)

    return df_campaigns_report_processed


# Функция расчета метрик из отчета РК по артикулам
def calc_campaigns_by_article(
        df_campaigns_report
):
    # Создаем копию для избежания изменений в оригинальном df
    df_campaigns_report_by_article = df_campaigns_report.copy()
    # Убираем сшитые товары
    df_campaigns_report_by_article = (
        df_campaigns_report_by_article
        .loc[df_campaigns_report_by_article['Тип товара'].isin(['Товар из РК', 'Товар из заказов']), :]
        .reset_index(drop=True)
    )
    # Переименовываем некоторые колонки
    df_campaigns_report_by_article = df_campaigns_report_by_article.rename(columns={
        'Расход по кампаниям с НДС, руб.': 'РК',
        'ДРР по РК': 'ДРР % по размеру'
    })
    # Считаем параметры РК с группировкой по артикулу
    df_campaigns_report_by_article = (
        df_campaigns_report_by_article
        .groupby('Артикул', as_index=False)
        .agg({
            'РК': 'sum',
            'ДРР % по размеру': 'first'
        })
    )
    # Переводим Артикул в строку
    df_campaigns_report_by_article['Артикул'] = df_campaigns_report_by_article['Артикул'].astype(str)

    return df_campaigns_report_by_article




# Функция обработки отчета ОЗЗ
def process_payment_orders_report(df_payment_orders_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_payment_orders_report_processed = df_payment_orders_report.copy()
    # Считаем параметры ОЗЗ с группировкой по артикулу
    df_payment_orders_report_processed = (
        df_payment_orders_report_processed
        .groupby('Артикул', as_index=False)
        .agg(**{
            'Заказы ОЗЗ 5%, шт.': ('Количество', 'sum'),
            'Расход ОЗЗ 5%, руб.': ('Расход, ₽', 'sum')
        })
    )
    # Переводим Артикул в строку
    df_payment_orders_report_processed['Артикул'] = df_payment_orders_report_processed['Артикул'].astype(str)

    return df_payment_orders_report_processed

# Функция обработка списка товаров АПИ
def process_product_list(df_products):
    # Создаем копию для избежания изменений в оригинальном df
    df_products_processed = df_products.copy()
    # Выбираем нужные колонки
    df_products_processed = df_products_processed.loc[:, df_products_processed.columns.isin([
        'Артикул',
        'Название товара',
        'SKU',
        # 'Отзывы',
        # 'Текущая цена с учетом скидки, ₽',
    ])]
    # Удаляем лишние символы из артикула
    df_products_processed['Артикул'] = df_products_processed['Артикул'].str.replace("'", "", regex=False)
    # Переводим баркод в число
    # df_products_processed['Barcode'] = pd.to_numeric(df_products_processed['Barcode'], errors='coerce')
    # df_products_processed['Barcode'] = df_products_processed['Barcode'].apply(lambda x: format(x, 'f') if pd.notnull(x) else x)
    # df_products_processed['Barcode'] = df_products_processed['Barcode'].apply(lambda x: str(x).split('.')[0] if pd.notnull(x) else x).astype('Int64')
    # Переименовываем колонки с ценами, которые приходят из отчета по товарам из апи
    # чтобы потом не было путаницы с ценами, которые мы получаем из отдельного метода
    df_products_processed = df_products_processed.rename(columns={
        'Текущая цена с учетом скидки, ₽': 'Текущая цена с учетом скидки (из отчета по товарам)',
        # 'Цена до скидки (перечеркнутая цена), ₽': 'Цена до скидки (перечеркнутая цена) (из отчета по товарам)',
        # 'Рыночная цена, ₽': 'Рыночная цена (из отчета по товарам)',
        # 'Цена Premium, ₽': 'Цена Premium (из отчета по товарам)',
    })
    # Переименовываем колонки с нужными ценами
    df_products_processed = df_products_processed.rename(columns={
        # 'Цена до скидки (перечеркнутая цена) (из отчета по товарам)': 'Цена до скидки',
        'Текущая цена с учетом скидки (из отчета по товарам)': 'Текущая цена с учетом скидки',
        # 'Минимальная цена после применения всех скидок': 'Мин. цена Ozon',
        # 'Цена до учета скидок (зачеркнутая)': 'Цена до скидки',
        # 'Цена с учетом скидок (на карточке товара)': 'Цена после скидки',
        # 'Минимальная цена после применения всех скидок': 'Мин. цена Ozon',
        # 'Цена с учетом акций продавца': 'Цена по акции',
        # 'Цена с учетом всех акций': 'Цена с баллами Озон',
    })
    # Переводим SKU в строку чтобы не было ошибок при merge
    df_products_processed['SKU'] = df_products_processed['SKU'].astype(str)
    # Переводим Артикул в строку
    df_products_processed['Артикул'] = df_products_processed['Артикул'].astype(str)


    return df_products_processed

# Функция получения списка моделей товаров
def get_model_list(df_products, df_campaigns_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_models = df_products.copy()
    # Выбираем нужные колонки
    df_models = df_models.loc[:, [
        'model_id',
        'SKU',
        'Артикул',
        'Название товара',

    ]]
    # Удаляем лишние символы из артикула
    df_models['Артикул'] = df_models['Артикул'].str.replace("'", "", regex=False)
    # Переводим Артикул в строку
    df_models['Артикул'] = df_models['Артикул'].astype(str)
    # Переименовываем колонки для удобства
    df_models = df_models.rename(columns={
        'model_id': 'ID модели',
    })
    # Добавляем данные из отчета РК
    df_campaigns_products = (
        df_campaigns_report
        .drop_duplicates(subset=['Артикул', 'Тип товара'])
        .loc[
            df_campaigns_report['Тип товара'].isin(['Товар из РК', 'Товар из заказов']),
            ['Артикул', 'Тип товара', 'Остаток']
        ]
    )
    df_models = df_models.merge(
        df_campaigns_products,
        how='left',
        on='Артикул',
    )
    # Делаем сортировку по ID модели
    df_models = df_models.sort_values(by=['ID модели'], ignore_index=True)

    return df_models

# Функция выгрузки заказов
def get_orders_for_svod(df_dates, headers):
    # Получаем даты для отчета по заказам
    date_start_orders, date_end_orders = get_start_end_date(df_dates, type_dates='orders')
    # Выгружаем заказы
    df_orders_all = get_orders(
        headers,
        date_start_orders,
        date_end_orders
    )

    return df_orders_all


# Функция расчета заказов по артикулам
def calc_orders(df_orders):
    # Убираем отмененные отправления
    df_orders_filtered = (
        df_orders
        .loc[~df_orders['Статус'].isin(['Отменён', 'Отменен']), :]
        .copy()
    )
    # Считаем заказы по артикулам
    df_orders_by_products = (
        df_orders_filtered
        .groupby(['Артикул'], as_index=False)
        .agg(**{
            'Сумма заказов': ('Заказы руб', 'sum'),
            'Кол-во заказов': ('Заказы шт', 'sum')
        })
    )

    return df_orders_by_products

# Функция добавления данных к списку товаров
def add_data_to_product_list(
        df_products_processed,
        df_campaigns_report_by_article,
        df_payment_orders_report_processed,
        df_reminders,
        df_orders_by_products,
    ):
    # Создаем список датафреймов, которые объединяем
    merge_params = [
        {"df": df_reminders, "columns": ["Остаток",], "on": "Артикул"},
        {"df": df_campaigns_report_by_article, "columns": ["РК"], "on": "Артикул"},
        {"df": df_orders_by_products, "columns": ["Сумма заказов", "Кол-во заказов"], "on": "Артикул"},
        {"df": df_payment_orders_report_processed, "columns": ["Заказы ОЗЗ 5%, шт.", "Расход ОЗЗ 5%, руб."], "on": "Артикул"},
    ]

    # Начинаем с основного датафрейма
    df_products_svod = df_products_processed.copy()

    for params in merge_params:
        # Ключ, по которому мерджим
        key = params["on"]
        # Берем указанные колонки из текущего df + колонку, по которой мерджим
        df_to_merge = params["df"][params["columns"] + [key]].copy()
        df_products_svod = df_products_svod.merge(df_to_merge, how='left', on=key, suffixes=('', f'_{key}_right'))

    return df_products_svod

# Функция расчета доп колонок
def calc_additional_columns(df_products_svod):
    # Создаем копию для избежания изменений в оригинальном df
    df_client_svod = df_products_svod.copy()
    # Заполняем пропуски
    columns_to_fillna = ['Остаток', 'Сумма заказов', 'Кол-во заказов']
    df_client_svod[columns_to_fillna] = (
        df_client_svod[columns_to_fillna]
        .fillna(0)
        .infer_objects(copy=False)
    )
    # Переводим некоторые колонки в integer
    columns_to_int = ['Остаток', 'Кол-во заказов']
    df_client_svod[columns_to_int] = (
        df_client_svod[columns_to_int]
        .astype(int)
    )
    # Средняя цена продажи
    df_client_svod['Средняя цена продажи'] = np.where(
        df_client_svod['Кол-во заказов'] > 0,
        df_client_svod['Сумма заказов'] / df_client_svod['Кол-во заказов'],
        0
    )
    # ДРР %
    df_client_svod['ДРР'] = (
        df_client_svod[['РК', 'Расход ОЗЗ 5%, руб.']].sum(axis=1)
        / df_client_svod['Сумма заказов'].replace(0, np.nan) * 100
    )
    df_client_svod['ДРР'] = df_client_svod['ДРР'].replace([-np.inf, np.inf], np.nan)
    return df_client_svod

# Функция создания df для записи в excel
def create_client_svod_excel(df_client_svod, date_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_client_excel_svod = df_client_svod.copy()
    # Добавляем дату в колонку заказов
    date_report_formatted = pd.to_datetime(date_report).strftime('%d.%m')
    df_client_excel_svod = df_client_excel_svod.rename(columns={
        'Остаток': f'Остаток {date_report_formatted}'
    })
    # Делаем порядок колонок
    column_order = [
        'SKU', 'Артикул', 'Название товара',
        'Средняя цена продажи',
        'РК',
        'Заказы ОЗЗ 5%, шт.', 'Расход ОЗЗ 5%, руб.',
        'Кол-во заказов', 'Сумма заказов',
        'ДРР',
        f'Остаток {date_report_formatted}',
        # 'Тип кампании',
        # 'Отзывы',
        # 'Текущая цена с учетом скидки (из отчета по товарам)',
        # 'ДРР % по размеру',

    ]
    df_client_excel_svod = df_client_excel_svod[column_order]

    # Сортировка по артикулу
    df_client_excel_svod = df_client_excel_svod.sort_values(by=['Артикул'], ignore_index=True)

    # Заменяем все значения 0 и 0.0 на NaN
    df_client_excel_svod = df_client_excel_svod.replace({0: np.nan, 0.0: np.nan})

    return df_client_excel_svod


# Функция форматирования столбца с номером РК
def format_campaigns_columns(ws, df, column_names=['Номер кампании', 'ID Кампании']):

    # 1. Палитра из 15 неярких расцветок
    color_palette = [
        "FFB3C6", "FFE6E6", "B6E2D3", "B7D3DF", "DEEAF6",
        "F7E1D7", "F5F8C4", "D4E6A5", "A2D5AB", "F9D6C1",
        "F5B5FC", "D9D9D9", "ABC4FF", "C6F1E7", "EEE3CB"
    ]

    # 3. Находим имена колонок кампании в df
    df_campaign_cols = [
        col for col in df.columns if any(col.strip().lower() == name.strip().lower() for name in column_names)
    ]
    if not df_campaign_cols:
        return  # Нет нужных колонок — ничего делать не надо

    # 4. Получаем все уникальные значения кампании
    unique_campaigns = []
    for col in df_campaign_cols:
        unique_campaigns.extend(df[col].dropna().unique())
    unique_campaigns = list(dict.fromkeys(unique_campaigns))

    # 5. Цвета кампаниям (циклично)
    campaign_to_color = {
        val: color_palette[i % len(color_palette)]
        for i, val in enumerate(unique_campaigns)
    }

    # 6. Индексы колонок кампании в Excel (по строке заголовков листа)
    header_cells = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    campaign_col_idxs = [
        idx for idx, header in enumerate(header_cells) if header in df_campaign_cols
    ]
    if not campaign_col_idxs:
        return

    # 7. Форматируем ячейки
    for idx in campaign_col_idxs:
        prev_camp = None
        for excel_row_idx, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2
        ):
            cell = row[idx]
            camp = cell.value
            # Заливаем цвет
            if camp in campaign_to_color:
                cell.fill = PatternFill(
                    start_color=campaign_to_color[camp],
                    end_color=campaign_to_color[camp],
                    fill_type='solid'
                )
            # Проверяем переход к новой кампании
            is_last_row = (excel_row_idx == ws.max_row)
            # Доступ к значению кампании в следующей строке (если есть)
            next_camp = ws[excel_row_idx + 1][idx].value if not is_last_row else None
            if camp != next_camp:
                # Последняя строка блока — нижняя граница жирная у всей строки
                for c in row:
                    c.border = Border(
                        left=c.border.left,
                        right=c.border.right,
                        top=c.border.top,
                        bottom=Side(style='medium')
                    )



def format_all_sheets_common(wb, exclude_columns=None):
    """
    Общее форматирование для всех листов книги:
    - границы
    - заголовки: центр, перенос, жирный, цвет заливки по списку
    - данные: вертикальное выравнивание по центру, без переноса (по умолчанию)
    - числовой формат (целое / с двумя знаками) для чисел, кроме исключённых колонок
    - автоширина столбцов с небольшим запасом
    - автофильтр по всей таблице
    - если есть столбец 'Тип товара' — условное форматирование:
      значение 'Товар из РК' -> жёлтая заливка
    - дополнительное форматирование столбцов по объединённому словарю
      CLIENT_SUMMARY_COL_FORMAT + CAMPAIGNS_REPORT_COL_FORMAT:
        * поиск по startswith(ключ)
        * для совпавших — выставляем width и alignment из словаря
        * для остальных — остаётся автоподбор ширины и базовое выравнивание
    """
    if exclude_columns is None:
        exclude_columns = ["SKU", "SKU продвигаемого товара", "ID заказа", "ID модели", "model_id"]
    exclude_columns_lower = {col.lower() for col in exclude_columns}

    # --- объединяем два словаря форматирования ---
    # при конфликте ключей приоритет у CAMPAIGNS_REPORT_COL_FORMAT
    combined_col_format = dict(CLIENT_SUMMARY_COL_FORMAT)
    for k, v in CAMPAIGNS_REPORT_COL_FORMAT.items():
        combined_col_format[k] = v

    # Граница
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Выравнивание: заголовок
    header_alignment = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )

    # Базовое выравнивание данных (для всех, затем может быть переопределено словарём)
    data_alignment_default = Alignment(
        vertical='center',
        wrap_text=False
    )

    # Цвета заголовков
    BLUE_1 = "8DB4E2"  # 141, 180, 226
    ORANGE = "FCD5B4"  # 252, 213, 180
    BLUE_2 = "DAEEF3"  # 218, 238, 243

    blue1_headers = {
        "Номер кампании",
        "SKU",
        "Артикул",
        "Название товара",
        "Тип товара",
        "Цена товара, руб.",
        "Показы",
        "Клики",
        "CTR (%)",
        "В корзину",
        "Ср. цена клика, руб.",
        "Заказы шт. (РК)",
        "Заказы руб. (РК)",
        "Заказы шт. (Заказы)",
        "Заказы руб. (Заказы)",
        "Заказы модели шт. (РК)",
        "Заказы модели руб. (РК)",
        "Заказы модели шт. (Заказы)",
        "Заказы модели руб. (Заказы)",
        "Кол-во размеров в артикуле",
        "Заказы размеров шт. (Заказы)",
        "Заказы размеров руб. (Заказы)",
        "Заказы всего шт.",
        "Заказы всего руб.",
    }
    orange_headers = {
        "Заказы ОЗЗ 5%, шт.",
        "Расход ОЗЗ 5%, руб.",
        "Заказы по размерам ОЗЗ 5%, шт.",
        "Расход по размерам ОЗЗ 5%, руб.",
    }
    blue2_headers = {
        "Расход РК, руб.",
        "Расход РК и ОЗЗ, руб.",
        "Расход РК по размерам, руб.",
        "Расход РК и ОЗЗ по размерам, руб.",
    }

    header_color_map = {}
    for h in blue1_headers:
        header_color_map[h] = BLUE_1
    for h in orange_headers:
        header_color_map[h] = ORANGE
    for h in blue2_headers:
        header_color_map[h] = BLUE_2

    def get_header_fill(header_value):
        if not isinstance(header_value, str) or not header_value.strip():
            return None
        color = header_color_map.get(header_value)
        if color is None:
            color = BLUE_1
        return PatternFill(start_color=color, end_color=color, fill_type="solid")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        min_row = ws.min_row
        max_row = ws.max_row
        min_col = ws.min_column
        max_col = ws.max_column

        # Заголовки
        header_row = next(ws.iter_rows(min_row=1, max_row=1))
        headers = [cell.value for cell in header_row]

        # Индексы столбцов, которые надо исключить из числового форматирования
        exclude_idx = {
            idx
            for idx, header in enumerate(headers)
            if isinstance(header, str) and header.lower() in exclude_columns_lower
        }

        # Высота первой строки
        ws.row_dimensions[1].height = 45

        # Базовое оформление: границы, заголовки, числа, базовое выравнивание
        for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                min_col=min_col, max_col=max_col):
            for idx, cell in enumerate(row):
                cell.border = thin_border

                if cell.row == 1:
                    # Заголовок
                    cell.font = Font(name="Calibri", size=11, bold=True, color="000000")
                    cell.alignment = header_alignment
                    header_val = cell.value
                    fill = get_header_fill(header_val)
                    if fill is not None:
                        cell.fill = fill
                else:
                    # Данные — базовое выравнивание
                    cell.alignment = data_alignment_default

                    # Числовой формат
                    if idx not in exclude_idx and isinstance(cell.value, (int, float)):
                        if float(cell.value).is_integer():
                            cell.number_format = '#,##0'
                        else:
                            cell.number_format = '#,##0.00'

        # Базовый автофит ширины
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_len + 2

        # Автофильтр
        ws.auto_filter.ref = ws.dimensions

        # Условное форматирование по "Тип товара"
        if "Тип товара" in headers:
            type_col_idx = headers.index("Тип товара") + 1
            type_col_letter = get_column_letter(type_col_idx)
            cell_range = f"{type_col_letter}2:{type_col_letter}{ws.max_row}"

            rule = CellIsRule(
                operator='equal',
                formula=['"Товар из РК"'],
                stopIfTrue=False,
                fill=PatternFill(
                    start_color="F2E422",
                    end_color="F2E422",
                    fill_type="solid"
                )
            )
            ws.conditional_formatting.add(cell_range, rule)

        # Дополнительное форматирование по объединённому словарю
        # Для каждой колонки листа ищем первый key в combined_col_format,
        # который является префиксом заголовка (header.startswith(key)).
        for col_idx, header in enumerate(headers, start=1):
            if not isinstance(header, str):
                continue

            fmt = None
            for key, cfg in combined_col_format.items():
                if header.startswith(key):
                    fmt = cfg
                    break

            if fmt is None:
                continue  # для этой колонки остаётся автофит и базовое выравнивание

            width = fmt.get('width')
            align_kwargs = fmt.get('alignment', {})
            alignment = Alignment(**align_kwargs)

            col_letter = get_column_letter(col_idx)

            # переопределяем ширину из словаря
            if width is not None:
                ws.column_dimensions[col_letter].width = width

            # переопределяем выравнивание только для данных (со 2-й строки)
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.alignment = alignment


def format_client_summary_sheet(wb, sheet_name, df_client_excel_svod, percent_cols):
    """
    Форматирование листа сводного отчёта по клиенту:
    - делим значения ДРР на 100
    - задаём формат процентов
    - условное форматирование (градиент) по % столбцам
    - подгоняем ширину и выравнивание для конкретных столбцов по маппингу
    """
    if sheet_name not in wb.sheetnames:
        return

    # маппинг ширины и выравнивания
    # CLIENT_SUMMARY_COL_FORMAT = {
    #     'SKU': {
    #         'width': 12.0,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Артикул': {
    #         'width': 12.0,
    #         'alignment': {'horizontal': 'left', 'vertical': 'center'}
    #     },
    #     'Название товара': {
    #         'width': 47.8,  # округлено до одного знака
    #         'alignment': {'horizontal': 'left', 'vertical': 'center'}
    #     },
    #     'Средняя цена продажи': {
    #         'width': 11.4,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'РК': {
    #         'width': 12.1,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Заказы ОЗЗ 5%, шт.': {
    #         'width': 11.4,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Расход ОЗЗ 5%, руб.': {
    #         'width': 13.0,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Кол-во заказов': {
    #         'width': 10.1,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Сумма заказов': {
    #         'width': 12.1,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'ДРР': {
    #         'width': 9.0,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    #     'Остаток': {
    #         'width': 11.4,
    #         'alignment': {'horizontal': 'center', 'vertical': 'center'}
    #     },
    # }

    ws = wb[sheet_name]

    # --- 1. Процентный формат ---
    for col_name in percent_cols:
        if col_name not in df_client_excel_svod.columns:
            continue
        col_idx = list(df_client_excel_svod.columns).index(col_name) + 1  # Excel: 1-based
        for row in range(2, ws.max_row + 1):  # пропускаем заголовок
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                cell.value = cell.value * 0.01  # делим на 100 для корректного отображения %
            cell.number_format = '0.0%'

    # --- 2. Условное форматирование (градиент) ---
    gradient_cols = percent_cols
    for col_name in gradient_cols:
        if col_name not in df_client_excel_svod.columns:
            continue
        col_idx = list(df_client_excel_svod.columns).index(col_name) + 1
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        cell_range = f"{col_letter}2:{col_letter}{ws.max_row}"
        rule = ColorScaleRule(
            start_type='min', start_color='63BE7B',   # зелёный
            mid_type='percentile', mid_value=50, mid_color='FFEB84',  # жёлтый
            end_type='max', end_color='E06967',       # красный
        )
        ws.conditional_formatting.add(cell_range, rule)

    # --- 3. Ширина столбцов и выравнивание данных по маппингу ---
    # Берём заголовки с листа
    # header_row = next(ws.iter_rows(min_row=1, max_row=1))
    # headers = [cell.value for cell in header_row]

    # for col_idx, header in enumerate(headers, start=1):
    #     if header not in CLIENT_SUMMARY_COL_FORMAT:
    #         continue

    #     fmt = None
    #     # ищем первый ключ-«префикс», который подходит под startswith
    #     for key, cfg in CLIENT_SUMMARY_COL_FORMAT.items():
    #         if header.startswith(key):
    #             fmt = cfg
    #             break

    #     if fmt is None:
    #         continue

    #     # fmt = CLIENT_SUMMARY_COL_FORMAT[header]
    #     width = fmt.get('width')
    #     align_kwargs = fmt.get('alignment', {})
    #     alignment = Alignment(**align_kwargs)

    #     col_letter = get_column_letter(col_idx)

    #     # устанавливаем ширину столбца
    #     if width is not None:
    #         ws.column_dimensions[col_letter].width = width

    #     # выравнивание только для данных (со 2-й строки)
    #     for row in range(2, ws.max_row + 1):
    #         cell = ws.cell(row=row, column=col_idx)
    #         cell.alignment = alignment


def format_campaigns_report_sheet(
        wb,
        sheet_name,
        df_campaigns_report=None,
        percent_cols=None,
):
    """
    Форматирование листа "Отчет РК {client_name}":
    - подгоняем ширину столбцов и выравнивание данных по маппингу заголовков
    - условное форматирование по столбцу "Тип товара" (значение "Товар из РК" -> жёлтая заливка)
    - процентный формат и градиентное условное форматирование для percent_cols
    """
    if sheet_name not in wb.sheetnames:
        return

    ws = wb[sheet_name]


    # Заголовки с листа
    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    headers = [cell.value for cell in header_row]

    # --- 1. Ширина столбцов и выравнивание данных ---
    for col_idx, header in enumerate(headers, start=1):
        if header not in CAMPAIGNS_REPORT_COL_FORMAT:
            continue

        fmt = CAMPAIGNS_REPORT_COL_FORMAT[header]
        width = fmt.get('width')
        align_kwargs = fmt.get('alignment', {})
        alignment = Alignment(**align_kwargs)

        col_letter = get_column_letter(col_idx)

        if width is not None:
            ws.column_dimensions[col_letter].width = width

        # данные (со 2-й строки)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.alignment = alignment

    # # --- 2. Условное форматирование по столбцу "Тип товара" ---
    # # Значение "Товар из РК" -> жёлтая заливка
    # if df_campaigns_report is not None and "Тип товара" in df_campaigns_report.columns:
    #     try:
    #         type_col_idx = list(df_campaigns_report.columns).index("Тип товара") + 1
    #         type_col_letter = get_column_letter(type_col_idx)
    #         cell_range = f"{type_col_letter}2:{type_col_letter}{ws.max_row}"

    #         rule = CellIsRule(
    #             operator='equal',
    #             formula=['"Товар из РК"'],
    #             stopIfTrue=False,
    #             fill=PatternFill(
    #                 start_color="f2e422",  # мягкий жёлтый
    #                 end_color="f2e422",
    #                 fill_type="solid"
    #             )
    #         )
    #         ws.conditional_formatting.add(cell_range, rule)
    #     except ValueError:
    #         pass

    # --- 3. Процентный формат и градиент для percent_cols ---
    if percent_cols is None:
        percent_cols = []

    # 3.1. Процентный формат
    for col_name in percent_cols:
        if df_campaigns_report is None or col_name not in df_campaigns_report.columns:
            continue
        col_idx = list(df_campaigns_report.columns).index(col_name) + 1  # Excel: 1-based
        for row in range(2, ws.max_row + 1):  # пропускаем заголовок
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                cell.value = cell.value * 0.01
            cell.number_format = '0.0%'

    # 3.2. Градиентное условное форматирование (по тем же столбцам)
    for col_name in percent_cols:
        if df_campaigns_report is None or col_name not in df_campaigns_report.columns:
            continue
        col_idx = list(df_campaigns_report.columns).index(col_name) + 1
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        cell_range = f"{col_letter}2:{col_letter}{ws.max_row}"

        rule = ColorScaleRule(
            start_type='min', start_color='63BE7B',   # зелёный
            mid_type='percentile', mid_value=50, mid_color='FFEB84',  # жёлтый
            end_type='max', end_color='E06967',       # красный
        )
        ws.conditional_formatting.add(cell_range, rule)

# Функция сохранения и форматирования excel
def save_and_format_excel(
        df_client_excel_svod,
        df_campaigns_report,
        df_payment_orders_report,
        df_models,
        client_rk_svod_dir,
        date_report,
        df_dates
    ):
    # Формируем директорию для сохранения итогового отчета
    path_report = client_rk_svod_dir
    # Формируем имя файла отчета
    name_report = f"{date_report}_Сводная_РК_{client_name}_Ozon.xlsx"
    file_name_report = path_report / name_report
    path_report.mkdir(parents=True, exist_ok=True)

    # Формируем даты для названия листа
    date_start_excel = (
        pd.to_datetime(df_dates['date_start'])
        .iloc[0]
        .strftime('%d.%m')
    )
    date_end_excel = (
        pd.to_datetime(df_dates['date_end'])
        .iloc[0]
        .strftime('%d.%m')
    )

    # Колонки с процентами в сводном файле
    percent_cols = [
        'ДРР',
        'ДРР по РК',
        'ДРР артикула',
    ]
    sheet_name_client_svod = f"{client_name} {date_start_excel}-{date_end_excel}"

    # Сохраняем книгу
    with pd.ExcelWriter(file_name_report) as w:
        # Сводный отчет
        df_client_excel_svod.to_excel(w, index=False, sheet_name=sheet_name_client_svod)
        # Основной отчет РК
        df_campaigns_report.to_excel(w, index=False, sheet_name=f'Отчет РК {client_name}')
        # Отчет оплаты за заказ
        df_payment_orders_report.to_excel(w, index=False, sheet_name='ОЗЗ 5%')
        # Отчет по моделям товаров
        df_models.to_excel(w, index=False, sheet_name='Модели товаров')

    # ---- Форматируем файл с отчетом ----
    wb = load_workbook(file_name_report)

    # 1) Общее форматирование для всех листов
    format_all_sheets_common(wb)

    # 2) Специальное форматирование листа "Отчет РК {client_name}" (заготовка)
    sheet_campaigns = f'Отчет РК {client_name}'
    format_campaigns_report_sheet(
        wb,
        sheet_name=sheet_campaigns,
        df_campaigns_report=df_campaigns_report,
        percent_cols=percent_cols
    )

    # 3) Форматирование столбцов кампаний на листе "Отчет РК ..."
    if sheet_campaigns in wb.sheetnames:
        ws_campaigns = wb[sheet_campaigns]
        format_campaigns_columns(ws_campaigns, df_campaigns_report)

    # 4) Форматирование столбцов с моделями товаров на листе "Модели товаров"
    sheet_models = 'Модели товаров'
    if sheet_models in wb.sheetnames:
        ws_models = wb[sheet_models]
        format_campaigns_columns(ws_models, df_models, column_names=['ID модели', 'model_id'])

    # 5) Форматирование сводного листа клиента (проценты, градиент)
    format_client_summary_sheet(
        wb,
        sheet_name=sheet_name_client_svod,
        df_client_excel_svod=df_client_excel_svod,
        percent_cols=percent_cols,
    )

    # Сохранение обновленного файла
    wb.save(file_name_report)
    return file_name_report


# %% Вызов всех функций
if __name__ == '__main__':

   # Парсим аргументы командной строки
    params_companies = cli_companies(parser)

    # Получаем даты из аргументов командной строки
    date_start_cli = params_companies["start_date"]
    date_end_cli = params_companies["end_date"]

    # Получаем список кампаний из аргументов командной строки
    cli_companies = params_companies["promotion_companies"]

    # Парсим список кампаний командной строки
    cli_companies_parsed = parse_cli_companies(cli_companies)

    # Дата, за который считался (или считается) отчет
    date_report=str(date.today())
    date_report = '2026-03-31'  # для теста
    # Директория, куда будет сохраняться итоговый отчет
    client_rk_svod_dir = create_client_rk_svod_dir(client_name, date_report)
    # Директория, куда будет загружена статистика кампаний
    upload_path_statistic = create_upload_statistic_dir(client_name, date_report)

    # Генерируем диапазон дат
    df_dates = generate_dates(
        # reference_date=dates_cli["reference_date"],
        start_date=date_start_cli,
        end_date=date_end_cli,
        # start_date='2025-10-27',
        # end_date='2025-11-06'
    )
    # Выбираем начальную и конечную дату
    date_start_print = df_dates['date_start'].iloc[0]
    date_end_print = df_dates['date_end'].iloc[0]
    logger.info(
        f"\nCreating client rk svod for client {client_name}\n"
        f"dates: {date_start_print} - {date_end_print}"
    )

    # Формируем список РК
    input_companies = create_input_companies(
        client_name,
        df_dates,
        cli_companies_parsed
    )

    # Получаем отчет РК
    result_companies_report = create_campaigns_report(
        client_name,
        client_id_performance,
        client_secret_performance,
        headers,
        df_dates,
        input_companies,
        upload_path_statistic,
        date_campaigns_report=date_report,
        task_id=None,
        to_save=True,
        delete_files=False
    )
    # Получаем нужные датафреймы из результата функции create_campaigns_report
    df_campaigns_report = result_companies_report['Отчет РК']
    df_payment_orders_report = result_companies_report['ОЗЗ 5%']

    # Обрабатываем отчет РК
    df_campaigns_report_processed = process_campaigns_report(df_campaigns_report)
    # Считаем метрики РК по артикулам
    df_campaigns_report_by_article = calc_campaigns_by_article(df_campaigns_report)
    # Обрабатываем отчет по заказам ОЗЗ 5%
    df_payment_orders_report_processed = process_payment_orders_report(df_payment_orders_report)

    # Получаем список товаров АПИ
    df_products = get_ozon_product(headers, to_save=False)
    # Обрабатываем список товаров
    df_products_processed = process_product_list(df_products)
    # Получаем список моделей товаров
    df_models = get_model_list(df_products, df_campaigns_report)

    # Получаем заказы
    df_orders = get_orders_for_svod(df_dates, headers)
    # Получаем остатки
    df_reminders = get_reminders_by_product(df_products, headers)

    # Считаем заказы по артикулу
    df_orders_by_products = calc_orders(df_orders)
    # Добавляем данные к списку товаров
    df_products_svod = add_data_to_product_list(
        df_products_processed,
        df_campaigns_report_by_article,
        df_payment_orders_report_processed,
        df_reminders,
        df_orders_by_products,
    )
    # Расчитываем доп. колонки
    df_client_svod = calc_additional_columns(df_products_svod)
    # Формируем таблицу для записи в excel
    df_client_excel_svod = create_client_svod_excel(
        df_client_svod,
        date_report
    )
    # Сохраняем excel
    file_name_report = save_and_format_excel(
        df_client_excel_svod,
        df_campaigns_report_processed,
        df_payment_orders_report,
        df_models,
        client_rk_svod_dir,
        date_report,
        df_dates
    )
    logger.info(f"Done creating rk client svod for client {client_name}")
    logger.info(f"RK client svod saved in: {file_name_report}")

# %%
