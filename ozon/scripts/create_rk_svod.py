
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
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
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
    PROMOTION_COMPANIES, PROMOTION_ORDERS_THRESHOLD, RK_BUYOUT_PERCENT, shop_type,
    CANCEL_PERCENT_THRESHOLD,
    CLIENT_SUMMARY_COL_FORMAT,
    CAMPAIGNS_REPORT_COL_FORMAT,
    enable_ai_recommendations,
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

# ИИ-помощник для рекомендаций по РК
from ozon.scripts.ai_assistant_rk import add_ai_recommendations

# Некоторые вспомогательные функции
# Добавляем рекомендации ИИ-помощника (отдельный шаг; при ошибке отчёт формируется без них)
from ozon.scripts.ai_assistant_rk import add_ai_recommendations


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
        'Заказы ОЗЗ ст. по размерам, шт.': 'Заказы по размерам ОЗЗ ст., шт.',
        'Расход ОЗЗ ст. по размерам, руб.': 'Расход по размерам ОЗЗ ст., руб.',
    })
    # Добавляем колонку для добавлений в корзину из заказов
    # df_campaigns_report_processed['Добавления в корзину (Заказы)'] = np.nan
    # Добавляем столбец с комментарием
    df_campaigns_report_processed['Комментарий'] = np.nan
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
        'Добавления в корзину',
        # 'Добавления в корзину (Заказы)',
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
        'Заказы ОЗЗ ст., шт.',
        'Расход ОЗЗ ст., руб.',
        'Заказы по размерам ОЗЗ ст., шт.',
        'Расход по размерам ОЗЗ ст., руб.',
        'Расход РК, руб.',
        'Расход РК и ОЗЗ, руб.',
        'Расход РК по размерам, руб.',
        'Расход РК и ОЗЗ по размерам, руб',
        'ДРР',
        'ДРР артикула',
        'Остаток',
        'Дата добавления',
        'Комментарий',
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
        'Заказы по размерам ОЗЗ ст., шт.',
        'Расход по размерам ОЗЗ ст., руб.',
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
            'Заказы ОЗЗ 5%, шт.': ('Заказы ОЗЗ 5%, шт.', 'sum'),
            'Расход ОЗЗ 5%, руб.': ('Расход ОЗЗ 5%, руб.', 'sum'),
            'Расход ОЗЗ ст., руб.': ('Расход ОЗЗ ст., руб.', 'sum'),
            'Заказы ОЗЗ ст., шт.': ('Заказы ОЗЗ ст., шт.', 'sum'),
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
def get_model_list(df_products, df_campaigns_report, df_reminders):
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
            ['Артикул', 'Тип товара']
        ]
    )
    df_models = (
        df_models
        # Добавляем данные из отчета РК
        .merge(
            df_campaigns_products,
            how='left',
            on='Артикул',
        )
        # Добавляем данные по остаткам
        .merge(
            df_reminders[['Артикул', 'Остатки FBO', 'Остатки FBS']],
            how='left',
            on='Артикул',
        )
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
    # Всего заказов по артикулам (включая отмененные)
    df_orders_total = (
        df_orders
        .groupby('Артикул', as_index=False)
        .agg(**{'Всего заказов': ('Заказы шт', 'sum')})
    )
    # Отмененные заказы по артикулам
    df_orders_cancelled = (
        df_orders
        .loc[df_orders['Статус'].isin(['Отменён', 'Отменен']), :]
        .groupby('Артикул', as_index=False)
        .agg(**{'Кол-во отмен': ('Заказы шт', 'sum')})
    )
    # Считаем заказы по артикулам (без отмененных)
    df_orders_by_products = (
        df_orders
        .loc[~df_orders['Статус'].isin(['Отменён', 'Отменен']), :]
        .groupby(['Артикул'], as_index=False)
        .agg(**{
            'Сумма заказов': ('Заказы руб', 'sum'),
            'Кол-во заказов': ('Заказы шт', 'sum')
        })
    )
    # Добавляем общее число заказов и число отмен
    df_orders_by_products = (
        df_orders_by_products
        .merge(df_orders_total, how='left', on='Артикул')
        .merge(df_orders_cancelled, how='left', on='Артикул')
    )
    # Заполняем пропуски по отменам
    df_orders_by_products['Кол-во отмен'] = df_orders_by_products['Кол-во отмен'].fillna(0)
    # Процент отмененных заказов (доля от общего числа заказов)
    df_orders_by_products['Процент отмен'] = np.where(
        df_orders_by_products['Всего заказов'] > 0,
        df_orders_by_products['Кол-во отмен'] / df_orders_by_products['Всего заказов'],
        0
    )
    # Убираем вспомогательную колонку
    df_orders_by_products = df_orders_by_products.drop(columns=['Всего заказов'])

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
        {"df": df_reminders, "columns": ["Остатки FBO", "Остатки FBS"], "on": "Артикул"},
        {"df": df_campaigns_report_by_article, "columns": ["РК"], "on": "Артикул"},
        {"df": df_orders_by_products, "columns": ["Сумма заказов", "Кол-во заказов", "Процент отмен"], "on": "Артикул"},
        {"df": df_payment_orders_report_processed, "columns": ["Заказы ОЗЗ 5%, шт.", "Расход ОЗЗ 5%, руб.", "Заказы ОЗЗ ст., шт.", "Расход ОЗЗ ст., руб."], "on": "Артикул"},
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
    columns_to_fillna = ['Остатки FBO', 'Остатки FBS', 'Сумма заказов', 'Кол-во заказов']
    df_client_svod[columns_to_fillna] = (
        df_client_svod[columns_to_fillna]
        .fillna(0)
        .infer_objects(copy=False)
    )
    # Переводим некоторые колонки в integer
    columns_to_int = ['Остатки FBO', 'Остатки FBS', 'Кол-во заказов']
    df_client_svod[columns_to_int] = (
        df_client_svod[columns_to_int]
        .astype(int)
    )
    # Суммируем остатки в один столбец
    df_client_svod['Остаток'] = df_client_svod['Остатки FBO'] + df_client_svod['Остатки FBS']
    # Удаляем колонки с остатками FBO и FBS
    df_client_svod = df_client_svod.drop(columns=['Остатки FBO', 'Остатки FBS'], errors='ignore')
    # Средняя цена продажи
    df_client_svod['Средняя цена продажи'] = np.where(
        df_client_svod['Кол-во заказов'] > 0,
        df_client_svod['Сумма заказов'] / df_client_svod['Кол-во заказов'],
        0
    )
    # ДРР %
    df_client_svod['ДРР'] = (
        df_client_svod[['РК', 'Расход ОЗЗ 5%, руб.', 'Расход ОЗЗ ст., руб.']].sum(axis=1)
        / df_client_svod['Сумма заказов'].replace(0, np.nan) * 100
    )
    df_client_svod['ДРР'] = df_client_svod['ДРР'].replace([-np.inf, np.inf], np.nan)

    # Процент отмен считаем только для товаров, у которых есть расходы на РК
    # (РК, ОЗЗ 5%, ОЗЗ ст.). У остальных оставляем пропуск (NaN).
    has_rk_costs = (
        df_client_svod['РК'].fillna(0)
        + df_client_svod['Расход ОЗЗ 5%, руб.'].fillna(0)
        + df_client_svod['Расход ОЗЗ ст., руб.'].fillna(0)
    ) != 0
    df_client_svod['Процент отмен'] = (
        df_client_svod['Процент отмен']
        .where(has_rk_costs)
    )

    # Процент выкупа клиента для расчёта фактического ДРР (константа для всех строк).
    # Если для клиента не задан — оставляем NaN.
    df_client_svod['Процент выкупа'] = RK_BUYOUT_PERCENT.get(client_name, np.nan)
    # Столбец фактического ДРР заполняется формулой уже при записи в Excel
    df_client_svod['Фактический ДРР'] = np.nan

    # Добавляем столбец с комментарием
    # 1. Собираем условия для комментариев: список (текст_комментария, условие-строка)
    # Порог процента отмен — константа для клиента, по умолчанию 20%.
    cancel_threshold = CANCEL_PERCENT_THRESHOLD.get(client_name, 0.2)
    comment_conditions = [
        (
            'Большой процент отмен, влияет на ДРР',
            (df_client_svod['Процент отмен'] > cancel_threshold) & has_rk_costs
        ),
    ]
    # Условие: нет затрат РК, но достаточно заказов для рассмотрения продвижения.
    # Если для клиента не задан порог заказов — условие не добавляем.
    orders_threshold = PROMOTION_ORDERS_THRESHOLD.get(client_name)
    if orders_threshold is not None:
        comment_conditions.append((
            'Нет затрат РК, можно рассмотреть включение в продвижение',
            ~has_rk_costs
            & (df_client_svod['Кол-во заказов'] >= orders_threshold)
        ))

# Условие: остатков менее чем на 2 недели (z = Остаток - 2 × Кол-во заказов ≤ 0),
    # только для товаров, у которых были заказы.
    comment_conditions.append((
        'Остатков менее чем на 2 недели',
        ((df_client_svod['Остаток'] - 2 * df_client_svod['Кол-во заказов']) <= 0)
        & (df_client_svod['Кол-во заказов'] > 0)
    ))

    # 2. Инициализируем столбец комментария
    df_client_svod['Комментарий'] = ''

    # 3. Для каждого комментария находим строки, где условие выполнено, и записываем его
    for comment_text, condition in comment_conditions:
        matched_rows = df_client_svod.index[condition]
        df_client_svod.loc[matched_rows, 'Комментарий'] = (
            df_client_svod.loc[matched_rows, 'Комментарий'] + comment_text + '\n'
        )

    # 4. Убираем завершающий перенос строки
    df_client_svod['Комментарий'] = df_client_svod['Комментарий'].str.rstrip('\n')

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
        'Заказы ОЗЗ ст., шт.', 'Расход ОЗЗ ст., руб.',
        'Кол-во заказов', 'Сумма заказов',
        'Процент выкупа',
        'ДРР',
        'Фактический ДРР',
        f'Остаток {date_report_formatted}',
        'Комментарий',
        'Процент отмен',
        # 'Тип кампании',
        # 'Отзывы',
        # 'Текущая цена с учетом скидки (из отчета по товарам)',
        # 'ДРР % по размеру',

    ]
    # Столбцы с результатами ИИ добавляются только при включённом флаге
    if enable_ai_recommendations:
        column_order.append('Комментарий ИИ')
        column_order.append('Проблемы ИИ')
    df_client_excel_svod = df_client_excel_svod[column_order]

    # Сортировка
    df_client_excel_svod = df_client_excel_svod.sort_values(by=['Сумма заказов'], ascending=[False], ignore_index=True)

    # Заменяем все значения 0 и 0.0 на NaN
    df_client_excel_svod = df_client_excel_svod.replace({0: np.nan, 0.0: np.nan})

    return df_client_excel_svod


# Функция форматирования столбца с номером РК
def format_campaigns_columns(ws, df=None, column_names=None):
    """
    Форматирование столбцов с номером кампании:
    - каждой кампании назначается цвет из палитры (циклично)
    - строки, где заканчивается блок кампании, получают жирную нижнюю границу

    Параметры
    ---------
    ws : openpyxl.worksheet.worksheet.Worksheet
        Лист Excel, на котором нужно применить форматирование.
    df : pandas.DataFrame | None
        Если передан, имена колонок берём из df.columns.
        Если None — имена колонок берём из заголовков листа (первая строка).
    column_names : list[str] | None
        Список возможных названий колонок кампании (точное совпадение, без учёта регистра).
        По умолчанию ['Номер кампании', 'ID Кампании'].
    """
    if column_names is None:
        column_names = ['Номер кампании', 'ID Кампании']

    # 1. Палитра из 15 неярких расцветок
    color_palette = [
        "FFB3C6", "FFE6E6", "B6E2D3", "B7D3DF", "DEEAF6",
        "F7E1D7", "F5F8C4", "D4E6A5", "A2D5AB", "F9D6C1",
        "F5B5FC", "D9D9D9", "ABC4FF", "C6F1E7", "EEE3CB"
    ]

    # 2. Определяем имена колонок кампании
    if df is not None:
        # берём названия колонок из df
        df_campaign_cols = [
            col for col in df.columns
            if any(col.strip().lower() == name.strip().lower() for name in column_names)
        ]
    else:
        # df не передан — берём заголовки с первой строки листа
        header_cells = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        df_campaign_cols = [
            header for header in header_cells if isinstance(header, str) and
            any(header.strip().lower() == name.strip().lower() for name in column_names)
        ]

    if not df_campaign_cols:
        return  # Нет нужных колонок — ничего не делаем

    # 3. Получаем все уникальные значения кампании
    unique_campaigns = []

    if df is not None:
        # Используем данные из df
        for col in df_campaign_cols:
            unique_campaigns.extend(df[col].dropna().unique())
    else:
        # Берём данные с листа Excel (по колонкам, найденным по имени)
        header_cells = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        # Индексы столбцов кампании на листе
        campaign_col_idxs_for_values = [
            idx for idx, header in enumerate(header_cells)
            if header in df_campaign_cols
        ]
        for idx in campaign_col_idxs_for_values:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                    min_col=1, max_col=ws.max_column):
                cell = row[idx]
                if cell.value is not None:
                    unique_campaigns.append(cell.value)

    # Сохраняем порядок появления
    unique_campaigns = list(dict.fromkeys(unique_campaigns))
    if not unique_campaigns:
        return

    # 4. Цвета кампаниям (циклично)
    campaign_to_color = {
        val: color_palette[i % len(color_palette)]
        for i, val in enumerate(unique_campaigns)
    }

    # 5. Индексы колонок кампании в Excel (по строке заголовков листа)
    header_cells = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    campaign_col_idxs = [
        idx for idx, header in enumerate(header_cells) if header in df_campaign_cols
    ]
    if not campaign_col_idxs:
        return

    # 6. Форматируем ячейки
    for idx in campaign_col_idxs:
        for excel_row_idx, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column),
            start=2
        ):
            cell = row[idx]
            camp = cell.value

            # Заливаем цвет по кампании
            if camp in campaign_to_color:
                cell.fill = PatternFill(
                    start_color=campaign_to_color[camp],
                    end_color=campaign_to_color[camp],
                    fill_type='solid'
                )

            # Проверяем переход к новой кампании
            is_last_row = (excel_row_idx == ws.max_row)
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
    BROWN = "C4BD97"  # 196, 189, 151

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
    brown_headers = {
        "Расход ОЗЗ ст., руб.",
        "Заказы ОЗЗ ст., шт.",
        "Заказы по размерам ОЗЗ ст., шт.",
        "Расход по размерам ОЗЗ ст., руб.",
    }

    header_color_map = {}
    for h in blue1_headers:
        header_color_map[h] = BLUE_1
    for h in orange_headers:
        header_color_map[h] = ORANGE
    for h in blue2_headers:
        header_color_map[h] = BLUE_2
    for h in brown_headers:
        header_color_map[h] = BROWN

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

    # --- 0. Зашиваем ДРР и Фактический ДРР формулами ---
    # ДРР = СУММ(расходы) / Сумма заказов.
    # Фактический ДРР = СУММ(расходы) / (Процент выкупа × Кол-во заказов × Средняя цена продажи).
    # Числитель (РК + ОЗЗ 5% + ОЗЗ ст.) — по аналогии с обычным ДРР.
    if 'ДРР' in df_client_excel_svod.columns or 'Фактический ДРР' in df_client_excel_svod.columns:
        cost_cols = ['РК', 'Расход ОЗЗ 5%, руб.', 'Расход ОЗЗ ст., руб.']
        cost_letters = [
            get_column_letter(list(df_client_excel_svod.columns).index(c) + 1)
            for c in cost_cols if c in df_client_excel_svod.columns
        ]

        # Обычный ДРР = СУММ(расходы) / Сумма заказов
        if 'ДРР' in df_client_excel_svod.columns:
            sum_orders_letter = get_column_letter(list(df_client_excel_svod.columns).index('Сумма заказов') + 1)
            drr_col_idx = list(df_client_excel_svod.columns).index('ДРР') + 1
            for row in range(2, ws.max_row + 1):
                numerator = "SUM(" + ",".join(f"{l}{row}" for l in cost_letters) + ")"
                ws.cell(row=row, column=drr_col_idx).value = (
                    f'=IFERROR(IF({numerator}=0,"",{numerator}/{sum_orders_letter}{row}),"")'
                )

        # Фактический ДРР = СУММ(расходы) / (Процент выкупа × Кол-во заказов × Средняя цена продажи)
        if 'Фактический ДРР' in df_client_excel_svod.columns:
            buyout_letter = get_column_letter(list(df_client_excel_svod.columns).index('Процент выкупа') + 1)
            orders_letter = get_column_letter(list(df_client_excel_svod.columns).index('Кол-во заказов') + 1)
            price_letter = get_column_letter(list(df_client_excel_svod.columns).index('Средняя цена продажи') + 1)
            fdrr_col_idx = list(df_client_excel_svod.columns).index('Фактический ДРР') + 1
            for row in range(2, ws.max_row + 1):
                numerator = "SUM(" + ",".join(f"{l}{row}" for l in cost_letters) + ")"
                ws.cell(row=row, column=fdrr_col_idx).value = (
                    f'=IFERROR(IF({numerator}=0,"",'
                    f'({numerator})/'
                    f'({buyout_letter}{row}*{orders_letter}{row}*{price_letter}{row})),"")'
                )

    # --- 1. Процентный формат ---
    for col_name in percent_cols:
        if col_name not in df_client_excel_svod.columns:
            continue
        if col_name == 'ДРР':
            continue  # теперь это формула (доля) — формат задаётся отдельно
        col_idx = list(df_client_excel_svod.columns).index(col_name) + 1  # Excel: 1-based
        for row in range(2, ws.max_row + 1):  # пропускаем заголовок
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                cell.value = cell.value * 0.01  # делим на 100 для корректного отображения %
            cell.number_format = '0.0%'

    # Процентный формат для колонок-формул (значение — доля) и "Процент выкупа"
    for col_name in ['ДРР', 'Фактический ДРР', 'Процент выкупа']:
        if col_name not in df_client_excel_svod.columns:
            continue
        col_idx = list(df_client_excel_svod.columns).index(col_name) + 1
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col_idx).number_format = '0.0%'

    # --- 2. Условное форматирование (градиент) ---
    gradient_cols = list(percent_cols)
    if 'Фактический ДРР' in df_client_excel_svod.columns:
        gradient_cols.append('Фактический ДРР')
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

    # --- 3. Закрепление области до столбца "РК" (включительно) ---
    if "РК" in df_client_excel_svod.columns:
        name_idx = list(df_client_excel_svod.columns).index("РК") + 1  # 1-based индекс "РК"
        # нужно зафиксировать всё, что левее следующей колонки → +1
        freeze_col_idx = name_idx + 1
        freeze_col_letter = get_column_letter(freeze_col_idx)
        # Закрепляем область выше 2-й строки и левее freeze_col_idx:
        # строка 1 + все столбцы до и включая "РК"
        ws.freeze_panes = f"{freeze_col_letter}2"

# --- 4. Подсветка "Кол-во заказов" по проценту отмен (выше порога) ---
    # Формула ссылается на скрытый столбец "Процент отмен"
    if "Процент отмен" in df_client_excel_svod.columns and "Кол-во заказов" in df_client_excel_svod.columns:
        canc_col_idx = list(df_client_excel_svod.columns).index("Процент отмен") + 1
        orders_col_idx = list(df_client_excel_svod.columns).index("Кол-во заказов") + 1
        canc_col_letter = get_column_letter(canc_col_idx)
        orders_col_letter = get_column_letter(orders_col_idx)

        # Порог процента отмен — константа для клиента, по умолчанию 20%
        cancel_threshold = CANCEL_PERCENT_THRESHOLD.get(client_name, 0.2)

        # Условное форматирование оранжевой заливкой для строк с процентом отмен выше порога
        cell_range = f"{orders_col_letter}2:{orders_col_letter}{ws.max_row}"
        rule = FormulaRule(
            formula=[f"${canc_col_letter}2>{cancel_threshold}"],
            stopIfTrue=False,
            fill=PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        )
        ws.conditional_formatting.add(cell_range, rule)

        # Процентный формат для скрытого столбца с процентом отмен
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=canc_col_idx).number_format = '0.0%'

        # Скрываем столбец с процентом отмен
        ws.column_dimensions[canc_col_letter].hidden = True

    # --- 5. Скрываем столбцы фактического ДРР, если для клиента не задан Процент выкупа ---
    if 'Процент выкупа' in df_client_excel_svod.columns:
        has_buyout = df_client_excel_svod['Процент выкупа'].notna().any()
        if not has_buyout:
            buyout_col_idx = list(df_client_excel_svod.columns).index('Процент выкупа') + 1
            buyout_col_letter = get_column_letter(buyout_col_idx)
            ws.column_dimensions[buyout_col_letter].hidden = True
            if 'Фактический ДРР' in df_client_excel_svod.columns:
                drr_col_idx = list(df_client_excel_svod.columns).index('Фактический ДРР') + 1
                drr_col_letter = get_column_letter(drr_col_idx)
                ws.column_dimensions[drr_col_letter].hidden = True

    # --- 6. Подсветка "Остаток" бледно-красным при недостаточных остатках ---
    # z = Остаток - 2 × Кол-во заказов; если z ≤ 0 и были заказы — остатков менее чем на 2 недели
    stock_cols = [c for c in df_client_excel_svod.columns if c.startswith('Остаток')]
    if stock_cols and 'Кол-во заказов' in df_client_excel_svod.columns:
        stock_col_name = stock_cols[0]
        stock_col_idx = list(df_client_excel_svod.columns).index(stock_col_name) + 1
        orders_col_idx = list(df_client_excel_svod.columns).index('Кол-во заказов') + 1
        stock_col_letter = get_column_letter(stock_col_idx)
        orders_col_letter = get_column_letter(orders_col_idx)

        cell_range = f"{stock_col_letter}2:{stock_col_letter}{ws.max_row}"
        rule = FormulaRule(
            formula=[f"AND(${stock_col_letter}2-2*${orders_col_letter}2<=0,${orders_col_letter}2>0)"],
            stopIfTrue=False,
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        )
        ws.conditional_formatting.add(cell_range, rule)

    # --- 7. Ширина столбцов и выравнивание данных по маппингу ---
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
        reverse_gradient_cols=None,
):
    """
    Форматирование листа "Отчет РК {client_name}":
    - ширина столбцов и выравнивание по CAMPAIGNS_REPORT_COL_FORMAT
    - процентный формат для percent_cols
    - градиентное условное форматирование:
        * для столбцов из reverse_gradient_cols: min -> красный, max -> зелёный
        * для остальных столбцов, попавших в percent_cols или reverse_gradient_cols:
          min -> зелёный, max -> красный
      (т.е. градиент ставится, если колонка есть хотя бы в одном из списков,
       а тип градиента определяется наличием в reverse_gradient_cols)
    - закрепление области до столбца "Цена товара, руб." (включительно)
    """
    if sheet_name not in wb.sheetnames:
        return

    if percent_cols is None:
        percent_cols = []

    if reverse_gradient_cols is None:
        reverse_gradient_cols = ['CTR (%)']

    percent_set = set(percent_cols)
    reverse_set = set(reverse_gradient_cols)
    # Все колонки, для которых нужен градиент — объединение двух множеств
    gradient_cols = list(percent_set | reverse_set)

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

    # --- 2. Процентный формат для percent_cols ---
    for col_name in percent_set:
        if df_campaigns_report is None or col_name not in df_campaigns_report.columns:
            continue
        col_idx = list(df_campaigns_report.columns).index(col_name) + 1  # Excel: 1-based
        for row in range(2, ws.max_row + 1):  # пропускаем заголовок
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                cell.value = cell.value * 0.01
            cell.number_format = '0.0%'

    # --- 3. Градиентное условное форматирование для gradient_cols ---
    for col_name in gradient_cols:
        if df_campaigns_report is None or col_name not in df_campaigns_report.columns:
            continue

        col_idx = list(df_campaigns_report.columns).index(col_name) + 1
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        cell_range = f"{col_letter}2:{col_letter}{ws.max_row}"

        # Приоритет: если колонка в reverse_gradient_cols → обратный градиент
        if col_name in reverse_set:
            rule = ColorScaleRule(
                start_type='min', start_color='E06967',              # красный (минимум)
                mid_type='percentile', mid_value=50, mid_color='FFEB84',  # жёлтый
                end_type='max', end_color='63BE7B',                  # зелёный (максимум)
            )
        else:
            rule = ColorScaleRule(
                start_type='min', start_color='63BE7B',              # зелёный
                mid_type='percentile', mid_value=50, mid_color='FFEB84',  # жёлтый
                end_type='max', end_color='E06967',                  # красный
            )

        ws.conditional_formatting.add(cell_range, rule)

    # --- 4. Закрепление области до столбца (включительно) ---
    rk_col_name = "Тип товара"
    rk_col_idx = None
    if df_campaigns_report is not None and rk_col_name in df_campaigns_report.columns:
        rk_col_idx = list(df_campaigns_report.columns).index(rk_col_name) + 1  # 1-based
    elif rk_col_name in headers:
        rk_col_idx = headers.index(rk_col_name) + 1

    if rk_col_idx is not None:
        freeze_col_idx = rk_col_idx + 1
        freeze_col_letter = get_column_letter(freeze_col_idx)
        ws.freeze_panes = f"{freeze_col_letter}2"


def format_models_sheet(ws, df_models=None, column_names=None):
    """
    Форматирование листа отчёта по моделям:
    - закрепление области до столбца 'Тип товара' (включительно).

    Параметры
    ---------
    ws : openpyxl.worksheet.worksheet.Worksheet
        Лист Excel, на котором нужно применить форматирование.
    df_models : pandas.DataFrame | None
        Если передан, имена колонок берём из df_models.columns для поиска 'Тип товара'.
        Если None — имена колонок берём из заголовков листа (первая строка).
    column_names : list[str] | None
        Зарезервировано для единообразия с format_campaigns_columns, сейчас не используется.
    """
    # Определяем индекс столбца "Тип товара"
    type_col_idx = None

    # 1. Если передан df_models — ищем в нём
    if df_models is not None and "Тип товара" in df_models.columns:
        type_col_idx = list(df_models.columns).index("Тип товара") + 1  # 1-based

    # 2. Если не нашли или df_models нет — ищем по заголовкам листа
    if type_col_idx is None:
        header_row = next(ws.iter_rows(min_row=1, max_row=1))
        headers = [cell.value for cell in header_row]
        if "Тип товара" in headers:
            type_col_idx = headers.index("Тип товара") + 1  # 1-based

    # 3. Если столбец найден — закрепляем до него включительно
    if type_col_idx is not None:
        # нужно зафиксировать всё, что левее следующей колонки → +1
        freeze_col_idx = type_col_idx + 1
        freeze_col_letter = get_column_letter(freeze_col_idx)
        # Закрепляем область: строка 1 + все столбцы до и включая "Тип товара"
        ws.freeze_panes = f"{freeze_col_letter}2"


# Функция скрытия столбца с заданным заголовком
def hide_column_by_header(ws, header):
    """Скрывает на листе столбец с указанным заголовком (если он есть)."""
    header_cells = next(ws.iter_rows(min_row=1, max_row=1))
    for cell in header_cells:
        if cell.value == header:
            ws.column_dimensions[get_column_letter(cell.column)].hidden = True
            break


# Функция сохранения и форматирования excel
def save_and_format_excel(
        df_client_excel_svod,
        df_campaigns_report,
        df_payment_orders_report,
        df_models,
        client_rk_svod_dir,
        date_report,
        df_dates,
        ai_data_issues=None
    ):
    # Формируем директорию для сохранения итогового отчета
    path_report = client_rk_svod_dir
    path_report.mkdir(parents=True, exist_ok=True)

    # Формируем даты для названия листа и файла
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

    # Формируем имя файла отчета по шаблону:
    # "Отчет РК {client_name} {date_start}-{date_end}.xlsx"
    name_report = f"Отчет РК {client_name} {date_start_excel}-{date_end_excel}.xlsx"
    file_name_report = path_report / name_report

    # Колонки с процентами в сводном файле
    percent_cols = [
        # 'CTR (%)',
        'ДРР',
        'ДРР по РК',
        'ДРР артикула',
    ]
    reverse_gradient_cols = ['CTR (%)',]
    sheet_name_client_svod = f"{client_name} {date_start_excel}-{date_end_excel}"

    # Сохраняем книгу
    with pd.ExcelWriter(file_name_report) as w:
        # Сводный отчет
        df_client_excel_svod.to_excel(w, index=False, sheet_name=sheet_name_client_svod)
        # Основной отчет РК
        df_campaigns_report.to_excel(w, index=False, sheet_name=f'Отчет РК {client_name}')
        # Отчет по моделям товаров
        df_models.to_excel(w, index=False, sheet_name='Модели товаров')
        # Отчет оплаты за заказ
        df_payment_orders_report.to_excel(w, index=False, sheet_name='ОЗЗ 5%')
        # Глобальные проблемы анализа от ИИ-помощника (если есть)
        if ai_data_issues:
            df_ai_issues = pd.DataFrame({'Проблемы ИИ': ai_data_issues})
            df_ai_issues.to_excel(w, index=False, sheet_name='Проблемы ИИ')

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
        percent_cols=percent_cols,
        reverse_gradient_cols=reverse_gradient_cols
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
        format_models_sheet(ws_models, df_models)

    # 5) Форматирование сводного листа клиента (проценты, градиент)
    format_client_summary_sheet(
        wb,
        sheet_name=sheet_name_client_svod,
        df_client_excel_svod=df_client_excel_svod,
        percent_cols=percent_cols,
    )

    # 6) Скрываем служебные столбцы и лист с проблемами ИИ
    if sheet_name_client_svod in wb.sheetnames:
        hide_column_by_header(wb[sheet_name_client_svod], 'Проблемы ИИ')
    if sheet_campaigns in wb.sheetnames:
        hide_column_by_header(wb[sheet_campaigns], 'Проблемы ИИ')
    if 'Проблемы ИИ' in wb.sheetnames:
        wb['Проблемы ИИ'].sheet_state = 'hidden'

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
    # date_report = '2026-03-31'  # для теста
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

    # Получаем заказы
    df_orders = get_orders_for_svod(df_dates, headers)
    # Получаем остатки
    df_reminders = get_reminders_by_product(df_products, headers, to_agg=True, split_stocks=True)

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

    # Формируем лист с моделями товаров
    df_models = get_model_list(df_products, df_campaigns_report, df_reminders)

    # Добавляем рекомендации ИИ-помощника, если включено для клиента
    # (отдельный шаг; при ошибке отчёт формируется без них)
    ai_data_issues = []
    # %%
    if enable_ai_recommendations:
        df_client_svod, df_campaigns_report_processed, ai_data_issues = add_ai_recommendations(
            df_client_svod,
            df_campaigns_report_processed
        )
    # %%

    # Формируем таблицу для записи в excel (последний шаг, подбивает все предыдущие)
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
        df_dates,
        ai_data_issues=ai_data_issues
    )
    logger.info(f"Done creating rk client svod for client {client_name}")
    logger.info(f"RK client svod saved in: {file_name_report}")

# %%
