
# %% Определение всех функций
import argparse
import os
import re
import requests
import json
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.options.mode.chained_assignment = None  # default='warn'
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

# Некоторые доп. функции
from generic_functions import move_columns
from ozon.scripts.ozon_generic_functions import get_common_article
# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_finance
# Некоторые константы
from ozon.scripts.constants import (
    ozon_seller_api_url,
    default_path_download_finance_report,
    catalog_finace_svod_columns,
    client_id_performance, client_name,
    client_secret_performance, headers,
    marketplace_dir_path
    )
# Функция выгрузки товаров в поставках
from ozon.scripts.uploadDataFromOzon import (
    get_supply_order_list,
    get_supply_orders_details,
    get_supply_bundle_items,
    get_ozon_product as upload_ozon_product
)
# Функция выгрузки данных о рекламных кампаниях
from ozon.scripts.uploadDataFromOzonPerformance import (
    upload_data_performane,
    getCompanyStatistics
)
# Функция выгрузки отчета о релизации товаров
from ozon.scripts.get_realization import get_products_realization_report

parser = argparse.ArgumentParser(
    description="Create finance report (Ozon)",
    usage=(
        "python.exe -m ozon.scripts.calc_svod_finance_report_new [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD] "
        "[-c N | --client N] "
        "[--file PATH_TO_REPORT.xlsx]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new 2024-11-13 2024-11-20\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new 2024-11-13 2024-11-20 -c 3\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new --file \"Report_Озон_2024-11-13_2024-11-20.xlsx\"\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new --start_date 2024-11-13 --end_date 2024-11-20 --file \"Report_2024-11-13_2024-11-20.xlsx\"\n\n"
        "Если путь или имя файла содержит пробелы, обязательно заключайте его в кавычки!\n"
        "(например: --file \"file name with spaces.xlsx\")"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


# Функция генерации диапазона дат
def generate_date_range(date_start, date_end):
    # Если не заданы аргументы командной строки, то генерируем диапазон за предыдущий месяц
    if date_start is None and date_end is None:
        today = datetime.today()
        # Первый день текущего месяца
        first_day_this_month = today.replace(day=1)
        # Последний день предыдущего месяца
        last_day_prev_month = first_day_this_month - timedelta(days=1)
        # Первый день предыдущего месяца
        first_day_prev_month = last_day_prev_month.replace(day=1)
        # Форматируем даты в строку YYYY-MM-DD
        date_start = first_day_prev_month.strftime("%Y-%m-%d")
        date_end = last_day_prev_month.strftime("%Y-%m-%d")
        return date_start, date_end
    else:
        return date_start, date_end

# Функция создания директории с текущим отчетом
def create_current_report_dir(client_name, date_start, date_end):
    # Переводим дату в нужный формат
    date_folder = (
        pd.to_datetime(date_start)
        .strftime('%Y_%m')
    )
    # Формируем путь к директории
    finance_reports_dir = (
        marketplace_dir_path
        / 'Clients'
        / client_name
        / 'FinanceReports'
        / date_folder
    )
    # Создаем папку, если её нет
    finance_reports_dir.mkdir(parents=True, exist_ok=True)

    return finance_reports_dir, date_folder


# Функция поиска отчета с файлом в директории
def rename_finance_report_file(finance_report_dir, client_name, date_start, date_end):
    try:
        start_dot = datetime.strptime(date_start, "%Y-%m-%d").strftime("%d.%m.%Y")
        end_dot = datetime.strptime(date_end, "%Y-%m-%d").strftime("%d.%m.%Y")
        start_us = datetime.strptime(date_start, "%Y-%m-%d").strftime("%d_%m_%Y")
        end_us = datetime.strptime(date_end, "%Y-%m-%d").strftime("%d_%m_%Y")
    except Exception as e:
        return f"Ошибка формата даты: {e}"

    reports_dir = Path(finance_report_dir)
    final_name = f"{client_name}_Отчет_по_начислениям_{date_start}_{date_end}.xlsx"
    final_path = reports_dir / final_name

    if final_path.exists():
        return final_path

    # Паттерн для точечного формата с тире
    pattern1 = re.compile(
        rf"отчет[ _]по[ _]начислениям[ _]{start_dot}-{end_dot}.*\.(xlsx|xls)$",
        re.I
    )
    # Исправленный паттерн для подчеркиваний (более одного _ или пробела)
    pattern2 = re.compile(
        rf"отчет[ _]по[ _]начислениям[ _]{start_us}[ _]+{end_us}.*\.(xlsx|xls)$",
        re.I
    )

    matched_files = []

    for f in reports_dir.iterdir():
        if f.is_file() and (pattern1.match(f.name) or pattern2.match(f.name)):
            matched_files.append((f, f.stat().st_mtime))

    if not matched_files:
        return "Файл отчета по заданным датам не найден."

    matched_files.sort(key=lambda x: x[1], reverse=True)
    chosen_file = matched_files[0][0]
    chosen_file.rename(final_path)
    return final_path

# Функция парсинга аргументов командной строки пути к файлу по акциям
def get_ozon_report_file_path(arg_path: str | None, default_dir: Path, default_name: str | None) -> Path | None:
    """
    Возвращает итоговый Path к файлу акций по логике:
      - Если передан arg_path (через --file), возвращает абсолютный путь или default_dir/относительный путь.
      - Если не передан arg_path и default_name не задан, возвращает None.
      - Если не передан arg_path, но задан default_name, возвращает default_dir / default_name.
    """
    if arg_path:  # Аргумент командной строки был передан
        file_path = Path(arg_path)
        if file_path.is_absolute():
            # Передан абсолютный путь — используем как есть
            return file_path
        else:
            # Передано имя или относительный путь — склеиваем с default_dir
            return default_dir / file_path
    else:
        # Путь не задан — если имя не указано, возвращаем None
        if not default_name:
            return None
        return default_dir / default_name

# Функция чтения файла с отчетом
def read_finance_report_file(finance_report_filepath):

    df_finance_report = pd.read_excel(
        finance_report_filepath,
        sheet_name='Начисления',
        skiprows=1
    )

    df_finance_report = df_finance_report.rename(columns={
        'Ozon SKU': 'SKU',
        'Сумма итого, руб.': 'Итого руб'
    })
    # Добавляем уникальный идентификатор строки
    df_finance_report['id'] = np.arange(1, df_finance_report.shape[0] + 1)

    return df_finance_report


# Функция выгрузки Продаж, Возвратов и Баллов Озон
def calc_sales_returns(date_start, date_end):
    # Создание диапазона дат выгрузки отчета о реализации
    dt_start = pd.to_datetime(date_start, format='%Y-%m-%d')
    dt_end = pd.to_datetime(date_end, format='%Y-%m-%d')
    date_range = (pd
                  .date_range(start=dt_start, end=dt_end, freq='MS')
                  .to_frame()
                  .rename(columns={0: 'date'})
                  .reset_index(drop=True)
                #   .assign(
                #       date_end=date_range['date_start'].shift(-1)
                #       )
    )
    df_realization_report_total = pd.DataFrame()
    for i in range(date_range.shape[0]):
        # Определение месяца и года, за который будет выгружен отчет о реализации
        month = date_range['date'][i].month
        year = date_range['date'][i].year

        logger.info(f"Uploading realization report year: {year} month: {month}")

        # Выгрузка отчета о релизации товаров
        df_realization_report = get_products_realization_report(headers, year, month, to_save=False)
        # Колонка с месяцем и годом для информации
        df_realization_report = df_realization_report.assign(
            month=month,
            year=year
        )
        # Расчет итоговых столбцов по продажам и возвратам
        df_realization_report['Продажи руб'] = df_realization_report[['Сумма_продажи',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_звёзды_продажи',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_продажи',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_апвз_продажи']] \
                                                                        .sum(axis=1)
        df_realization_report['Продажи шт'] = df_realization_report['Количество_продажи']
        df_realization_report['Возвраты руб'] = -abs(df_realization_report[['Сумма_возвраты',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_звёзды_возвраты',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_возвраты',
                                                                    'Выплаты_по_механикам_лояльности_партнёров_апвз_возвраты']] \
                                                                        .sum(axis=1))
        df_realization_report['Возвраты шт'] = df_realization_report['Количество_возвраты']
        df_realization_report['Продажи - Возвраты (руб)'] = df_realization_report['Продажи руб'] - abs(df_realization_report['Возвраты руб'])
        df_realization_report['Продажи - Возвраты (шт)'] = df_realization_report['Продажи шт'] - df_realization_report['Возвраты шт']

        # Расчет итогового столбца по баллам озон
        df_realization_report['Баллы Озон Итог'] = df_realization_report['Баллы_за_скидки_продажи'] - df_realization_report['Баллы_за_скидки_возвраты']

        # Выбор нужных колонок
        df_realization_report = df_realization_report.loc[:, ['SKU', 'Артикул_продавца', 'Наименование_товара', 'Продажи шт', 'Возвраты шт', 'Продажи - Возвраты (шт)', 'Продажи руб', 'Сумма_продажи',  'Возвраты руб', 'Сумма_возвраты',  'Продажи - Возвраты (руб)', 'Баллы Озон Итог']]
        # Объединение с предыдущим проходом цикла
        df_realization_report_total = pd.concat([df_realization_report, df_realization_report_total])

    # Расчет итоговых сумм за все периоды
    df_realization_report_stats = (
        df_realization_report_total
        .groupby(['SKU', 'Артикул_продавца', 'Наименование_товара'])
        .sum()
        .reset_index()
        .rename(columns={
            'Артикул_продавца': 'Артикул продавца',
            'Наименование_товара': 'Наименование товара'
        })
    )

    # Расчет общего количества продаж, возвратов и баллов озон
    df_sales_returns_by_operation = (
        df_realization_report_stats
        .loc[:, ~df_realization_report_stats.columns.isin(['SKU', 'Артикул продавца', 'Наименование товара'])]
        .sum()
        .reset_index()
        .rename(columns={
            'index': 'Тип начисления',
            0: 'Итого'
        })
    )

    # Помещаем результат в словарь
    result_sales_returns = {
        'df_realization_report_stats': df_realization_report_stats,
        'df_sales_returns_by_operation': df_sales_returns_by_operation
    }

    return result_sales_returns


# Функция расчета расходов, привязанных к SKU
def calc_sku_costs(df_finance_report, groupby_cols = ['SKU', 'Артикул', 'Название товара']):
    # Выбираем затраты, которые идут с привязкой к SKU
    df_finance_report_sku = df_finance_report.loc[~df_finance_report['SKU'].isna(), :]
    # Список номеров строк с расходами
    sku_costs_ids = df_finance_report_sku['id'].to_list()
    # Колонки, по которым делаем группировку
    groupby_cols = groupby_cols
    # Переводим артикул в строку
    df_finance_report_sku['Артикул'] = df_finance_report_sku['Артикул'].astype(str)


    # Типы начислений, относящиеся к продажам и возвратам
    sales_and_returns_ids = df_finance_report_sku.loc[df_finance_report_sku['Тип начисления'].isin(['Выручка', 'Возврат выручки']), 'id'].to_list()
    # Делаем выборку только по типам начислений с продажами и возвратами
    df_sales_returns = df_finance_report_sku.loc[df_finance_report_sku['id'].isin(sales_and_returns_ids), :]
    # Заменяем значения на "Продажа" и "Возврат"
    df_sales_returns['Тип начисления'] = (
        df_sales_returns['Тип начисления']
        .replace({
            'Выручка': 'Продажи',
            'Возврат выручки': 'Возвраты'
        })
    )
    # Считаем продажи и возвраты в штуках
    df_sku_costs_sales_returns_stats = (
        pd.pivot_table(
            df_sales_returns,
            values=['Количество', 'Итого руб'],
            index=groupby_cols,
            columns='Тип начисления',
            aggfunc='sum',
            fill_value=0
        )
    )
    # Избавление от MultiIndex и добавление префиксов
    df_sku_costs_sales_returns_stats.columns = [
        f"{col[1]} шт" if col[0] == 'Количество' else f"{col[1]} руб"
        for col in df_sku_costs_sales_returns_stats.columns
    ]
    # Получаем столбцы из Index
    # df_sku_costs_sales_returns_stats = df_sku_costs_sales_returns_stats.reset_index()

    # # Расчет возвратов и продаж в штуках и рублях
    # df_products_grouped = df_finance_report_sku.groupby(by=groupby_cols, dropna=False)
    # df_with_sku_stats = pd.DataFrame()
    # df_with_sku_stats['tmp_col'] = df_products_grouped.size()
    # # Продажи и возвраты считаем по столбцу "Выручка"
    # df_with_sku_stats['Продажи шт'] = df_finance_report_sku.loc[df_finance_report_sku['Выручка'] > 0, :] \
    #     .groupby(by=groupby_cols, dropna=False) \
    #         .size()
    # df_with_sku_stats['Продажи рубли'] = df_finance_report_sku.loc[df_finance_report_sku['Выручка'] > 0, :] \
    #     .groupby(by=groupby_cols, dropna=False)['Выручка'] \
    #         .sum()
    # df_with_sku_stats['Возвраты шт'] = df_finance_report_sku.loc[df_finance_report_sku['Выручка'] < 0, :] \
    #     .groupby(by=groupby_cols, dropna=False) \
    #         .size()
    # df_with_sku_stats['Возвраты рубли'] = df_finance_report_sku.loc[df_finance_report_sku['Выручка'] < 0, :] \
    #     .groupby(by=groupby_cols, dropna=False)['Выручка'] \
    #         .sum()

    # Делаем выборку по типам начислений, не относящимся к продажам и возвратам
    df_sku_other_costs = df_finance_report_sku.loc[~df_finance_report_sku['id'].isin(sales_and_returns_ids), :]
    # Считаем прочие расходы в рублях
    df_sku_other_costs_stats = (
        pd.pivot_table(
            df_sku_other_costs,
            values='Итого руб',
            index=groupby_cols,
            columns='Тип начисления',
            aggfunc='sum',
            fill_value=0
        )
    )
    # Объединяем все расходы в один df
    df_sku_costs_all = (
        pd.concat([
            df_sku_costs_sales_returns_stats,
            df_sku_other_costs_stats
        ])
        #.reset_index()
        # .groupby(level=[0, 1, 2]) # 0 - SKU, 1 - Артикул, 2 - Наименование
        .groupby(level=[1])
        .sum()
        .reset_index()
        .rename(columns={
            'Артикул': 'Артикул продавца',
            'Название товара': 'Наименование товара'
        })
    )

    # Если не было продаж или возвратов, искуственно создаем эти колонки
    for col in ['Продажи шт', 'Продажи руб', 'Возвраты шт', 'Возвраты руб']:
        if col not in df_sku_costs_all.columns:
            df_sku_costs_all[col] = 0

    # Продажи минус возвраты
    df_sku_costs_all['Продажи - Возвраты (шт)'] = df_sku_costs_all['Продажи шт'] - df_sku_costs_all['Возвраты шт']
    df_sku_costs_all['Продажи - Возвраты (руб)'] = df_sku_costs_all['Продажи руб'] - abs(df_sku_costs_all['Возвраты руб'])

    # Перемещаем колонки в начало df
    df_sku_costs_all = move_columns(
        df_sku_costs_all,
        columns_to_move=[
            'Продажи шт', 'Возвраты шт',
            'Продажи руб', 'Возвраты руб',
            'Продажи - Возвраты (шт)',
            'Продажи - Возвраты (руб)'
        ],
        position='Артикул продавца',
        insert_type='after'
    )

    # Считаем итоговые суммы на каждый вид расходов
    df_sku_costs_by_operation = (
        df_sku_costs_all
        .loc[:, ~df_sku_costs_all.columns.isin(['SKU', 'Артикул продавца', 'Наименование товара'])]
        .sum()
        .reset_index()
        .rename(columns={
            'index': 'Тип начисления',
            0: 'Итого'
        })
    )

    # Продажи и возвраты берем из отчета о реализации
    # df_realization_report = result_sales_returns['df_realization_report_stats'].copy()
    # df_with_sku_stats = pd.concat([df_with_sku_stats, df_realization_report])
    # Считаем суммы после concat
    # df_with_sku_stats = df_with_sku_stats.groupby(groupby_cols).sum().reset_index()
    # Удаляем ненужную колонку
    # df_with_sku_stats = df_with_sku_stats.drop(columns=['tmp_col'])
    # Итоговый словарь с результатами
    result_sku_costs = {
        'df_sku_costs_all': df_sku_costs_all,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids
        }

    return result_sku_costs


# Функция проверки наличия расходов по рекламным кампаниям
def check_companies_costs(df_finance_report):
    # Флаг наличия рекламных расходов
    companies_costs_flag = False
    # Выбираем затраты по рекламе, где есть кампании
    # df_companies_costs = df_finance_report.loc[
    #     (df_finance_report['Название товара'].str.contains('Заказ №', na=False) )
    #     | (df_finance_report['ID начисления'].str.match('^\d{8}$', na=False)),
    #     ['Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    # ]
    df_companies_costs = df_finance_report.loc[
        df_finance_report['Тип начисления'].str.contains('Оплата за клик|Продвижение с оплатой за заказ', na=False),
        ['Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]

    if not df_companies_costs.empty:
        companies_costs_flag = True

    return companies_costs_flag


# Функция нахождения id строк с расходами по РК
def get_companies_costs_ids(df_finance_report):
    companies_str_pattern = (
            'Оплата за клик'
            '|Продвижение с оплатой за заказ'
            # '|Продвижение в поиске'
            # '|Трафареты'
    )
    df_companies_costs = df_finance_report.loc[
        df_finance_report['Тип начисления'].str.contains(
            companies_str_pattern,
            na=False
            ),
        ['id', 'Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]
    # id строк с расходами, которые относим к РК
    companies_costs_ids = df_companies_costs['id'].to_list()

    return companies_costs_ids

# Функция нахождения идентификаторов рекламных кампаний
def get_companies_ids(df_finance_report, companies_costs_ids):
    """
    Возвращает df с колонками:
    - Номер кампании
    - Тип начисления
    - Итого_из_отчета

    Для типа "Продвижение с оплатой за заказ" все строки
    агрегируются в одну кампанию с Номер кампании = 'ORDERS'.
    """
    ORDERS_TYPE = "Продвижение с оплатой за заказ"

    # Выбираем только строки с нужными id (рекламные расходы)
    df_companies_costs = df_finance_report.loc[
        df_finance_report['id'].isin(companies_costs_ids),
        ['id', 'Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]

    if df_companies_costs.shape[0] > 0:
        df_companies_costs = df_companies_costs.copy()

        # Базово номер кампании = ID начисления
        df_companies_costs['Номер кампании'] = df_companies_costs['ID начисления'].astype(str)

        # Для "оплаты за заказ" собираем всё в одну виртуальную кампанию 'ORDERS'
        mask_orders = df_companies_costs['Тип начисления'] == ORDERS_TYPE
        df_companies_costs.loc[mask_orders, 'Номер кампании'] = 'ORDERS'

        # Считаем расходы на каждую кампанию и тип начисления
        df_companies_costs_total = (
            df_companies_costs
            .groupby(['Номер кампании', 'Тип начисления'], as_index=False)
            .agg(
                Итого_из_отчета=('Итого руб', 'sum')
            )
        )
    else:
        df_companies_costs_total = pd.DataFrame()

    return df_companies_costs_total

# Функция парсинга файлов со статистикой по кампаниям
def parse_companies_files(current_report_dir, date_report):
    # Директория, где лежат данные по рекламным кампаниям
    companies_stats_dir = current_report_dir / f"{date_report}_Кампании"
    # Директория, куда будем складывать обработанные файлы
    companies_parsed_dir = current_report_dir / f"{date_report}_Кампании_parsed"

    # Префикс файлов отчёта по продвижению с оплатой за заказ
    orders_report_prefix = "search_promo_all_skus_orders_report_"

    # ==============================
    # 1. ПАРСИМ ОБЫЧНЫЕ КАМПАНИИ (ПО ID)
    # ==============================

    filenames_companies = {"path": []}
    companies_stats_extensions = ['.csv', '.xlsx']

    # Осуществляем поиск файлов с заданным расширением
    for ext in companies_stats_extensions:
        for file in companies_stats_dir.glob(f'*{ext}'):
            # Пропускаем файлы отчёта по продвижению с оплатой за заказ —
            # у них нет ID кампании
            if file.name.startswith(orders_report_prefix):
                continue
            filenames_companies['path'].append(file)

    # Если файлов кампаний нет — логируем и переходим к отчёту по заказам
    if not filenames_companies['path']:
        logger.info(
            f"В директории {companies_stats_dir} не найдено файлов статистики кампаний "
            f"(исключая отчёты по оплате за заказ)"
        )
    else:
        # Считывание только имени файла
        filenames_companies['file_name'] = [file.name for file in filenames_companies['path']]
        # Размер файла
        filenames_companies['file_size'] = [file.stat().st_size for file in filenames_companies['path']]
        # Перевод в df
        filenames_companies = pd.DataFrame(filenames_companies)
        # Номер кампании (берём всё до первого подчёркивания в имени файла)
        filenames_companies['company_id'] = filenames_companies['file_name'].str.split('_').str[0]
        # Сортировка по размеру (сначала более крупные)
        filenames_companies = filenames_companies.sort_values(
            by='file_size',
            ascending=False,
            ignore_index=True
        )

        # Обработка csv/xlsx по статистике кампаний в цикле
        for idx, row in filenames_companies.iterrows():
            # Считываем файл (первую строку с описанием пропускаем)
            df_company = pd.read_csv(row['path'], sep=';', decimal=',', skiprows=1)

            # Если в отчёте по кампании есть строки, начинаем их обрабатывать
            if df_company.shape[0] >= 2:
                # Определяем тип отчета (по имени первой колонки)
                if df_company.columns[0] == 'Дата':
                    # Убираем строку "Всего" и "Корректировка"
                    df_company = df_company.loc[
                        ~df_company['Дата'].isin(['Всего', 'Корректировка']),
                        :
                    ]
                    # Ozon ID = SKU в данном отчете
                    df_company = df_company.rename(columns={'Ozon ID': 'sku'})
                    # Переименовываем колонку с расходами для удобства
                    df_company = df_company.rename(columns={'Расход, ₽': 'Расходы'})
                    # Переименовываем колонку с названием товара
                    df_company = df_company.rename(columns={'Наименование': 'Наименование товара'})
                    # Выбираем нужные колонки
                    df_company = df_company.loc[:, ['Дата', 'sku', 'Наименование товара', 'Расходы']]
                    # Сохраняем обработанный файл по ID кампании
                    df_company.to_csv(
                        companies_parsed_dir / f"{row['company_id']}.csv",
                        sep=';',
                        encoding='utf-8-sig',
                        index=False
                    )
                else:
                    # Убираем строку "Всего" и "Корректировка"
                    df_company = df_company.loc[
                        ~df_company['sku'].isin(['Всего', 'Корректировка']),
                        :
                    ]
                    # Переименовываем колонку с расходами для удобства
                    df_company = df_company.rename(columns={'Расход, ₽, с НДС': 'Расходы'})
                    # Переименовываем колонку с названием товара
                    df_company = df_company.rename(columns={'Название товара': 'Наименование товара'})
                    # Выбираем нужные колонки
                    df_company = df_company.loc[:, ['sku', 'Наименование товара', 'Расходы']]
                    # Сохраняем обработанный файл по ID кампании
                    df_company.to_csv(
                        companies_parsed_dir / f"{row['company_id']}.csv",
                        sep=';',
                        encoding='utf-8-sig',
                        index=False
                    )

    # ==============================
    # 2. ПАРСИМ ОТЧЁТ ПО ПРОДВИЖЕНИЮ С ОПЛАТОЙ ЗА ЗАКАЗ
    # ==============================

    orders_files = list(companies_stats_dir.glob(f"{orders_report_prefix}*.csv"))
    if not orders_files:
        logger.info(
            f"Отчёт по продвижению с оплатой за заказ "
            f"({orders_report_prefix}*.csv) в директории {companies_stats_dir} не найден"
        )
        return

    # Предполагаем один отчёт на период; если их несколько — берём первый по имени
    orders_file = sorted(orders_files)[0]
    logger.info(f"Найден отчёт по оплате за заказ: {orders_file.name}, начинаем парсинг")

    # В первой строке файла — описание отчёта, во второй — заголовок, поэтому skiprows=1
    df_orders = pd.read_csv(orders_file, sep=';', decimal=',', skiprows=1)

    if df_orders.empty:
        logger.warning(f"Отчёт по оплате за заказ пуст: {orders_file.name}")
        return

    # На всякий случай убираем строки "Всего" / "Корректировка", если они присутствуют в колонке 'Дата'
    if 'Дата' in df_orders.columns:
        df_orders = df_orders.loc[
            ~df_orders['Дата'].isin(['Всего', 'Корректировка']),
            :
        ]

    # Переименовываем колонки под общий формат:
    # - SKU продвигаемого товара -> 'sku'
    # - Название товара -> 'Наименование товара'
    # - Расход, ₽ -> 'Расходы'
    rename_map = {}
    if 'SKU продвигаемого товара' in df_orders.columns:
        rename_map['SKU продвигаемого товара'] = 'sku'
    elif 'SKU' in df_orders.columns:
        rename_map['SKU'] = 'sku'

    if 'Название товара' in df_orders.columns:
        rename_map['Название товара'] = 'Наименование товара'

    if 'Расход, ₽' in df_orders.columns:
        rename_map['Расход, ₽'] = 'Расходы'

    df_orders = df_orders.rename(columns=rename_map)

    # Оставляем только нужные для дальнейших расчётов колонки
    needed_cols = [c for c in [
        'Дата',
        'ID заказа',
        'Номер заказа',
        'sku',
        'Наименование товара',
        'Расходы'
    ] if c in df_orders.columns]

    df_orders = df_orders[needed_cols]

    # Сохраняем обработанный отчёт по оплате за заказ в отдельный файл
    # (без ID кампании, т.к. его нет в этом отчёте)
    orders_parsed_path = companies_parsed_dir / "orders_report.csv"
    df_orders.to_csv(
        orders_parsed_path,
        sep=';',
        encoding='utf-8-sig',
        index=False
    )

    logger.info(
        f"Обработанный отчёт по продвижению с оплатой за заказ сохранён в {orders_parsed_path}"
    )

# Функция расчета расходов по рекламным кампаниям (в т.ч. оплата за заказ)
def calc_companies_cost_by_sku(date_report, df_finance_report, companies_costs_ids):
    # df, куда будем помещать итоговый результат (разбивку расходов по SKU)
    df_result_companies_stats = pd.DataFrame()
    # df с рекламными расходами из отчета ЛК (агрегировано по Номер кампании + Тип начисления)
    df_companies_stats_from_report = pd.DataFrame()
    # df со сравнением расходов по рекламе из АПИ и отчета ЛК
    df_companies_api_report_stats = pd.DataFrame()

    ORDERS_TYPE = "Продвижение с оплатой за заказ"

    # Агрегированные расходы по кампаниям из финансового отчёта
    df_companies_costs_total = get_companies_ids(df_finance_report, companies_costs_ids)

    if df_companies_costs_total.shape[0] > 0:
        # Сохраняем копию для вывода/аналитики
        df_companies_stats_from_report = df_companies_costs_total.copy()

        # Разделяем на обычные кампании и "оплату за заказ"
        df_regular_costs = df_companies_costs_total[
            df_companies_costs_total['Тип начисления'] != ORDERS_TYPE
        ].copy()
        df_orders_costs = df_companies_costs_total[
            df_companies_costs_total['Тип начисления'] == ORDERS_TYPE
        ].copy()

        # Сопоставление кампаний/заказов с расходами по SKU из API
        df_companies_api = pd.DataFrame()

        # 1. Обычные кампании (по ID кампании)
        for idx, row in df_regular_costs.iterrows():
            # Файл с данными по кампании
            company_file_path = (
                current_report_dir
                / f"{date_report}_Кампании_parsed"
                / f"{row['Номер кампании']}.csv"
            )
            # Если данный файл есть, начинаем считать расходы по данной кампании
            if company_file_path.exists():
                df_company = pd.read_csv(company_file_path, sep=';')
            else:
                df_company = pd.DataFrame(columns=['sku', 'Наименование товара', 'Расходы'])
                logger.warning(
                    f"Файл статистики по кампании {row['Номер кампании']} не найден "
                    f"({company_file_path})"
                )

            # Удаляем лишнюю колонку, если есть
            if 'Unnamed: 0' in df_company.columns:
                df_company = df_company.drop(columns='Unnamed: 0')

            if not df_company.empty:
                # Считаем итоговые расходы по каждому SKU
                tmp_df = (
                    df_company
                    .groupby(['sku', 'Наименование товара'], as_index=False)
                    .agg(Итого=('Расходы', 'sum'))
                )
                # Добавляем столбцы с информацией о кампании
                tmp_df['Номер кампании'] = row['Номер кампании']
                tmp_df['Тип начисления'] = row['Тип начисления']
                # Объединяем с предыдущими кампаниями
                df_companies_api = pd.concat([df_companies_api, tmp_df], ignore_index=True)

        # 2. Продвижение с оплатой за заказ (нет ID кампании, используем orders_report.csv)
        if not df_orders_costs.empty:
            orders_parsed_path = (
                current_report_dir
                / f"{date_report}_Кампании_parsed"
                / "orders_report.csv"
            )
            if orders_parsed_path.exists():
                df_orders = pd.read_csv(orders_parsed_path, sep=';')

                # Удаляем лишнюю колонку, если есть
                if 'Unnamed: 0' in df_orders.columns:
                    df_orders = df_orders.drop(columns='Unnamed: 0')

                if not df_orders.empty:
                    # Считаем итоговые расходы по каждому SKU из отчёта по заказам
                    # Ожидаем колонки: 'sku', 'Наименование товара', 'Расходы'
                    tmp_orders = (
                        df_orders
                        .groupby(['sku', 'Наименование товара'], as_index=False)
                        .agg(Итого=('Расходы', 'sum'))
                    )
                    # Добавляем виртуальный номер кампании 'ORDERS' и тип начисления
                    tmp_orders['Номер кампании'] = 'ORDERS'
                    tmp_orders['Тип начисления'] = ORDERS_TYPE

                    df_companies_api = pd.concat(
                        [df_companies_api, tmp_orders],
                        ignore_index=True
                    )
            else:
                logger.warning(
                    f"Файл orders_report.csv с отчётом по оплате за заказ не найден по пути "
                    f"{orders_parsed_path}"
                )

        if not df_companies_api.empty:
            # Считаем затраты на каждый тип начисления на каждый SKU
            df_companies_api_stats = (
                df_companies_api
                .groupby(['sku', 'Тип начисления'], as_index=False)
                .agg(Итого=('Итого', 'sum'))
            )

            # Заменяем знак расходов с плюса на минус
            df_companies_api_stats['Итого'] = np.where(
                df_companies_api_stats['Итого'] > 0,
                -abs(df_companies_api_stats['Итого']),
                df_companies_api_stats['Итого']
            )

            # Переводим в широкий формат (по типам начисления)
            df_companies_api_stats = (
                pd.pivot_table(
                    df_companies_api_stats,
                    index=['sku'],
                    columns='Тип начисления',
                    values='Итого',
                    aggfunc='sum',
                    fill_value=0
                )
                .reset_index()
            )

            # Переименовываем колонку с SKU
            df_companies_api_stats = df_companies_api_stats.rename(columns={'sku': 'SKU'})
            # Убираем имя у index
            df_companies_api_stats.index.name = None
            # Помещаем итоговый результат в переменную из начала функции
            df_result_companies_stats = df_companies_api_stats.copy()

            # Считаем расходы по типам начисления и кампаниям из апи
            df_companies_api_stats_by_operation = (
                df_companies_api
                .groupby(['Номер кампании', 'Тип начисления'], as_index=False)
                .agg(Итого_из_апи=('Итого', 'sum'))
            )

            # Заменяем знак расходов с плюса на минус
            df_companies_api_stats_by_operation['Итого_из_апи'] = np.where(
                df_companies_api_stats_by_operation['Итого_из_апи'] > 0,
                -abs(df_companies_api_stats_by_operation['Итого_из_апи']),
                df_companies_api_stats_by_operation['Итого_из_апи']
            )

            # Мерджим с расходами из отчета ЛК
            df_companies_api_report_stats = df_companies_costs_total.merge(
                df_companies_api_stats_by_operation,
                on=['Номер кампании', 'Тип начисления'],
                how='left'
            )
        else:
            logger.warning(
                "Не удалось сформировать данные по расходам из API "
                "(нет данных ни по кампаниям, ни по оплате за заказ)"
            )

    # Словарь с результатами
    result_companies_costs_no_article = {
        'df_companies_api_stats': df_result_companies_stats,
        'df_companies_stats_from_report': df_companies_stats_from_report,
        'df_companies_api_report_stats': df_companies_api_report_stats,
        'companies_costs_ids': companies_costs_ids
    }
    return result_companies_costs_no_article

# Функция добавления артикулов и наименований товаров к расходам по рекламным кампаниям
def add_articles_to_companies_costs(result_companies_costs_no_article, df_products_processed):
    # Создаем копию словаря, куда будем помещать итоговый результат
    result_companies_costs = result_companies_costs_no_article.copy()
    # Получаем df с результатами расходов SKU по рекламным кампаниям
    df_result_companies_stats = result_companies_costs_no_article['df_companies_api_stats'].copy()
    # Если были расходы по РК, то добавляем к SKU названия артикулов
    if df_result_companies_stats.shape[0] > 0:
        # Мерджим со списком артикулов
        df_companies_stats_with_articles = (
            df_result_companies_stats
            .merge(
                df_products_processed[['SKU', 'Артикул продавца']],
                on='SKU',
                how='left'
            )
        )
        # Перемещаем колонку с артикулом в начало df
        df_companies_stats_with_articles = move_columns(
            df=df_companies_stats_with_articles,
            columns_to_move=['Артикул продавца'],
            position='SKU',
            insert_type='after'
        )
        # Заполняем пропуски
        df_companies_stats_with_articles = df_companies_stats_with_articles.fillna({
            'Артикул продавца': 'Неизвестный артикул'
        })
        # Удаляем колонку с SKU, т.к. она была нужна только для мерджа
        df_companies_stats_with_articles = df_companies_stats_with_articles.drop(columns=['SKU'])
    # Если расходов по РК не было, возвращаем исходный пустой df из словаря
    else:
        df_companies_stats_with_articles = df_result_companies_stats.copy()

    # Изменяем df в словаре рекламных расходов на тот, что с артикулами
    result_companies_costs['df_companies_api_stats'] = df_companies_stats_with_articles.copy()

    return result_companies_costs


# Функция расчета данных по рекламным кампаниям
def calc_companies_costs(
        date_start,
        date_end,
        date_report,
        df_finance_report,
        df_products_processed,
        client_id_performance,
        client_secret_performance,
        task_id=None
    ):
    # Директория, где будут лежать данные по рекламным кампаниям
    companies_stats_dir: Path = current_report_dir / f"{date_report}_Кампании"
    companies_stats_dir.mkdir(parents=True, exist_ok=True)

    # Директория, куда будем складывать обработанные файлы
    companies_parsed_dir: Path = current_report_dir / f"{date_report}_Кампании_parsed"
    companies_parsed_dir.mkdir(parents=True, exist_ok=True)

    # Проверяем, что в отчете были расходы по рекламе
    companies_costs_ids = get_companies_costs_ids(df_finance_report)
    if not companies_costs_ids:
        logger.info("В финансовом отчёте не найдено расходов по рекламным кампаниям")
        return

    logger.info("Найдены расходы по рекламным кампаниям")

    # Все кампании, по которым есть расходы
    df_companies_ids = get_companies_ids(df_finance_report, companies_costs_ids)

    ACCRUAL_COL = "Тип начисления"
    ORDERS_TYPE = "Продвижение с оплатой за заказ"
    CAMPAIGN_ID_COL = "Номер кампании"

    # --- Разделяем на обычные кампании и "оплату за заказ" ---
    df_regular = df_companies_ids.copy()
    df_orders = pd.DataFrame()

    if ACCRUAL_COL in df_companies_ids.columns:
        mask_orders = df_companies_ids[ACCRUAL_COL] == ORDERS_TYPE
        df_orders = df_companies_ids[mask_orders].copy()
        df_regular = df_companies_ids[~mask_orders].copy()

    # Общие даты для Performance API
    date_start_performance = ''   # при необходимости: f"{date_start}T00:00:00Z"
    date_end_performance = ''     # при необходимости: f"{date_end}T23:59:59Z"
    date_from_performance = date_start
    date_to_performance = date_end

    # ==========================
    # 1. ОПРЕДЕЛЯЕМ, КАКИЕ КАМПАНИИ ДОКАЧАТЬ
    # ==========================
    df_missing_regular = pd.DataFrame()

    if not df_regular.empty:
        # Все требуемые ID кампаний
        required_campaign_ids = set(df_regular[CAMPAIGN_ID_COL].astype(str))

        # Уже имеющиеся .csv по кампаниям: <id>_dd.mm.yyyy-dd.mm.yyyy.csv
        csv_files = list(companies_stats_dir.glob("*.csv"))
        existing_campaign_ids = set()

        for path in csv_files:
            stem = path.stem  # например: "23010536_01.02.2026-28.02.2026"
            parts = stem.split("_", 1)
            if not parts:
                continue
            campaign_id = parts[0]
            if campaign_id.isdigit():
                existing_campaign_ids.add(campaign_id)

        # Кампании, по которым ещё нет файлов
        missing_campaign_ids = required_campaign_ids - existing_campaign_ids

        if missing_campaign_ids:
            df_missing_regular = df_regular[
                df_regular[CAMPAIGN_ID_COL].astype(str).isin(missing_campaign_ids)
            ]

    need_regular_upload = not df_missing_regular.empty

    # ==========================
    # 2. НУЖНО ЛИ ДОКАЧИВАТЬ "ОПЛАТУ ЗА ЗАКАЗ"
    # ==========================
    has_orders_expenses = not df_orders.empty

    # Есть ли уже файл отчёта по оплате за заказ:
    # search_promo_all_skus_orders_report_24.02.2026-01.03.2026.csv
    orders_report_exists = bool(
        list(companies_stats_dir.glob("search_promo_all_skus_orders_report_*.csv"))
    )

    need_orders_upload = has_orders_expenses and not orders_report_exists

    # ==========================
    # 3. ОДИН ВЫЗОВ upload_data_performane С НУЖНЫМ mode
    # ==========================
    if need_regular_upload or need_orders_upload:
        if need_regular_upload and need_orders_upload:
            mode = "all"          # кампании + оплата за заказ
            df_for_upload = df_missing_regular
            logger.info(
                f"Выгружаем статистику по {len(df_missing_regular)} недостающим кампаниям "
                f"и отчёт по продвижению с оплатой за заказ (mode='all')"
            )
        elif need_regular_upload:
            mode = "campaigns"    # только кампании
            df_for_upload = df_missing_regular
            logger.info(
                f"Выгружаем статистику по {len(df_missing_regular)} недостающим кампаниям "
                f"(mode='campaigns')"
            )
        else:  # need_orders_upload only
            mode = "orders"       # только оплата за заказ
            df_for_upload = None  # df не нужен для отчёта по заказам
            logger.info(
                "Выгружаем только отчёт по продвижению с оплатой за заказ (mode='orders')"
            )

        upload_data_performane(
            date_start_performance,
            date_end_performance,
            client_id_performance,
            client_secret_performance,
            companies_stats_dir,
            date_from_performance,
            date_to_performance,
            df_for_upload,
            task_id,
            mode=mode,
        )
    else:
        logger.info(
            "Нет недостающих файлов статистики по кампаниям "
            "и отчёта по продвижению с оплатой за заказ — догрузка не требуется"
        )

    # ==========================
    # 4. Парсим всё, что есть
    # ==========================
    parse_companies_files(current_report_dir, date_report)

    # Считаем расходы по рекламным кампаниям
    result_companies_costs_no_article = calc_companies_cost_by_sku(
        date_report,
        df_finance_report,
        companies_costs_ids
    )

    # Добавляем артикул к SKU с рекламными расходами
    result_companies_costs = add_articles_to_companies_costs(
        result_companies_costs_no_article,
        df_products_processed
    )

    return result_companies_costs

# Функция товаров в заявках на поставку
def upload_supplies_items(df_supplies_from_report):
    # Получаем уникальный список номеров заявок
    supplies_numbers = (
        df_supplies_from_report['Номер заявки на поставку']
        .astype(int)
        .unique()
        .tolist()
    )
    # Формируем df с уникальным списком номеров заявок
    df_supplies_numbers = pd.DataFrame({'Номер заявки на поставку': supplies_numbers})
    # Разбиваем df на равные части по 50 заявок в виде отдельного столбца
    df_supplies_numbers['chunk'] = df_supplies_numbers.index // 50 + 1
    # Цикл по каждому чанку
    for chunk in df_supplies_numbers['chunk'].unique():
        order_ids = (
            df_supplies_numbers
            .loc[df_supplies_numbers['chunk'] == chunk, 'Номер заявки на поставку']
            .to_list()
        )
        # Формируем параметры запроса
        # params_supplies_numbers = {
        #     "orders_ids": order_ids[0]
        # }
        # # Отправляем запрос
        # resp_data_supplies = requests.post(
        #     f"{ozon_seller_api_url}/v3/supply-order/get",
        #     headers=headers,
        #     json=params_supplies_numbers
        # ).json()
        # Формируем параметры запроса
        params_supplies_numbers = {
            'filter':{
                "orders_number_search": [order_ids[0]],
                "states": [
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
            },
            'last_id': 'CMDE0xcSBgjQhqDABg==',
            'limit': 100,
            'sort_by': "ORDER_CREATION"

        }
        # Отправляем запрос
        resp_data_supplies = requests.post(
            f"{ozon_seller_api_url}/v3/supply-order/list",
            headers=headers,
            json=params_supplies_numbers
        ).json()

# Функция расчета расходов по поставкам
def calc_supplies_costs(df_finance_report):

    # Итоговый df, который будет возвращен функцией
    supplies_costs_matched = pd.DataFrame()
    # Расходы по поставкам из отчета
    supplies_costs_from_report = pd.DataFrame()
    # id строк отчета с поставками, которые нашлись в апи
    supplies_costs_ids_matched = []

    # Выбираем расходы, которые относятся к поставкам
    df_supplies_from_report = df_finance_report.loc[
        # Все заявки на поставку начинаются с цифры 2
        (df_finance_report['ID начисления'].str.startswith('2', na=False))
        # Не содержат -
        & ~(df_finance_report['ID начисления'].str.contains('-', na=False))
        # Все заявки на поставку содержат ровно 13 цифр
        & (df_finance_report['ID начисления'].str.len() == 13),
        ['id', 'Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]


    # Если есть расходы на поставки в отчете
    if df_supplies_from_report.shape[0] > 0:
        # Список номеров строк с расходами на поставки, расходы по которым можно разбить на sku
        supplies_costs_ids_matched = []
        # Считаем расходы по конкретной заявке и по типу начисления
        df_supplies_from_report = df_supplies_from_report.rename(columns={'ID начисления': 'Номер заявки на поставку'})
        # Конвертация в Int64 для избежания ошибок при join
        df_supplies_from_report['Номер заявки на поставку'] = df_supplies_from_report['Номер заявки на поставку'].astype('Int64')
        df_supplies_from_report_total = df_supplies_from_report.groupby(['Номер заявки на поставку', 'Тип начисления']).agg(
            Расходы_поставки_из_отчета_ЛК=('Итого руб', 'sum')
        ).reset_index()
        # Помещаем расходы на поставки из ЛК в отдельный df
        supplies_costs_from_report = df_supplies_from_report_total.copy()
        # df_supplies_from_report_total['Номер заявки на поставку'] = df_supplies_from_report_total['Номер заявки на поставку'].astype('Int64')
        # Составляем список уникальных заявок на поставку (только те, которые были в отчете)
        unique_supply_ids = (
            df_supplies_from_report_total['Номер заявки на поставку']
            # .astype(str)
            .unique()
            .tolist()
        )
        # --- Выгрузка информации о поставках --- #
        # Список заявок на поставку
        df_supply_order_list = get_supply_order_list(
            headers,
            states = [
                # "UNSPECIFIED",
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
            # order_number_search=unique_supply_ids
        )
        # Информация о заявках на поставку
        df_supply_orders_info = get_supply_orders_details(
            df_supply_order_list,
            headers,
            batch_size=50
        )
        # Выбираем только те заявки на поставку, которые были в отчете по начислениям
        df_supply_orders_info_filtered = (
            df_supply_orders_info
            .loc[df_supply_orders_info['supply_id'].isin(unique_supply_ids), :]
        )
        # Получаем товары в заявках на поставку
        df_supplies_items = get_supply_bundle_items(
            df_supply_orders_info_filtered,
            headers
        )
        # df_supplies_items.to_excel(f"{finance_reports_dir}/{client_name}_Товары_в_поставках.xlsx")
        # Переименовываем колонки для удобства
        df_supplies_items = df_supplies_items.rename(columns={
            'sku': 'SKU',
            'Артикул': 'Артикул продавца',
        })
        # Считаем итоговое количество sku в заявке на поставку
        df_supplies_items_total = (
            df_supplies_items
            # .groupby(['SKU', 'Артикул продавца', 'Наименование товара', 'supply_id'])
             .groupby(['Артикул продавца', 'supply_id'])
            .agg(
                Количество=('Количество', 'sum')
            )
            .reset_index()
        )
        # Количество товаров в заявке
        df_supplies_items_total['Количество товаров в поставке'] = df_supplies_items_total.groupby('supply_id')['Количество'].transform('sum')

        # Конвертация в Int64 для избежания ошибок при join
        df_supplies_items_total['supply_id'] = df_supplies_items_total['supply_id'].astype('Int64')

        # Джойним таблицу с расходами из отчета для получения типа начисления
        # BUG: После джойна одна из колонок конвертируется во float
        # Из-за этого пришлось для соединяемых колонок делать конвертацию в Int64
        df_supplies_all = df_supplies_items_total.merge(
            df_supplies_from_report_total[['Номер заявки на поставку', 'Тип начисления', 'Расходы_поставки_из_отчета_ЛК']],
            left_on='supply_id',
            right_on='Номер заявки на поставку',
            how='outer',
            indicator=True
        )

        # Номера заявок на поставку, которые совпали с апи
        df_supplies_found = df_supplies_all.loc[df_supplies_all['_merge'] == 'both', :]
        # Номера заявок на поставку, которые не совпали с апи
        # df_supplies_mismatched = df_supplies_all.loc[df_supplies_all['_merge'] == 'right_only', :]

        # Если есть совпадающие заявки, считаем по ним расходы
        if df_supplies_found.shape[0] > 0:
            # Получаем id строк отчета из ЛК с поставками, расходы по которым мы разбили на sku
            supplies_costs_ids_matched = df_supplies_from_report.loc[df_supplies_from_report['Номер заявки на поставку'].isin(df_supplies_found['Номер заявки на поставку'].dropna().unique()), 'id'].to_list()
            # Считаем число товаров в каждой поставке
            # df_supplies_found['Количество товаров в поставке'] = df_supplies_found.groupby('supply_id')['quantity'].transform('sum')
            # Считаем расходы на отдельный sku (распределение от количества sku в поставке)
            df_supplies_found['Расходы по поставкам'] = df_supplies_found['Количество'] / df_supplies_found['Количество товаров в поставке'] * df_supplies_found['Расходы_поставки_из_отчета_ЛК']
            # Переводим в широкий формат
            df_supplies_found_total = (
                pd.pivot_table(
                    df_supplies_found,
                    index=['Артикул продавца'],
                    columns='Тип начисления',
                    values='Расходы по поставкам',
                    aggfunc='sum',
                    fill_value=0
                )
                .reset_index()
            )
            # Переименовываем колонку с SKU
            # df_supplies_found_total = df_supplies_found_total.rename(columns={'sku': 'SKU'})
            # Переименовываем index
            df_supplies_found_total.index.name = None
            # Переводим артикул в строку
            df_supplies_found_total['Артикул продавца'] = df_supplies_found_total['Артикул продавца'].astype(str)
            # Сохраняем результат в отдельную переменную, которую будет возвращать функция
            supplies_costs_matched = df_supplies_found_total.copy()

        # Если есть поставки, которых нет в апи, их будем считать отдельно в другой функции
        # if df_supplies_mismatched.shape[0] > 0:
        #     supplies_costs_ids_matched = df_supplies_from_report.loc[df_supplies_from_report['Номер заявки на поставку'].isin(df_supplies_mismatched['Номер заявки на поставку'].dropna().unique()), 'id'].to_list()

    # Формируем итоговый словарь с датафреймом расходов по поставкам
    # и id поставок, которых нет в апи
    result_supplies_costs = {
        'supplies_costs_matched': supplies_costs_matched,
        'supplies_costs_from_report': supplies_costs_from_report,
        'supplies_costs_ids_matched': supplies_costs_ids_matched
    }

    return result_supplies_costs


# Функция чтения справочной таблицы
def read_catalog(client_name):
    # Задаем путь к справочной таблице
    path_catalog = (
        marketplace_dir_path
        / "Clients"
        / client_name
        / "catalog"
        / f"Справочная_таблица_{client_name}_Ozon.xlsx"
    )
    # Считывание справочной таблицы
    catalog = pd.read_excel(path_catalog)

    # Из справочной таблицы убираем товары, у которых нет SKU или Артикула
    # catalog_ = catalog.dropna(subset=['SKU', 'Артикул продавца'])
    # В справочной таблице переводим себестоимость в число, т.к. там иногда бывают строки
    catalog['Себестоимость'] = catalog['Себестоимость'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
    # Переводим артикул в строку
    catalog['Артикул продавца'] = catalog['Артикул продавца'].astype(str)
    # Переименовываем колонку с наименованием товара
    catalog = catalog.rename(columns={
        'Название': 'Наименование товара'
    })

    return catalog


# Функция выгрузки списка товаров Ozon
def get_ozon_product(headers, to_save=False):
    # Выгружаем список товаров
    df_products = upload_ozon_product(headers=headers, to_save=to_save)
    # Создаем копию
    df_products_processed = df_products.copy()
    # Выбираем нужные колонки
    df_products_processed = (
        df_products_processed
        .rename(columns={
                'Артикул': 'Артикул продавца',
                'Название товара': 'Наименование товара'
        })
        .loc[:, ['SKU', 'Артикул продавца', 'Наименование товара']]
    )
    # Переводим артикул в строку
    df_products_processed['Артикул продавца'] = df_products_processed['Артикул продавца'].astype(str)
    # Переводим SKU в int64
    df_products_processed['SKU'] = df_products_processed['SKU'].astype('int64')

    return df_products_processed


# Функция создания списка артикулов из расходов, которые уже разбиты по артикулам
def create_all_sku_df(result_sku_costs, result_supplies_costs, catalog):
    # Получаем df с расходами, которые уже разбиты по SKU и Артикулу
    df_sku_costs = result_sku_costs['df_sku_costs_all'].copy()
    df_supplies_costs = result_supplies_costs['supplies_costs_matched'].copy()
    # catalog_ = catalog.copy()
    # # Из справочной таблицы убираем товары, у которых нет SKU или Артикула
    # catalog_ = catalog_.dropna(subset=['SKU', 'Артикул продавца'])
    # Формируем лист с датафреймами
    df_list = [
        df_sku_costs,
        df_supplies_costs,
        # catalog_
    ]
    # Переводи артикул в каждом df в строку
    for df in df_list:
        if 'Артикул продавца' in df.columns:
            df['Артикул продавца'] = df['Артикул продавца'].astype(str)
    # Выбираем нужные колонки
    sku_columns = [
        # 'SKU',
        'Артикул продавца',
        # 'Наименование товара'
    ]
    # Используем reindex для добавления отсутствующих колонок
    df_list = [df.reindex(columns=sku_columns) for df in df_list]
    # Объединяем в один df и убираем дубликаты
    df_sku_list_all = (
        pd.concat(df_list)
        .drop_duplicates(subset=['Артикул продавца'])
        .dropna()
    )

    return df_sku_list_all


# Функция объединения затрат, которые удалось распределить sku
def merge_sku_costs(result_sku_costs, result_companies_costs, result_supplies_costs):
    # Получаем df посчитанных затрат
    df_sales_returns = result_sku_costs['df_sku_costs_all']
    df_companies_stats = result_companies_costs['df_companies_api_stats']
    df_supplies_costs = result_supplies_costs['supplies_costs_matched']

    # Объединяем всё в один df
    df_sku_costs_all = pd.concat([df_sales_returns, df_supplies_costs, df_companies_stats])
    # Считаем сумму после union
    df_sku_costs_all_stats = (
        df_sku_costs_all
        .groupby(['SKU', 'Артикул продавца', 'Наименование товара'])
        .sum()
        .reset_index()
    )

    return df_sku_costs_all_stats

# Функция объединения затрат, которые удалось распределить sku
def merge_sku_costs_v2(
        result_sku_costs,
        result_companies_costs,
        result_supplies_costs,
        df_products_processed
        ):
    # Получаем df посчитанных затрат
    df_sales_returns = result_sku_costs['df_sku_costs_all']
    df_companies_stats = result_companies_costs['df_companies_api_stats']
    df_supplies_costs = result_supplies_costs['supplies_costs_matched']

    # Объединяем всё в один df
    df_sku_costs_all = pd.concat([df_sales_returns, df_supplies_costs, df_companies_stats])
    # Считаем сумму после union
    df_sku_costs_all_stats = (
        df_sku_costs_all
        # .groupby(['SKU', 'Артикул продавца', 'Наименование товара'])
        .groupby(['Артикул продавца'])
        .sum()
        .reset_index()
    )
    # Из списка товаров добавляем SKU и Наименование товара
    df_sku_costs_all_stats = df_sku_costs_all_stats.merge(
        df_products_processed[['Артикул продавца', 'SKU', 'Наименование товара']],
        on='Артикул продавца',
        how='left'
    )
    # Перемещаем колонки в начало df
    df_sku_costs_all_stats = move_columns(
        df_sku_costs_all_stats,
        ['Наименование товара', 'SKU'],
        position='Артикул продавца'
    )

    return df_sku_costs_all_stats



# Функция расчета расходов, которые распределяются в зависимости от количества продаж sku
def calc_other_costs(df_finance_report, df_sku_costs_all_stats, result_sku_costs, result_companies_stats, result_supplies_costs):

    # Переменные, в которые будем сохранять итоговый результат
    df_result_all_costs = pd.DataFrame()
    result_operations_other_costs = pd.DataFrame()

    # Копируем df с расходами по SKU для избежания изменений в оригинальном df
    df_sku_costs_all_stats_ = df_sku_costs_all_stats.copy()
    # Получаем id строк, по которым уже посчитали затраты, разбив их на sku
    sales_returns_ids = result_sku_costs['sku_costs_ids']
    companies_stats_ids = result_companies_stats['companies_costs_ids']
    supplies_costs_ids = result_supplies_costs['supplies_costs_ids_matched']
    calculated_costs_ids = sales_returns_ids + companies_stats_ids + supplies_costs_ids

    # Получаем из фин. отчета оставшиеся строки
    df_other_costs = df_finance_report.loc[~df_finance_report['id'].isin(calculated_costs_ids), :]

    if df_other_costs.shape[0] > 0:
        # Считаем суммарные расходы по каждому из типов начисления в прочих расходах
        df_other_costs_stats = df_other_costs.groupby('Тип начисления').agg(
            Итого_прочие_расходы=('Итого руб', 'sum')
        ).reset_index()

        # Для каждого из типов начислений разбиваем расходы на каждый SKU
        # Получаем список типов начислений
        operation_list_other = df_other_costs_stats['Тип начисления'].unique().tolist()
        # Всего продаж по всем SKU
        total_sales = df_sku_costs_all_stats_['Продажи шт'].sum()
        for operation_type in operation_list_other:
            # Расходы по данному типу операции
            operation_costs = df_other_costs_stats.loc[df_other_costs_stats['Тип начисления'] == operation_type, 'Итого_прочие_расходы'].values[0]
            df_sku_costs_all_stats_[operation_type] = df_sku_costs_all_stats_['Продажи шт'] / total_sales * operation_costs
            # Добавляем 0.0, чтобы избавиться от -0.0
            df_sku_costs_all_stats_[operation_type] = df_sku_costs_all_stats_[operation_type] + 0.0

        # Помещаем результаты расчетов в переменные из начала функции
        df_result_all_costs = df_sku_costs_all_stats_
        result_operations_other_costs = df_other_costs_stats

    # Формируем словарь с результатами
    result_all_costs = {'df_result_all_costs': df_result_all_costs,
                        'result_operations_other_costs': result_operations_other_costs}

    return result_all_costs


# Функция добавления данных из справочной таблицы
def add_data_from_catalog(result_all_costs):
    # Задаем путь к справочной таблице
    path_catalog = (
        marketplace_dir_path
        / "Clients"
        / client_name
        / "catalog"
        / f"Справочная_таблица_{client_name}_Ozon.xlsx"
    )
    # Считывание справочной таблицы
    catalog_reference = pd.read_excel(path_catalog)
    # Из справочной таблицы убираем товары, у которых нет SKU
    catalog_reference_ = catalog_reference.loc[~catalog_reference['SKU'].isna(), :]
    # В справочной таблице переводим себестоимость в число, т.к. там иногда бывают строки
    catalog_reference_['Себестоимость'] = catalog_reference_['Себестоимость'].apply(lambda x: pd.to_numeric(x, errors='coerce'))
    # Переводим артикул в строку
    catalog_reference_['Артикул продавца'] = catalog_reference_['Артикул продавца'].astype(str)
    # Переименовываем некоторые колонки
    catalog_reference_ = catalog_reference_.rename(columns={
        'Название': 'Наименование товара'
    })
    # Переводим SKU в float64, чтобы нормально проходило объединение
    catalog_reference_['SKU'] = catalog_reference_['SKU'].astype('float64')
    # Достаем df с результатами расчетов
    df_all_costs = result_all_costs['df_result_all_costs'].copy()
    # Колонки, которые берем из справочной таблицы
    catalog_columns = catalog_finace_svod_columns
    # Объединяем справочную таблицу с результатами расчетов
    df_all_costs_with_catalog = df_all_costs.merge(
        catalog_reference_[catalog_columns],
        on='Артикул продавца',
        how='left'
    )
    # # Перемещаем справочные колонки в начало df
    # df_all_costs_with_catalog = df_all_costs_with_catalog[ catalog_columns + [col for col in df_all_costs_with_catalog.columns if col not in catalog_columns]]
    # # Перемещаем колонку с себестоимостью в конец df
    # df_all_costs_with_catalog.insert(len(df_all_costs_with_catalog.columns) - 1, 'Себестоимость', df_all_costs_with_catalog.pop('Себестоимость'))
    columns_to_move=[col for col in catalog_columns if col not in ['SKU', 'Артикул продавца']]
    df_all_costs_with_catalog = move_columns(
        df=df_all_costs_with_catalog,
        columns_to_move=columns_to_move,
        position='Наименование товара',
        insert_type='after'
    )

    return df_all_costs_with_catalog


# Функция расчета колонки итоговых расходов
def calc_final_costs(df_all_costs_with_catalog):
    # Достаем df с результатами расчетов
    df_final_costs = df_all_costs_with_catalog.copy()

    # Колонка, от которой считаем долю маржинальности
    marg_sales = 'Продажи - Возвраты (руб)'
    marg_amount = 'Продажи - Возвраты (шт)'
    marg_sales_column = df_final_costs[marg_sales]


    # Считаем себестоимость
    df_final_costs['Себестоимость * Факт продаж'] = (-1) * abs(df_final_costs['Себестоимость'] * df_final_costs[marg_amount])

    # Считаем все расходы на рекламные кампании
    # Определяем колонки, которые будут включены в затраты на рекламу
    promotion_columns = [
        'Вывод в топ',
        'Трафареты',
        'Продвижение с оплатой за заказ',
        'Оплата за клик',
        'Абонентское обслуживание товаров на продвижении',
        'Подписка Premium',
        'Подписка Premium Plus',
        'Бонусы продавца',
        'Бейдж Оригинал',
        'Баллы за отзывы',
        'Продвижение бренда',
        'Закрепление отзыва',
        'Рассылка пуш-уведомлений',
    ]
    # Находим колонки по рекламе, которые реально есть в отчете реализации
    existing_promotion_columns = [col for col in df_final_costs.columns if col in promotion_columns]
    if existing_promotion_columns:
        df_final_costs['Затраты на рекламу (руб)'] = df_final_costs[existing_promotion_columns].sum(axis=1)
    else:
        df_final_costs['Затраты на рекламу (руб)'] = 0.0
    # Считаем % по рекламным расходам
    df_final_costs['% Затрат на рекламу'] = (
        abs(
            df_final_costs['Затраты на рекламу (руб)'] /
            marg_sales_column
            * 100
        )
    )

    # Определяем колонки, которые будут исключены из затрат
    columns_to_exclude_from_final_costs = ([
        'SKU', 'Наименование товара', 'Артикул продавца',
        'Продажи шт',
        'Возвраты шт',
        'Продажи - Возвраты (шт)',
        'Продажи руб',
        'Продажи - Возвраты (руб)',
        'Возвраты руб',
        # 'Затраты на рекламу (руб)',
        '% Затрат на рекламу',
        'Маржинальность %',
        'Маржинальность на ед. товара руб',
        'Маржинальность руб',
        'Маржа от себестоимости %',
        'Маржа по артикулу %',
        'Приход денег на РС',
        'Все затраты (руб)',
        ]
        + catalog_finace_svod_columns + promotion_columns
    )
    # Выбираем итоговые колонки для расчета маржинальности
    final_costs_columns = df_final_costs.columns[~df_final_costs.columns.isin(columns_to_exclude_from_final_costs)]

    # Считаем сумму по всем затратам
    df_final_costs['Все затраты (руб)'] = df_final_costs[final_costs_columns].sum(axis=1)
    # Вычитаем из продаж в рублях все затраты
    df_final_costs['Маржинальность руб'] = df_final_costs['Продажи - Возвраты (руб)'] + df_final_costs['Все затраты (руб)']
    df_final_costs = df_final_costs.drop(columns=['Все затраты (руб)'])

    # Приход денег на РС (как маржинальность, но не учитываем Себестоимость * Факт продаж)
    df_final_costs['Приход денег на РС'] = (
        df_final_costs['Маржинальность руб']
        - df_final_costs['Себестоимость * Факт продаж'].fillna(0)
    )
    # Там, где нет себестоимости, маржинальность не считаем
    df_final_costs.loc[df_final_costs['Себестоимость'].isna(), 'Маржинальность руб'] = np.nan

    # Считаем маржинальность в %
    df_final_costs['Маржинальность %'] = df_final_costs['Маржинальность руб'] / marg_sales_column

    # Считаем маржинальность на ед. товара
    df_final_costs['Маржинальность на ед. товара руб'] = (
        df_final_costs['Маржинальность руб'] /
        df_final_costs['Продажи - Возвраты (шт)']
    )
    # Маржинальность от себестоимости (%)
    df_final_costs['Маржа от себестоимости %'] = (
        df_final_costs['Маржинальность руб']
        / abs(df_final_costs['Себестоимость * Факт продаж'])
        * 100
    )
    # Выделяем артикул размера
    df_final_costs['Артикул размерного ряда'] = get_common_article(df_final_costs)
    # Считаем маржинальность по артикулу
    sales_rub_mean_by_article = (
        df_final_costs
        .groupby(['Артикул размерного ряда'])
        [marg_sales].transform('mean')
    )
    marg_rub_mean_by_article = (
        df_final_costs
        .groupby(['Артикул размерного ряда'])
        ['Маржинальность руб'].transform('mean')
    )
    df_final_costs['Маржа по артикулу %'] = (
        marg_rub_mean_by_article / sales_rub_mean_by_article * 100
    )
    # Удаляем колонку с артикулом размерного ряда
    df_final_costs = df_final_costs.drop(columns=['Артикул размерного ряда'])

    # Там, где маржинальность в рублях > 0, у маржинальности на единицу товара так же будет знак "+"
    # в ином случае "-"
    for col in ['Маржинальность на ед. товара руб', 'Маржинальность %']:
        df_final_costs[col] = np.where(
            df_final_costs['Маржинальность руб'] < 0,
            (-1) * abs(df_final_costs[col]),
            abs(df_final_costs[col])
        )
    # for col in ['Маржа от себестоимости %']:
    #     df_final_costs[col] = np.where(
    #         df_final_costs['Себестоимость * Факт продаж'] <= 0,
    #         (-1) * abs(df_final_costs[col]),
    #         abs(df_final_costs[col])
    #     )

    # Меняем inf на nan
    columns_to_replace_inf = [
        'Маржинальность %',
        'Маржинальность на ед. товара руб',
        'Маржа от себестоимости %',
        'Маржа по артикулу %',
        '% Затрат на рекламу',
    ]
    df_final_costs[columns_to_replace_inf] = df_final_costs[columns_to_replace_inf].replace([-np.inf, np.inf], np.nan)


    return df_final_costs, promotion_columns, columns_to_exclude_from_final_costs


# Функция расчета сводной таблицы по всем расходам
def create_svod_by_operations(df_final_costs):

    # Создаем копию для избежания изменений в оригинальном df
    df_final_costs_ = df_final_costs.copy()
    # Список колонок, по котором мы не считаем суммы
    columns_to_exclude = [
        'SKU',
        'Артикул продавца',
        'Основной артикул',
        'Размер',
        'Цвет',
        'Сезон',
        'Статус',
        'Наименование товара',
        'Себестоимость',
        'Категория',
        'РРЦ',
        '% Затрат на рекламу',
        # 'Затраты на рекламу (руб)',
        # 'Себестоимость * Факт продаж',
        # 'Все затраты (руб)',
        # 'Маржинальность руб',
        'Маржинальность %',
        'Маржа по артикулу %',

    ]
    # Считаем сумму, от которой будем считать процент
    sum_columns = [
        'Продажи - Возвраты (руб)',
        'Баллы за скидки',
        'Программы партнеров',
        'Программы партнёров'
    ]
    # Считается сумма по тем колонкам, которые есть в df_final_costs
    df_final_costs_['Продажа (Выручка) руб'] = (
        df_final_costs_
        .loc[:, df_final_costs_.columns.isin(sum_columns)]
        .sum(axis=1)
    )
    # Перемещаем колонку в начало df
    df_final_costs_ = move_columns(
        df_final_costs_,
        'Продажа (Выручка) руб',
        position=0,
        insert_type='before'
    )
    # Считаем суммы по остальным столбцам
    df_final_costs_by_operation = (
        df_final_costs_
        # Убираем колонки, которые не должны быть включены в сумму
        .loc[
            :,
            ~df_final_costs_.columns.isin(columns_to_exclude)
        ]
        # Считаем суммы по остальным
        .sum()
        # Переворачиваем расходы в столбцы
        .to_frame()
        .T
        .reset_index(drop=True)

        # # Считаем суммы по остальным
        # .sum(skipna=True)
        # # Достаем столбцы из index
        # .reset_index()
        # # Переименовываем колонки
        # .rename(columns={
        #     'index': 'Тип операции',
        #     0: 'Итого руб'
        # })
    )
    # Пересчитываем некоторые столбцы
    df_final_costs_by_operation['Маржинальность на ед. товара'] = (
        df_final_costs_by_operation['Маржинальность руб']
        / df_final_costs_by_operation['Продажи - Возвраты (шт)']
    )
    # Делаем расходы снова в строках
    df_final_costs_by_operation = (
        df_final_costs_by_operation
        .melt(var_name='Тип операции', value_name='Итого руб')
    )

    # Сумма, от которой считаем процент
    sum_for_percent = df_final_costs_['Продажа (Выручка) руб'].sum()

    # Считаем процент
    df_final_costs_by_operation['Итого, %'] = (
        df_final_costs_by_operation['Итого руб'] / sum_for_percent * 100
    )
    # Делаем процент пропуском в некоторых расходах
    no_percent_operations = [
        'Продажи шт',
        'Возвраты шт',
        'Продажи - Возвраты (шт)',
    ]
    # Маска для loc
    mask_no_operations = df_final_costs_by_operation['Тип операции'].isin(no_percent_operations)
    # Создание пропусков
    df_final_costs_by_operation.loc[mask_no_operations, 'Итого, %'] = np.nan

    return df_final_costs_by_operation

# Функция сохранения расчетов в excel
def save_finance_svod_to_excel(
        date_report,
        df_finance_report,
        df_final_costs_by_operation,
        result_sku_costs,
        df_final_costs,
        result_all_costs,
        result_companies_stats,
        result_supplies_costs,
        date_start,
        date_end
    ):
    # Получаем нужные df из расчетов
    df_all_costs = df_final_costs.copy()
    df_sku_costs_by_operation = result_sku_costs['df_sku_costs_by_operation']
    result_operations_other_costs = result_all_costs['result_operations_other_costs']
    df_companies_api_report_stats = result_companies_stats['df_companies_api_report_stats']
    df_supplies_costs_from_report = result_supplies_costs['supplies_costs_from_report']

    # Формируем даты отчета
    df_report_dates = pd.DataFrame({'Дата начала периода': date_start, 'Дата окончания периода': date_end}, index=[0])

    # Путь для сохранения файла (с именем)
    filepath_finance_svod = (
        current_report_dir
        / f'{date_report}_Отчет_маржинальности_Ozon_{client_name}.xlsx'
    )
    # Сохраняем отчет
    with pd.ExcelWriter(filepath_finance_svod) as w:
        # Объединенные расходы с распределением по sku
        df_all_costs.to_excel(w, sheet_name='Итоговая таблица', index=False)
        # Сводная по начислениям
        df_final_costs_by_operation.to_excel(w, sheet_name='Сводная по начислениям', index=False)
        # Исходный отчет
        df_finance_report.to_excel(w, sheet_name='Исходный отчет', index=False)
        # Даты формирования исходного отчета
        df_report_dates.to_excel(w, sheet_name='Даты формирования отчета', index=False)
        # Расходы, которые уже были распределены по SKU в исходном отчете
        df_sku_costs_by_operation.to_excel(w, sheet_name='Расходы по SKU', index=False)
        # Расходы по типам начисления, которые распределялись в зависимости от продаж SKU
        result_operations_other_costs.to_excel(w, sheet_name='Прочие расходы', index=False)
        # Расходы на поставки из отчета ЛК
        df_supplies_costs_from_report.to_excel(w, sheet_name='Расходы на поставки из ЛК', index=False)
        # Сопоставление расходов по рекламе из АПИ и из отчета
        df_companies_api_report_stats.to_excel(w, sheet_name='Расходы по рекламе', index=False)

    return filepath_finance_svod

# Функция форматирования excel
def format_excel(
        date_report,
        df_final_costs,
        promotion_columns,
        columns_to_exclude_from_final_costs,
    ):

    logger.info(f"Formatting Finance Report for client {client_name}")

    # Формируем путь к итоговому файлу Excel
    filepath_finance_svod = (
        current_report_dir
        / f'{date_report}_Отчет_маржинальности_Ozon_{client_name}.xlsx'
    )
    # Открываем итоговый Excel-файл
    wb = openpyxl.load_workbook(filepath_finance_svod)

    # --- ШАГ 1. Форматируем лист "Итоговой таблицы" ---
    final_ws = wb['Итоговая таблица']

    # Готовим dataframe для сопоставления: название -> номер столбца -> буква Excel
    # excel_columns = pd.DataFrame({
    #     "column": df_final_costs.columns,
    #     "column_number": np.arange(1, len(df_final_costs.columns) + 1)
    # })
    # excel_columns['excel_column'] = excel_columns['column_number'].apply(get_column_letter)

    # svod_len = df_final_costs.shape[0]
    # row_start = 2  # С какой строки начинаются данные (1 — только заголовок)

    # min_col = excel_columns.excel_column.iloc[0]
    # max_col = excel_columns.excel_column.iloc[-1]

    # --- ШАГ 2. Формирование формул для рекламных расходов и маржинальности ---

    # Список реально присутствующих в файле колонок по рекламе
    existing_promotion_columns = [col for col in df_final_costs.columns if col in promotion_columns]
    # Колонки для расчёта затрат (все, кроме явно исключённых)
    cost_columns = [col for col in df_final_costs.columns if col not in columns_to_exclude_from_final_costs]

    # Маппинг название столбца -> буква столбца Excel
    col_map = {col: get_column_letter(idx + 1) for idx, col in enumerate(df_final_costs.columns)}
    start_row = 2
    end_row = final_ws.max_row

    # Колонка, от которой считаем маржинальность
    marg_sales = 'Продажи - Возвраты (руб)'
    marg_amount = 'Продажи - Возвраты (шт)'

    for row in range(start_row, end_row + 1):
        # --- "Затраты на рекламу (руб)": сумма по вычисленным колонкам
        if existing_promotion_columns:
            promo_cells = ','.join(f"{col_map[col]}{row}" for col in existing_promotion_columns)
            promo_formula = f"=SUM({promo_cells})"
            final_ws[f"{col_map['Затраты на рекламу (руб)']}{row}"].value = promo_formula
        else:
            final_ws[f"{col_map['Затраты на рекламу (руб)']}{row}"].value = "0"

        # --- "% Затрат на рекламу" (через IFERROR, чтобы не было #ЗНАЧ!)
        cell_promo = f"{col_map['Затраты на рекламу (руб)']}{row}"
        cell_sales_minus_returns = f"{col_map[marg_sales]}{row}"
        formula_percent = (
            f'=IFERROR(ABS({cell_promo}/{cell_sales_minus_returns}*100), "")'
        )
        final_ws[f"{col_map['% Затрат на рекламу']}{row}"].value = formula_percent

        # Столбец с себестоимостью
        cell_sebestoimost = f"{col_map['Себестоимость']}{row}"

        # --- Себестоимость * Факт продаж (IFERROR — пропускать ошибки)
        cell_fact_qty = f"{col_map[marg_amount]}{row}"
        formula_sebestoimost_fact = (
            f'=IFERROR(IF(OR({cell_sebestoimost}=0, {cell_sebestoimost}=""), "", (-1)* ABS({cell_sebestoimost}*{cell_fact_qty})), "")'
        )
        final_ws[f"{col_map['Себестоимость * Факт продаж']}{row}"].value = formula_sebestoimost_fact

        # --- "Маржинальность руб": продажи-возвраты минус сумма всех затрат (IFERROR — пропускать ошибки)
        base_sales = f"{col_map[marg_sales]}{row}"
        cost_sum = '+'.join(f"{col_map[col]}{row}" for col in cost_columns)
        margin_formula = (
            f'=IFERROR(IF(OR({cell_sebestoimost}=0,{cell_sebestoimost}=""),"",{base_sales}+{cost_sum}), "")'
        )
        final_ws[f"{col_map['Маржинальность руб']}{row}"].value = margin_formula
        # sum_args = [base_sales] + [f"{col_map[col]}{row}" for col in cost_columns]
        # cost_sum_str = ",".join(sum_args)
        # margin_formula = (
        #     f'=IFERROR(IF(OR({cell_sebestoimost}=0, {cell_sebestoimost}=""), "", SUM({cost_sum_str})), "")'
        # )
        # final_ws[f"{col_map['Маржинальность руб']}{row}"].value = margin_formula

        # --- "Маржинальность %" (через IFERROR)
        cell_marg_rub = f"{col_map['Маржинальность руб']}{row}"
        cell_sales = f"{col_map[marg_sales]}{row}"
        formula_percent_margin = (
            f'=IFERROR(ABS({cell_marg_rub}/{cell_sales})*SIGN({cell_marg_rub}), "")'
        )
        final_ws[f"{col_map['Маржинальность %']}{row}"].value = formula_percent_margin

        # --- "Маржинальность на ед. товара руб" (через IFERROR)
        cell_qty = f"{col_map[marg_amount]}{row}"
        formula_na_ed = (
            f'=IFERROR(ABS({cell_marg_rub}/{cell_qty})*SIGN({cell_marg_rub}), "")'
        )
        final_ws[f"{col_map['Маржинальность на ед. товара руб']}{row}"].value = formula_na_ed        # Форматирование столбца Маржинальность % в формат процентного числа

        # --- Формула: "Маржа от себестоимости %" (IFERROR)
        cell_marg_rub = f"{col_map['Маржинальность руб']}{row}"
        cell_cost_fact = f"{col_map['Себестоимость * Факт продаж']}{row}"

        formula_margin_from_cost = (
            f'=IFERROR({cell_marg_rub} / ABS({cell_cost_fact}) * 100, "")'
        )
        final_ws[f"{col_map['Маржа от себестоимости %']}{row}"].value = formula_margin_from_cost

        # --- Формула: "Приход денег на РС"
        # Для "Приход денег на РС" исключаем себестоимость * факт продаж:
        prihod_cols = [
            col for col in cost_columns if col != 'Себестоимость * Факт продаж'
        ]
        prihod_base_formula = (
            f"{base_sales}+" + f"+".join(f"{col_map[col]}{row}" for col in prihod_cols)
            if prihod_cols else base_sales
        )
        prihod_rs_formula = (
            f'=IFERROR({prihod_base_formula}, "")'
        )
        final_ws[f"{col_map['Приход денег на РС']}{row}"].value = prihod_rs_formula

        try:
            # Получаем букву столбца для "Маржинальность %"
            percent_col_letter = col_map['Маржинальность %']
            # Диапазон строк (от row_start до svod_len + 1)
            cell_range = final_ws[f"{percent_col_letter}{start_row}:{percent_col_letter}{end_row + 1}"]
            for row in cell_range:
                for cell in row:
                    cell.number_format = '0%'
        except KeyError:
            logger.warning('Столбец "Маржинальность %" не найден в col_map для форматирования процентов')

    # --- Условное форматирование столбцов с маржинальностью ---

    # Список названий столбцов маржинальности
    margin_columns = [
        "Маржинальность руб",
        "Маржинальность %",
        "Маржинальность на ед. товара руб",
        "Маржа от себестоимости %",
        "Маржа по артикулу, %"
        "Приход денег на РС"
    ]

    # Цвета
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    green_font = Font(color='006100', bold=True)
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    red_font = Font(color='9C0006', bold=True)

    for col_name in margin_columns:
        if col_name in col_map:
            col_letter = col_map[col_name]
            rng = f"{col_letter}{start_row}:{col_letter}{end_row}"

            # Формула для первой строки диапазона (например, A2)
            first_cell = f"{col_letter}{start_row}"
            rule_pos = FormulaRule(
                formula=[f'AND(NOT(ISTEXT({first_cell})), {first_cell}>0)'],
                fill=green_fill, font=green_font)
            rule_neg = FormulaRule(
                formula=[f'AND(NOT(ISTEXT({first_cell})), {first_cell}<0)'],
                fill=red_fill, font=red_font)
            final_ws.conditional_formatting.add(rng, rule_pos)
            final_ws.conditional_formatting.add(rng, rule_neg)


    # Переименовываем некоторые заголовки столбцов
    rename_map = {
        'Маржинальность %': 'Маржинальность руб / Продажи минус возвраты руб с СПП * 100',
        'Маржинальность на ед. товара': 'Маржинальность руб / Продажи минус возвраты шт * 100',
        'Маржа по артикулу %': 'средняя маржинальность по артикулу в руб., / на среднюю сумму по Продажи минус возвраты руб с СПП'
    }

    # Закрепляем первую строку
    final_ws.freeze_panes = 'A2'




    # --- ШАГ 4. Форматирование всех листов (без формул) ---
    # Задаем стиль границы
    thin_border = Side(border_style="thin", color="000000")
    # Исключаем колонки из числового форматирования
    exclude_columns = ["SKU", "Артикул продавца", 'Размер']

    for worksheet_name in wb.sheetnames:
        ws = wb[worksheet_name]

        # Подбор ширины столбцов
        for column in ws.columns:
            column_letter = column[0].column_letter
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column)
            ws.column_dimensions[column_letter].width = (max_length + 2) * 1.2

        # Включаем фильтр Excel
        ws.auto_filter.ref = ws.dimensions

        # Получаем заголовки и их индексы
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        exclude_idx = {idx for idx, header in enumerate(headers) if header in exclude_columns}

        # Границы диапазона
        min_row = ws.min_row
        max_row = ws.max_row
        min_col = ws.min_column
        max_col = ws.max_column

        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for idx, cell in enumerate(row):
                # Граница для каждой ячейки
                cell.border = Border(top=thin_border, bottom=thin_border, left=thin_border, right=thin_border)
                # Шапка: жирный текст
                if cell.row == 1:
                    cell.font = Font(name="Calibri", size=11, bold=True)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                # Формат "число" только если не исключён, не шапка и это действительно число
                elif idx not in exclude_idx and isinstance(cell.value, (int, float)):
                    if float(cell.value).is_integer():
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = '#,##0.00'

    # Сохраняем файл
    wb.save(filepath_finance_svod)
    # Выводим сообщение об успешном сохранении
    # logger.info(f"Finance report saved in: {filepath_finance_svod}")

    return filepath_finance_svod


# %% Вызов всех функций
if __name__ == "__main__":

    # Имя файла с отчетом
    # report_file_name = None
    report_file_name = "Отчет_по_начислениям_01_02_2026_28_02_20266.xlsx"

    # Парсим аргументы командной строки
    finance_cli = cli_finance(parser)

    # Получаем даты из командной строки
    date_start = finance_cli['start_date']
    date_end = finance_cli['end_date']

    # Получаем имя файла отчета по начислениям Ozon из командной строки
    report_file_path = finance_cli['finance_report_file_path']

    # Генерируем диапазон дат для текущего отчета
    date_start, date_end = generate_date_range(
        # date_start,
        # date_end,
        date_start = '2026-02-01',
        date_end='2026-02-28'
    )

    # Создаем директорию для текущего отчета
    current_report_dir, date_report = create_current_report_dir(client_name, date_start, date_end)
    # Ищем файл с отчетом
    # finance_report_filepath = rename_finance_report_file(current_report_dir, client_name, date_start, date_end)
    finance_report_filepath = get_ozon_report_file_path(
        report_file_path,
        default_path_download_finance_report,
        report_file_name
    )
    # Считывание файла с фин. отчетом
    df_finance_report = read_finance_report_file(finance_report_filepath)
    # Выгрузка отчета о реализации продаж, возвратов и баллов Озон
    # df_realization_report = calc_sales_returns(date_start, date_end)
    # Расчет расходов с привязкой к SKU
    result_sku_costs = calc_sku_costs(df_finance_report)
    # Расчет расходов по поставкам с разбивкой на SKU
    result_supplies_costs = calc_supplies_costs(df_finance_report)
    # Выгружаем список товаров (текущий)

    df_products_processed = get_ozon_product(headers, to_save=False)

    # Формируем список SKU с артикулами из расходов, которые уже разбиты по SKU
    # df_sku_list_all = create_all_sku_df(
    #     result_sku_costs,
    #     result_supplies_costs,
    #     df_products_processed
    # )
    # Расчет расходов по рекламным кампаниям с привязкой к SKU
    # task_id = 'bb511716-e864-443d-94c7-66b326c042eb'
    result_companies_costs = calc_companies_costs(
        date_start,
        date_end,
        date_report,
        df_finance_report,
        df_products_processed,
        client_id_performance,
        client_secret_performance,
        # task_id=task_id
    )
    # Объединение расходов по SKU
    df_sku_costs_all_stats = merge_sku_costs_v2(
        result_sku_costs,
        result_companies_costs,
        result_supplies_costs,
        df_products_processed
    )
    # Расчет расходов по типам начислений, которые не удалось привязать к SKU
    result_all_costs = calc_other_costs(
        df_finance_report,
        df_sku_costs_all_stats,
        result_sku_costs,
        result_companies_costs,
        result_supplies_costs
    )

    # Считываем справочную таблицу
    # catalog = read_catalog(client_name)
    # Добавляем данные из справочной таблицы
    df_all_costs_with_catalog = add_data_from_catalog(result_all_costs)
    # Расчет колонки с итоговыми расходами
    df_final_costs, promotion_columns, columns_to_exclude_from_final_costs = calc_final_costs(df_all_costs_with_catalog)
    # Расчет сводной по всем операциям
    df_final_costs_by_operation = create_svod_by_operations(df_final_costs)
    # Сохранение в excel
    filepath_finance_svod = save_finance_svod_to_excel(
        date_report,
        df_finance_report,
        df_final_costs_by_operation,
        result_sku_costs,
        df_final_costs,
        result_all_costs,
        result_companies_costs,
        result_supplies_costs,
        date_start,
        date_end
    )
    # Функция форматирования Excel (жестко)
    format_excel(date_report, df_final_costs, promotion_columns, columns_to_exclude_from_final_costs)
    # Логируем информацию о завершении расчета
    print(f"\033[32m\033[47mDone calculating Finance Report for Client {client_name}")
    logger.info(f"Finance Report saved in: {filepath_finance_svod}")

# %%
