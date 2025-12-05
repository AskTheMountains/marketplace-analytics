
# %% Определение функций
import argparse
import glob
import os
import time
from datetime import datetime, timedelta

import numpy as np
import openpyxl
import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter
pd.set_option('future.no_silent_downcasting', True)
pd.options.mode.chained_assignment = None  # default='warn'

# Вспомогательные функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_date_range
# Некоторые константы
from wb.scripts.constants import (
    client_name, headers, marketplace_dir_path, k_currency_to_rub, catalog_finace_svod_columns
)
# Уже написанные ранее функции выгрузки данных из АПИ
from wb.scripts.uploadDataFromWB import getOrdersWB
# Функция выгрузки для АПИ Продвижение
from wb.scripts.uploadDataPerformanceWB import get_costs_history
from wb.scripts.wb_generic_functions import get_company_stats_v2
from wb.scripts.wb_generic_functions import(
    upload_realization_report as upload_realization_report_new,
    convert_currency
)

# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Create finance report (WB)",
    usage=(
        "python.exe -m wb.scripts.calc_svod_finance_report_WB_v2 [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.calc_svod_finance_report_WB_v2 2024-11-13 2024-11-20\n\n"
        "  python.exe -m wb.scripts.calc_svod_finance_report_WB_v2 --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m wb.scripts.calc_svod_finance_report_WB_v2 2024-11-13 2024-11-20 -c 3"
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
def create_finance_report_dir(client_name, date_start, date_end):
    # Переводим дату в нужный формат
    date_folder = (
        pd.to_datetime(date_start)
        .strftime('%Y_%m')
    )
    # Формируем путь к директории
    current_report_dir = (
        marketplace_dir_path
        / 'Clients'
        / client_name
        / 'FinanceReports'
        / date_folder
    )
    # Создаем папку, если её нет
    current_report_dir.mkdir(parents=True, exist_ok=True)

    return current_report_dir, date_folder


# Функция выгрузки отчета о реализации по апи
def upload_realization_report(headers, date_start, date_end):
    logger.info(f"Uploading realization report for client {client_name} for dates {date_start} - {date_end}")

    # Параметры запроса
    params_realization_report = {
    "dateFrom": date_start,
    "dateTo": date_end,
    "rrdid": 0
}
    # Запрос к апи
    resp_data_realization_report = requests.get("https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod",
                                                headers=headers, params=params_realization_report).json()
    # Создаем df с отчетом
    df_realization_report = pd.DataFrame(resp_data_realization_report)
    # Добавляем столбцы года и месяца, за который был сформирован отчет
    df_realization_report['year'] = datetime.strptime(date_start, '%Y-%m-%dT%H:%M:%S').year
    df_realization_report['month'] = datetime.strptime(date_start, '%Y-%m-%dT%H:%M:%S').month
    # Добавляем столбец с id строки
    df_realization_report['id'] = np.arange(1, df_realization_report.shape[0] + 1)
    # Перемещаем столбец с id в начало отчета
    df_realization_report.insert(0, 'id', df_realization_report.pop('id'))

    logger.info("Finished uploading realization report")

    return df_realization_report


# Функция выгрузки отчета о реализации по апи
def upload_realization_report_v2(headers, date_start, date_end):

    # df, куда будем помещать итоговый результат
    df_realization_report = pd.DataFrame()
    # df, возвращаемый в методе API (делаем его не пустым для начала цикла)
    df_realization_report_api = pd.DataFrame({'A': [1]})
    rrd_id = 0

    # Цикл до тех пор, пока выгружаемый отчет не будет пустым
    while not df_realization_report_api.empty:
        # Параметры запроса
        params_realization_report = {
            "dateFrom": date_start,
            "dateTo": date_end,
            "rrdid": rrd_id
        }
        logger.info(
            f"\nUploading realization report for client\n"
            f"{client_name}\n"
            f"Request dates: {date_start} - {date_end}\n"
            f"rrd_id: {rrd_id}"
        )

        # Запрос к апи
        resp_data_realization_report = (
            requests
            .get(
                "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod",
                headers=headers,
                params=params_realization_report
            )
            .json()
        )
        # Создаем df с отчетом
        df_realization_report_api = pd.DataFrame(resp_data_realization_report)

        # Добавляем инфу о периоде выгрузки текущего отчета
        if not df_realization_report_api.empty:
            # Переводим колонки в формат datetime
            date_from_column = pd.to_datetime(df_realization_report_api['date_from'])
            date_to_column = pd.to_datetime(df_realization_report_api['date_to'])
            # Переводим в читаемый строковый формат
            min_date = date_from_column.min().strftime('%Y-%m-%d')
            max_date = date_to_column.max().strftime('%Y-%m-%d')
            # Получаем id строки отчета
            rrd_id = int(df_realization_report_api['rrd_id'].iloc[-1])
            logger.info(
                f"\nCurrent chunk:\n"
                f"date_from min: {min_date}\n"
                f"date_to max: {max_date}\n"
                f"rrd_id: {rrd_id}"
            )

        #  Добавляем столбцы года и месяца, за который был сформирован отчет
        # df_realization_report_api['year'] = datetime.strptime(date_start, '%Y-%m-%dT%H:%M:%S').year
        # df_realization_report_api['month'] = datetime.strptime(date_start, '%Y-%m-%dT%H:%M:%S').month
        # Добавляем столбец с id строки
        df_realization_report_api['id'] = np.arange(1, df_realization_report_api.shape[0] + 1)
        # Перемещаем столбец с id в начало отчета
        df_realization_report_api.insert(0, 'id', df_realization_report_api.pop('id'))
        # Объединяем с предыдущим проходом цикла
        df_realization_report = pd.concat([
            df_realization_report,
            df_realization_report_api
        ])

        # Число строк отчета, которое вернулось из API
        report_row_length = df_realization_report_api.shape[0]
        logger.info(f"Realization report rows: {report_row_length}")
        # Если количество строк отчета превышает 100к,
        # то подключаем логику выгрузки отчета частями
        if report_row_length >= 100000:
            logger.info(f"Returned 100 000 rows, uploading realization report by parts")
            # Получаем строку отчета, которую будем передавать в следующий запрос
            rrd_id = int(df_realization_report_api['rrd_id'].max())
                # Ждем одну минуту перед следующим запросом
            logger.info(f"Waiting 1 minute before next request")
            time.sleep(60)
        # Если отчет пришел пустой, то выгружены все строки, выходим из цикла
        else:
            break

    # Сбрасываем index после concat
    df_realization_report = df_realization_report.reset_index(drop=True)

    logger.info("Finished uploading realization report")

    return df_realization_report


# Функция чтения минимальной и максимальной даты отчета
def read_report_dates(df_realization_report):
    # Выбираем мин. и макс. даты из отчета о реализации
    date_start_ = min(df_realization_report['date_from'])
    date_end_ = max(df_realization_report['date_to'])
    # Создаем нужный формат даты (нужна для некоторых методов апи)
    date_start_api = date_start_ + 'T00:00:00'
    date_end_api = date_end_ + 'T23:59:59'
    # Создаем df с датами
    df_report_dates = pd.DataFrame({'date_start': date_start_,
                                    'date_end': date_end_,
                                    'date_start_api': date_start_api,
                                    'date_end_api': date_end_api},
                                    index=[0])

    return df_report_dates


# Функция переименования столбцов отчета по реализации на русский
def rename_report_columns(df_realization_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_ = df_realization_report.copy()
    # Считывание файла с соответствием колонки из апи её описанию из документации
    api_report_columns = pd.read_csv(f'{marketplace_dir_path}/scripts/Столбцы отчета о реализации из апи.csv', sep=';')
    # Создаем df с колонками из отчета по апи
    df_realization_report_columns = pd.DataFrame({'api_name': df_realization_report_.columns})
    # Мерджим с колонками отчета, который выгрузили по апи
    all_columns = df_realization_report_columns.merge(api_report_columns,
                                                            how='outer',
                                                            on='api_name',
                                                            indicator=True)
    # Находим общие колонки
    columns_to_rename = all_columns.loc[all_columns['_merge'] == 'both', ['api_name', 'ru_name']]
    # Переименовываем колонки
    df_realization_report_ = df_realization_report_.rename(columns=columns_to_rename.set_index('api_name')['ru_name'])
    # Старый способ переименовывания колонок через цикл по каждой колонке
    # for i in range(api_report_columns.shape[0]):
    #     if api_report_columns['api_name'][i] in df_realization_report.columns_:
    #         df_realization_report_ = df_realization_report_.rename(columns={api_report_columns['api_name'][i]: api_report_columns['ru_name'][i]})

    df_realization_report_renamed = df_realization_report_.copy()

    return df_realization_report_renamed


# Функция создания доп. колонок для расчетов
def add_new_columns(df_realization_report_renamed):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_new_columns = df_realization_report_renamed.copy()
    # Создаем столбец с id строки
    df_realization_report_new_columns['id'] = np.arange(1, df_realization_report_new_columns.shape[0] + 1)
    # Создание колонок, которые не всегда приходят с апи
    variable_api_columns = [
        'Обоснование штрафов и доплат',
        'Возмещение издержек по перевозке',
        'Возмещение издержек по эквайрингу',
        'Тип платежа за Эквайринг/Комиссии за организацию платежей',
        'Вознаграждение с продаж до вычета услуг поверенного, без НДС',
        'Сумма, удержанная за начисленные баллы программы лояльности',
        'Компенсация скидки по программе лояльности'
    ]
    for col in variable_api_columns:
        if col not in df_realization_report_new_columns.columns:
            df_realization_report_new_columns[col] = 0

    # Перевод в datetime некоторых колонок с датой
    for col in ['Дата начала отчётного периода', 'Дата конца отчётного периода', 'Дата заказа', 'Дата продажи', 'Дата операции']:
        # Преобразуем всё в строку для корректной конвертации (это уберет смешение типов)
        df_realization_report_new_columns[col] = df_realization_report_new_columns[col].astype(str)
        # Переводим в datetime
        df_realization_report_new_columns[col] = pd.to_datetime(
            df_realization_report_new_columns[col],
            format='mixed',
            utc=True
        )
        # Убираем миллисекунды
        df_realization_report_new_columns[col] = df_realization_report_new_columns[col].dt.floor('s')
        # Убираем тайм зоны
        df_realization_report_new_columns[col] = df_realization_report_new_columns[col].dt.tz_localize(None)
    # Год и месяц из даты продажи
    df_realization_report_new_columns['Год даты продажи'] = df_realization_report_new_columns['Дата продажи'].dt.year
    df_realization_report_new_columns['Месяц даты продажи'] = df_realization_report_new_columns['Дата продажи'].dt.month

    # Переносим строки с датами в начало df
    df_realization_report_new_columns = move_columns(
        df_realization_report_new_columns,
        [
            'Номер отчёта', 'Номер строки',
            'Дата начала отчётного периода', 'Дата конца отчётного периода',
            'Дата заказа', 'Дата продажи','Дата операции',
            'Дата формирования отчёта',
            'Дата начала действия фиксации', 'Дата окончания действия фиксации',
            'Артикул WB',
            'Артикул продавца',
            'Размер',
            'Предмет',
            'Бренд',
            'Тип документа',
            'Обоснование для оплаты',
            'Обоснование штрафов и доплат',
            'Цена розничная',
            'Цена розничная с учетом согласованной скидки',
            'Сумма продаж (возвратов)',
            'К перечислению продавцу за реализованный товар',
            'Стоимость логистики',
            'Тип платежа за Эквайринг/Комиссии за организацию платежей',
            'Возмещение издержек по эквайрингу',
            'Возмещение за выдачу и возврат товаров на ПВЗ',
            'Возмещение издержек по перевозке',
            'Вознаграждение WB без НДС',
            'НДС с вознаграждения WB',
            'Штрафы',
            'Доплаты',
            'Прочие удержания/выплаты',
            'Стоимость хранения',
            'Стоимость платной приёмки',
            'Сумма, удержанная за начисленные баллы программы лояльности',
            'Компенсация скидки по программе лояльности'
        ],
        position=0,
        insert_type='after'
    )
    # Заменяем некоторые название групп расходов
    other_costs_replace_patterns = {
        # r'^Списание за отзыв .+: акция №\d+, товар \d+': 'Списания за отзывы',
        'Платная приемка': 'Платная приемка (Отчет реализации)',
        'Хранение': 'Хранение (Отчет реализации)',
        'Возмещение за выдачу и возврат товаров на ПВЗ': 'Возмещение за выдачу и возврат товаров на ПВЗ (не разбито по SKU)',
    }
    # Колонки, в которых заменяем
    cols_to_replace_patterns = [
        'Обоснование для оплаты',
        'Обоснование штрафов и доплат'
    ]
    df_realization_report_new_columns[cols_to_replace_patterns] = (
        df_realization_report_new_columns[cols_to_replace_patterns]
        .replace(
            other_costs_replace_patterns,
            regex=False
        )
    )
    # У некоторых групп расходов заменяем лишь часть значений
    other_costs_combo_replace_patterns = {
        'Возмещение издержек по перевозке/по складским операциям с товаром': 'Возмещение издержек по перевозке/по складским операциям с товаром (не разбито по SKU)'
    }
    # Там, где нет Артикула продавца или он равен '', делаем замену расходов на (не разбито по SKU)
    mask = (df_realization_report_new_columns['Артикул продавца'].isna()) | \
        (df_realization_report_new_columns['Артикул продавца'] == '')

    # Итерация по колонкам, где делаем замену
    for col in cols_to_replace_patterns:
        df_realization_report_new_columns.loc[
            mask & (df_realization_report_new_columns[col].isin(other_costs_combo_replace_patterns.keys())),
            col
        ] = df_realization_report_new_columns.loc[
            mask & (df_realization_report_new_columns[col].isin(other_costs_combo_replace_patterns.keys())),
            col
        ].map(other_costs_combo_replace_patterns)


    # Колонки, по которым считаем прочие расходы
    other_costs_columns = [
        'Стоимость логистики',
        # 'К перечислению продавцу за реализованный товар',
        'Возмещение за выдачу и возврат товаров на ПВЗ',
        # 'Возмещение издержек по эквайрингу',
        # 'Размер комиссии за эквайринг без НДС, %',
        # 'Вознаграждение WB без НДС',
        # 'НДС с вознаграждения WB',
        'Возмещение издержек по перевозке',
        'Стоимость хранения',
        'Прочие удержания/выплаты',
        'Стоимость платной приёмки',
        'Компенсация скидки по программе лояльности',
        'Штрафы',
        'Доплаты'
    ]

    # Итоговая колонка с прочими расходами
    df_realization_report_new_columns['Итого руб'] = df_realization_report_new_columns.loc[:, other_costs_columns].sum(axis=1)

    return df_realization_report_new_columns

# Функция исключения переходных недель из отчета
def filter_report_dates(date_start, date_end, df_realization_report_new_columns, filter_dates=False):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_date_filtered = df_realization_report_new_columns.copy()
    # Если указан флаг фильтра по датам, то делаем фильтр по датам
    if filter_dates:
        # Переводим даты выгрузки в Timestamp
        dt_start = pd.to_datetime(date_start, format='mixed')
        dt_end = pd.to_datetime(date_end, format='mixed')
        # Ищем id строк переходных недель
        tmp_df_transitional_weeks = df_realization_report_date_filtered.loc[
            (df_realization_report_date_filtered['Дата начала отчётного периода'] < dt_start) | (df_realization_report_date_filtered['Дата конца отчётного периода'] > dt_end),
            :]
        # Внутри переходных недель ищем id строк, у которых дата продажи не соответствует датам начала и окончания выгрузки
        transitional_weeks_ids = tmp_df_transitional_weeks.loc[(tmp_df_transitional_weeks['Дата продажи'] < dt_start) | (tmp_df_transitional_weeks['Дата продажи'] > dt_end), 'id'].to_list()
        # Убираем найденные строки из отчета
        df_realization_report_date_filtered = df_realization_report_date_filtered.loc[~df_realization_report_date_filtered['id'].isin(transitional_weeks_ids), :]
        # Делаем фильтр по дате продажи в пределах указанного диапазона
        # df_realization_report_date_filtered = df_realization_report_date_filtered.loc[(df_realization_report_date_filtered['Дата продажи'] >= dt_start) & (df_realization_report_date_filtered['Дата продажи'] <= dt_end), :]
        # Делаем reset_index после concat
        df_realization_report_date_filtered = df_realization_report_date_filtered.reset_index(drop=True)

    return df_realization_report_date_filtered


# Функция получения дат отчета
def get_report_dates(df_realization_report_date_filtered):
    # Создание копии для избежания изменений в оригинальном df
    df_dates = df_realization_report_date_filtered.copy()
    # Убираем время в Дате продажи, иначе получается очень большой список
    df_dates['Дата продажи (без времени)'] = df_dates['Дата продажи'].dt.date
    # Формируем df со списком дат в отчете
    report_dates = (
        df_dates
        .groupby(['Дата начала отчётного периода', 'Дата конца отчётного периода', 'Дата продажи (без времени)'])
        .size()
        .reset_index()
        .drop(columns=0)
    )

    return report_dates


# Функция выгрузки заказов
def get_orders(headers, date_start, date_end, to_save=False):
    # Выгружаем отчет по заказам из АПИ
    df_orders_api = getOrdersWB(headers, date_start, date_end, to_save)
    # Создаем колонку, которую считаем за заказы
    df_orders_api['Заказы шт'] = 1
    # Выбираем нужные колонки
    df_orders = df_orders_api.loc[:, df_orders_api.columns.isin([
        'date',
        'Артикул WB',
        'Артикул продавца',
        'Размер',
        'Заказы шт'
        ])
    ]

    # Переводим артикул продавца с верхний регистр
    df_orders['Артикул продавца'] = df_orders['Артикул продавца'].str.upper()

    return df_orders


# Функция расчета заказов по Артикулу и Размеру
def calc_orders_by_size(df_orders):
    df_orders_stats = (
        df_orders
        .loc[:, ['Артикул WB', 'Артикул продавца', 'Размер', 'Заказы шт']]
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'])
        .agg(**{
            'Заказы шт': ('Заказы шт', 'sum')
        })
        .reset_index()
    )

    return df_orders_stats

# Функция расчета расходов, которые уже разделены по артикулам в отчете апи
def calc_sku_costs(df_realization_report_date_filtered, groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']):
    # Создаем копию для избежания изменений в оригинальном df
    df_sku_costs = df_realization_report_date_filtered.copy()
    # Выбираем расходы, где есть артикул
    df_sku_costs = df_realization_report_date_filtered.loc[df_realization_report_date_filtered['Артикул продавца'] != '', :]

    # Заменяем пропуски в размере на 0
    df_sku_costs['Размер'] = df_sku_costs['Размер'].replace("", "0")
    # Колонки, по которым будет производиться группировка
    # groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']

    df_sku_sizes_costs_stats = pd.DataFrame()

    # Временная колонка с размерами групп
    df_sku_sizes_costs_stats['tmp_col'] = df_sku_costs.groupby(groupby_cols).size()

    # Колонка, по которой считаем продажи и возвраты в рублях
    # if client_name in ['Orsk_Combinat', 'Neva_Metall']:
    sales_returns_spp = 'Цена розничная с учетом согласованной скидки'
    # else:
    sales_returns_without_spp = 'Сумма продаж (возвратов)'
    sales_returns_to_client = 'К перечислению продавцу за реализованный товар'

    # Расчет продаж
    sales_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Продажа', 'id'].to_list()
    df_sku_sizes_costs_stats['Продажи шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб к перечислению'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Расчет возвратов
    returns_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Возврат', 'id'].to_list()
    df_sku_sizes_costs_stats['Возвраты шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб к перечислению'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )

    # Считаем итоговы суммы по Продажам и Возвратам (без разбивки по Артикулам)
    df_sku_sales_returns_by_operation = (
        df_sku_costs
        .loc[df_sku_costs['id'].isin(sales_ids + returns_ids), :]
        .groupby('Обоснование для оплаты')
        .agg(**{
            'Расходы руб с СПП': (sales_returns_spp, 'sum'),
            'Расходы руб без СПП': (sales_returns_without_spp, 'sum'),
            'Расходы руб к перечислению': (sales_returns_to_client, 'sum'),
        })
        .reset_index()
        .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    )

    # Заполняем пропуски
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_stats['Продажи минус возвраты шт'] = df_sku_sizes_costs_stats['Продажи шт'] - df_sku_sizes_costs_stats['Возвраты шт']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] = df_sku_sizes_costs_stats['Продажи руб с СПП'] - df_sku_sizes_costs_stats['Возвраты руб с СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] = df_sku_sizes_costs_stats['Продажи руб без СПП'] - df_sku_sizes_costs_stats['Возвраты руб без СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб к перечислению'] = df_sku_sizes_costs_stats['Продажи руб к перечислению'] - df_sku_sizes_costs_stats['Возвраты руб к перечислению']

    # Считаем комиссию WB
    df_sku_sizes_costs_stats['Вознаграждение WB без НДС'] = (
        df_sku_costs
        .groupby(groupby_cols)
        ['Вознаграждение WB без НДС']
        .sum()
    )
    df_sku_sizes_costs_stats['НДС с вознаграждения WB'] = (
        df_sku_costs
        .groupby(groupby_cols)
        ['НДС с вознаграждения WB'].sum()
    )

    # Изменяем знак некоторых расходов
    # change_sign_columns = [
    #     'Вознаграждение WB без НДС',
    #     'НДС с вознаграждения WB'
    # ]
    # for col in df_sku_sizes_costs_stats:
    #     if col in change_sign_columns:
    #         df_sku_sizes_costs_stats[col] = df_sku_sizes_costs_stats[col] * (-1)

    # Считаем полную комиссию WB
    df_sku_sizes_costs_stats['Полная комиссия WB'] = (
        df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] -
        df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] +
        df_sku_sizes_costs_stats['Вознаграждение WB без НДС']
    )
    # Считаем комиссию WB, %
    df_sku_sizes_costs_stats['Комиссия WB, %'] = (
        (df_sku_sizes_costs_stats['Вознаграждение WB без НДС'] + df_sku_sizes_costs_stats['НДС с вознаграждения WB']) /
        df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП']
        * 100
    )
    # Заменяем деление на 0 на пропуски
    df_sku_sizes_costs_stats['Комиссия WB, %'] = df_sku_sizes_costs_stats['Комиссия WB, %'].replace([np.inf, -np.inf], np.nan)

    # Какая-то параша
    strange_cols = ['Вознаграждение с продаж до вычета услуг поверенного, без НДС']
    for col in strange_cols:
        if col in df_sku_costs.columns:
            df_sku_sizes_costs_stats[col] = (
                df_sku_costs
                .groupby(groupby_cols)
                [col].sum()
            )
    # Считаем суммы по колонкам, которые частично разбиты по SKU
    combo_sku_columns = [
        'Возмещение за выдачу и возврат товаров на ПВЗ',
        'Возмещение издержек по перевозке'
    ]
    for col in combo_sku_columns:
        if col in df_sku_costs.columns:
            df_sku_sizes_costs_stats[f'{col} (разбито по SKU)'] = (
                df_sku_costs
                .groupby(groupby_cols)
                [col].sum()
            )
    # Считаем расходы на эквайринг
    acquiring_type_column = 'Тип платежа за Эквайринг/Комиссии за организацию платежей'
    acquiring_costs_column = 'Возмещение издержек по эквайрингу'
    acquiring_costs_ids = []
    if acquiring_type_column in df_sku_costs.columns:
        acquiring_costs_ids = (
            df_sku_costs
            .loc[
                (df_sku_costs[acquiring_type_column] != 0)
                & (df_sku_costs[acquiring_type_column] != '')
                & (~df_sku_costs[acquiring_type_column].isna()),
                'id'
            ]
            .to_list()
        )
        df_sku_acquiring_costs = df_sku_costs.loc[df_sku_costs['id'].isin(acquiring_costs_ids), :]
        for operation_type in df_sku_acquiring_costs[acquiring_type_column].unique():
            df_sku_sizes_costs_stats[f'Эквайринг: {operation_type}'] = (
                df_sku_acquiring_costs
                .loc[df_sku_acquiring_costs[acquiring_type_column] == operation_type, :]
                .groupby(groupby_cols)
                [acquiring_costs_column].sum()
                .fillna(0)
            )

    # Список с названиями и затратами по Эквайрингу
    df_sku_acquiring_costs_by_operation = (
        df_sku_acquiring_costs
        .groupby('Обоснование для оплаты')
        .agg(**{
            'Расходы руб': ('Итого руб', 'sum')
        })
        .reset_index()
        .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    )

    # Считаем суммы по Штрафам и Доплатам (по ним есть доп. детализация в столбце Обоснование штрафов и доплат)
    penalty_costs_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'].isin(['Штрафы и доплаты', 'Штраф']), 'id'].to_list()
    df_sku_penalty_costs = df_sku_costs.loc[df_sku_costs['id'].isin(penalty_costs_ids), :]
    # Тут берем столбец 'Обоснование штрафов и доплат', а не 'Обоснование для оплаты', т.к. в нем есть доп. детализация
    for operation_type in df_sku_penalty_costs['Обоснование штрафов и доплат'].unique():
        df_sku_sizes_costs_stats[operation_type] = df_sku_costs.loc[df_sku_costs['Обоснование штрафов и доплат'] == operation_type, :] \
        .groupby(groupby_cols) \
        ['Итого руб'].sum()
    # Список с названиями и затратами по Штрафам и Доплатам
    df_sku_penalty_costs_by_operation = (
        df_sku_penalty_costs
        .groupby('Обоснование штрафов и доплат')
        .agg(**{
            'Расходы руб': ('Итого руб', 'sum')
        })
        .reset_index()
        .rename(columns={'Обоснование штрафов и доплат': 'Группа расходов'})
    )

    # Считаем суммы по оставшимся расходам
    other_costs_ids = df_sku_costs.loc[~df_sku_costs['id'].isin(sales_ids + returns_ids + penalty_costs_ids + acquiring_costs_ids), 'id'].to_list()
    df_sku_other_costs = df_sku_costs.loc[df_sku_costs['id'].isin(other_costs_ids), :]
    for operation_type in df_sku_other_costs['Обоснование для оплаты'].unique():
        df_sku_sizes_costs_stats[operation_type] = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == operation_type, :] \
        .groupby(groupby_cols) \
        ['Итого руб'].sum()
    # Список с названиями и затратами по оставшимся расходам
    df_sku_other_costs_by_operation = (
        df_sku_other_costs
        .groupby('Обоснование для оплаты')
        .agg(**{
            'Расходы руб': ('Итого руб', 'sum')
        })
        .reset_index()
        .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    )

    # Удаляем временную колонку
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.drop(columns=['tmp_col'])

    # Достаем доп. столбцы из index
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.reset_index()

    # Переводим артикул продавца с верхний регистр
    df_sku_sizes_costs_stats['Артикул продавца'] = df_sku_sizes_costs_stats['Артикул продавца'].str.upper()


    # Формируем list с id строк расходов по артикулам
    sku_costs_ids = sales_ids + returns_ids + penalty_costs_ids + other_costs_ids

    # Формируем df со списком расходов, которые привязаны к Артикулам
    df_sku_costs_by_operation = (
        pd.concat([
                df_sku_sales_returns_by_operation,
                df_sku_penalty_costs_by_operation,
                df_sku_other_costs_by_operation
            ],
            ignore_index=True)
        #.reset_index(drop=True)
    )

    # Формируем словарь с результатами
    result_sku_costs = {
        'df_sku_sizes_costs_stats': df_sku_sizes_costs_stats,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids
    }

    return result_sku_costs


# Функция расчета расходов, которые уже разделены по артикулам в отчете апи
def calc_sku_costs_v2(df_realization_report_date_filtered, groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']):
    # Создаем копию для избежания изменений в оригинальном df
    df_sales_returns = df_realization_report_date_filtered.copy()
    # Выбираем тип документа продажа или возврат
    # df_sales_returns = df_sales_returns.loc[df_sales_returns['Тип документа'].isin(['Продажа', 'Возврат']), :]
    # Выбираем расходы, где есть артикул
    df_sku_costs = df_sales_returns.loc[df_sales_returns['Артикул продавца'] != '', :]

    # Заменяем пропуски в размере на 0
    df_sku_costs['Размер'] = df_sku_costs['Размер'].replace("", "0")
    # Колонки, по которым будет производиться группировка
    # groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']

    df_sku_sizes_costs_stats = pd.DataFrame()

    # Временная колонка с размерами групп
    df_sku_sizes_costs_stats['tmp_col'] = df_sku_costs.groupby(groupby_cols).size()

    # Колонка, по которой считаем продажи и возвраты в рублях
    # if client_name in ['Orsk_Combinat', 'Neva_Metall']:
    sales_returns_spp = 'Цена розничная с учетом согласованной скидки'
    # else:
    sales_returns_without_spp = 'Сумма продаж (возвратов)'
    sales_returns_to_client = 'К перечислению продавцу за реализованный товар'

    # Расчет продаж
    sales_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Продажа', 'id'].to_list()
    df_sku_sizes_costs_stats['Продажи шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб к перечислению'] = (
        df_sku_costs
        #.loc[df_sku_costs['id'].isin(sales_ids), :]
        .loc[df_sku_costs['Тип документа'] == 'Продажа', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Расчет возвратов
    returns_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Возврат', 'id'].to_list()
    df_sku_sizes_costs_stats['Возвраты шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб к перечислению'] = (
        df_sku_costs
        # .loc[df_sku_costs['id'].isin(returns_ids), :]
        .loc[df_sku_costs['Тип документа'] == 'Возврат', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Находим id начислений, где сумма к перечислению не равна 0
    # и обоснование для оплаты не равна продажи или возврату,
    # поскольку мы уже включили их в продажи и возвраты к перечислению
    # чтобы не учесть эти расходы 2 раза
    sales_returns_to_client_ids = (
        df_sku_costs
        .loc[df_sku_costs[sales_returns_to_client] != 0, :]
        .loc[~df_sku_costs['Обоснование для оплаты'].isin(['Продажа', 'Возврат']), :]
        ['id']
        .to_list()
    )

    # Считаем итоговы суммы по Продажам и Возвратам (без разбивки по Артикулам)
    df_sku_sales_returns_by_operation = (
        df_sku_costs
        .loc[df_sku_costs['id'].isin(sales_ids + returns_ids + sales_returns_to_client_ids), :]
        # .groupby('Обоснование для оплаты')
        .groupby('Тип документа')
        .agg(**{
            'С СПП': (sales_returns_spp, 'sum'),
            'Без СПП': (sales_returns_without_spp, 'sum'),
            'К перечислению': (sales_returns_to_client, 'sum'),
        })
        .rename_axis('Группа расходов')
        .reset_index()
        # .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    )

    # Заполняем пропуски
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_stats['Продажи минус возвраты шт'] = df_sku_sizes_costs_stats['Продажи шт'] - df_sku_sizes_costs_stats['Возвраты шт']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] = df_sku_sizes_costs_stats['Продажи руб с СПП'] - df_sku_sizes_costs_stats['Возвраты руб с СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] = df_sku_sizes_costs_stats['Продажи руб без СПП'] - df_sku_sizes_costs_stats['Возвраты руб без СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб к перечислению'] = df_sku_sizes_costs_stats['Продажи руб к перечислению'] - df_sku_sizes_costs_stats['Возвраты руб к перечислению']

    # Считаем расходы по типу документа продажи минус возвраты
    def sales_minus_returns(df, col_sum, groupby_cols, doc_type_col='Тип документа'):
        sales = df[df[doc_type_col] == 'Продажа'].groupby(groupby_cols)[col_sum].sum()
        returns = df[df[doc_type_col] == 'Возврат'].groupby(groupby_cols)[col_sum].sum()
        return sales.sub(returns, fill_value=0)

    columns_to_calc = [
        'Вознаграждение WB без НДС',
        'НДС с вознаграждения WB',
        'Возмещение за выдачу и возврат товаров на ПВЗ',
        'Возмещение издержек по эквайрингу',
        'Компенсация скидки по программе лояльности',
    ]

    for col in columns_to_calc:
        df_sku_sizes_costs_stats[col] = sales_minus_returns(
            df_sku_costs, col, groupby_cols
        )
    # Изменяем знак некоторых расходов
    # change_sign_columns = [
    #     'Вознаграждение WB без НДС',
    #     'НДС с вознаграждения WB'
    # ]
    # for col in df_sku_sizes_costs_stats:
    #     if col in change_sign_columns:
    #         df_sku_sizes_costs_stats[col] = df_sku_sizes_costs_stats[col] * (-1)

    # Считаем полную комиссию WB
    df_sku_sizes_costs_stats['Полная комиссия WB'] = (
        df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] -
        df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] +
        df_sku_sizes_costs_stats['Вознаграждение WB без НДС']
    )
    # Считаем комиссию WB, %
    df_sku_sizes_costs_stats['Комиссия WB, %'] = (
        (df_sku_sizes_costs_stats['Вознаграждение WB без НДС'] + df_sku_sizes_costs_stats['НДС с вознаграждения WB']) /
        df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП']
        * 100
    )
    # Заменяем деление на 0 на пропуски
    df_sku_sizes_costs_stats['Комиссия WB, %'] = df_sku_sizes_costs_stats['Комиссия WB, %'].replace([np.inf, -np.inf], np.nan)

    # Считаем затраты на логистику
    logistics_column = 'Стоимость логистики'
    logistics_costs_ids = (
        df_sku_costs.loc[
            df_sku_costs[logistics_column] != 0, 'id'
        ]
        .to_list()
    )
    df_sku_sizes_costs_stats['Логистика'] = (
        df_sku_costs
        # .loc[
        #     df_sku_costs['id'].isin(logistics_costs_ids),
        #     :
        # ]
        .groupby(groupby_cols)
        [logistics_column]
        .sum()
    )


    # Какая-то параша
    strange_cols = ['Вознаграждение с продаж до вычета услуг поверенного, без НДС']
    for col in strange_cols:
        if col in df_sku_costs.columns:
            df_sku_sizes_costs_stats[col] = (
                df_sku_costs
                .groupby(groupby_cols)
                [col].sum()
            )
    # Считаем суммы по колонкам, которые частично разбиты по SKU
    combo_sku_columns = [
        # 'Возмещение за выдачу и возврат товаров на ПВЗ',
        'Возмещение издержек по перевозке'
    ]
    # Выбираем id, по которым считаем частичные расходы, чтобы не учитывать их второй раз
    partial_expense_ids = df_sku_costs.loc[
        df_sku_costs['Обоснование для оплаты'].isin(['Возмещение издержек по перевозке/по складским операциям с товаром']),
        'id'
    ].to_list()
    for col in combo_sku_columns:
        if col in df_sku_costs.columns:
            df_sku_sizes_costs_stats[f'{col} (разбито по SKU)'] = (
                df_sku_costs
                .groupby(groupby_cols)
                [col].sum()
            )

    # Собираем расходы, которые мы уже учли
    df_sku_other_costs = (
        df_sku_costs.loc[
            ~df_sku_costs['id'].isin(
                sales_ids + returns_ids + sales_returns_to_client_ids + partial_expense_ids + logistics_costs_ids
            )
        ]
    )
    # Убираем некоторые типы расходов из обоснования для оплаты, поскольку мы их уже учли
    df_sku_other_costs = (
        df_sku_other_costs
        .loc[
            ~df_sku_other_costs['Обоснование для оплаты'].isin([
                'Возмещение издержек по перевозке/по складским операциям с товаром',
                'Компенсация скидки по программе лояльности'
            ])
        ]
    )
    # Получаем id прочих расходов
    sku_other_costs_ids = df_sku_other_costs['id'].tolist()

    # В прочих расходам создаем колонку групп расходов, по которой будем считать
    df_sku_other_costs['Группа расходов'] = np.where(
        # Где столбец штрафы и доплаты не пустой и не пустая строка, берем его
        df_sku_other_costs['Обоснование штрафов и доплат'].notna() & (df_sku_other_costs['Обоснование штрафов и доплат'].astype(str).str.strip() != ''),
        df_sku_other_costs['Обоснование штрафов и доплат'],
        df_sku_other_costs['Обоснование для оплаты']
    )


    # # Считаем суммы по Штрафам и Доплатам (по ним есть доп. детализация в столбце 'Обоснование штрафов и доплат')
    # penalty_costs_ids = df_sku_costs.loc[
    #     df_sku_costs['Обоснование для оплаты'].isin(['Штрафы и доплаты', 'Штраф']),
    #     'id'
    # ].to_list()
    # df_sku_penalty_costs = df_sku_costs.loc[
    #     df_sku_costs['id'].isin(penalty_costs_ids), :
    # ]

    # # Словарь: {группа расходов: имя столбца для суммы}
    # custom_sum_cols = {
    #     'Название группы расходов': 'Столбец для суммы',
    #     # 'Штраф за что-то': 'Другая метрика',
    #     # и т.д.
    # }
    # Словарь: {группа расходов: имя столбца для суммы}
    custom_sum_cols = {
        'Компенсация ущерба': 'К перечислению продавцу за реализованный товар',
        'Добровольная компенсация при возврате': 'К перечислению продавцу за реализованный товар',
        # и т.д.
    }
    rows = []
    for op_type, df_op in df_sku_other_costs.groupby('Обоснование штрафов и доплат'):
        # Выбираем столбец для суммы
        sum_col = custom_sum_cols.get(op_type, 'Итого руб')

        # Для итоговой таблицы (без группировки по артикулу)
        doc_types_not_empty = df_op['Тип документа'].dropna().astype(str).str.strip()
        doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
        has_sales = 'Продажа' in doc_types_not_empty.values
        has_returns = 'Возврат' in doc_types_not_empty.values

        sum_sales = df_op.loc[df_op['Тип документа'] == 'Продажа', sum_col].sum()
        sum_returns = df_op.loc[df_op['Тип документа'] == 'Возврат', sum_col].sum()
        sum_empty = df_op.loc[
            df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == ''),
            sum_col
        ].sum()

        if has_sales and has_returns:
            total = sum_sales - sum_returns + sum_empty
        else:
            total = df_op[sum_col].sum()

        rows.append({
            'Группа расходов': op_type,
            'Продажа': sum_sales,
            'Возврат': sum_returns,
            'Расходы руб': total
        })

        # ---- Формируем столбцы в итоговом df с группировкой ----
        group_sum_sales = df_op[df_op['Тип документа'] == 'Продажа'].groupby(groupby_cols)[sum_col].sum()
        group_sum_returns = df_op[df_op['Тип документа'] == 'Возврат'].groupby(groupby_cols)[sum_col].sum()
        group_sum_empty = df_op[
            df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
        ].groupby(groupby_cols)[sum_col].sum()

        if has_sales and has_returns:
            result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
        else:
            result = df_op.groupby(groupby_cols)[sum_col].sum()

        # Вставляем расчёт в итоговый df (столбец с названием текущей группы расходов)
        df_sku_sizes_costs_stats[op_type] = result

    df_sku_other_costs_by_operation = pd.DataFrame(rows)

    # rows = []
    # for op_type, df_op in df_sku_penalty_costs.groupby('Обоснование штрафов и доплат'):
    #     # Выбираем столбец для суммы
    #     sum_col = custom_sum_cols.get(op_type, 'Итого руб')

    #     # Для итоговой таблицы (без группировки по артикулу)
    #     doc_types_not_empty = df_op['Тип документа'].dropna().astype(str).str.strip()
    #     doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
    #     has_sales = 'Продажа' in doc_types_not_empty.values
    #     has_returns = 'Возврат' in doc_types_not_empty.values

    #     sum_sales = df_op.loc[df_op['Тип документа'] == 'Продажа', sum_col].sum()
    #     sum_returns = df_op.loc[df_op['Тип документа'] == 'Возврат', sum_col].sum()
    #     sum_empty = df_op.loc[
    #         df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == ''),
    #         sum_col
    #     ].sum()

    #     if has_sales and has_returns:
    #         total = sum_sales - sum_returns + sum_empty
    #     else:
    #         total = df_op[sum_col].sum()

    #     rows.append({
    #         'Группа расходов': op_type,
    #         'Продажа': sum_sales,
    #         'Возврат': sum_returns,
    #         'Расходы руб': total
    #     })

    #     # ---- Формируем столбцы в итоговом df с группировкой ----
    #     group_sum_sales = df_op[df_op['Тип документа'] == 'Продажа'].groupby(groupby_cols)[sum_col].sum()
    #     group_sum_returns = df_op[df_op['Тип документа'] == 'Возврат'].groupby(groupby_cols)[sum_col].sum()
    #     group_sum_empty = df_op[
    #         df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
    #     ].groupby(groupby_cols)[sum_col].sum()

    #     if has_sales and has_returns:
    #         result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
    #     else:
    #         result = df_op.groupby(groupby_cols)[sum_col].sum()

    #     # Вставляем расчёт в итоговый df (столбец с названием текущей группы расходов)
    #     df_sku_sizes_costs_stats[op_type] = result

    # df_sku_penalty_costs_by_operation = pd.DataFrame(rows)

    # # Считаем суммы по оставшимся расходам с учётом типа документа
    # other_costs_ids = df_sku_costs.loc[
    #     ~df_sku_costs['id'].isin(sales_ids + returns_ids + penalty_costs_ids),
    #     'id'
    # ].to_list()
    # df_sku_other_costs = df_sku_costs.loc[
    #     df_sku_costs['id'].isin(other_costs_ids), :
    # ]

    # rows = []
    # for op_type, df_op in df_sku_other_costs.groupby('Обоснование для оплаты'):
    #     # Столбец для суммы
    #     sum_col = custom_sum_cols.get(op_type, 'Итого руб')

    #     # Для таблицы по всем расходам (итоговая по категориям)
    #     doc_types_not_empty = df_op['Тип документа'].dropna().astype(str).str.strip()
    #     doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
    #     has_sales = 'Продажа' in doc_types_not_empty.values
    #     has_returns = 'Возврат' in doc_types_not_empty.values

    #     sum_sales = df_op.loc[df_op['Тип документа'] == 'Продажа', sum_col].sum()
    #     sum_returns = df_op.loc[df_op['Тип документа'] == 'Возврат', sum_col].sum()
    #     sum_empty = df_op.loc[
    #         df_op['Тип документа'].isna() |
    #         (df_op['Тип документа'].astype(str).str.strip() == ''),
    #         sum_col
    #     ].sum()

    #     if has_sales and has_returns:
    #         total = sum_sales - sum_returns + sum_empty
    #     else:
    #         total = df_op[sum_col].sum()

    #     rows.append({
    #         'Группа расходов': op_type,
    #         'Продажа': sum_sales,
    #         'Возврат': sum_returns,
    #         'Расходы руб': total
    #     })

    #     # --- Формируем столбцы в df с группировкой ---
    #     group_sum_sales = df_op[df_op['Тип документа'] == 'Продажа'].groupby(groupby_cols)[sum_col].sum()
    #     group_sum_returns = df_op[df_op['Тип документа'] == 'Возврат'].groupby(groupby_cols)[sum_col].sum()
    #     group_sum_empty = df_op[
    #         df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
    #     ].groupby(groupby_cols)[sum_col].sum()

    #     if has_sales and has_returns:
    #         result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
    #     else:
    #         result = df_op.groupby(groupby_cols)[sum_col].sum()

    #     df_sku_sizes_costs_stats[op_type] = result

    # df_sku_other_costs_by_operation = pd.DataFrame(rows)

    # Удаляем временную колонку
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.drop(columns=['tmp_col'])

    # Достаем доп. столбцы из index
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.reset_index()

    # Переводим артикул продавца с верхний регистр
    df_sku_sizes_costs_stats['Артикул продавца'] = df_sku_sizes_costs_stats['Артикул продавца'].str.upper()


    # Формируем list с id строк расходов по артикулам
    # sku_costs_ids = sales_ids + returns_ids + penalty_costs_ids + other_costs_ids
    sku_costs_ids = (
        sales_ids + returns_ids + sales_returns_to_client_ids
        + partial_expense_ids + logistics_costs_ids
    )
    # Удаляем дубликаты на всякий
    sku_costs_ids = list(set(sku_costs_ids))

    # Формируем df со списком расходов, которые привязаны к Артикулам
    df_sku_costs_by_operation = (
        pd.concat([
                df_sku_sales_returns_by_operation,
                # df_sku_penalty_costs_by_operation,
                df_sku_other_costs_by_operation
            ],
            ignore_index=True)
        #.reset_index(drop=True)
    )

    # Формируем словарь с результатами
    result_sku_costs = {
        'df_sku_sizes_costs_stats': df_sku_sizes_costs_stats,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids
    }

    return result_sku_costs


# Функция расчета расходов, которые уже разделены по артикулам в отчете апи
def calc_sku_costs_v3(df_realization_report_date_filtered, groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']):
    # Создаем копию для избежания изменений в оригинальном df
    df_sales_returns = df_realization_report_date_filtered.copy()
    # Выбираем тип документа продажа или возврат
    # df_sales_returns = df_sales_returns.loc[df_sales_returns['Тип документа'].isin(['Продажа', 'Возврат']), :]
    # Выбираем расходы, где есть артикул
    df_sku_costs = df_sales_returns.loc[df_sales_returns['Артикул продавца'] != '', :]
    # Формируем id строк с расходами по артикулам
    sku_costs_ids = df_sku_costs['id'].to_list()

    # Заменяем пропуски в размере на 0
    df_sku_costs['Размер'] = df_sku_costs['Размер'].replace("", "0")
    # Колонки, по которым будет производиться группировка
    # groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']

    df_sku_sizes_costs_stats = pd.DataFrame()

    # Временная колонка с размерами групп
    df_sku_sizes_costs_stats['tmp_col'] = df_sku_costs.groupby(groupby_cols).size()

    # Колонка, по которой считаем продажи и возвраты в рублях
    # if client_name in ['Orsk_Combinat', 'Neva_Metall']:
    sales_returns_spp = 'Цена розничная с учетом согласованной скидки'
    # else:
    sales_returns_without_spp = 'Сумма продаж (возвратов)'
    sales_returns_to_client = 'К перечислению продавцу за реализованный товар'

    # Расчет продаж
    sales_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Продажа', 'id'].to_list()
    df_sku_sizes_costs_stats['Продажи шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Продажи руб к перечислению'] = (
        df_sku_costs
        #.loc[df_sku_costs['id'].isin(sales_ids), :]
        .loc[df_sku_costs['Тип документа'] == 'Продажа', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Расчет возвратов
    returns_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == 'Возврат', 'id'].to_list()
    df_sku_sizes_costs_stats['Возвраты шт'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб с СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб без СПП'] = (
        df_sku_costs.loc[df_sku_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_stats['Возвраты руб к перечислению'] = (
        df_sku_costs
        # .loc[df_sku_costs['id'].isin(returns_ids), :]
        .loc[df_sku_costs['Тип документа'] == 'Возврат', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Находим id начислений, где сумма к перечислению не равна 0
    # и обоснование для оплаты не равна продажи или возврату,
    # поскольку мы уже включили их в продажи и возвраты к перечислению
    # чтобы не учесть эти расходы 2 раза
    sales_returns_to_client_ids = (
        df_sku_costs
        .loc[df_sku_costs[sales_returns_to_client] != 0, :]
        .loc[~df_sku_costs['Обоснование для оплаты'].isin(['Продажа', 'Возврат']), :]
        ['id']
        .to_list()
    )

    # Считаем итоговы суммы по Продажам и Возвратам (без разбивки по Артикулам)
    df_sku_sales_returns_by_operation = (
        df_sku_costs
        .loc[df_sku_costs['id'].isin(sales_ids + returns_ids + sales_returns_to_client_ids), :]
        # .groupby('Обоснование для оплаты')
        .groupby('Тип документа')
        .agg(**{
            'С СПП': (sales_returns_spp, 'sum'),
            'Без СПП': (sales_returns_without_spp, 'sum'),
            'К перечислению': (sales_returns_to_client, 'sum'),
        })
        .rename_axis('Группа расходов')
        .reset_index()
        # .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    )

    # Заполняем пропуски
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_stats['Продажи минус возвраты шт'] = df_sku_sizes_costs_stats['Продажи шт'] - df_sku_sizes_costs_stats['Возвраты шт']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] = df_sku_sizes_costs_stats['Продажи руб с СПП'] - df_sku_sizes_costs_stats['Возвраты руб с СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] = df_sku_sizes_costs_stats['Продажи руб без СПП'] - df_sku_sizes_costs_stats['Возвраты руб без СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб к перечислению'] = df_sku_sizes_costs_stats['Продажи руб к перечислению'] - df_sku_sizes_costs_stats['Возвраты руб к перечислению']

    # Считаем расходы по типу документа продажи минус возвраты
    def sales_minus_returns(df, col_sum, groupby_cols, doc_type_col='Тип документа'):
        sales = df[df[doc_type_col] == 'Продажа'].groupby(groupby_cols)[col_sum].sum()
        returns = df[df[doc_type_col] == 'Возврат'].groupby(groupby_cols)[col_sum].sum()
        return sales.sub(returns, fill_value=0)

    custom_sum_cols = {
        # Столбец для суммирования: [Обоснования для оплаты, которые учитывает этот столбец]
        'Вознаграждение WB без НДС': [''],
        'НДС с вознаграждения WB': [''],
        'Стоимость логистики': ['Логистика', 'Коррекция логистики'],
        'Возмещение издержек по эквайрингу': [''],
        'Возмещение за выдачу и возврат товаров на ПВЗ': [''],
        'Возмещение издержек по перевозке': ['Возмещение издержек по перевозке/по складским операциям с товаром'],
        # 'Компенсация ущерба': ['Компенсация ущерба'],
        'Компенсация скидки по программе лояльности': ['Компенсация скидки по программе лояльности'],
    }

    rows = []
    used_cost_reasons = []

    for sum_col, reasons_list in custom_sum_cols.items():
        # Игнорируем пустые строки и NaN
        valid_reasons = [r for r in reasons_list if r is not None and str(r).strip() != '' and not pd.isna(r)]

        # Если reasons_list пустой или после фильтрации пуст — суммируем столбец по всему df_sku_costs без фильтра по "Обоснование для оплаты"
        if not valid_reasons:
            df_op = df_sku_costs.copy()
        else:
            df_op = df_sku_costs[df_sku_costs['Обоснование для оплаты'].isin(valid_reasons)]
            if df_op.empty:
                continue

        sales = df_op[df_op['Тип документа'] == 'Продажа'][sum_col].sum()
        returns = df_op[df_op['Тип документа'] == 'Возврат'][sum_col].sum()
        empty = df_op[
            df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
        ][sum_col].sum()

        doc_types_not_empty = df_op['Тип документа'].dropna().astype(str).str.strip()
        doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
        has_sales = 'Продажа' in doc_types_not_empty.values
        has_returns = 'Возврат' in doc_types_not_empty.values

        if has_sales and has_returns:
            total = sales - returns + empty
        else:
            total = df_op[sum_col].sum()

        rows.append({
            'Столбец': sum_col,
            'Обоснования для оплаты': valid_reasons if valid_reasons else None,  # None — значит рассчитываем по всему столбцу!
            'Обоснование штрафов и доплат': None,
            'Продажа': sales,
            'Возврат': returns,
            'Расходы руб': total
        })
        used_cost_reasons.extend(valid_reasons)

        # --- По группировочным колонкам (если нужно) ---
        group_sum_sales = df_op[df_op['Тип документа'] == 'Продажа'].groupby(groupby_cols)[sum_col].sum()
        group_sum_returns = df_op[df_op['Тип документа'] == 'Возврат'].groupby(groupby_cols)[sum_col].sum()
        group_sum_empty = df_op[
            df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
        ].groupby(groupby_cols)[sum_col].sum()

        if has_sales and has_returns:
            result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
        else:
            result = df_op.groupby(groupby_cols)[sum_col].sum()

        df_sku_sizes_costs_stats[sum_col] = result

    df_sku_main_costs_by_operation = pd.DataFrame(rows)

    # Добавляем к учтенным обоснованиям для оплаты те,
    # которые мы уже учли в столбце К перечислению продавцу за реализованный товар
    sales_returns_to_client_reasons = (
        df_sku_costs
        .loc[
            df_sku_costs['id'].isin(sales_returns_to_client_ids),
            'Обоснование для оплаты'
        ]
        .unique()
        .tolist()
    )

    # Ищем оставшиеся расходы
    df_sku_other_costs = (
        df_sku_costs.loc[
            ~df_sku_costs['Обоснование для оплаты'].isin(used_cost_reasons)
        ]
    )

    # Список столбцов, которые ищем для детализации (по приоритету)
    other_sum_columns = ['Штрафы', 'Доплаты', 'Прочие удержания\\Выплаты']
    rows = []

    for _, df_group in df_sku_other_costs.groupby(['Обоснование для оплаты', 'Обоснование штрафов и доплат']):
        obsh_pay = df_group['Обоснование для оплаты'].iloc[0]
        obsh_fine = df_group['Обоснование штрафов и доплат'].iloc[0]

        # Название для столбца в итоговом DataFrame
        col_name = obsh_fine if pd.notna(obsh_fine) and str(obsh_fine).strip() else obsh_pay
        if pd.isna(col_name) or str(col_name).strip() == '':
            continue

        # Проверяем: есть ли хоть где-то расходы по основным спец-столбцам
        # Если в группе ни в одном столбце нет суммы > 0, пропускаем
        sum_by_other_cols = [
            df_group[col].sum() for col in other_sum_columns
            if col in df_group.columns and not df_group[col].isnull().all()
        ]
        if not any([s > 0 for s in sum_by_other_cols]):
            # Если нет ни одной положительной суммы — пропускаем расчет этой группы
            continue

        # Столбец для суммы — тот, где он есть и хоть что-то не NaN
        used_col = next(
            (col for col in other_sum_columns if col in df_group.columns and not df_group[col].isnull().all()),
            'Итого руб'
        )

        # Определяем наличия "Продажа" и "Возврат"
        doc_types_not_empty = df_group['Тип документа'].dropna().astype(str).str.strip()
        doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
        has_sales = 'Продажа' in doc_types_not_empty.values
        has_returns = 'Возврат' in doc_types_not_empty.values

        sum_sales = df_group.loc[df_group['Тип документа'] == 'Продажа', used_col].sum()
        sum_returns = df_group.loc[df_group['Тип документа'] == 'Возврат', used_col].sum()
        sum_empty = df_group.loc[
            df_group['Тип документа'].isna() | (df_group['Тип документа'].astype(str).str.strip() == ''),
            used_col
        ].sum()

        if has_sales and has_returns:
            total = sum_sales - sum_returns + sum_empty
        else:
            total = df_group[used_col].sum()

        rows.append({
            'Столбец': used_col,
            'Обоснования для оплаты': obsh_pay,
            'Обоснование штрафов и доплат': obsh_fine,
            'Продажа': sum_sales,
            'Возврат': sum_returns,
            'Расходы руб': total
        })

        # --- Добавляем столбец в df_sku_sizes_costs_stats ---
        group_sum_sales = (
            df_group[df_group['Тип документа'] == 'Продажа']
            .groupby(groupby_cols)[used_col].sum()
        )
        group_sum_returns = (
            df_group[df_group['Тип документа'] == 'Возврат']
            .groupby(groupby_cols)[used_col].sum()
        )
        group_sum_empty = (
            df_group[
                df_group['Тип документа'].isna() |
                (df_group['Тип документа'].astype(str).str.strip() == '')
            ].groupby(groupby_cols)[used_col].sum()
        )

        if has_sales and has_returns:
            result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
        else:
            result = df_group.groupby(groupby_cols)[used_col].sum()

        df_sku_sizes_costs_stats[col_name] = result

    df_sku_other_costs_by_operation = pd.DataFrame(rows)

    # Удаляем временную колонку с размерами групп
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.drop(columns=['tmp_col'])
    # Сбрасываем индекс после группировки
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.reset_index()
    # Переводим артикул продавца в верхний регистр
    df_sku_sizes_costs_stats['Артикул продавца'] = df_sku_sizes_costs_stats['Артикул продавца'].str.upper()

    # Объединяем Итоговые суммы по SKU в один датафрейм
    df_sku_costs_by_operation = pd.concat([
        df_sku_main_costs_by_operation,
        df_sku_other_costs_by_operation
    ], ignore_index=True)

    # Переводим некоторые расходы из KZT в RUB
    columns_to_ignore_transfer = [
        'Артикул WB',
        'Артикул продавца',
        'Размер',
        'Продажи шт',
        'Возвраты шт',
        'Продажи минус возвраты шт',
    ]
    # Конвертируем валюту в рубли (если необходимо)
    df_sku_sizes_costs_stats = convert_currency(
        df_sku_sizes_costs_stats,
        # columns_to_ignore=columns_to_ignore_transfer,
        k_exchange_rate=k_currency_to_rub
    )
    df_sku_costs_by_operation = convert_currency(
        df_sku_costs_by_operation,
        k_exchange_rate=k_currency_to_rub
    )

    # Формируем словарь с результатами
    result_sku_costs = {
        'df_sku_sizes_costs_stats': df_sku_sizes_costs_stats,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids
    }

    return result_sku_costs


# Функция добавления заказов в отчет о реализации
def add_orders_to_sku_costs(result_sku_costs, df_orders):
    # Достаем df с расходами по SKU
    df_sku_costs_stats = result_sku_costs['df_sku_sizes_costs_stats'].copy()
    # Создаем копию для избежания изменений в оригинальном df
    df_orders_ = df_orders.copy()
    # Создаем колонку, по которой считаем заказы
    df_orders_['Заказы'] = 1
    # На всякий случай заполняем пропуски в размерах
    df_orders_['Размер'] = df_orders_['Размер'].fillna(0)
    # Считаем заказы на каждый артикул и размер
    df_orders_stats = (
        df_orders_
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'])
        .agg(**{
            'Заказы шт': ('Заказы', 'sum')
        })
        .reset_index()
    )
    # Объединяем заказы с расходами по SKU
    df_sku_and_orders = (
        pd.
        concat([df_sku_costs_stats, df_orders_stats])
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'])
        .sum()
        .reset_index()
    )
    # Перемещаем заказы в начало df для удобства
    df_sku_and_orders = move_columns(df_sku_and_orders, ['Заказы шт'], 'Продажи шт', 'before')
    # Считаем долю продаж в артикуле после добавления товаров
    df_sku_and_orders['Продажи в артикуле'] = df_sku_and_orders.groupby('Артикул WB')['Продажи шт'].transform('sum')
    df_sku_and_orders['Доля продаж в артикуле'] = df_sku_and_orders['Продажи шт'] / df_sku_and_orders['Продажи в артикуле']
    df_sku_and_orders['Количество размеров в артикуле'] = df_sku_and_orders.groupby('Артикул WB').transform('size')
    # Перемещаем колонки в начало df
    df_sku_and_orders = move_columns(
        df_sku_and_orders,
        ['Количество размеров в артикуле', 'Продажи в артикуле', 'Доля продаж в артикуле'],
        position='Продажи шт',
        insert_type='before'
    )

    return df_sku_and_orders



# Функция выгрузки и расчета расходов по платому хранению
def calc_paid_storage_costs(
        headers,
        df_realization_report_date_filtered,
        date_start,
        date_end
    ):

    # df, в который будем помещать итоговый результат
    df_paid_storage_stats = pd.DataFrame()

    # df, в который будем помещать результаты выгрузки отчетов апи
    df_paid_storage = pd.DataFrame()

    # Формируем недели, из которых был собран отчет о реализации и где были расходы на хранение
    paid_storage_dates = (
        df_realization_report_date_filtered
        .loc[
            df_realization_report_date_filtered['Обоснование для оплаты'].isin(['Хранение (Отчет реализации)']),
            :
        ]
        .groupby(['Дата начала отчётного периода', 'Дата конца отчётного периода'])
        .size()
        .reset_index()
        .rename(columns={
            0: 'row_count',
            'Дата начала отчётного периода': 'date_from',
            'Дата конца отчётного периода': 'date_to'
            }
        )
    )
    paid_storage_dates = paid_storage_dates.loc[paid_storage_dates['row_count'] > 0, :]
    # Если были расходу по платному хранению, начинаем выгрузку отчетов
    if paid_storage_dates.shape[0] > 0:
        # df, в который будем помещать результат апи
        df_paid_storage = pd.DataFrame()

        # # -- GPT START --
        # # Преобразование строковых дат в тип данных datetime
        # start_date = pd.to_datetime(date_start)
        # end_date = pd.to_datetime(date_end)
        # # Создаем список для хранения диапазона дат
        # date_range_paid_storage = []
        # # Генерируем даты с разницей не более 8 дней и без пересечений
        # current_date = start_date
        # while current_date <= end_date:
        #     # Устанавливаем конечную дату как текущую дату плюс 7 дней или дату окончания, если она раньше
        #     new_end_date = min(current_date + pd.Timedelta(days=7), end_date)
        #     date_range_paid_storage.append({'date_from': current_date, 'date_to': new_end_date})
        #     # Переходим к следующей начальной дате (добавляем 9 дней)
        #     current_date = new_end_date + pd.Timedelta(days=1)
        # # Создаем DataFrame
        # df_paid_storage_date_range = pd.DataFrame(date_range_paid_storage)
        # # -- GPT END --

        # Находим минимальную дату
        min_date = min(paid_storage_dates['date_from'])
        # Считаем разницу с минимальной датой в днях
        paid_storage_dates['days_diff'] = abs(paid_storage_dates['date_to'] - min_date).dt.days + 1
        # Разбиваем на интервалы по 30 дней
        paid_storage_dates['chunks'] = paid_storage_dates['days_diff'].apply(lambda x: int(x / 7) + 1)
        # Сколько всего получилось интервалов
        total_chunks = max(paid_storage_dates['chunks'])
        # total_chunks =  df_paid_acceptance_date_range.shape[0]

        # Делаем цикл по каждой из недель
        # for i in range(df_paid_storage_date_range.shape[0]):
        last_paid_storage_generation_time = None
        last_paid_storage_download_time = None

        for idx, row in paid_storage_dates.iterrows():
            # Выбираем одну неделю
            date_start_paid_storage = row['date_from'].isoformat()
            date_end_paid_storage = row['date_to'].replace(hour=23, minute=59, second=59).isoformat()

            logger.info(f"Uploading paid storage for dates {date_start_paid_storage} - {date_end_paid_storage}")

            # Параметры запроса на создание отчета
            params_paid_storage = {
                'dateFrom': date_start_paid_storage,
                'dateTo': date_end_paid_storage
            }

            # --- ДО генерации отчета ждем лимит ---
            if last_paid_storage_generation_time is not None:
                gen_wait_time = 62 - (time.time() - last_paid_storage_generation_time)
                if gen_wait_time > 0:
                    logger.info(f"Waiting {gen_wait_time:.1f} seconds before generating next paid storage report (WB limit)")
                    time.sleep(gen_wait_time)

            start_time = time.time()

            # Формируем задание на генерацию отчета
            resp_data_task_id = requests.get(
                "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage",
                headers=headers,
                params=params_paid_storage
            ).json()
            last_paid_storage_generation_time = time.time()

            # Получаем ID задания
            task_id = resp_data_task_id['data']['taskId']

            # Проверяем статус отчета
            report_status = ''
            while report_status not in ['done', 'purged', 'canceled']:
                resp_data_report_status = requests.get(
                    f"https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{task_id}/status",
                    headers=headers
                ).json()
                report_status = resp_data_report_status['data']['status']
                logger.info(
                    f"\n"
                    f"Paid storage report status: {report_status}\n"
                    f"Dates: {date_start_paid_storage} - {date_end_paid_storage}\n"
                    f"taskID: {task_id}"
                )
                # Ограничение на запросы раз в 5-10 секунд
                time.sleep(10)

            # Если отчет выгружен успешно, начинаем процесс загрузки
            if report_status == 'done':
                # --- ДО скачивания отчёта ждем лимит ---
                if last_paid_storage_download_time is not None:
                    download_wait_time = 62 - (time.time() - last_paid_storage_download_time)
                    if download_wait_time > 0:
                        logger.info(f"Waiting {download_wait_time:.1f} seconds before downloading paid storage report (WB limit)")
                        time.sleep(download_wait_time)

                # Получаем отчет
                resp_data_paid_storage_report = requests.get(
                    f"https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{task_id}/download",
                    headers=headers
                ).json()
                # Засекаем время, когда скачали отчет
                last_paid_storage_download_time = time.time()

                # Переводим отчет в df
                tmp_df_paid_storage = pd.DataFrame(resp_data_paid_storage_report)
                # Объединяем с предыдущим проходом цикла
                df_paid_storage = pd.concat([df_paid_storage, tmp_df_paid_storage])
                logger.info("Done uploading chunk")
            else:
                print(f"Error uploading paid storage report, info: {resp_data_report_status}")

            # Фиксируем, сколько времени ушло на генерацию и скачивание отчета
            elapsed_time = time.time() - start_time
            logger.info(
                f"Time spent: {elapsed_time:.1f} seconds"
            )
            # После этого ожидание уже не нужно — всё разруливается контролем времени до вызова!
        # Переименовываем некоторые столбцы для удобства
        df_paid_storage = df_paid_storage.rename(columns={
            'nmId': 'Артикул WB',
            'vendorCode': 'Артикул продавца',
            'size': 'Размер',
            'warehousePrice': 'Сумма хранения'
        })

        # Считаем стоимость хранения для каждого товара
        df_paid_storage_stats = df_paid_storage.groupby(['Артикул WB', 'Артикул продавца', 'Размер']).agg(
            Платное_хранение=('Сумма хранения', 'sum')
        ).reset_index()
        df_paid_storage_stats = df_paid_storage_stats.rename(columns={'Платное_хранение': 'Платное хранение (отдельный отчет)'})

        # Переводим артикул продавца с верхний регистр
        df_paid_storage_stats['Артикул продавца'] = df_paid_storage_stats['Артикул продавца'].str.upper()

    # Формируем словарь с результатами
    result_paid_storage = {
        'df_paid_storage': df_paid_storage,
        'df_paid_storage_stats': df_paid_storage_stats,
    }

    return result_paid_storage


# Функция объединения расходов по SKU, которые уже разбиты по размерам
def union_sku_sizes_costs(result_sku_costs, result_paid_storage):
    # Получаем df с результатами расходов по SKU
    df_sku_sizes_costs_stats = result_sku_costs['df_sku_sizes_costs_stats'].copy()
    # Получаем df с расходами на пл. хранение
    df_paid_storage_stats = result_paid_storage['df_paid_storage_stats'].copy()
    # Объединяем расходы по SKU из всех отчетов
    df_sku_sizes_costs_union = pd.concat([
        df_sku_sizes_costs_stats,
        #df_orders,
        df_paid_storage_stats
    ])
    # Колонки, по которым будет производиться группировка
    groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']
    # Считаем суммы
    df_sku_sizes_costs_merged = (
        df_sku_sizes_costs_union
        .groupby(groupby_cols)
        .sum()
        .reset_index()
    )
    # Считаем долю продаж в каждом артикуле в каждом размере (изменится, если добавились новые артикулы)
    df_sku_sizes_costs_merged['Продажи в артикуле'] = df_sku_sizes_costs_merged.groupby('Артикул WB')['Продажи шт'].transform('sum')
    df_sku_sizes_costs_merged['Доля продаж в артикуле'] = df_sku_sizes_costs_merged['Продажи шт'] / df_sku_sizes_costs_merged['Продажи в артикуле']
    df_sku_sizes_costs_merged['Количество размеров в артикуле'] = df_sku_sizes_costs_merged.groupby('Артикул WB').transform('size')
    # Перемещаем колонки в начало df
    df_sku_sizes_costs_merged = move_columns(
        df_sku_sizes_costs_merged,
        ['Количество размеров в артикуле', 'Продажи в артикуле', 'Доля продаж в артикуле'],
        position='Продажи шт',
        insert_type='before'
    )

    return df_sku_sizes_costs_merged


# Функция выгрузки и расчета расходов по платной приемке
def calc_paid_acceptance_costs(
        headers,
        df_sku_sizes_costs_merged,
        df_realization_report_date_filtered,
        date_start,
        date_end
):

     # df, в который будем помещать итоговый результат
    df_paid_acceptance_stats = pd.DataFrame()

    # df, в который будем помещать результаты выгрузки апи
    df_paid_acceptance = pd.DataFrame()

    # df, в который будем помещать результат разбивки платной приемки по артикулам
    df_paid_acceptance_sku_sizes = pd.DataFrame()

    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_ = df_realization_report_date_filtered.copy()

    # Проверяем, были ли расходы по платной приемке
    paid_acceptance_costs = df_realization_report_.loc[df_realization_report_['Обоснование для оплаты'].isin(['Платная приемка (Отчет реализации)']), :]

    if paid_acceptance_costs.shape[0] > 0:
        # Переводим в datetime
        for col in ['Дата начала отчётного периода', 'Дата конца отчётного периода']:
            paid_acceptance_costs[col] = pd.to_datetime(paid_acceptance_costs[col])

        # Разбиваем df на равные интервалы по 30 дней
        paid_acceptance_dates = (
            paid_acceptance_costs
            .groupby(['Дата начала отчётного периода', 'Дата конца отчётного периода']) \
            .size()
            .reset_index()
            .rename(columns={
                0: 'row_count',
                'Дата начала отчётного периода': 'date_from',
                'Дата конца отчётного периода': 'date_to'
            })
        )
        # Переводим в datetime
        for col in ['date_from', 'date_to']:
            paid_acceptance_dates[col] = pd.to_datetime(paid_acceptance_dates[col])

        # -- GPT START --
        # # Преобразование строковых дат в тип данных datetime
        # start_date = pd.to_datetime(date_start)
        # end_date = pd.to_datetime(date_end)
        # # Создаем список для хранения диапазона дат
        # date_range_paid_storage = []
        # # Генерируем даты с разницей не более 30 дней и без пересечений
        # current_date = start_date
        # while current_date <= end_date:
        #     # Устанавливаем конечную дату как текущую дату плюс 29 дней или дату окончания, если она раньше
        #     new_end_date = min(current_date + pd.Timedelta(days=30), end_date)
        #     date_range_paid_storage.append({'date_from': current_date, 'date_to': new_end_date})
        #     # Переходим к следующей начальной дате
        #     current_date = new_end_date + pd.Timedelta(days=1)
        # # Создаем DataFrame
        # df_paid_acceptance_date_range = pd.DataFrame(date_range_paid_storage)
        # -- GPT END --

        # Находим минимальную дату
        min_date = min(paid_acceptance_dates['date_from'])
        # Считаем разницу с минимальной датой в днях
        paid_acceptance_dates['days_diff'] = abs(paid_acceptance_dates['date_to'] - min_date).dt.days + 1
        # Разбиваем на интервалы по 30 дней
        paid_acceptance_dates['chunks'] = paid_acceptance_dates['days_diff'].apply(lambda x: int(x / 30) + 1)
        # Сколько всего получилось интервалов
        total_chunks = max(paid_acceptance_dates['chunks'])
        # total_chunks =  df_paid_acceptance_date_range.shape[0]

        # Переменные для трекинга лимитов WB
        last_report_generation_time = None
        last_report_download_time = None

        # Цикл по каждому интервалу
        for idx, row in paid_acceptance_dates.iterrows():
            # Выбираем временной интервал не больше 30 дней
            # df_dates_tmp = paid_acceptance_dates.loc[paid_acceptance_dates['chunks'] == chunk, :]
            # Ищем мин. и макс дату
            min_date_chunk = str(row['date_from'].date())
            max_date_chunk = str(row['date_to'].date())

            logger.info(f"Uploading paid acceptance report for dates {min_date_chunk} - {max_date_chunk}")
            # Переводим даты в нужный формат

            # Параметры запроса
            params_paid_acceptance = {
                'dateFrom': min_date_chunk,
                'dateTo': max_date_chunk
            }

            # --- ДО генерации отчета ждем лимит ---
            if last_report_generation_time is not None:
                gen_wait_time = 62 - (time.time() - last_report_generation_time)
                if gen_wait_time > 0:
                    logger.info(f"Waiting {gen_wait_time:.1f} seconds before generating next report due to WB limit")
                    time.sleep(gen_wait_time)

            # Засекаем время, когда отправили задание на генерацию
            start_time = time.time()

            # Формируем задание на генерацию отчета
            resp_data_paid_acceptance = (
                requests
                .get(
                    "https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report",
                    headers=headers,
                    params=params_paid_acceptance
                )
                .json()
            )
            # Засекаем время, когда отчет сгенерировался
            last_report_generation_time = time.time()

            # Получаем ID задания на генерацию отчета
            task_id = resp_data_paid_acceptance['data']['taskId']

            # Проверяем статус отчета
            report_status = ''
            while report_status not in ['done', 'purged', 'canceled']:
                report_status = (requests
                    .get(
                        f'https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/{task_id}/status',
                        headers=headers
                    )
                    .json()
                    ['data']['status']
                )
                logger.info(
                    f"\n"
                    f"Paid acceptance report status: {report_status}\n"
                    f"Dates: {min_date_chunk} - {max_date_chunk}\n"
                    f"taskID: {task_id}"
                )
                # Ограничение на запросы раз в 5 секунд
                time.sleep(6)

            # Если отчет выгружен успешно, то начинаем процесс загрузки
            if report_status == 'done':
                # --- ДО скачивания отчёта ждем лимит ---
                if last_report_download_time is not None:
                    download_wait_time = 62 - (time.time() - last_report_download_time)
                    if download_wait_time > 0:
                        logger.info(f"Waiting {download_wait_time:.1f} seconds before downloading report due to WB limit")
                        time.sleep(download_wait_time)

                # Получаем отчет
                resp_data_paid_acceptance_report = (
                    requests
                    .get(
                        f'https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/{task_id}/download',
                        headers=headers
                    )
                    .json()
                )
                # Засекаем время, когда скачали отчет
                last_report_download_time = time.time()

                # Переводим отчет в df
                tmp_df_paid_acceptance = pd.DataFrame(resp_data_paid_acceptance_report)
                # Объединяем с предыдущим проходом цикла
                df_paid_acceptance = pd.concat([df_paid_acceptance, tmp_df_paid_acceptance])
                logger.info("Done uploading chunk")
            else:
                print(f"Error uploading paid acceptance report, info: {resp_data_paid_acceptance}")

            # Фиксируем, сколько времени ушло на генерацию и скачивание отчета
            elapsed_time = time.time() - start_time
            logger.info(
                f"Time spent: {elapsed_time:.1f} seconds"
            )
            # После этого ожидание уже не нужно — всё разруливается контролем времени до вызова!

        # Если df с отчетами не пустой, начинаем обработку отчетов
        if not df_paid_acceptance.empty:
            # Переименовываем колонку с номером артикула
            df_paid_acceptance = df_paid_acceptance.rename(columns={
                'nmID': 'Артикул WB',
                'total': 'Суммарная стоимость приемки руб'
            })
            # Считаем итоговые расходы по платной приемке
            df_paid_acceptance_stats = (
                df_paid_acceptance
                .groupby('Артикул WB')
                .agg(
                    Платная_приемка=('Суммарная стоимость приемки руб', 'sum')
                )
                .reset_index()
            )
            df_paid_acceptance_stats = df_paid_acceptance_stats.rename(columns={'Платная_приемка': 'Платная приемка (отдельный отчет)'})

            # # Меняем знак на итоговых расходах на минус
            # df_paid_acceptance_stats['Платная_приемка'] = -df_paid_acceptance_stats['Платная_приемка']

            # # Считаем расходы на каждую поставку
            # df_paid_acceptance_by_incomes = df_paid_acceptance.groupby('incomeId').agg(
            #     Расходы_на_поставку=('total', 'sum')
            # ).reset_index()
            # df_paid_acceptance_by_incomes['Расходы_на_поставку'] = -df_paid_acceptance_by_incomes['Расходы_на_поставку']

            # Разбиваем список поставок на диапазоны по 1000 поставок
            # step = 1000
            # df_paid_acceptance_by_incomes['chunks'] = df_paid_acceptance_by_incomes

            # Переименовываем колонки в списке товаров для удобства
            # df_products_ = df_products.copy()
            # df_products_ = df_products_.rename(columns={'nmID': 'Артикул WB'})
            # df_products_ = df_products_.loc[:, ['Артикул WB', 'Артикул продавца', 'Размер']]

            # Получаем df с расходами по SKU, где уже есть все размеры
            df_sku_sizes_costs_merged_ = df_sku_sizes_costs_merged.copy()
            # Мерджим с расходами на приемку, чтобы получить количество размеров в одном артикуле
            df_paid_acceptance_sku_sizes = df_paid_acceptance_stats.merge(
                df_sku_sizes_costs_merged_[[
                    'Артикул WB',
                    'Артикул продавца',
                    'Размер', 'Продажи шт',
                    'Количество размеров в артикуле',
                    'Продажи в артикуле',
                    'Доля продаж в артикуле'
                ]],
                on='Артикул WB',
                how='left',
                indicator='Наличие артикула'
            )
            # Если какого-то артикула не было в списке SKU по размерам, выводим предупреждение
            missing_articles = df_paid_acceptance_sku_sizes.loc[
                df_paid_acceptance_sku_sizes['Наличие артикула'] == 'left_only',
                ['Артикул WB']
            ]
            if not missing_articles.empty:
                logger.warning(
                    f"No size found for articles \n"
                    f"{missing_articles}"
                )

            # Заполняем пропуски в артикуле и размере
            fill_values = {
                'Артикул продавца': 'Неизвестный артикул',
                'Количество размеров в артикуле': 1,
                'Продажи в артикуле': 0,
                'Доля продаж в артикуле': 0,
                'Размер': 0
            }
            for col, fill_value in fill_values.items():
                df_paid_acceptance_sku_sizes[col] = df_paid_acceptance_sku_sizes[col].fillna(fill_value)

            # Разбиваем расходы по платной приемке на долю продаж в артикуле
            df_paid_acceptance_sku_sizes['Платная приемка (отдельный отчет)'] =(
                df_paid_acceptance_sku_sizes['Платная приемка (отдельный отчет)'] *
                # df_paid_acceptance_sku_sizes['Количество размеров в артикуле']
                df_paid_acceptance_sku_sizes['Доля продаж в артикуле']
            )

            # Выбираем нужные колонки
            df_paid_acceptance_sku_sizes = df_paid_acceptance_sku_sizes[['Артикул WB', 'Артикул продавца', 'Размер', 'Платная приемка (отдельный отчет)']]

    # Выгружаем список поставок
    #     logger.info('Uploading incomes list')
    #     df_incomes = pd.DataFrame()
    #     # Цикл по каждому интервалу
    #     for chunk in paid_acceptance_dates['chunks'].unique():
    #         # Выбираем временной интервал не больше 30 дней
    #         df_dates_tmp = paid_acceptance_dates.loc[paid_acceptance_dates['chunks'] == chunk, :]
    #         # Ищем мин. и макс дату, переводим в unix timestamp
    #         min_date_chunk = min(df_dates_tmp['date_from'])
    #         max_date_chunk = max(df_dates_tmp['date_to'])
    #         min_date_chunk_unix = min_date_chunk.timestamp()
    #         max_date_chunk_unix = max_date_chunk.timestamp()

    #         logger.info(f"Uploading paid incomes for dates {str(min_date_chunk.date())} - {str(max_date_chunk.date())}")
    #         # Переводим даты в нужный формат
    #         # date_start_paid_acceptance = datetime.strptime(date_start, '%Y-%m-%dT%H:%M:%S').date().strftime('%Y-%m-%d')
    #         # date_end_paid_acceptance = datetime.strptime(date_end, '%Y-%m-%dT%H:%M:%S').date().strftime('%Y-%m-%d')

    #         # Параметры запроса
    #         next = 1 # начальное значение для while
    #         params_incomes = {'limit': 1000,
    #             'next': 0,
    #             'dateFrom': int(min_date_chunk_unix),
    #             'dateTo': int(max_date_chunk_unix)
    #             }
    #         while next > 0:
    #             print(f"Next page incomes: {next not in [0, 1]}")
    #             # Делаем запрос, получаем отчет
    #             resp_data_incomes = requests.get("https://marketplace-api.wildberries.ru/api/v3/orders", headers=headers, params=params_incomes).json()
    #             # Переводим отчет в df
    #             tmp_df_incomes = pd.DataFrame(resp_data_incomes['orders'])
    #             # Объединяем с предыдущим проходом цикла
    #             df_incomes = pd.concat([df_incomes, tmp_df_incomes])
    #             # Пагинация
    #             next = resp_data_incomes['next']
    #             params_incomes = {'limit': 1000,
    #                             'next': next,
    #                             'dateFrom': int(min_date_chunk_unix),
    #                             'dateTo': int(max_date_chunk_unix)
    #                             }

    # Если расходов по платной приемке нет, то оставляем пустой df
    # else:
    #     df_paid_acceptance_sku_sizes = pd.DataFrame()

    # Формируем словарь с результатами
    result_paid_acceptance = {
        'df_paid_acceptance_sku_sizes': df_paid_acceptance_sku_sizes,
        'df_paid_acceptance_stats': df_paid_acceptance_stats,
        'df_paid_acceptance': df_paid_acceptance
    }

    return result_paid_acceptance


# Функция объединения расходов по SKU, пл.хранению и приемке
def union_sku_paid_storage_acceptance_costs(df_sku_and_paid_storage_stats, result_paid_acceptance):
    # Получаем df с результатами расходов по SKU пл. хранению и приемку
    df_sku_and_paid_storage_stats_ = df_sku_and_paid_storage_stats.copy()
    df_paid_acceptance_sku_ = result_paid_acceptance['df_paid_acceptance_sku'].copy()
    # Объединяем расходы по SKU пл. хранению и приемке
    df_sku_paid_storage_acceptance = pd.concat([df_sku_and_paid_storage_stats_, df_paid_acceptance_sku_])
    # Если есть расходы по пл. приемке, объединяем все расходы в один df
    if not df_paid_acceptance_sku_.empty:
        # Колонки, по которым будет производиться группировка
        groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']
        # Считаем суммы
        df_sku_paid_storage_acceptance_stats = (
            df_sku_paid_storage_acceptance
            # Убираем колонки с количеством размеров в артикуле, т.к. их будем считать заново
            .loc[:, df_sku_and_paid_storage_stats_.columns.isin(['Продажи в артикуле', 'Доля продаж в артикуле', 'Количество размеров в артикуле'])]
            .groupby(groupby_cols)
            .sum()
            .reset_index()
        )
        # Считаем долю продаж в каждом артикуле в каждом размере (изменится, если добавились новые артикулы)
        df_sku_paid_storage_acceptance['Продажи в артикуле'] = df_sku_paid_storage_acceptance.groupby('Артикул WB')['Продажи шт'].transform('sum')
        df_sku_paid_storage_acceptance['Доля продаж в артикуле'] = df_sku_paid_storage_acceptance['Продажи шт'] / df_sku_paid_storage_acceptance['Продажи в артикуле']
        df_sku_paid_storage_acceptance['Количество размеров в артикуле'] = df_sku_paid_storage_acceptance.groupby('Артикул WB').transform('size')
        # Перемещаем колонки в начало df
        df_sku_paid_storage_acceptance_stats = move_columns(df_sku_paid_storage_acceptance_stats,
                                        ['Количество размеров в артикуле', 'Продажи в артикуле', 'Доля продаж в артикуле'],
                                        position='Продажи шт',
                                        insert_type='before'
        )
    # Если расходов по пл. приемке нет, то возвращаем исходный df с
    else:
        df_sku_paid_storage_acceptance_stats = df_sku_paid_storage_acceptance.copy()

    return df_sku_paid_storage_acceptance_stats


# Функция получения дат расходов по ВБ Продвижению
def get_promotion_dates(df_realization_report_date_filtered):
    # Ищем даты расходов по ВБ Продвижению
    promotion_dates = df_realization_report_date_filtered.loc[
        df_realization_report_date_filtered['Обоснование штрафов и доплат'].isin([
            'Оказание услуг «ВБ.Продвижение»',
            'Оказание услуг «WB Продвижение»'
            ]),
        ['Дата начала отчётного периода', 'Дата конца отчётного периода', 'Итого руб']
    ]
    if not promotion_dates.empty:
        # Переводим дату в нужный формат для выгрузки по апи рекламы
        date_start_promotion = (
            pd
            .to_datetime(min(promotion_dates['Дата начала отчётного периода']))
            .strftime('%Y-%m-%d')
        )
        date_end_promotion = (
            pd
            .to_datetime(max(promotion_dates['Дата конца отчётного периода']))
            .strftime('%Y-%m-%d')
        )
        # Переводим дату в нужный формат для выгрузки заказов
        date_start_orders = date_start_promotion + 'T00:00:00.000Z'
        date_end_orders = date_end_promotion + 'T23:59:59.000Z'
    else:
        date_start_promotion = None
        date_end_promotion = None
        date_start_orders = None
        date_end_orders = None

    # Формируем словарь с резульататми
    promotion_dates = {
        'date_start_promotion': date_start_promotion,
        'date_end_promotion': date_end_promotion,
        'date_start_orders': date_start_orders,
        'date_end_orders': date_end_orders,
    }

    return promotion_dates


# Функция обработки файлов по рекламным кампаниям
def parse_companies_files(date_report):

    # df, куда будем помещать итоговый результат
    df_companies_stats = pd.DataFrame()

    # Директория, в которой лежат данные по рекламным кампаниям
    companies_stats_dir = f"{current_report_dir}/{date_report}_Кампании/"
    if not os.path.exists(companies_stats_dir):
        os.mkdir(companies_stats_dir)  # Создаем, если директория не существует
    # Получаем список файлов в директории
    companies_stats_files = os.listdir(companies_stats_dir)
    # Если папка пустая, выводим сообщение об ошибке и возвращаем пустой df
    if len(companies_stats_files) == 0:
        logger.warning("Нет данных по кампаниям из ЛК")
        return df_companies_stats

    # Если в папке есть файлы с кампаниями, начинаем их обработку
    else:
        logger.info(f"Parsing companies files")
        filenames_companies = {"path": []}
        path_companies = f"{companies_stats_dir}/Статистика-*.xlsx"
        # Считывание файла с путем до него
        for file in glob.glob(path_companies):
            filenames_companies['path'].append(file)
        # Считывание только имени файла
        filenames_companies['file_name'] = [os.path.basename(x) for x in glob.glob(path_companies)]
        # Размер файла
        filenames_companies['file_size'] = [os.path.getsize(x) for x in glob.glob(path_companies)]
        # Перевод в df
        filenames_companies = pd.DataFrame(filenames_companies)
        # ID кампании
        filenames_companies['id'] = filenames_companies['file_name'].str.extract(r'-(\d+)\.xlsx$')

        # df, куда будем помещать результаты
        df_companies_parsed = pd.DataFrame()
        # Цикл по каждой кампании
        for i in range(filenames_companies.shape[0]):
            # Считываем файл со статистикой кампании
            tmp_df_company = pd.read_excel(filenames_companies['path'][i])
            # Добавляем id кампании
            tmp_df_company['id'] = filenames_companies['id'][i]
            # Удаляем последнюю строку из df
            tmp_df_company = tmp_df_company.iloc[:-1]
            # Объединяем с предыдущим проходом цикла
            df_companies_parsed = pd.concat([df_companies_parsed, tmp_df_company])

        # Сбрасываем index после concat
        df_companies_parsed = df_companies_parsed.reset_index(drop=True)

        # Считаем суммы по кампании
        df_companies_stats = (
            df_companies_parsed
            .rename(columns={
                'id': 'ID Кампании',
                'Номенклатура': 'Артикул WB',
                'Заказанные товары, шт.': 'Заказы из кампаний шт'
            })
            .groupby(['ID Кампании', 'Артикул WB'])
            .agg(**{
                'Заказы из кампаний шт': ('Заказы из кампаний шт', 'sum')
            })
            .reset_index()
        )
        # Добавляем столбец с суммарными заказами по каждой кампании
        df_companies_stats['Заказы внутри кампании шт'] = df_companies_stats.groupby(['ID Кампании'])['Заказы из кампаний шт'].transform('sum')
        # Добавляем столбец с количеством артикулов в кампании
        df_companies_stats['Количество артикулов в кампании'] = df_companies_stats.groupby(['ID Кампании']).transform('size')

        # Сохраняем обработанные данные по кампаниям в excel
        with pd.ExcelWriter(f"{companies_stats_dir}/{date_report}_Кампании_parsed.xlsx") as w:
            df_companies_parsed.to_excel(w, sheet_name='Кампании парсинг', index=False)
            df_companies_stats.to_excel(w, sheet_name='Кампании суммы', index=False)

        return df_companies_stats


# Функция обработки датафрейма со статистикой РК
def process_company_stats_df(df_companies_stats):
    # Создаем копию для избежания изменений в оригинальном df
    df_companies_stats_processed = df_companies_stats.copy()
    df_companies_stats_processed = df_companies_stats_processed.rename(columns={
        'Номер кампании': 'ID Кампании',
        'Заказы по рекламе шт (из отчета РК)': 'Заказы из кампаний шт',
        'Затраты руб': 'Затраты (Статистика РК)'
    })
    # # Добавляем столбец с количеством артикулов в кампании
    # df_companies_stats_processed['Количество артикулов в кампании'] = (
    #     df_companies_stats_processed
    #     .groupby(['ID Кампании'])['Артикул WB']
    #     .transform('nunique')
    # )
    # # Переводим колонку с ID кампании в int64
    # df_companies_stats_processed['ID Кампании'] = df_companies_stats_processed['ID Кампании'].astype(np.int64)

    # Считаем суммы затрат по артикулу
    df_companies_stats_processed = (
        df_companies_stats_processed
        .groupby(['Артикул WB'], as_index=False)
        .agg(**{
            'Заказы из кампаний шт': ('Заказы из кампаний шт', 'sum'),
            'Затраты (Статистика РК)': ('Затраты (Статистика РК)', 'sum')
        })
    )
    return df_companies_stats_processed


# Функция распределения расходов рекламных кампаний по артикулам WB
def calc_companies_by_sku_and_size(df_companies_stats_processed, df_costs_history, df_sku_sizes_costs_merged):
    # Создаем копию для избежания изменений в оригинальном df
    # df_costs_history_ = df_costs_history.copy()
    df_companies_stats_by_sku = df_companies_stats_processed.copy()
    df_sku_sizes_costs_merged_ = df_sku_sizes_costs_merged.copy()
    # Переименовываем колонки для соответствия
    # df_costs_history_ = df_costs_history_.rename(columns={
    #     'advertId': 'ID Кампании',
    #     'Расходы': 'Затраты (История затрат)'
    # })
    # Мерджим историю затрат (там, где "правильные расходы")
    # со статистикой кампаний (где содержатся артикулы кампаний)
    # df_companies_stats_by_sku = df_companies_stats_processed.merge(
    #     df_costs_history_,
    #     on='ID Кампании',
    #     how='left',
    #     indicator='Наличие кампаний'
    # )
    # # Выводим предупреждение, если расходы по каким-то кампаниям не найдены
    # missing_companies = df_companies_stats_by_sku.loc[
    #     df_companies_stats_by_sku['Наличие кампаний'] == 'left_only',
    #     ['ID Кампании', 'Заказы из кампаний шт']
    # ]
    # if not missing_companies.empty:
    #     logger.warning(f"No companies costs for companies\n"
    #                    f"{missing_companies}")
    # # Сортировка
    # df_companies_stats_by_sku = df_companies_stats_by_sku.sort_values(by=['ID Кампании', 'Артикул WB'], ignore_index=True)
    # # Заполняем пропуски в расходах
    # # (если кампании не было в Истории затрат, то считаем, что расходы по ней равны 0)
    # df_companies_stats_by_sku['Затраты (История затрат)'] = df_companies_stats_by_sku['Затраты (История затрат)'].fillna(0)
    # # Убираем кампании, которых не было в Истории затрат
    # df_companies_stats_by_sku = df_companies_stats_by_sku.loc[df_companies_stats_by_sku['Наличие кампаний'] == 'both', :]
    # # Распределяем расходы из Истории затрат на один артикул
    # df_companies_stats_by_sku['Расходы по артикулу'] = (
    #     # df_companies_stats_by_sku['Заказы из кампаний шт']
    #     # / df_companies_stats_by_sku['Заказы внутри кампании шт']
    #     # * df_companies_stats_by_sku['Расходы (из Истории затрат)']
    #     # df_companies_stats_by_sku['Расходы (из Истории затрат)'] /
    #     # df_companies_stats_by_sku['Количество артикулов в кампании']
    #     df_companies_stats_by_sku['Затраты (Статистика РК)']
    # )
    # # Заполняем пропуски
    # df_companies_stats_by_sku['Расходы по артикулу'] = df_companies_stats_by_sku['Расходы по артикулу'].fillna(0)
    # Распределяем расходы на один размер
    df_companies_stats_by_size = df_companies_stats_by_sku.merge(
        df_sku_sizes_costs_merged_[[
            'Артикул WB',
            'Артикул продавца',
            'Размер',
            'Количество размеров в артикуле',
            'Продажи в артикуле',
            'Доля продаж в артикуле',
            'Продажи шт'
        ]],
        how='left',
        on='Артикул WB',
        indicator='Наличие артикулов'
    )
    # Если какого-то артикула не было в списке SKU по размерам, выводим предупреждение
    missing_articles = df_companies_stats_by_size.loc[
        df_companies_stats_by_size['Наличие артикулов'] == 'left_only',
        ['Артикул WB']
    ]
    if not missing_articles.empty:
        logger.warning(f"No size found for articles\n"
                       f"{missing_articles}")

    # Заполняем пропуски
    fill_values = {
        'Артикул продавца': 'Неизвестный артикул',
        'Размер': 0,
        'Количество размеров в артикуле': 1,
    }
    df_companies_stats_by_size = df_companies_stats_by_size.fillna(fill_values)

    # Сортировка
    df_companies_stats_by_size = df_companies_stats_by_size.sort_values(by=['Артикул WB', 'Артикул продавца', 'Размер'], ignore_index=True)
    # Считаем общее количество продаж
    # total_sales = df_sku_sizes_costs_merged_['Продажи шт'].sum()
    # Распределяем расходы на один размер
    df_companies_stats_by_size['Расходы по размеру'] = (
        df_companies_stats_by_size['Затраты (Статистика РК)'] *
        # df_companies_stats_by_size['Количество размеров в артикуле']
        (df_companies_stats_by_size['Доля продаж в артикуле'] )
    )
    # Выбираем колонку, которая будет считаться итоговыми расходами
    df_companies_stats_by_size['Расходы по рекламным кампаниям'] = df_companies_stats_by_size['Расходы по размеру']
    # Выбираем нужные колонки
    df_companies_stats_by_size = df_companies_stats_by_size.loc[:, ['Артикул WB', 'Артикул продавца', 'Размер', 'Расходы по рекламным кампаниям']]

    return df_companies_stats_by_size


# Функция расчета расходов по рекламным кампаниям
def calc_companies_costs(df_realization_report_date_filtered, df_sku_sizes_costs_merged, date_report):
    # Ищем даты, в которых были расходы на ВБ Продвижение
    promotion_dates = get_promotion_dates(df_realization_report_date_filtered)

    # Если расходов на ВБ Продвижение не было, возвращаем пустой df
    if any(date_promotion is None for date_promotion in promotion_dates.values()):
        logger.info("No promotion costs have been found")
        df_companies_stats_by_sku = pd.DataFrame()
    # Если расходы были, то начинаем расчет расходов по рекламным кампаниям
    else:
        # Обработка данных рекламных кампаний
        df_companies_stats = parse_companies_files(date_report)
        # # Выгружаем заказы
        # df_orders = get_orders(
        #     headers,
        #     promotion_dates['date_start_orders'],
        #     promotion_dates['date_end_orders'],
        #     to_save=False
        # )
        # # Считаем заказы по размерам
        # df_orders_stats = calc_orders_by_size(df_orders)
        # Получение истории затрат
        df_costs_history = get_costs_history(
            headers,
            promotion_dates['date_start_promotion'],
            promotion_dates['date_end_promotion']
        )
        # Распределяем расходы из истории затрат по артикулам в кампании
        df_companies_stats_by_sku = calc_companies_by_sku_and_size(
            df_companies_stats,
            df_costs_history,
            df_sku_sizes_costs_merged
        )

    return df_companies_stats_by_sku


# Функция расчета расходов по рекламным кампаниям (v2)
def calc_companies_costs_v2(df_realization_report_date_filtered, df_sku_sizes_costs_merged, date_report):
    # Ищем даты, в которых были расходы на ВБ Продвижение
    promotion_dates = get_promotion_dates(df_realization_report_date_filtered)

    # Если расходов на ВБ Продвижение не было, возвращаем пустой df
    if any(date_promotion is None for date_promotion in promotion_dates.values()):
        logger.info("No promotion costs have been found")
        df_companies_stats_by_sku = pd.DataFrame()
        df_companies_stats = pd.DataFrame()
    # Если расходы были, то начинаем расчет расходов по рекламным кампаниям
    else:
        # # Обработка данных рекламных кампаний
        # df_companies_stats = parse_companies_files(date_report)
        # # Выгружаем заказы
        # df_orders = get_orders(
        #     headers,
        #     promotion_dates['date_start_orders'],
        #     promotion_dates['date_end_orders'],
        #     to_save=False
        # )
        # # Считаем заказы по размерам
        # df_orders_stats = calc_orders_by_size(df_orders)
        date_start_promotion = promotion_dates['date_start_promotion']
        date_end_promotion = promotion_dates['date_end_promotion']
        # Получение истории затрат
        df_costs_history = get_costs_history(
            headers,
            date_start_promotion,
            date_end_promotion
        )
        # Создание датафрейма с номерами кампаний для статистики
        companies_list = (
            df_costs_history['ID Кампании']
            .astype(int)
            .unique()
            .tolist()
        )
        df_input_companies = pd.DataFrame({
            'Номер кампании': companies_list,
            'date_start': [date_start_promotion] * len(companies_list),
            'date_end': [date_end_promotion] * len(companies_list),
        })
        # Получаем статистику кампаний
        df_companies_stats = get_company_stats_v2(
            headers,
            df_input_companies,
            # groupby_cols=[]
        )

        # Обрабатываем df со статистикой кампаний
        df_companies_stats_processed = process_company_stats_df(df_companies_stats)
        # Распределяем расходы из истории затрат по артикулам в кампании
        df_companies_stats_by_sku = calc_companies_by_sku_and_size(
            df_companies_stats_processed,
            df_costs_history,
            df_sku_sizes_costs_merged
        )

    # Формируем словарь с результатами
    result_companies_costs = {
        'df_companies_stats': df_companies_stats,
        'df_companies_stats_by_sku': df_companies_stats_by_sku
    }


    return result_companies_costs


# Функция расчета расходов списаний за отзывы
def calc_reviews_costs(df_realization_report_date_filtered, df_sku_sizes_costs_merged):
    # Ищем расходы списаний за отзывы
    # Создаём маску для строк по паттерну
    reviews_costs_text_pattern = r'Списание за отзыв.*акция №\d+, товар \d+$'
    # Колонка, в которой ищем списание за отзывы
    column_to_search = 'Обоснование штрафов и доплат'
    mask_reviews_costs = (
        df_realization_report_date_filtered
        [column_to_search]
        .str.contains(
            reviews_costs_text_pattern,
            na=False,
            regex=True
        )
    )
    # Делаем выборку по списаниям за отзывы
    df_reviews_sku_costs = (
        df_realization_report_date_filtered
        .loc[mask_reviews_costs, [column_to_search, 'Итого руб']]
        # Выделяем артикул из строки списания за отзыв
        .assign(
            **{
                "Артикул WB": lambda df: pd.to_numeric(
                    df[column_to_search].str.extract(r'(\d+)$')[0], errors='coerce'
                ).astype('Int64')
        })
        # Считаем суммы списаний за отзывы
        .groupby(['Артикул WB'], as_index=False)
        .agg(**{
            'Списания за отзывы': ('Итого руб', 'sum')
        })
    )
    # Если есть расходы по списаниям за отзывы, то разбиваем их по размерам
    if df_reviews_sku_costs.shape[0] > 0:
        # Получаем df с расходами по SKU
        df_sku_sizes = df_sku_sizes_costs_merged.copy()
        # Мерджим с расходами на приемку, чтобы получить количество размеров в одном артикуле
        df_reviews_sku_sizes_costs = df_reviews_sku_costs.merge(
            df_sku_sizes[[
                'Артикул WB',
                'Артикул продавца',
                'Размер', 'Продажи шт',
                'Количество размеров в артикуле',
                'Продажи в артикуле',
                'Доля продаж в артикуле'
            ]],
            on='Артикул WB',
            how='left',
            indicator='Наличие артикула'
        )
        # Если какого-то артикула не было в списке SKU по размерам, выводим предупреждение
        missing_articles = df_reviews_sku_sizes_costs.loc[
            df_reviews_sku_sizes_costs['Наличие артикула'] == 'left_only',
            ['Артикул WB']
        ]
        if not missing_articles.empty:
            logger.warning(
                f"No size found for articles \n"
                f"{missing_articles}"
            )

        # Заполняем пропуски в артикуле и размере
        fill_values = {
            'Артикул продавца': 'Неизвестный артикул',
            'Количество размеров в артикуле': 0,
            'Продажи в артикуле': 0,
            'Доля продаж в артикуле': 0,
            'Размер': 0
        }
        for col, fill_value in fill_values.items():
            df_reviews_sku_sizes_costs[col] = df_reviews_sku_sizes_costs[col].fillna(fill_value)

        # Разбиваем расходы по платной приемке на долю продаж в артикуле
        df_reviews_sku_sizes_costs['Списания за отзывы'] =(
            df_reviews_sku_sizes_costs['Списания за отзывы'] /
            df_reviews_sku_sizes_costs['Количество размеров в артикуле']
        )

        # Выбираем нужные колонки
        df_reviews_sku_sizes_costs = df_reviews_sku_sizes_costs[['Артикул WB', 'Артикул продавца', 'Размер', 'Списания за отзывы']]
    # Если расходы по списаниям за отзывы найдены не были, то возвращаем пустой df
    else:
        df_reviews_sku_sizes_costs = pd.DataFrame()

    # Ищем id строк расходов за отзывы
    reviews_sku_costs_ids = (
        df_realization_report_date_filtered
        .loc[
            df_realization_report_date_filtered[column_to_search]
            .str.contains(reviews_costs_text_pattern, na=False, regex=True),
            'id'
        ]
        .to_list()
    )

    # Формируем словарь с результатами
    result_reviews_costs = {
        'df_reviews_sku_sizes_costs': df_reviews_sku_sizes_costs,
        'reviews_sku_costs_ids': reviews_sku_costs_ids
    }

    return result_reviews_costs


# Функция объединения расходов
def merge_all_costs(
        df_sku_sizes_costs_merged,
        result_paid_acceptance,
        result_companies_costs,
        result_reviews_costs,
):
    # Получаем df с расходами по платной приемке
    df_paid_acceptance_costs = result_paid_acceptance['df_paid_acceptance_sku_sizes']
    # Получаем df с расходами по рекламным кампаниям
    df_companies_costs = result_companies_costs['df_companies_stats_by_sku']
    # Получаем df с расходами по списаниям за отзывы
    df_reviews_sku_sizes_costs = result_reviews_costs['df_reviews_sku_sizes_costs']
    # Объединяем полученные расходы в один df
    df_sku_sizes_costs_union = pd.concat([
        df_sku_sizes_costs_merged,
        df_paid_acceptance_costs,
        df_companies_costs,
        df_reviews_sku_sizes_costs,
    ])
    # Сводим все по артикулам после concat
    df_sku_sizes_costs_all = (
        df_sku_sizes_costs_union
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'])
        .sum()
        .reset_index()
    )

    return df_sku_sizes_costs_all


# Функция расчета расходов, которые не привязаны к артикулам (v2)
def calc_other_costs_v2(result_sku_costs, df_sku_sizes_costs_all):
    # Получаем list с id строк расходов по артикулам
    sku_costs_ids = result_sku_costs['sku_costs_ids']
    # df, в который будем помещать итоговый результат (df уже содержит расходы по sku)
    df_sku_and_other_costs = df_sku_sizes_costs_all.copy()

    # Выбираем строки из отчета, которые не привязаны к артикулам
    df_other_costs = df_realization_report_date_filtered.loc[~df_realization_report_date_filtered['id'].isin(sku_costs_ids), :]
    other_costs_ids = df_realization_report_date_filtered.loc[~df_realization_report_date_filtered['id'].isin(sku_costs_ids), 'id'].to_list()

    # Создаем столбец с типом расходов по прочим расходам
    df_other_costs['Тип расходов'] = np.where(
         (df_other_costs['Обоснование штрафов и доплат'] != '') & (~df_other_costs['Обоснование штрафов и доплат'].isna()),
         df_other_costs['Обоснование штрафов и доплат'],
         df_other_costs['Обоснование для оплаты']
    )
    # Считаем суммы по прочим расходам
    df_other_costs_by_operation = (
        df_other_costs
        .groupby('Тип расходов')
        .agg(
            Итого_прочие_расходы=('Итого руб', 'sum')
        )
        .reset_index()
        # .rename(columns={'Обоснование штрафов и доплат': 'Тип расходов'})
    )

    # Выбираем расходы, у которых есть доп. детализация в столбце "Обоснование штрафов и доплат"
    # df_other_costs_fees = (
    #     df_other_costs
    #     .loc[
    #         (df_other_costs['Обоснование штрафов и доплат'] != '') & (~df_other_costs['Обоснование штрафов и доплат'].isna()),
    #         :
    #     ]
    #     # .rename(columns={'Обоснование штрафов и доплат': 'Тип расходов'})
    # )
    # # Определяем id строк, у которых есть доп. детализация
    # other_costs_fees_ids = df_other_costs_fees['id'].to_list()

    # # Считаем итоговые суммы по Штрафам и доплатам
    # df_other_costs_fees_by_operation = (
    #     df_other_costs_fees
    #     .groupby('Обоснование штрафов и доплат')
    #     .agg(
    #         Итого_прочие_расходы=('Итого руб', 'sum')
    #     )
    #     .reset_index()
    #     .rename(columns={'Обоснование штрафов и доплат': 'Тип расходов'})
    # )

    # # Выбираем расходы, у которых нет доп. детализации
    # df_other_costs_leftover = df_other_costs.loc[~df_other_costs['id'].isin(other_costs_fees_ids), :]
    # other_costs_leftover_ids = df_other_costs_leftover['id'].to_list()

    # # Считаем итоговые суммы по прочим расходам
    # df_other_costs_leftover_by_operation = (
    #     df_other_costs_leftover
    #     .groupby('Обоснование для оплаты')
    #     .agg(
    #         Итого_прочие_расходы=('Итого руб', 'sum')
    #         )
    #     .reset_index()
    #     .rename(columns={'Обоснование для оплаты': 'Тип расходов'})
    # )

    # # Объединяем группы расходов, полученные выше
    # df_other_costs_by_operation = pd.concat([
    #         df_other_costs_fees_by_operation,
    #         df_other_costs_leftover_by_operation
    #     ],
    #     axis=0,
    #     ignore_index=True
    # )
    # Удаляем некоторые группы расходов, которые мы уже учли
    exclude_costs_list = [
        r'Списание за отзыв.*акция №\d+, товар \d+$',
        'Хранение (Отчет реализации)',
        'Платная приемка (Отчет реализации)',
        'Оказание услуг «ВБ.Продвижение»',
        'Оказание услуг «WB Продвижение»',

    ]
    exclude_costs_filter = '|'.join(f"(?:{p})" for p in exclude_costs_list)
    df_other_costs_by_operation = (
        df_other_costs_by_operation
        .loc[
            ~df_other_costs_by_operation['Тип расходов']
            .str.contains(exclude_costs_filter, regex=True, na=False),
            :
        ]
        .reset_index(drop=True)
    )
    # Добавляем прочие расходы в df с расходами по sku, разбивая их от количества продаж
    total_sales = df_sku_and_other_costs['Продажи шт'].sum()
    for operation_type in df_other_costs_by_operation['Тип расходов']:
        operation_costs = df_other_costs_by_operation.loc[df_other_costs_by_operation['Тип расходов'] == operation_type, 'Итого_прочие_расходы'].values[0]
        df_sku_and_other_costs[operation_type] = df_sku_and_other_costs['Продажи шт'] / total_sales * operation_costs

    # Удаляем некоторые группы расходов, т.к. по ним мы формируем отдельные отчеты
    # Поскольку мы выгружали отдельные отчеты по некоторым расходам,
    # мы должны убрать их из прочих расходов, чтобы не было дубликатов
    # Да, удаляем два раза, потому что я дебил и мне пох
    other_costs_duplicate_flags = {
        'Платное хранение (отдельный отчет)': ['Хранение (Отчет реализации)'],
        'Платная приемка (отдельный отчет)': ['Платная приемка (Отчет реализации)'],
        'Расходы по рекламным кампаниям': ['Оказание услуг «ВБ.Продвижение»', 'Оказание услуг «WB Продвижение»']
    }
    for separate_report_cost, realization_report_cost in other_costs_duplicate_flags.items():
        if separate_report_cost in df_sku_and_other_costs.columns:
            df_sku_and_other_costs.drop(columns=realization_report_cost, inplace=True, errors='ignore')

    # Старый способ удаления повторяющихся групп расходов
    # df_sku_and_other_costs = df_sku_and_other_costs.loc[:, ~df_sku_and_other_costs.columns.isin([
    #     'Оказание услуг «ВБ.Продвижение»',
    #     # 'Платная приемка',
    #     # 'Корректировка приемки',
    #     'Хранение',
    #     'Корректировка хранения',
    #     'Удержания',
    #     ])]

    # Формируем словарь с результатами
    result_other_costs = {
        'df_sku_and_other_costs': df_sku_and_other_costs,
        'df_other_costs_by_operation': df_other_costs_by_operation,
        'other_costs_ids': other_costs_ids
    }

    return result_other_costs


# Функция расчета расходов, которые не привязаны к артикулам (v2)
def calc_other_costs_v3(result_sku_costs, df_sku_sizes_costs_all):
    # Получаем list с id строк расходов по артикулам
    sku_costs_ids = result_sku_costs['sku_costs_ids']
    # df, в который будем помещать итоговый результат (df уже содержит расходы по sku)
    df_sku_and_other_costs = df_sku_sizes_costs_all.copy()

    # Выбираем строки из отчета, которые не привязаны к артикулам
    df_other_costs = df_realization_report_date_filtered.loc[
        ~df_realization_report_date_filtered['id'].isin(sku_costs_ids), :
    ]
    other_costs_ids = df_realization_report_date_filtered.loc[
        ~df_realization_report_date_filtered['id'].isin(sku_costs_ids), 'id'
    ].to_list()

    # Создаем столбец с типом расходов по прочим расходам
    df_other_costs['Тип расходов'] = np.where(
         (df_other_costs['Обоснование штрафов и доплат'] != '') & (~df_other_costs['Обоснование штрафов и доплат'].isna()),
         df_other_costs['Обоснование штрафов и доплат'],
         df_other_costs['Обоснование для оплаты']
    )

    # --- Новый блок: считаем суммы с учетом типа документа (Продажа или Возврат) ---
    # Считаем суммы по типу документа и виду расходов
    other_costs_result = []

    for op_type, df_op in df_other_costs.groupby('Тип расходов'):
        # Определяем типы документов (без пустых и NaN)
        doc_types = df_op['Тип документа'].dropna()
        doc_types = [str(x).strip() for x in doc_types]
        doc_types = [x for x in doc_types if x]
        has_sales = 'Продажа' in doc_types
        has_returns = 'Возврат' in doc_types

        if has_sales and has_returns:
            # Если есть оба – вычитаем (суммируем только непустые типы)
            sum_sales = df_op.loc[df_op['Тип документа'] == 'Продажа', 'Итого руб'].sum()
            sum_returns = df_op.loc[df_op['Тип документа'] == 'Возврат', 'Итого руб'].sum()
            # Отдельно — суммируем строки с пустым/NaN типом документа
            sum_empty = df_op.loc[df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == ''), 'Итого руб'].sum()
            total = sum_sales - sum_returns + sum_empty
            sale = sum_sales
            ret = sum_returns
        else:
            # Иначе — просто сумма по всей группе
            total = df_op['Итого руб'].sum()
            sale = df_op.loc[df_op['Тип документа'] == 'Продажа', 'Итого руб'].sum()
            ret = df_op.loc[df_op['Тип документа'] == 'Возврат', 'Итого руб'].sum()

        other_costs_result.append({
            'Тип расходов': op_type,
            'Продажа': sale,
            'Возврат': ret,
            'Итого прочие расходы': total
        })

    df_other_costs_by_operation = pd.DataFrame(other_costs_result)
    # --- конец нового блока ---

    # Удаляем некоторые группы расходов, которые мы уже учли
    exclude_costs_list = [
        r'списание за отзыв.*акция №\d+, товар \d+$',
        'хранение (отчет реализации)',
        'платная приемка (отчет реализации)',
        'оказание услуг «вб.продвижение»',
        'оказание услуг «wb продвижение»',
    ]
    exclude_costs_filter = '|'.join(f"(?:{p})" for p in exclude_costs_list)

    # Фильтрация с нормализацией, но без изменения исходного столбца
    mask = (
        ~df_other_costs_by_operation['Тип расходов'] \
        .str.strip().str.lower() \
        .str.contains(exclude_costs_filter, regex=True, na=False)
    )

    df_other_costs_by_operation = df_other_costs_by_operation.loc[mask].reset_index(drop=True)
    # Переводим в нужную валюту
    df_other_costs_by_operation = convert_currency(
        df_other_costs_by_operation,
        k_exchange_rate=k_currency_to_rub
    )
    # Добавляем прочие расходы в df с расходами по sku, разбивая их от количества продаж
    total_sales = df_sku_and_other_costs['Продажи шт'].sum()
    for operation_type in df_other_costs_by_operation['Тип расходов']:
        operation_costs = df_other_costs_by_operation.loc[
            df_other_costs_by_operation['Тип расходов'] == operation_type, 'Итого прочие расходы'
        ].values[0]
        df_sku_and_other_costs[operation_type] = df_sku_and_other_costs['Продажи шт'] / total_sales * operation_costs

    # Удаляем некоторые группы расходов, т.к. по ним мы формируем отдельные отчеты
    # Поскольку мы выгружали отдельные отчеты по некоторым расходам,
    # мы должны убрать их из прочих расходов, чтобы не было дубликатов
    # Да, удаляем два раза, потому что я дебил и мне пох
    other_costs_duplicate_flags = {
        'Платное хранение (отдельный отчет)': ['Хранение (Отчет реализации)'],
        'Платная приемка (отдельный отчет)': ['Платная приемка (Отчет реализации)'],
        'Расходы по рекламным кампаниям': ['Оказание услуг «ВБ.Продвижение»', 'Оказание услуг «WB Продвижение»']
    }
    for separate_report_cost, realization_report_cost in other_costs_duplicate_flags.items():
            if separate_report_cost in df_sku_and_other_costs.columns:
                df_sku_and_other_costs.drop(columns=realization_report_cost, inplace=True, errors='ignore')

    # Формируем словарь с результатами
    result_other_costs = {
        'df_sku_and_other_costs': df_sku_and_other_costs,
        'df_other_costs_by_operation': df_other_costs_by_operation,
        'other_costs_ids': other_costs_ids
    }

    return result_other_costs

# Функция расчета расходов, которые не привязаны к артикулам
def calc_other_costs(df_realization_report_date_filtered, result_sku_costs):
    # Получаем list с id строк расходов по артикулам
    sku_costs_ids = result_sku_costs['sku_costs_ids']
    # df, в который будем помещать итоговый результат (df уже содержит расходы по sku)
    df_sku_and_other_costs = result_sku_costs['df_sku_costs_stats'].copy()

    # Выбираем строки из отчета, которые не привязаны к артикулам
    df_other_costs = df_realization_report_date_filtered.loc[~df_realization_report_date_filtered['id'].isin(sku_costs_ids), :]
    other_costs_ids = df_realization_report_date_filtered.loc[~df_realization_report_date_filtered['id'].isin(sku_costs_ids), 'id'].to_list()

    # Выбираем расходы, у которых есть доп. детализация в столбце "Обоснование штрафов и доплат"
    df_other_costs_fees = df_other_costs.loc[
        (df_other_costs['Обоснование штрафов и доплат'] != '') & (~df_other_costs['Обоснование штрафов и доплат'].isna())
        , :]
    other_costs_fees_ids = df_other_costs_fees['id'].to_list()

    # Считаем итоговые суммы по Штрафам и доплатам
    df_other_costs_fees_by_operation = (
        df_other_costs_fees
        .groupby('Обоснование штрафов и доплат')
        .agg(
            Итого_прочие_расходы=('Итого руб', 'sum')
            )
            .reset_index()
            .rename(columns={'Обоснование штрафов и доплат': 'Тип расходов'})
    )

    # Выбираем расходы, у которых нет доп. детализации
    df_other_costs_leftover = df_other_costs.loc[~df_other_costs['id'].isin(other_costs_fees_ids), :]
    other_costs_leftover_ids = df_other_costs_leftover['id'].to_list()

    # Считаем итоговые суммы по прочим расходам
    df_other_costs_leftover_by_operation = (
        df_other_costs_leftover
        .groupby('Обоснование для оплаты')
        .agg(
            Итого_прочие_расходы=('Итого руб', 'sum')
            )
        .reset_index()
        .rename(columns={'Обоснование для оплаты': 'Тип расходов'})
    )

    # Объединяем группы расходов, полученные выше
    df_other_costs_by_operation = pd.concat([df_other_costs_fees_by_operation, df_other_costs_leftover_by_operation], axis = 0).reset_index(drop=True)

    # Добавляем прочие расходы в df с расходами по sku, разбивая их от количества продаж
    total_sales = df_sku_and_other_costs['Продажи шт'].sum()
    for operation_type in df_other_costs_by_operation['Тип расходов']:
        operation_costs = df_other_costs_by_operation.loc[df_other_costs_by_operation['Тип расходов'] == operation_type, 'Итого_прочие_расходы'].values[0]
        df_sku_and_other_costs[operation_type] = df_sku_and_other_costs['Продажи шт'] / total_sales * operation_costs

    # Удаляем столбцы с приемкой и хранением, т.к. расходы по ним мы получаем отдельным отчетом
    # for col in ['Корректировка приемки','Корректировка хранения', 'Платная приемка', 'Удержания', 'Хранение']:
    #     if col in df_sku_and_other_costs.columns:
    #         df_sku_and_other_costs = df_sku_and_other_costs.drop(columns=col)
    df_sku_and_other_costs = df_sku_and_other_costs.loc[:, df_sku_and_other_costs.columns.isin([
        'Платная приемка',
        'Корректировка приемки',
        # 'Хранение'
        # 'Корректировка хранения',
        'Удержания'
        ])]

    # Формируем словарь с результатами
    result_other_costs = {
        'df_sku_and_other_costs': df_sku_and_other_costs,
        'df_other_costs_by_operation': df_other_costs_by_operation,
        'other_costs_ids': other_costs_ids
    }

    return result_other_costs


# Функция добавления данных из справочной таблицы
def add_data_from_catalog(result_other_costs):
    # Чтение справочной таблицы
    catalog = pd.read_excel(
        marketplace_dir_path
        / "Clients"
        / client_name
        / "catalog"
        / f"Справочная_таблица_{client_name}_WB.xlsx"
    )
    # Создаем колонку размера, если её нет в справочной таблице
    if 'Размер' not in catalog.columns:
        catalog['Размер'] = '0'

    # catalog['Размер'] = pd.to_numeric(catalog['Размер'], errors='coerce').fillna(0).astype(int)
    # Создаем остальные колонки, если их нет в справочной таблице
    for col in catalog_finace_svod_columns:
        if col not in catalog.columns:
            catalog[col] = np.nan

    # Переводим артикул продавца в верхний регистр
    catalog['Артикул продавца'] = catalog['Артикул продавца'].str.upper()

    # В каждой группе артикулов искуственно создаем артикул с размером 0, если такого размера нет
    # Словарь, где для каждого нужного столбца указана агрегация
    if client_name in ['SENS', 'SENS_IP', 'SENS_KZ']:
        agg_dict = {
            'Себестоимость': 'first',
            'Предмет': 'first',
            'Цвет': 'first',
            'Статус': 'first'
    }

        for art, group in catalog.groupby('Артикул продавца'):
            if not (group['Размер'].astype(str).str.strip().replace(',', '.') == '0').any():
                new_row = {col: np.nan for col in catalog.columns}
                new_row['Артикул продавца'] = art
                new_row['Размер'] = 0

                for col, func in agg_dict.items():
                    if isinstance(func, str):
                        # Стандартные строковые pandas-функции
                        if func == 'first':
                            new_row[col] = group[col].iloc[0]
                        elif func == 'last':
                            new_row[col] = group[col].iloc[-1]
                        else:
                            new_row[col] = getattr(group[col], func)()
                    elif callable(func):
                        new_row[col] = func(group[col])
                    else:
                        new_row[col] = np.nan

                catalog = pd.concat([catalog, pd.DataFrame([new_row])], ignore_index=True)

    # Сортировка
    catalog = catalog.sort_values(['Артикул продавца', 'Размер']).reset_index(drop=True)


    # Получаем таблицу с расходами из результатов по всем расходам
    df_sku_and_other_costs = result_other_costs['df_sku_and_other_costs'].copy()

    # Создаем колонки для мерджа
    catalog['Артикул_Размер'] = catalog[['Артикул продавца', 'Размер']].apply(lambda row: '_size_'.join(row.values.astype(str)), axis=1)
    # Удаляем дубликаты на всякий случай
    catalog = catalog.drop_duplicates(subset=['Артикул_Размер'], keep='last')
    df_sku_and_other_costs['Артикул_Размер'] = df_sku_and_other_costs[['Артикул продавца', 'Размер']].apply(lambda row: '_size_'.join(row.values.astype(str)), axis=1)

    # Убираем артикул и размер
    catalog = catalog.loc[:, ~catalog.columns.isin(['Артикул продавца', 'Размер'])]
    df_sku_and_other_costs = df_sku_and_other_costs.loc[:, ~df_sku_and_other_costs.columns.isin(['Артикул продавца', 'Размер'])]

    # Мерджим расходы и справочную таблицу
    df_sku_and_other_costs_with_catalog = df_sku_and_other_costs.merge(
        catalog[['Артикул_Размер'] + catalog_finace_svod_columns],
        how='left',
        on='Артикул_Размер'
    )

    # Достаем колонки с Артикулом и Размером из общей колонки
    df_sku_and_other_costs_with_catalog = df_sku_and_other_costs_with_catalog.assign(**{
        'Артикул продавца': df_sku_and_other_costs_with_catalog['Артикул_Размер'].str.split('_size_', expand=True)[0],
        'Размер': df_sku_and_other_costs_with_catalog['Артикул_Размер'].str.split('_size_', expand=True)[1]
    })
    # Перемещаем справочные колонки в начало df
    df_sku_and_other_costs_with_catalog = move_columns(
        df_sku_and_other_costs_with_catalog,
        columns_to_move=['Артикул_Размер', 'Артикул продавца', 'Размер'] + catalog_finace_svod_columns,
        position='Артикул WB',
        insert_type='after'
    )
    # У артикулов с размером без себестоимости перем первую себестоимость в группе артикулов
    df_sku_and_other_costs_with_catalog['Себестоимость'] = np.where(
        df_sku_and_other_costs_with_catalog['Себестоимость'].isna(),
        df_sku_and_other_costs_with_catalog.groupby('Артикул продавца')['Себестоимость'].transform('first'),
        df_sku_and_other_costs_with_catalog['Себестоимость']
    )
    # Считаем себестоимость
    df_sku_and_other_costs_with_catalog['Себестоимость * Факт продаж'] = abs(
        df_sku_and_other_costs_with_catalog['Себестоимость'] * df_sku_and_other_costs_with_catalog['Продажи минус возвраты шт']
    )

    return df_sku_and_other_costs_with_catalog


# Функция расчета финальных расходов
def calc_final_costs(df_all_costs_stats_with_catalog):
    # Создание копии для избежания изменений в оригинальном df
    df_final_costs = df_all_costs_stats_with_catalog.copy()
    # Колонка, от которой будем считать маржинальность в %
    sales_for_percent = 'Продажи минус возвраты руб с СПП'
    sales_for_rub = 'Продажи минус возвраты руб к перечислению'
    marg_amount = 'Продажи минус возвраты шт'
    sales_for_percent_column = pd.to_numeric(df_final_costs[sales_for_percent], errors='coerce')
    sales_for_rub_column = pd.to_numeric(df_final_costs[sales_for_rub], errors='coerce')
    sales_amount_column = df_final_costs[marg_amount]
    # # Колонки, по которым нужно изменить знак расходов
    # change_sign_columns = [
    #     'Возмещение издержек по перевозке',
    #     'Добровольная компенсация при возврате',
    #     'Компенсация ущерба',
    #     'Логистика сторно',
    #     # 'Вознаграждение WB без НДС',  # Знак уже изменен в расходах по SKU
    #     # 'НДС с вознаграждения WB'  # Знак уже изменен в расходах по SKU
    # ]
    # # Изменяем знак некоторых расходов с + на -
    # for col in change_sign_columns:
    #     if col in df_final_costs.columns:
    #         df_final_costs[col] = df_final_costs[col] * (-1)

    # # Считаем комиссию WB
    # df_final_costs['Комиссия WB, %'] = (
    #     (df_final_costs['Вознаграждение WB без НДС'] + df_final_costs['НДС с вознаграждения WB']) /
    #     df_final_costs['Продажи минус возвраты руб с СПП']
    #     * 100
    # )
    # # Заменяем деление на 0 на пропуски
    # df_final_costs['Комиссия WB, %'] = df_final_costs['Комиссия WB, %'].replace([np.inf, -np.inf], np.nan)
    # # Считаем полную комиссию WB
    # df_final_costs['Полная комиссия WB'] = (
    #     df_final_costs['Продажи минус возвраты руб с СПП'] -
    #     df_final_costs['Продажи минус возвраты руб без СПП'] +
    #     df_final_costs['Вознаграждение WB без НДС']
    # )
    # # Помещаем колонки после продаж минус возвратов
    # df_final_costs = move_columns(
    #     df_final_costs,
    #     ['Полная комиссия WB', 'Комиссия WB, %'],
    #     position='НДС с вознаграждения WB'
    # )

    # Считаем все расходы на рекламные кампании
    promotion_columns = [
        'Расходы по рекламным кампаниям',
        'Витрина Магазина',
        'Оказание услуг «ВБ.Медиа»',
        'Предоставление услуг по подписке «Джем»',
        'Стоимость участия в программе лояльности',
    ]
    # Находим колонки по рекламе, которые реально есть в отчете реализации
    existing_promotion_columns = [col for col in df_final_costs.columns if col in promotion_columns]
    if existing_promotion_columns:
        df_final_costs['Затраты на рекламу руб'] = df_final_costs[existing_promotion_columns].sum(axis=1)
    else:
        df_final_costs['Затраты на рекламу руб'] = 0.0

    # Считаем % по рекламным расходам
    df_final_costs['% Затрат на рекламные кампании'] = (
        abs(
            df_final_costs['Затраты на рекламу руб'] /
            sales_for_percent_column
            * 100
        )
    )
    # Переносим % по рекламным расходам до маржинальности
    df_final_costs = move_columns(
        df_final_costs,
        ['Затраты на рекламу руб', '% Затрат на рекламные кампании'],
        position=df_final_costs.columns.shape[0] - 2,
        insert_type='before'
    )

    # Колонки, которые нужно исключить из затрат
    columns_for_costs = ([
        'Артикул WB',
        'Артикул продавца',
        'Размер',
        'Артикул_Размер',
        'Статус',
        'Заказы шт',
        'Продажи шт',
        'Продажи руб с СПП',
        'Продажи руб без СПП',
        'Возвраты руб с СПП',
        'Возвраты руб без СПП',
        'Продажи минус возвраты руб с СПП',
        'Продажи минус возвраты руб без СПП',
        'Продажи минус возвраты шт',
        'Возвраты шт',
        'Количество размеров в артикуле',
        'Продажи в артикуле',
        'Доля продаж в артикуле',
        '% Затрат на рекламные кампании',
        'Маржинальность руб',
        'Маржинальность %',
        'Маржа от себестоимости %',
        'Маржа по артикулу %',
        'Маржинальность на ед. товара руб',
        'Себестоимость',
        'Вознаграждение WB без НДС',  # Уже учтено в полной комиссии WB
        # 'НДС с вознаграждения WB',
        'Комиссия WB, %',
        # 'Платное хранение (отдельный отчет)',
        # 'Платная приемка (отдельный отчет)'
        'Приход денег на РС'
        ] + promotion_columns
    )
    columns_for_final_costs = ([
        'Стоимость логистики',
        'Затраты на рекламу руб',
        'Платное хранение возвратов на ПВЗ более 3 дней',
        'Платное хранение (отдельный отчет)',
        'Платная приемка (отдельный отчет)',
        'Выявленные расхождения после приемки на складе WB',
        'Компенсация скидки по программе лояльности',
        'Занижение фактических габаритов упаковки товара',
        'Себестоимость * Факт продаж'

    ])
    # Расчет маржинальности в рублях
    df_final_costs['Маржинальность руб'] = sales_for_rub_column
    # for col in df_final_costs.columns[~df_final_costs.columns.isin(columns_to_exclude_from_final_costs)]:
    for col in df_final_costs.columns[df_final_costs.columns.isin(columns_for_final_costs)]:
        df_final_costs['Маржинальность руб'] = (
            df_final_costs['Маржинальность руб'] - df_final_costs[col].fillna(0)
        )
    # Приход денег на РС (как маржинальность, но не учитываем Себестоимость * Факт продаж)
    df_final_costs['Приход денег на РС'] = (
        df_final_costs['Маржинальность руб']
        + df_final_costs['Себестоимость * Факт продаж'].fillna(0)
    )
    # Расчет Маржинальности в процентах
    # Там, где нет себестоимости, маржинальность не считаем
    df_final_costs.loc[df_final_costs['Себестоимость'].isna(), 'Маржинальность руб'] = np.nan
    df_final_costs['Маржинальность %'] = np.where(
        (sales_for_percent_column != 0) & (~np.isnan(sales_for_percent_column)),
        df_final_costs['Маржинальность руб'] / sales_for_percent_column * 100,
        np.nan
    )
    # Расчет маржинальности на единицу товара
    df_final_costs['Маржинальность на ед. товара'] = np.where(
        (sales_amount_column != 0) & (~pd.isna(sales_amount_column)),
        df_final_costs['Маржинальность руб'] / sales_amount_column,
        np.nan
    )
    # Маржинальность от себестоимости (%)
    df_final_costs['Маржа от себестоимости %'] = (
        df_final_costs['Маржинальность руб']
        / abs(df_final_costs['Себестоимость * Факт продаж'])
        * 100
    )
    # Считаем маржинальность по артикулу
    sales_rub_mean_by_article = (
        df_final_costs
        .groupby(['Артикул WB'])
        [sales_for_percent].transform('mean')
    )
    marg_rub_mean_by_article = (
        df_final_costs
        .groupby(['Артикул WB'])
        ['Маржинальность руб'].transform('mean')
    )
    df_final_costs['Маржа по артикулу %'] = (
        marg_rub_mean_by_article / sales_rub_mean_by_article * 100
    )
    # У некоторых колонок с маржинальностью меняем знак по принципу:
    # там, где маржинальность в рублях > 0, знак "+  в ином случае "-"
    for col in ['Маржинальность на ед. товара']:
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

    # Меняем inf на nan в маржинальностях
    marg_columns = [
        'Маржинальность %',
        'Маржинальность на ед. товара',
        'Маржа от себестоимости %',
        'Маржа по артикулу %',
    ]
    for col in marg_columns:
        df_final_costs[col] = df_final_costs[col].replace([-np.inf, np.inf], np.nan)


    return df_final_costs, promotion_columns, columns_for_final_costs


# Функция расчета сводной таблицы по всем расходам
def create_svod_by_operations(df_final_costs):
    # Список колонок, по котором мы не считаем суммы
    columns_to_exclude_from_operations_svod = [
        'Артикул WB',
        'Артикул_Размер',
        'Артикул продавца',
        'Артикул 1С',
        'Размер',
        'Цвет',
        'Статус',
        'Количество размеров в артикуле',
        'Продажи в артикуле',
        'Доля продаж в артикуле',
        'Комиссия WB, %',
        'Себестоимость',
        # 'Себестоимость * Факт продаж',
        # 'Все затраты (руб)',
        # 'Маржинальность руб',
        'Маржинальность %',
        'Маржа по артикулу %',
        'Маржа от себестоимости %',
        # 'Маржинальность на ед. товара',
        '% Затрат на рекламные кампании',
    ]
    # Создаем копию для избежения изменений в исходном df
    df_final_costs_by_operation = df_final_costs.copy()

    # Считаем суммы по остальным столбцам
    df_final_costs_by_operation = (
        df_final_costs_by_operation
        # Убираем колонки, которые не должны быть включены в сумму
        .loc[
            :,
            ~df_final_costs_by_operation.columns.isin(columns_to_exclude_from_operations_svod)
        ]
        # Считаем суммы по остальным
        .sum()
        # Переворачиваем расходы в столбцы
        .to_frame()
        .T
        .reset_index(drop=True)
        # # Достаем столбцы из index
        # .reset_index()
        # # Переименовываем колонки
        # .rename(columns={
        #     'index': 'Тип операции',
        #     0: 'Итого руб'
        # })
    )
    # Пересчитываем некоторые столбцы
    # df_final_costs_by_operation['Вознаграждение WB без НДС'] = (
    #     df_final_costs_by_operation['Продажи минус возвраты руб с СПП'] -
    #     df_final_costs_by_operation['Продажи минус возвраты руб без СПП']
    #     + df_final_costs_by_operation['Вознаграждение WB без НДС']
    # )
    df_final_costs_by_operation['Маржинальность на ед. товара'] = (
        df_final_costs_by_operation['Маржинальность руб']
        / df_final_costs_by_operation['Продажи минус возвраты шт']
    )
    # Делаем расходы снова в строках
    df_final_costs_by_operation = (
        df_final_costs_by_operation
        .melt(var_name='Тип операции', value_name='Итого руб')
    )

    # Считаем сумму, от которой будем считать процент
    sum_columns = [
        'Продажи минус возвраты руб с СПП',
    ]
    # Считается сумма по тем колонкам, которые есть в df_final_costs
    sum_for_percent = (
        df_final_costs
        .loc[:, df_final_costs.columns.isin(sum_columns)]
        .sum().sum()
    )
    # Считаем процент
    df_final_costs_by_operation['Итого, %'] = (
        df_final_costs_by_operation['Итого руб'] / sum_for_percent * 100
    )
    # Делаем процент пропуском в некоторых расходах
    no_percent_operations = [
        'Продажи шт',
        'Возвраты шт',
        'Продажи минус возвраты шт',
        'Маржинальность на ед. товара',
    ]
    # Маска для loc
    mask_no_operations = df_final_costs_by_operation['Тип операции'].isin(no_percent_operations)
    # Создание пропусков
    df_final_costs_by_operation.loc[mask_no_operations, 'Итого, %'] = np.nan

    return df_final_costs_by_operation


# Функция сохранения результатов в excel
def save_excel(
        current_report_dir, date_report,
        df_realization_report_new_columns,
        report_dates,
        df_final_costs,
        df_final_costs_by_operation,
        result_sku_costs,
        result_other_costs,
        result_companies_costs,
        result_paid_storage,
        result_paid_acceptance
        ):
    # Получаем результаты по расходам, которые уже были разбиты на артикулы
    df_sku_costs_by_operation = result_sku_costs['df_sku_costs_by_operation']
    # Получаем результаты по платному хранению
    df_paid_storage = result_paid_storage['df_paid_storage']
    # Получаем результаты по платной приемке
    df_paid_acceptance = result_paid_acceptance['df_paid_acceptance']
    # Получаем результаты по прочим расходам
    df_other_costs_by_operation = result_other_costs['df_other_costs_by_operation']
    # Получаем результаты по расходам компаний
    df_companies_costs = result_companies_costs['df_companies_stats']

    # Сохраняем результаты в excel
    filepath_finance_report = current_report_dir / f"{date_report}_Свод_реализации_WB_{client_name}.xlsx"
    with pd.ExcelWriter(filepath_finance_report) as w:
        df_final_costs.to_excel(w, sheet_name='Итоговая таблица', index=False)
        df_final_costs_by_operation.to_excel(w, sheet_name='Сводная', index=False)
        df_realization_report_new_columns.to_excel(w, sheet_name='Исходный отчет АПИ', index=False)
        report_dates.to_excel(w, sheet_name='Список отчетных недель', index=False)
        df_sku_costs_by_operation.to_excel(w, sheet_name='Расходы по SKU', index=False)
        df_other_costs_by_operation.to_excel(w, sheet_name='Прочие расходы', index=False)
        df_companies_costs.to_excel(w, sheet_name='Расходы по РК', index=False)
        # df_paid_acceptance.to_excel(w, sheet_name='Отчет пл. приемка', index=False)
        # df_paid_storage.to_excel(w, sheet_name='Отчет пл. хранение', index=False)

    return filepath_finance_report


# Функция форматирования excel
def format_excel(
    current_report_dir,
    date_report,
    client_name,
    df_final_costs,
    promotion_columns,
    columns_for_final_costs,
):

    logger.info(f"Formatting Finance report for client {client_name}")

    # Формируем путь к итоговому файлу Excel
    filename = current_report_dir / f"{date_report}_Свод_реализации_WB_{client_name}.xlsx"
    # Загружаем рабочую книгу Excel
    wb = openpyxl.load_workbook(filename)

    # Проходим по всем листам и настраиваем ширины столбцов + фильтр
    for ws in wb.worksheets:
        for column in ws.columns:
            column_letter = column[0].column_letter  # Получаем букву столбца (например, 'A')
            # Находим максимальную длину значения среди всех ячеек этого столбца
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column)
            # Делаем ширину побольше найденного максимума для удобства визуализации
            ws.column_dimensions[column_letter].width = (max_length + 2) * 1.2
        # Включаем автoфильтр на всех данных листа
        ws.auto_filter.ref = ws.dimensions

    # --- Форматирование листа с итоговой таблицей ---
    # Получаем лист, где будем работать с итоговой таблицей
    final_ws = wb['Итоговая таблица']

    # Создаём словарь: название столбца -> буква столбца в Excel
    col_map = {col: get_column_letter(idx + 1) for idx, col in enumerate(df_final_costs.columns)}

    # Готовим стили для ячеек: тонкая чёрная рамка
    thin_border = Side(border_style="thin", color="000000")

    svod_len = df_final_costs.shape[0]  # Количество строк с данными отчёта
    row_start = 2  # Обычно строка 1 — заголовки, а с 2 — данные

    # Границы диапазона столбцов по буквам Excel (например, A:Q)
    min_col = get_column_letter(1)
    max_col = get_column_letter(len(df_final_costs.columns))

    # Получаем диапазон заголовков по соответствующим ячейкам
    header_cells = final_ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    for row in header_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(top=thin_border, bottom=thin_border, left=thin_border, right=thin_border)

    # Получаем диапазон всех строк с данными
    all_cells = final_ws[f"{min_col}{row_start}:{max_col}{svod_len + 1}"]
    for row in all_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = Border(top=thin_border, bottom=thin_border, left=thin_border, right=thin_border)

    # Определяем, какие рекламные столбцы реально присутствуют в таблице
    existing_promotion_columns = [col for col in df_final_costs.columns if col in promotion_columns]

    # Определяем все колонки, учитываемые в расходах (исключаем явно ненужные)
    cost_columns = [col for col in df_final_costs.columns if col in columns_for_final_costs]

    start_row = 2
    end_row = final_ws.max_row  # Последняя строка с данными в финальной таблице

    # Колонка продаж минус возвратов
    sales_for_percent_col = 'Продажи минус возвраты руб с СПП'
    sales_for_rub_col = 'Продажи минус возвраты руб к перечислению'

    sales_minus_returns_amount_col = 'Продажи минус возвраты шт'

    for row in range(start_row, end_row + 1):
        # --- Формула: "Затраты на рекламу руб"
        if existing_promotion_columns:
            promo_cells = ','.join(f"{col_map[col]}{row}" for col in existing_promotion_columns)
            promo_formula = f"=SUM({promo_cells})"
            final_ws[f"{col_map['Затраты на рекламу руб']}{row}"].value = promo_formula
        else:
            final_ws[f"{col_map['Затраты на рекламу руб']}{row}"].value = '0'

        # --- Формула: "% Затрат на рекламные кампании" (IFERROR — чтобы не было #ЗНАЧ!)
        cell_promo = f"{col_map['Затраты на рекламу руб']}{row}"
        cell_sales_minus_returns = f"{col_map[sales_for_percent_col]}{row}"
        formula_percent = (
            f'=IFERROR(ABS({cell_promo}/{cell_sales_minus_returns}*100), "")'
        )
        final_ws[f"{col_map['% Затрат на рекламные кампании']}{row}"].value = formula_percent

        # --- Формула: "Маржинальность руб" (IFERROR — если что-то не так)
        cell_sebestoimost = f"{col_map['Себестоимость']}{row}"
        base = f"{col_map[sales_for_rub_col]}{row}"
        cost_formula = '-'.join(f"{col_map[col]}{row}" for col in cost_columns)
        margin_base_formula = f"{base}-" + f"-".join(f"{col_map[col]}{row}" for col in cost_columns) if cost_columns else base
        margin_formula = (
            f'=IF(OR({cell_sebestoimost}=0, {cell_sebestoimost}=""), "", IFERROR({margin_base_formula}, ""))'
        )
        final_ws[f"{col_map['Маржинальность руб']}{row}"].value = margin_formula

        # --- Формула: "Маржинальность %" (IFERROR)
        cell_marg_rub = f"{col_map['Маржинальность руб']}{row}"
        cell_sales_for_percent = f"{col_map[sales_for_percent_col]}{row}"
        formula_percent_margin = (
            f'=IFERROR({cell_marg_rub}/{cell_sales_for_percent} * 100, "")'
        )
        final_ws[f"{col_map['Маржинальность %']}{row}"].value = formula_percent_margin

        # --- Формула: "Маржинальность на ед. товара руб" (IFERROR)
        cell_qty = f"{col_map[sales_minus_returns_amount_col]}{row}"
        formula_na_ed = (
            f'=IFERROR(ABS({cell_marg_rub}/{cell_qty})*SIGN({cell_marg_rub}), "")'
        )
        final_ws[f"{col_map['Маржинальность на ед. товара']}{row}"].value = formula_na_ed

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
            f"{base}-" + f"-".join(f"{col_map[col]}{row}" for col in prihod_cols)
            if prihod_cols else base
        )
        prihod_rs_formula = (
            f'=IFERROR({prihod_base_formula}, "")'
        )
        final_ws[f"{col_map['Приход денег на РС']}{row}"].value = prihod_rs_formula

    # Добавляем комментарии в некоторые заголовки столбцов
    comment_map = {
        'Маржинальность %': 'Маржинальность руб / Продажи минус возвраты руб с СПП * 100',
        'Маржинальность на ед. товара': 'Маржинальность руб / Продажи минус возвраты шт * 100',
        'Маржа по артикулу %': 'средняя маржинальность по артикулу в руб., / на среднюю сумму по Продажи минус возвраты руб с СПП'
    }
    for col, comment in comment_map.items():
        final_ws[f"{col_map[col]}1"].comment = Comment(comment, "Artem")

    # --- Форматирование всех листов (без формул) ---
    # Задаем стиль границы
    thin_border = Side(border_style="thin", color="000000")
    # Исключаем колонки из числового форматирования
    exclude_columns = [
        "Артикул WB",
        "chrtId",
        "incomeId",
        'officeId',
        'Номер кампании',
        'Номер отчёта',
        'Номер строки'
    ]

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

    # Сохраняем итоговый файл с формулами и форматированием
    wb.save(filename)
    # Выводим сообщение
    # logger.info(f"Finance report saved in: {filename}")


# %% Вызов всех функций
if __name__ == '__main__':

    # Парсим аргументы командной строки
    dates_cli = cli_date_range(parser)

    # Получаем даты из командной строки
    date_start = dates_cli['start_date']
    date_end = dates_cli['end_date']

    # Генерируем диапазон дат для текущего отчета
    date_start, date_end = generate_date_range(
        # date_start,
        # date_end,
        date_start = '2025-09-29',
        date_end='2025-10-31'
    )
    # Создаем директорию для текущего отчета
    current_report_dir, date_report = create_finance_report_dir(client_name, date_start, date_end)
    # Получение отчета о реализации по АПИ
    df_realization_report = upload_realization_report_new(
        headers,
        date_start,
        date_end,
        verbose=True
    )
    # Переименование колонок отчета о реализации в русские названия
    df_realization_report_renamed = rename_report_columns(
        df_realization_report
    )
    # Расчет некоторых дополнительных колонок, которые будут участвовать в расчетах
    df_realization_report_new_columns = add_new_columns(
        df_realization_report_renamed
    )
    # Фильтр по переходным неделям
    df_realization_report_date_filtered = filter_report_dates(
        date_start,
        date_end,
        df_realization_report_new_columns,
        filter_dates=True
    )
    # Получаем список дат отчета
    report_dates = get_report_dates(
        df_realization_report_date_filtered
    )

    # Расчет затрат, которые уже разбиты по артикулам
    result_sku_costs = calc_sku_costs_v3(
        df_realization_report_date_filtered
    )
    # # Выгрузка заказов
    # df_orders = get_orders(
    #     headers,
    #     date_start,
    #     date_end,
    #     to_save=False
    # )
    # Добавляем заказы в отчет
    # df_sku_costs_and_orders = add_orders_to_sku_costs(result_sku_costs, df_orders)
    # Расходы на платное хранение (отдельный отчет)
    result_paid_storage = calc_paid_storage_costs(
        headers,
        df_realization_report_date_filtered,
        date_start,
        date_end
    )
    # Объединяем расходы, которые уже разбиты по размерам
    df_sku_sizes_costs_merged = union_sku_sizes_costs(
        result_sku_costs,
        result_paid_storage
    )
    # Расходы на платную приемку (отдельный отчет)
    result_paid_acceptance = calc_paid_acceptance_costs(
        headers,
        df_sku_sizes_costs_merged,
        df_realization_report_date_filtered,
        date_start,
        date_end
    )
    # Считаем расходы на рекламные кампании
    result_companies_costs = calc_companies_costs_v2(
        df_realization_report_date_filtered,
        df_sku_sizes_costs_merged,
        date_report
    )
    # Считаем расходы на списания за отзывы на товары
    result_reviews_costs = calc_reviews_costs(
        df_realization_report_date_filtered,
        df_sku_sizes_costs_merged,
    )
    # Объединяем расходы на пл. приемку с расходами по SKU и хранению
    # df_sku_paid_storage_acceptance_stats = union_sku_paid_storage_acceptance_costs(df_sku_and_paid_storage_stats, result_paid_acceptance)
    # Объединяем все расходы из отдельных отчетов в один df
    df_sku_sizes_costs_all = merge_all_costs(
        df_sku_sizes_costs_merged,
        result_paid_acceptance,
        result_companies_costs,
        result_reviews_costs,
    )
    # Расчет затрат, которые не привязаны к артикулам
    # result_other_costs = calc_other_costs(df_realization_report_date_filtered, result_sku_costs)
    result_other_costs = calc_other_costs_v3(
        result_sku_costs,
        df_sku_sizes_costs_all
    )
    # Список текущих товаров по апи
    # df_products = getWBProduct(headers, to_save=False)
    # Объединение расходов
    # df_all_costs_stats = merge_all_costs(result_other_costs,
    #                                      result_paid_storage,
    #                                      result_paid_acceptance)
    # Достаем получившиеся расходы после расчета прочих расходов
    df_all_costs_stats = result_other_costs['df_sku_and_other_costs'].copy()
    # Добавляем данные из справочной таблицы
    df_all_costs_stats_with_catalog = add_data_from_catalog(
        result_other_costs
    )
    # Расчет итоговой колонки по затратам
    df_final_costs, promotion_columns, columns_for_final_costs = calc_final_costs(
        df_all_costs_stats_with_catalog
    )
    # Считаем сводную по затратам
    df_final_costs_by_operation = create_svod_by_operations(df_final_costs)
    # Сохранение результатов в excel
    filepath_finance_report = save_excel(
        current_report_dir,
        date_report,
        # df_realization_report_new_columns,
        df_realization_report_date_filtered,
        report_dates,
        df_final_costs,
        df_final_costs_by_operation,
        result_sku_costs,
        result_other_costs,
        result_companies_costs,
        result_paid_storage,
        result_paid_acceptance
    )
    # Форматирование файла excel
    format_excel(
        current_report_dir,
        date_report,
        client_name,
        df_final_costs,
        promotion_columns,
        columns_for_final_costs,
    )
    print(f"\033[32m\033[47mDone calculating Finance Report for Client {client_name}")
    logger.info(f"Finance Report saved in {filepath_finance_report}")



# %%
