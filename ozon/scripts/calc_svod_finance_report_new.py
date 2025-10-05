
# %% Определение всех функций
import argparse
import os
import re
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from loguru import logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.options.mode.chained_assignment = None  # default='warn'
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

# Некоторые доп. функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_date_range
# Некоторые константы
from ozon.scripts.constants import (catalog_finace_svod_columns,
                                    client_id_performance, client_name,
                                    client_secret_performance, headers,
                                    marketplace_dir_name)
# Функция выгрузки товаров в поставках
from ozon.scripts.uploadDataFromOzon import get_supply_orders
# Функция выгрузки данных о рекламных кампаниях
from ozon.scripts.uploadDataFromOzonPerformance import upload_data_performane
# Функция выгрузки отчета о релизации товаров
from ozon.scripts.uploadFinanceReports import get_products_realization_report

parser = argparse.ArgumentParser(
    description="Create finance report (Ozon)",
    usage=(
        "python.exe -m ozon.scripts.calc_svod_finance_report_new [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new 2024-11-13 2024-11-20\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m ozon.scripts.calc_svod_finance_report_new 2024-11-13 2024-11-20 -c 3"
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
        marketplace_dir_name
        / 'Clients'
        / client_name
        / 'FinanceReports'
        / date_folder
    )
    # Создаем папку, если её нет
    finance_reports_dir.mkdir(parents=True, exist_ok=True)

    return finance_reports_dir, date_folder


# Функция поиска отчета с файлом в директории
def rename_finance_report_file(current_report_dir, client_name, date_start, date_end):
    # Преобразуем даты из формата YYYY-MM-DD в DD.MM.YYYY для поиска существующего файла
    try:
        start = datetime.strptime(date_start, "%Y-%m-%d").strftime("%d.%m.%Y")
        end = datetime.strptime(date_end, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception as e:
        f"Ошибка формата даты: {e}"

    reports_dir = Path(current_report_dir)

    # Маска финального (нужного) имени файла
    final_name = f"{client_name}_Отчет_по_начислениям_{date_start}_{date_end}.xlsx"
    final_path = reports_dir / final_name

    # Если файл уже есть — вернуть его путь сразу
    if final_path.exists():
        return final_path

    # Шаблон поиска исходного файла
    pattern = re.compile(
        rf"отчет[ _]по[ _]начислениям[ _]{start}-{end}.*\.(xlsx|xls)$",
        re.I
    )

    matched_files = []

    # Поиск файлов по маске
    for f in reports_dir.iterdir():
        if f.is_file() and pattern.match(f.name):
            matched_files.append((f, f.stat().st_mtime))

    if not matched_files:
        return "Файл отчета по заданным датам не найден."

    # Сортировка по дате изменения (от нового к старому)
    matched_files.sort(key=lambda x: x[1], reverse=True)
    chosen_file = matched_files[0][0]

    # Переименование
    chosen_file.rename(final_path)

    # Возврат нового пути
    return final_path

# Функция чтения файла с отчетом
def read_finance_report_file(finance_report_filepath):
    # Получения формата дат, которые содержатся в имени файла.
    # date_start_file = datetime.strptime(date_start, '%Y-%m-%d').strftime('%d_%m_%Y')
    # date_end_file = datetime.strptime(date_end, '%Y-%m-%d').strftime('%d_%m_%Y')
    # date_start_file = date_start
    # date_end_file = date_end
    # Определение пути к файлу с отчетом
    # report_file = (
    #     finance_reports_dir
    #     / f"{client_name}_Отчет_по_начислениям_{date_start_file}_{date_end_file}.xlsx"
    # )
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
        .groupby(level=[0, 1, 2]) # 0 - SKU, 1 - Артикул, 2 - Наименование
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
        position='Наименование товара',
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
    df_companies_costs = df_finance_report.loc[
        (df_finance_report['Название товара'].str.contains('Заказ №', na=False) )
        | (df_finance_report['ID начисления'].str.match('^\d{8}$', na=False)),
        ['id', 'Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]

    if not df_companies_costs.empty:
        companies_costs_flag = True

    return companies_costs_flag


# Функция нахождения идентификаторов рекламных кампаний
def get_companies_ids(df_finance_report):

    # Выбираем затраты по рекламе, где есть кампании
    df_companies_costs = df_finance_report.loc[
        (df_finance_report['Название товара'].str.contains('Заказ №', na=False) )
        | (df_finance_report['ID начисления'].str.match('^\d{8}$', na=False)),
        ['Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]

    if df_companies_costs.shape[0] > 0:
        # Создаем столбец с номером кампании
        conditions = [
            # Трафареты
            df_companies_costs['Название товара'].str.contains('Заказ №', na=False),
            # Выводы в топ
            df_companies_costs['ID начисления'].str.match('^\d{8}$', na=False)
        ]
        choices = [
            # Номер кампании для трафаретов
            (df_companies_costs['Название товара']
            .str.split('Заказ №').str[1]
            .str.split('.').str[0]),
            # Номер кампании для выводов в топ
            df_companies_costs['ID начисления']
        ]
        df_companies_costs['Номер кампании'] = np.select(conditions, choices, default=np.nan)

        # Считаем расходы на каждую кампанию
        df_companies_costs_total = (
            df_companies_costs
            .groupby(['Номер кампании', 'Тип начисления'])
            .agg(
                Итого_из_отчета=('Итого руб', 'sum')
            )
            .reset_index()
        )

    return df_companies_costs_total


# Функция парсинга файлов со статистикой по кампаниям
def parse_companies_files(current_report_dir, date_report):
    # Директория, где лежат данные по рекламным кампаниям
    companies_stats_dir = current_report_dir / f"{date_report}_Кампании"
    # Директория, куда будем складывать обработанные файлы
    companies_parsed_dir = current_report_dir / f"{date_report}_Кампании_parsed"

    # Получаем список csv файлов по кампаниям
    filenames_companies = {"path": []}
    # Расширения файлов для поиска
    companies_stats_extensions = ['.csv', '.xlsx']
    # Осуществляем поиск файлов с заданным расширением
    for ext in companies_stats_extensions:
        for file in companies_stats_dir.glob(f'*{ext}'):
            filenames_companies['path'].append(file)
    # Считывание только имени файла
    filenames_companies['file_name'] = [file.name for file in filenames_companies['path']]
    # Размер файла
    filenames_companies['file_size'] = [file.stat().st_size for file in filenames_companies['path']]
    # Перевод в df
    filenames_companies = pd.DataFrame(filenames_companies)
    # Номер заказа
    filenames_companies['order_number'] = filenames_companies['file_name'].str.split('_').str[0]
    # Сортировка по размеру
    filenames_companies = filenames_companies.sort_values(by='file_size', ascending=False, ignore_index=True)

    # Обработка csv по статистике кампаний в цикле
    for i in range(len(filenames_companies)):
        # Считываем csv
        df_company = pd.read_csv(filenames_companies['path'][i], sep=';', decimal=',', skiprows=1)
        # Если по кампании есть расходы, начинаем их обрабатывать
        if df_company.shape[0] >= 2:
            # Определяем тип отчета (по первой колонке в отчете)
            if df_company.columns[0] == 'Дата':
                # Убираем строку Всего и Корректировка
                df_company = df_company.loc[~df_company['Дата'].isin(['Всего', 'Корректировка']), :]
                # Ozon ID = SKU в данном отчете
                df_company = df_company.rename(columns={'Ozon ID': 'sku'})
                # Переименовываем колонку с расходами для удобства
                df_company = df_company.rename(columns={'Расход, ₽': 'Расходы'})
                # Переименовываем колонку с названием товара
                df_company = df_company.rename(columns={'Наименование': 'Наименование товара'})
                # Выбираем нужные колонки
                df_company = df_company.loc[:, ['Дата', 'sku', 'Наименование товара', 'Расходы']]
                # Сохраняем обработанный файл
                df_company.to_csv(
                    companies_parsed_dir / f"{filenames_companies['order_number'][i]}.csv",
                    sep=';',
                    encoding='utf-8-sig',
                    index=False
                )
            else:
                # Убираем строку Всего и Корректировка
                df_company = df_company.loc[~df_company['sku'].isin(['Всего', 'Корректировка']), :]
                # Переименовываем колонку с расходами для удобства
                df_company = df_company.rename(columns={'Расход, ₽, с НДС': 'Расходы'})
                # Переименовываем колонку с названием товара
                df_company = df_company.rename(columns={'Название товара': 'Наименование товара'})
                # Выбираем нужные колонки
                df_company = df_company.loc[:, ['sku', 'Наименование товара', 'Расходы']]
                # Сохраняем обработанный файл
                df_company.to_csv(
                    companies_parsed_dir / f"{filenames_companies['order_number'][i]}.csv",
                    sep=';',
                    encoding='utf-8-sig',
                    index=False
                )


# Функция расчета расходов по рекламным кампаниям
def calc_companies_cost_by_sku(date_report, df_finance_report):

    # df, куда будем помещать итоговый результат (разбивку расходов по SKU)
    df_result_companies_stats = pd.DataFrame()
    # df с рекламными расходами из отчета ЛК
    df_companies_stats_from_report = pd.DataFrame()
    # df с сравнением расходов по рекламе из АПИ и отчета ЛК
    df_companies_api_report_stats = pd.DataFrame()
    # id строк отчета, которые относятся к рекламным расходам
    companies_costs_ids = []

    # Выбираем затраты по рекламе, где есть кампании
    df_companies_costs = df_finance_report.loc[
        (df_finance_report['Название товара'].str.contains('Заказ №', na=False) )
        | (df_finance_report['ID начисления'].str.match('^\d{8}$', na=False)),
        ['id', 'Группа услуг', 'ID начисления', 'Тип начисления', 'Название товара', 'Итого руб']
    ]

    if df_companies_costs.shape[0] > 0:
        # Создаем столбец с номером кампании
        conditions = [
            # Трафареты
            df_companies_costs['Название товара'].str.contains('Заказ №', na=False),
            # Выводы в топ
            df_companies_costs['ID начисления'].str.match('^\d{8}$', na=False)
        ]
        choices = [
            # Номер кампании для трафаретов
            (df_companies_costs['Название товара']
            .str.split('Заказ №').str[1]
            .str.split('.').str[0]),
            # Номер кампании для выводов в топ
            df_companies_costs['ID начисления']
        ]
        df_companies_costs['Номер кампании'] = np.select(conditions, choices, default=np.nan)
        companies_costs_ids = df_companies_costs['id'].to_list()

        # Считаем расходы на каждую кампанию
        df_companies_costs_total = (
            df_companies_costs
            .groupby(['Номер кампании', 'Тип начисления'])
            .agg(
                Итого_из_отчета=('Итого руб', 'sum')
            )
            .reset_index()
        )
        # Помещаем расходы в df из начала функции
        df_companies_stats_from_report = df_companies_costs_total.copy()

        # Сопоставление номера заказа с расходами SKU из апи
        df_companies_api = pd.DataFrame()
        for i in range(len(df_companies_costs_total)):
            # Файл с данными по кампании
            company_file_path = (
                current_report_dir
                / f"{date_report}_Кампании_parsed"
                / f"{df_companies_costs_total['Номер кампании'][i]}.csv"
            )
            # Если данный файл есть, начинаем считать расходы по данной кампании
            if company_file_path.exists():
                # Считываем данные по данной кампании
                df_company = pd.read_csv(company_file_path, sep=';')
            # Если нет, считаем, что расходы по данной кампании были равны 0
            else:
                df_company = pd.DataFrame(columns=['sku', 'Наименование товара', 'Расходы'])
                # Выводим предупреждение, что файла с расходами по данной кампании нет
                logger.warning(f"Company file {df_companies_costs_total['Номер кампании'][i]} has not been found")
            # Удаляем лишнюю колонку, если есть
            if 'Unnamed: 0' in df_company.columns:
                df_company = df_company.drop(columns='Unnamed: 0')
            # Считаем итоговые расходы по каждому SKU
            tmp_df = df_company.groupby(['sku', 'Наименование товара']).agg(
                Итого=('Расходы', 'sum')
            ).reset_index()
            # Добавляем столбцы с информацией о названии заказа
            tmp_df['Номер кампании'] = df_companies_costs_total['Номер кампании'][i]
            tmp_df['Тип начисления'] = df_companies_costs_total['Тип начисления'][i]

            # Объединяем с предыдущей кампанией
            df_companies_api = pd.concat([df_companies_api, tmp_df])

        # Считаем затраты на каждый тип начисления на каждый SKU
        df_companies_api_stats = (
            df_companies_api
            .groupby(['sku', 'Наименование товара', 'Тип начисления'])
            .agg(
            Итого=('Итого', 'sum')
            )
        .reset_index()
        )
        # Заменяем знак расходов с плюса на минус
        df_companies_api_stats['Итого'] = np.where(
            df_companies_api_stats['Итого'] > 0,
            -abs(df_companies_api_stats['Итого']),
            df_companies_api_stats['Итого']
        )
        # Переводим в длинный формат
        df_companies_api_stats = (
            pd.pivot_table(
                df_companies_api_stats,
                index=['sku', 'Наименование товара'],
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

        # Считаем расходы по типам начисления из апи
        df_companies_api_stats_by_operation = df_companies_api.groupby(['Номер кампании', 'Тип начисления']).agg(
            Итого_из_апи=('Итого', 'sum')
        ).reset_index()
        # Заменяем знак расходов с плюса на минус
        df_companies_api_stats_by_operation['Итого_из_апи'] = np.where(
            df_companies_api_stats_by_operation['Итого_из_апи'] > 0,
            -abs(df_companies_api_stats_by_operation['Итого_из_апи']),
            df_companies_api_stats_by_operation['Итого_из_апи']
        )
        # Мерджим с расходами из отчета
        df_companies_api_report_stats = df_companies_costs_total.merge(
            df_companies_api_stats_by_operation,
            on=['Номер кампании', 'Тип начисления'],
            how='left'
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
def add_articles_to_companies_costs(result_companies_costs, df_sku_list_all):
    # Получаем df с результатами расходов SKU по рекламным кампаниям
    df_result_companies_stats = result_companies_costs['df_companies_api_stats'].copy()
    # Если были расходы по РК, то добавляем к SKU названия артикулов
    if df_result_companies_stats.shape[0] > 0:
        # Мерджим со списком артикулов
        df_companies_stats_with_articles = (
            df_result_companies_stats
            .merge(
                df_sku_list_all[['SKU', 'Артикул продавца']],
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
        df_sku_list_all,
        client_id_performance,
        client_secret_performance,
        task_id=None
    ):
    # Директория, где будут лежать данные по рекламным кампаниям
    companies_stats_dir = current_report_dir / f"{date_report}_Кампании"
    companies_stats_dir.mkdir(parents=True, exist_ok=True)  # Создаем, если директория не существует
    # Директория, куда будем складывать обработанные файлы
    companies_parsed_dir = current_report_dir / f"{date_report}_Кампании_parsed"
    companies_parsed_dir.mkdir(parents=True, exist_ok=True)

    # Проверяем, что в отчете были расходы по рекламе
    companies_costs_flag = check_companies_costs(df_finance_report)
    if companies_costs_flag:
        logger.info('Found companies costs')
        # Получаем идентификаторы рекламных кампаний
        df_companies_ids = get_companies_ids(df_finance_report)
        # Получаем список файлов в директории с кампаниями
        companies_stats_files = os.listdir(companies_stats_dir)

        # Если нет выгруженных файлов по кампаниям, то делаем выгрузку статистики кампаний
        if not companies_stats_files:
            logger.info('No companies files found, uploading companies stats')
            # Формируем даты для выгрузки по API Performance
            date_start_performance = date_start + 'T00:00:00Z'
            date_end_performance = date_end + 'T23:59:59Z'
            # Выгружаем данные по рекламным кампаниям
            upload_data_performane(
                date_start_performance,
                date_end_performance,
                client_id_performance,
                client_secret_performance,
                companies_stats_dir,
                df_companies_ids,
                task_id
            )
        else:
            logger.info('Companies stats already downloaded')

        # Обрабатываем выгруженные данные
        parse_companies_files(current_report_dir, date_report)

    # Считаем расходы по рекламным кампаниям
    result_companies_costs_no_article = calc_companies_cost_by_sku(date_report, df_finance_report)
    # Добавляем артикул к SKU с рекламными расходами
    result_companies_costs = add_articles_to_companies_costs(result_companies_costs_no_article, df_sku_list_all)

    return result_companies_costs


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

        # Выгрузка товаров в заявке на поставку
        df_supplies_items = get_supply_orders(headers, states=[], to_save=False)
        # df_supplies_items.to_excel(f"{finance_reports_dir}/{client_name}_Товары_в_поставках.xlsx")

        # Переименовываем колонки для удобства
        df_supplies_items = df_supplies_items.rename(columns={
            'sku': 'SKU',
            'Артикул': 'Артикул продавца',
        })
        # Считаем итоговое количество sku в заявке на поставку
        df_supplies_items_total = (
            df_supplies_items
            .groupby(['SKU', 'Артикул продавца', 'Наименование товара', 'supply_id'])
            .agg(
                quantity=('quantity', 'sum')
            )
            .reset_index()
        )
        # Количество товаров в заявке
        df_supplies_items_total['Количество товаров в поставке'] = df_supplies_items_total.groupby('supply_id')['quantity'].transform('sum')

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
            df_supplies_found['Расходы по поставкам'] = df_supplies_found['quantity'] / df_supplies_found['Количество товаров в поставке'] * df_supplies_found['Расходы_поставки_из_отчета_ЛК']
            # Переводим в широкий формат
            df_supplies_found_total = (
                pd.pivot_table(
                    df_supplies_found,
                    index=['SKU', 'Артикул продавца', 'Наименование товара'],
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
        marketplace_dir_name
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


# Функция создания списка артикулов из расходов, которые уже разбиты по артикулам
def create_all_sku_df(result_sku_costs, result_supplies_costs, catalog):
    # Получаем df с расходами, которые уже разбиты по SKU и Артикулу
    df_sku_costs = result_sku_costs['df_sku_costs_all'].copy()
    df_supplies_costs = result_supplies_costs['supplies_costs_matched'].copy()
    catalog_ = catalog.copy()
    # Из справочной таблицы убираем товары, у которых нет SKU или Артикула
    catalog_ = catalog_.dropna(subset=['SKU', 'Артикул продавца'])
    # Формируем лист с датафреймами
    df_list = [df_sku_costs, df_supplies_costs, catalog_]
    # Переводи артикул в каждом df в строку
    for df in df_list:
        if 'Артикул продавца' in df.columns:
            df['Артикул продавца'] = df['Артикул продавца'].astype(str)
    # Выбираем нужные колонки
    sku_columns = ['SKU', 'Артикул продавца', 'Наименование товара']
    # Используем reindex для добавления отсутствующих колонок
    df_list = [df.reindex(columns=sku_columns) for df in df_list]
    # Объединяем в один df и убираем дубликаты
    df_sku_list_all = (
        pd.concat(df_list)
        .drop_duplicates(subset=['SKU', 'Артикул продавца'])
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
        marketplace_dir_name
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
    columns_to_move=[col for col in catalog_columns if col!='SKU']
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
    marg_sales_column = df_final_costs['Продажи - Возвраты (руб)']

    # Считаем себестоимость
    df_final_costs['Себестоимость * Факт продаж'] = (-1) * abs(df_final_costs['Себестоимость'] * df_final_costs['Продажи - Возвраты (шт)'])

    # Считаем все расходы на рекламные кампании
    # Определяем колонки, которые будут включены в затраты на рекламу
    promotion_columns = [
        'Вывод в топ',
        'Трафареты',
        'Продвижение с оплатой за заказ',
        'Оплата за клик',
        'Подписка Premium',
        'Подписка Premium Plus',
        'Бонусы продавца',
        'Бейдж Оригинал',
        'Баллы за отзывы',
        'Продвижение бренда',
        'Закрепление отзыва',
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
    # Там, где нет себестоимости, маржинальность не считаем
    df_final_costs.loc[df_final_costs['Себестоимость'].isna(), 'Маржинальность руб'] = np.nan
    # Считаем маржинальность на ед. товара
    df_final_costs['Маржинальность на ед. товара руб'] = (
        df_final_costs['Маржинальность руб'] /
        df_final_costs['Продажи - Возвраты (шт)']
    )
    # Там, где маржинальность в рублях > 0, у маржинальности на единицу товара так же будет знак "+"
    # в ином случае "-"
    df_final_costs['Маржинальность на ед. товара руб'] = np.where(
        df_final_costs['Маржинальность руб'] < 0,
        (-1) * abs(df_final_costs['Маржинальность на ед. товара руб']),
        abs(df_final_costs['Маржинальность на ед. товара руб'])
    )

    # Считаем маржинальность в %
    df_final_costs['Маржинальность %'] = df_final_costs['Маржинальность руб'] / marg_sales_column
    # Меняем inf на nan
    columns_to_replace_inf = [
        'Маржинальность %',
        'Маржинальность на ед. товара руб',
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

    ]
    # Считаем сумму, от которой будем считать процент
    sum_columns = [
        'Продажи - Возвраты (руб)',
        'Баллы за скидки',
        'Программы партнеров'
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
        .sum(skipna=True)
        # Достаем столбцы из index
        .reset_index()
        # Переименовываем колонки
        .rename(columns={
            'index': 'Тип операции',
            0: 'Итого руб'
        })
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
        / f'{date_report}_Свод_из_фин_отчета_Озон_{client_name}.xlsx'
    )
    # Сохраняем отчет
    with pd.ExcelWriter(filepath_finance_svod) as w:
        # Даты формирования исходного отчета
        df_report_dates.to_excel(w, sheet_name='Даты формирования отчета', index=False)
        # Исходный отчет
        df_finance_report.to_excel(w, sheet_name='Исходный отчет', index=False)
        # Объединенные расходы с распределением по sku
        df_all_costs.to_excel(w, sheet_name='Итоговая таблица', index=False)
        # Суммы по всем начислениям
        df_final_costs_by_operation.to_excel(w, sheet_name='Сводная по начислениям', index=False)
        # Расходы, которые уже были распределены по SKU в исходном отчете
        df_sku_costs_by_operation.to_excel(w, sheet_name='Расходы по SKU', index=False)
        # Расходы по типам начисления, которые распределялись в зависимости от продаж SKU
        result_operations_other_costs.to_excel(w, sheet_name='Прочие расходы', index=False)
        # Расходы на поставки из отчета ЛК
        df_supplies_costs_from_report.to_excel(w, sheet_name='Расходы на поставки из ЛК', index=False)
        # Сопоставление расходов по рекламе из АПИ и из отчета
        df_companies_api_report_stats.to_excel(w, sheet_name='Расходы по рекламе', index=False)


# Функция форматирования excel
def format_excel(
        date_report,
        df_final_costs,
        promotion_columns,
        columns_to_exclude_from_final_costs,
    ):

    logger.info(f"Formatting FinanceReport for client {client_name}")

    # Формируем путь к итоговому файлу Excel
    filepath_finance_svod = (
        current_report_dir
        / f'{date_report}_Свод_из_фин_отчета_Озон_{client_name}.xlsx'
    )
    # Открываем итоговый Excel-файл
    wb = openpyxl.load_workbook(filepath_finance_svod)

    # --- ШАГ 1. Форматируем лист "Итоговой таблицы" ---
    final_costs_worksheet = wb['Итоговая таблица']

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
    end_row = final_costs_worksheet.max_row

    for row in range(start_row, end_row + 1):
        # --- "Затраты на рекламу (руб)": сумма по вычисленным колонкам
        if existing_promotion_columns:
            promo_cells = ','.join(f"{col_map[col]}{row}" for col in existing_promotion_columns)
            promo_formula = f"=SUM({promo_cells})"
            final_costs_worksheet[f"{col_map['Затраты на рекламу (руб)']}{row}"].value = promo_formula
        else:
            final_costs_worksheet[f"{col_map['Затраты на рекламу (руб)']}{row}"].value = "0"

        # --- "% Затрат на рекламу" (через IFERROR, чтобы не было #ЗНАЧ!)
        cell_promo = f"{col_map['Затраты на рекламу (руб)']}{row}"
        cell_sales_minus_returns = f"{col_map['Продажи - Возвраты (руб)']}{row}"
        formula_percent = (
            f'=IFERROR(ABS({cell_promo}/{cell_sales_minus_returns}*100), "")'
        )
        final_costs_worksheet[f"{col_map['% Затрат на рекламу']}{row}"].value = formula_percent

        # --- "Маржинальность руб": продажи-возвраты минус сумма всех затрат (IFERROR — пропускать ошибки)
        base_sales = f"{col_map['Продажи - Возвраты (руб)']}{row}"
        cost_sum = '+'.join(f"{col_map[col]}{row}" for col in cost_columns)
        cell_sebestoimost = f"{col_map['Себестоимость']}{row}"
        margin_formula = (
            f'=IFERROR(IF(OR({cell_sebestoimost}=0,{cell_sebestoimost}=""),"",{base_sales}+{cost_sum}), "")'
        )
        final_costs_worksheet[f"{col_map['Маржинальность руб']}{row}"].value = margin_formula

        # --- "Маржинальность %" (через IFERROR)
        cell_marg_rub = f"{col_map['Маржинальность руб']}{row}"
        cell_sales = f"{col_map['Продажи руб']}{row}"
        formula_percent_margin = (
            f'=IFERROR({cell_marg_rub}/{cell_sales}, "")'
        )
        final_costs_worksheet[f"{col_map['Маржинальность %']}{row}"].value = formula_percent_margin

        # --- "Маржинальность на ед. товара руб" (через IFERROR)
        cell_qty = f"{col_map['Продажи - Возвраты (шт)']}{row}"
        formula_na_ed = (
            f'=IFERROR(ABS({cell_marg_rub}/{cell_qty})*SIGN({cell_marg_rub}), "")'
        )
        final_costs_worksheet[f"{col_map['Маржинальность на ед. товара руб']}{row}"].value = formula_na_ed        # Форматирование столбца Маржинальность % в формат процентного числа

        try:
            # Получаем букву столбца для "Маржинальность %"
            percent_col_letter = col_map['Маржинальность %']
            # Диапазон строк (от row_start до svod_len + 1)
            cell_range = final_costs_worksheet[f"{percent_col_letter}{start_row}:{percent_col_letter}{end_row + 1}"]
            for row in cell_range:
                for cell in row:
                    cell.number_format = '0%'
        except KeyError:
            logger.warning('Столбец "Маржинальность %" не найден в col_map для форматирования процентов')

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
    logger.info(f"Finance report saved in: {filepath_finance_svod}")

    return filepath_finance_svod


# %% Вызов всех функций
if __name__ == "__main__":

    # Парсим аргументы командной строки
    dates_cli = cli_date_range(parser)

    # Получаем даты из командной строки
    date_start = dates_cli['start_date']
    date_end = dates_cli['end_date']

    # Генерируем диапазон дат для текущего отчета
    date_start, date_end = generate_date_range(
        # date_start,
        # date_end,
        date_start = '2025-07-01',
        date_end='2025-07-31'
    )
    # Создаем директорию для текущего отчета
    current_report_dir, date_report = create_current_report_dir(client_name, date_start, date_end)
    # Ищем файл с отчетом
    finance_report_filepath = rename_finance_report_file(current_report_dir, client_name, date_start, date_end)
    # Считывание файла с фин. отчетом
    df_finance_report = read_finance_report_file(finance_report_filepath)
    # Выгрузка отчета о реализации продаж, возвратов и баллов Озон
    # df_realization_report = calc_sales_returns(date_start, date_end)
    # Расчет расходов с привязкой к SKU
    result_sku_costs = calc_sku_costs(df_finance_report)
    # Расчет расходов по поставкам с разбивкой на SKU
    result_supplies_costs = calc_supplies_costs(df_finance_report)
    # Считываем справочную таблицу
    catalog = read_catalog(client_name)
    # Формируем список SKU с артикулами из расходов, которые уже разбиты по SKU
    df_sku_list_all = create_all_sku_df(
        result_sku_costs,
        result_supplies_costs,
        catalog
    )
    # Расчет расходов по рекламным кампаниям с привязкой к SKU
    # task_id = 'bb511716-e864-443d-94c7-66b326c042eb'
    result_companies_costs = calc_companies_costs(
        date_start,
        date_end,
        date_report,
        df_finance_report,
        df_sku_list_all,
        client_id_performance,
        client_secret_performance,
        # task_id=task_id
    )
    # Объединение расходов по SKU
    df_sku_costs_all_stats = merge_sku_costs(
        result_sku_costs,
        result_companies_costs,
        result_supplies_costs
    )
    # Расчет расходов по типам начислений, которые не удалось привязать к SKU
    result_all_costs = calc_other_costs(
        df_finance_report,
        df_sku_costs_all_stats,
        result_sku_costs,
        result_companies_costs,
        result_supplies_costs
    )

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
