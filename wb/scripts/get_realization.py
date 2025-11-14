
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
from wb.scripts.constants import client_name, headers, marketplace_dir_path
from wb.scripts.wb_generic_functions import get_company_stats_v2
from wb.scripts.wb_generic_functions import upload_realization_report
from wb.scripts.calc_svod_finance_report_WB_v2 import (  # filter_report_dates
    add_new_columns, rename_report_columns)
# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Get sales and returns from realization report (WB)",
    usage=(
        "python.exe -m wb.scripts.get_realization [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.get_realization 2024-11-13 2024-11-20\n\n"
        "  python.exe -m wb.scripts.get_realization --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m wb.scripts.get_realization 2024-11-13 2024-11-20 -c 3"
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


# Функция фильтра отчета о реализации по датам
def filter_report_dates(date_start, date_end, df_realization_report_new_columns, filter_dates=False):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_date_filtered = df_realization_report_new_columns.copy()
    if filter_dates:
        # # Переводим даты выгрузки в Timestamp
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


# Функция выгрузки отчета о реализации
def get_realization_report(headers, date_start, date_end):
    # Переводим даты в нужный формат
    # date_start_ = datetime.fromisoformat(date_start).strftime('%Y-%m-%d')
    # date_end_ = datetime.fromisoformat(date_end).strftime('%Y-%m-%d')
    # Получение отчета о реализации по АПИ
    df_realization_report = upload_realization_report(headers, date_start, date_end)
    # Переименование колонок отчета о реализации в русские названия
    df_realization_report_renamed = rename_report_columns(df_realization_report)
    # Расчет некоторых дополнительных колонок, которые будут участвовать в расчетах
    df_realization_report_new_columns = add_new_columns(df_realization_report_renamed)
    # Фильтр по переходным неделям
    df_realization_report_date_filtered = filter_report_dates(
        date_start, date_end,
        df_realization_report_new_columns,
        filter_dates=True
    )

    return df_realization_report_date_filtered


# Функция расчета продаж и возвратов
def calc_sales_returns(df_realization_report_date_filtered, groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']):
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

    # Считаем итоговы суммы по Продажам и Возвратам (без разбивки по Артикулам)
    # df_sku_sales_returns_by_operation = (
    #     df_sku_costs
    #     .loc[df_sku_costs['id'].isin(sales_ids + returns_ids), :]
    #     .groupby('Обоснование для оплаты')
    #     .agg(**{
    #         'Расходы руб с СПП': (sales_returns_spp, 'sum'),
    #         'Расходы руб без СПП': (sales_returns_without_spp, 'sum'),
    #     })
    #     .reset_index()
    #     .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    # )

    # Заполняем пропуски
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_stats['Продажи минус возвраты шт'] = df_sku_sizes_costs_stats['Продажи шт'] - df_sku_sizes_costs_stats['Возвраты шт']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП'] = df_sku_sizes_costs_stats['Продажи руб с СПП'] - df_sku_sizes_costs_stats['Возвраты руб с СПП']
    df_sku_sizes_costs_stats['Продажи минус возвраты руб без СПП'] = df_sku_sizes_costs_stats['Продажи руб без СПП'] - df_sku_sizes_costs_stats['Возвраты руб без СПП']

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
    df_sku_sizes_costs_stats['Комиссия WB, %'] = (
        (df_sku_sizes_costs_stats['Вознаграждение WB без НДС'] + df_sku_sizes_costs_stats['НДС с вознаграждения WB']) /
        df_sku_sizes_costs_stats['Продажи минус возвраты руб с СПП']
        * 100
    )


    # Заменяем деление на 0 на пропуски
    df_sku_sizes_costs_stats['Комиссия WB, %'] = df_sku_sizes_costs_stats['Комиссия WB, %'].replace(np.inf, np.nan).replace(-np.inf, np.nan)


    # # Считаем суммы по Штрафам и Доплатам (по ним есть доп. детализация в столбце Обоснование штрафов и доплат)
    # penalty_costs_ids = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'].isin(['Штрафы и доплаты', 'Штраф']), 'id'].to_list()
    # df_sku_penalty_costs = df_sku_costs.loc[df_sku_costs['id'].isin(penalty_costs_ids), :]
    # # Тут берем столбец 'Обоснование штрафов и доплат', а не 'Обоснование для оплаты', т.к. в нем есть доп. детализация
    # for operation_type in df_sku_penalty_costs['Обоснование штрафов и доплат'].unique():
    #     df_sku_sizes_costs_stats[operation_type] = df_sku_costs.loc[df_sku_costs['Обоснование штрафов и доплат'] == operation_type, :] \
    #     .groupby(groupby_cols) \
    #     ['Итого руб'].sum()
    # # Список с названиями и затратами по Штрафам и Доплатам
    # df_sku_penalty_costs_by_operation = (
    #     df_sku_penalty_costs
    #     .groupby('Обоснование штрафов и доплат')
    #     .agg(**{
    #         'Расходы руб': ('Итого руб', 'sum')
    #     })
    #     .reset_index()
    #     .rename(columns={'Обоснование штрафов и доплат': 'Группа расходов'})
    # )

    # # Считаем суммы по оставшимся расходам
    # other_costs_ids = df_sku_costs.loc[~df_sku_costs['id'].isin(sales_ids + returns_ids + penalty_costs_ids), 'id'].to_list()
    # df_sku_other_costs = df_sku_costs.loc[df_sku_costs['id'].isin(other_costs_ids), :]
    # for operation_type in df_sku_other_costs['Обоснование для оплаты'].unique():
    #     df_sku_sizes_costs_stats[operation_type] = df_sku_costs.loc[df_sku_costs['Обоснование для оплаты'] == operation_type, :] \
    #     .groupby(groupby_cols) \
    #     ['Итого руб'].sum()
    # # Список с названиями и затратами по оставшимся расходам
    # df_sku_other_costs_by_operation = (
    #     df_sku_other_costs
    #     .groupby('Обоснование для оплаты')
    #     .agg(**{
    #         'Расходы руб': ('Итого руб', 'sum')
    #     })
    #     .reset_index()
    #     .rename(columns={'Обоснование для оплаты': 'Группа расходов'})
    # )

    # Удаляем временную колонку
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.drop(columns=['tmp_col'])

    # Достаем доп. столбцы из index
    df_sku_sizes_costs_stats = df_sku_sizes_costs_stats.reset_index()

    # Переводим артикул продавца с верхний регистр
    df_sku_sizes_costs_stats['Артикул продавца'] = df_sku_sizes_costs_stats['Артикул продавца'].str.upper()


    # # Формируем list с id строк расходов по артикулам
    # sku_costs_ids = sales_ids + returns_ids + penalty_costs_ids + other_costs_ids

    # # Формируем df со списком расходов, которые привязаны к Артикулам
    # df_sku_costs_by_operation = (
    #     pd.concat([
    #             df_sku_sales_returns_by_operation,
    #             df_sku_penalty_costs_by_operation,
    #             df_sku_other_costs_by_operation
    #         ],
    #         ignore_index=True)
    #     #.reset_index(drop=True)
    # )

    return df_sku_sizes_costs_stats


# Функция вывода df с продажами и возвратами в консоль
def print_realization_reports(df_sales_returns_wb):
    # df_to_print = df_all_reports.to_string(index=False)
    # print(df_to_print)

    df_sales_returns_wb.to_clipboard(excel=True, index=False)



# %% Вызов всех функций
if __name__ == '__main__':

    # Парсим аргументы командной строки
    dates_cli = cli_date_range(parser)

    # Формируем диапазон дат
    date_start, date_end = generate_date_range(
        date_start=dates_cli['start_date'],
        date_end=dates_cli['end_date'],
        # start_date='2025-08-27',
        # end_date='2025-09-02'
    )
    # date_range_df = generate_date_range(date_start, date_end)

    # Выгружаем отчет реализации
    df_realization_report_date_filtered = get_realization_report(
        headers,
        date_start=date_start,
        date_end=date_end,
        # date_start='2025-10-01',
        # date_end='2025-10-31'

    )
    # Считаем продажи возвраты и прочую лабуду
    df_sales_returns_wb = calc_sales_returns(
        df_realization_report_date_filtered,
        groupby_cols = [
            'Год даты продажи', 'Месяц даты продажи',
            'Артикул WB',
            'Артикул продавца',
            'Размер'
        ]
    )
    # Выводим в консоль
    print_realization_reports(df_sales_returns_wb)

    logger.info(
        f"Отчет реализации WB по клиенту {client_name} в буфере обмена. "
        "Скопируйте его в Excel"
    )
