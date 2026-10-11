
# %% Определение функций
import argparse
import glob
import os
import re
import time
from datetime import datetime, timedelta


import numpy as np
import openpyxl
import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from openpyxl.formatting.rule import CellIsRule, FormulaRule
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
    client_name, headers, marketplace_dir_path, shop_type,
    realization_report_column_map,
    k_currency_to_rub,
    catalog_finace_svod_columns
)
# Уже написанные ранее функции выгрузки данных из АПИ
from wb.scripts.uploadDataFromWB import getWBProduct
# Функция выгрузки для АПИ Продвижение
from wb.scripts.uploadDataPerformanceWB import get_costs_history
from wb.scripts.wb_generic_functions import(
    upload_realization_report as upload_realization_report_new,
    rename_report_columns,
    add_new_columns,
    filter_report_dates,
    get_realization_report,
    add_sizes_to_realization_report,
    get_company_stats,
    convert_currency
)
# Функция получения справочной таблицы
from wb.scripts.get_reference_table import get_reference_table

# Колонки для расчета отдельных расходов по артикулам
CUSTOM_SUM_COLS = {
    # Столбец для суммирования: [Обоснования для оплаты]
    'Вознаграждение WB без НДС': [''],
    'НДС с вознаграждения WB': [''],
    'Стоимость логистики': ['Логистика', 'Коррекция логистики', 'Доставка', 'Коррекция стоимости доставки'],
    'Стоимость платной приёмки': ['Платная приемка (Отчет реализации)', 'Обработка товара'],
    # 'Стоимость платной приемки': ['Платная приемка (Отчет реализации)', 'Обработка товара'],
    'Возмещение издержек по эквайрингу': [''],
    'Возмещение за выдачу и возврат товаров на ПВЗ': [''],
    'Возмещение издержек по перевозке': [
        'Возмещение издержек по перевозке/по складским операциям с товаром',
        'Возмещение издержек по перемещению и операционной обработке товара'
    ],
    'Компенсация скидки по программе лояльности': ['Компенсация скидки по программе лояльности'],
    'Обработка товара': ['Обработка товара'],
}

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


# Функция получения дат отчета
def get_report_dates(df_realization_report_with_sizes):
    # Создание копии для избежания изменений в оригинальном df
    df_dates = df_realization_report_with_sizes.copy()
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


# Функция расчета расходов, которые уже разделены по артикулам в отчете апи
def calc_sku_costs_v3(df_realization_report_with_sizes, groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']):
    # Создаем копию для избежания изменений в оригинальном df
    df_sales_returns = df_realization_report_with_sizes.copy()
    # Выбираем тип документа продажа или возврат
    # df_sales_returns = df_sales_returns.loc[df_sales_returns['Тип документа'].isin(['Продажа', 'Возврат']), :]

    # ------ Расчет расходов, которые уже разбиты по размерам ------

    # Выбираем расходы, где есть артикул
    df_sku_costs = df_sales_returns.loc[df_sales_returns['Артикул продавца'] != '', :]
    # Формируем id строк с расходами по артикулам
    sku_costs_ids = df_sku_costs['id'].to_list()

    # Пересчитываем столбцы с НДС
    # df_sku_costs['Вознаграждение WB без НДС'] = (
    #     (-1) * df_sku_costs['Итоговый кВВ без НДС, %'] / 100
    #     * df_sku_costs['Цена розничная с учетом согласованной скидки']
    # )
    # df_sku_costs['НДС'] = (
    #     df_sku_costs['НДС с вознаграждения WB'] / df_sku_costs['Вознаграждение WB без НДС']
    # )
    # df_sku_costs['НДС с вознаграждения WB'] = (
    #     df_sku_costs['Вознаграждение WB без НДС'] * df_sku_costs['НДС']
    # )

    # Заполняем пропуски в размере
    if shop_type != 'clothes':
        df_sku_costs['Размер'] = df_sku_costs['Размер'].fillna('0')
    # Выбираем расходы, где есть размер
    df_sku_sizes_costs = df_sku_costs.loc[df_sku_costs['Размер'].notna(), :]
    # Формируем id строк с расходами по размерам
    sku_sizes_costs_ids = df_sku_sizes_costs['id'].to_list()

    # Заменяем пропуски в размере на 0
    # df_sku_costs['Размер'] = df_sku_costs['Размер'].replace("", "0")
    # Колонки, по которым будет производиться группировка
    # groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']

    df_sku_sizes_costs_result = pd.DataFrame()

    # Временная колонка с размерами групп
    df_sku_sizes_costs_result['tmp_col'] = df_sku_sizes_costs.groupby(groupby_cols).size()

    # Колонка, по которой считаем продажи и возвраты в рублях
    # if client_name in ['Orsk_Combinat', 'Neva_Metall']:
    sales_returns_spp = 'Цена розничная с учетом согласованной скидки'
    # else:
    sales_returns_without_spp = 'Сумма продаж (возвратов)'
    sales_returns_to_client = 'К перечислению продавцу за реализованный товар'

    # Расчет продаж
    sales_ids = df_sku_sizes_costs.loc[df_sku_sizes_costs['Обоснование для оплаты'] == 'Продажа', 'id'].to_list()
    df_sku_sizes_costs_result['Продажи шт'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб с СПП'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб без СПП'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб к перечислению'] = (
        df_sku_sizes_costs
        #.loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .loc[df_sku_sizes_costs['Тип документа'] == 'Продажа', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Расчет возвратов
    returns_ids = df_sku_sizes_costs.loc[df_sku_sizes_costs['Обоснование для оплаты'] == 'Возврат', 'id'].to_list()
    df_sku_sizes_costs_result['Возвраты шт'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        ['Количество']
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб с СПП'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб без СПП'] = (
        df_sku_sizes_costs.loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)
        [sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб к перечислению'] = (
        df_sku_sizes_costs
        # .loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .loc[df_sku_sizes_costs['Тип документа'] == 'Возврат', :]
        .groupby(groupby_cols)
        [sales_returns_to_client]
        .sum()
    )
    # Находим id начислений, где сумма к перечислению не равна 0
    # и обоснование для оплаты не равна продажи или возврату,
    # поскольку мы уже включили их в продажи и возвраты к перечислению
    # чтобы не учесть эти расходы 2 раза
    sales_returns_to_client_ids = (
        df_sku_sizes_costs
        .loc[
            df_sku_sizes_costs[sales_returns_to_client] != 0 & df_sku_sizes_costs['Обоснование для оплаты'].isin(['Продажа', 'Возврат']),
            :
        ]
        ['id']
        .to_list()
    )

    # Считаем итоговы суммы по Продажам и Возвратам (без разбивки по Артикулам)
    df_sku_sales_returns_by_operation = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(sales_ids + returns_ids + sales_returns_to_client_ids), :]
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
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_result['Продажи минус возвраты шт'] = df_sku_sizes_costs_result['Продажи шт'] - df_sku_sizes_costs_result['Возвраты шт']
    df_sku_sizes_costs_result['Продажи минус возвраты руб с СПП'] = df_sku_sizes_costs_result['Продажи руб с СПП'] - df_sku_sizes_costs_result['Возвраты руб с СПП']
    df_sku_sizes_costs_result['Продажи минус возвраты руб без СПП'] = df_sku_sizes_costs_result['Продажи руб без СПП'] - df_sku_sizes_costs_result['Возвраты руб без СПП']
    df_sku_sizes_costs_result['Продажи минус возвраты руб к перечислению'] = df_sku_sizes_costs_result['Продажи руб к перечислению'] - df_sku_sizes_costs_result['Возвраты руб к перечислению']

    rows = []
    # Список использованных обоснований для оплаты (будет обновлен далее)
    used_cost_reasons = []

    for sum_col, reasons_list in CUSTOM_SUM_COLS.items():
        # Игнорируем пустые строки и NaN
        valid_reasons = [r for r in reasons_list if r is not None and str(r).strip() != '' and not pd.isna(r)]

        # Если reasons_list пустой или после фильтрации пуст — суммируем столбец по всему df_sku_sizes_costs без фильтра по "Обоснование для оплаты"
        if not valid_reasons:
            df_op = df_sku_sizes_costs.copy()
        else:
            df_op = df_sku_sizes_costs[df_sku_sizes_costs['Обоснование для оплаты'].isin(valid_reasons)]
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

        df_sku_sizes_costs_result[sum_col] = result

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
    used_cost_reasons.extend(sales_returns_to_client_reasons)

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
        if not any([s > 0 or s < 0 for s in sum_by_other_cols]):
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

        # --- Добавляем столбец в df_sku_costs_result ---
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

        df_sku_sizes_costs_result[col_name] = result

    df_sku_other_costs_by_operation = pd.DataFrame(rows)

    # Удаляем временную колонку с размерами групп
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.drop(columns=['tmp_col'])
    # Сбрасываем индекс после группировки
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.reset_index()
    # Переводим артикул продавца в верхний регистр
    df_sku_sizes_costs_result['Артикул продавца'] = df_sku_sizes_costs_result['Артикул продавца'].str.upper()

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
    df_sku_sizes_costs_result = convert_currency(
        df_sku_sizes_costs_result,
        # columns_to_ignore=columns_to_ignore_transfer,
        k_exchange_rate=k_currency_to_rub
    )
    df_sku_costs_by_operation = convert_currency(
        df_sku_costs_by_operation,
        k_exchange_rate=k_currency_to_rub
    )

    # Формируем словарь с результатами
    result_sku_costs = {
        'df_sku_sizes_costs_result': df_sku_sizes_costs_result,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids
    }

    return result_sku_costs

def calc_sku_costs_v4(
    df_realization_report_with_sizes,
    groupby_cols=['Артикул WB', 'Артикул продавца', 'Размер']
):
    # Создаем копию для избежания изменений в оригинальном df
    df_sales_returns = df_realization_report_with_sizes.copy()

    # ------ Выбираем строки с артикулом ------
    df_sku_costs = df_sales_returns.loc[
        (df_sales_returns['Артикул продавца'] != '') & (df_sales_returns['Артикул продавца'].notna()),
        :
        ]

    # Формируем id строк с расходами по артикулам
    sku_costs_ids = df_sku_costs['id'].to_list()

    # --- Нормализация размера: пустые -> '0' ---
    df_sku_costs['Размер'] = (
        df_sku_costs['Размер']
        .astype(str)
        .str.strip()
        .replace({'nan': '', 'None': ''})
    )
    df_sku_costs['Размер'] = df_sku_costs['Размер'].replace('', '0')

    # Все строки с артикулом и "размером" (включая '0')
    df_sku_sizes_costs = df_sku_costs.copy()
    sku_sizes_costs_ids = df_sku_sizes_costs['id'].to_list()

    # ------ Заготовка результата по SKU ------
    df_sku_sizes_costs_result = pd.DataFrame()
    # Временная колонка с размерами групп
    df_sku_sizes_costs_result['tmp_col'] = df_sku_sizes_costs.groupby(groupby_cols).size()

    # Колонки, по которым считаем продажи и возвраты в рублях
    sales_returns_spp = 'Цена розничная с учетом согласованной скидки'
    sales_returns_without_spp = 'Сумма продаж (возвратов)'
    sales_returns_to_client = 'К перечислению продавцу за реализованный товар'

    # ------ Расчет продаж ------
    sales_ids = df_sku_sizes_costs.loc[
        df_sku_sizes_costs['Обоснование для оплаты'] == 'Продажа', 'id'
    ].to_list()

    df_sku_sizes_costs_result['Продажи шт'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)['Количество']
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб с СПП'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)[sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб без СПП'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(sales_ids), :]
        .groupby(groupby_cols)[sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Продажи руб к перечислению'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['Тип документа'] == 'Продажа', :]
        .groupby(groupby_cols)[sales_returns_to_client]
        .sum()
    )

    # ------ Расчет возвратов ------
    returns_ids = df_sku_sizes_costs.loc[
        df_sku_sizes_costs['Обоснование для оплаты'] == 'Возврат', 'id'
    ].to_list()

    df_sku_sizes_costs_result['Возвраты шт'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)['Количество']
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб с СПП'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)[sales_returns_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб без СПП'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['id'].isin(returns_ids), :]
        .groupby(groupby_cols)[sales_returns_without_spp]
        .sum()
    )
    df_sku_sizes_costs_result['Возвраты руб к перечислению'] = (
        df_sku_sizes_costs
        .loc[df_sku_sizes_costs['Тип документа'] == 'Возврат', :]
        .groupby(groupby_cols)[sales_returns_to_client]
        .sum()
    )

    # Находим id начислений, где сумма к перечислению != 0 и обоснование Продажа/Возврат
    sales_returns_to_client_ids = (
        df_sku_sizes_costs
        .loc[
            (df_sku_sizes_costs[sales_returns_to_client] != 0)
            & (df_sku_sizes_costs['Обоснование для оплаты'].isin(['Продажа', 'Возврат'])),
            :
        ]
        ['id']
        .to_list()
    )

    # Считаем итоговые суммы по Продажам и Возвратам (без разбивки по артикулам)
    df_sku_sales_returns_by_operation = (
        df_sku_sizes_costs
        .loc[
            df_sku_sizes_costs['id'].isin(
                sales_ids + returns_ids + sales_returns_to_client_ids
            ),
            :
        ]
        .groupby('Тип документа')
        .agg(**{
            'С СПП': (sales_returns_spp, 'sum'),
            'Без СПП': (sales_returns_without_spp, 'sum'),
            'К перечислению': (sales_returns_to_client, 'sum'),
        })
        .rename_axis('Группа расходов')
        .reset_index()
    )

    # Заполняем пропуски
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.fillna(0)

    # Итоговые колонки по продажам и возвратам
    df_sku_sizes_costs_result['Продажи минус возвраты шт'] = (
        df_sku_sizes_costs_result['Продажи шт'] - df_sku_sizes_costs_result['Возвраты шт']
    )
    df_sku_sizes_costs_result['Продажи минус возвраты руб с СПП'] = (
        df_sku_sizes_costs_result['Продажи руб с СПП'] - df_sku_sizes_costs_result['Возвраты руб с СПП']
    )
    df_sku_sizes_costs_result['Продажи минус возвраты руб без СПП'] = (
        df_sku_sizes_costs_result['Продажи руб без СПП'] - df_sku_sizes_costs_result['Возвраты руб без СПП']
    )
    df_sku_sizes_costs_result['Продажи минус возвраты руб к перечислению'] = (
        df_sku_sizes_costs_result['Продажи руб к перечислению'] - df_sku_sizes_costs_result['Возвраты руб к перечислению']
    )

    # ------ Основные расходы по столбцам ------
    custom_sum_cols = {
        'Вознаграждение WB без НДС': [''],
        'НДС с вознаграждения WB': [''],
        'Стоимость логистики': ['Логистика', 'Коррекция логистики', 'Доставка', 'Коррекция стоимости доставки'],
        'Стоимость платной приёмки': ['Платная приемка (Отчет реализации)', 'Обработка товара'],
        # 'Стоимость платной приемки': ['Платная приемка (Отчет реализации)', 'Обработка товара'],
        'Возмещение издержек по эквайрингу': [''],
        'Возмещение за выдачу и возврат товаров на ПВЗ': [''],
        'Возмещение издержек по перевозке': [
            'Возмещение издержек по перевозке/по складским операциям с товаром',
            'Возмещение издержек по перемещению и операционной обработке товара'
        ],
        'Компенсация скидки по программе лояльности': ['Компенсация скидки по программе лояльности'],
    }

    rows = []
    used_cost_reasons = []

    for sum_col, reasons_list in custom_sum_cols.items():
        valid_reasons = [
            r for r in reasons_list
            if r is not None and str(r).strip() != '' and not pd.isna(r)
        ]

        if not valid_reasons:
            df_op = df_sku_sizes_costs.copy()
        else:
            df_op = df_sku_sizes_costs[
                df_sku_sizes_costs['Обоснование для оплаты'].isin(valid_reasons)
            ]
            if df_op.empty:
                continue

        sales_sum = df_op[df_op['Тип документа'] == 'Продажа'][sum_col].sum()
        returns_sum = df_op[df_op['Тип документа'] == 'Возврат'][sum_col].sum()
        empty_sum = df_op[
            df_op['Тип документа'].isna()
            | (df_op['Тип документа'].astype(str).str.strip() == '')
        ][sum_col].sum()

        doc_types_not_empty = df_op['Тип документа'].dropna().astype(str).str.strip()
        doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
        has_sales = 'Продажа' in doc_types_not_empty.values
        has_returns = 'Возврат' in doc_types_not_empty.values

        if has_sales and has_returns:
            total = sales_sum - returns_sum + empty_sum
        else:
            total = df_op[sum_col].sum()

        rows.append({
            'Столбец': sum_col,
            'Обоснования для оплаты': valid_reasons if valid_reasons else None,
            'Обоснование штрафов и доплат': None,
            'Продажа': sales_sum,
            'Возврат': returns_sum,
            'Расходы руб': total
        })
        used_cost_reasons.extend(valid_reasons)

        group_sum_sales = (
            df_op[df_op['Тип документа'] == 'Продажа']
            .groupby(groupby_cols)[sum_col].sum()
        )
        group_sum_returns = (
            df_op[df_op['Тип документа'] == 'Возврат']
            .groupby(groupby_cols)[sum_col].sum()
        )
        group_sum_empty = (
            df_op[
                df_op['Тип документа'].isna()
                | (df_op['Тип документа'].astype(str).str.strip() == '')
            ]
            .groupby(groupby_cols)[sum_col].sum()
        )

        if has_sales and has_returns:
            result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
        else:
            result = df_op.groupby(groupby_cols)[sum_col].sum()

        df_sku_sizes_costs_result[sum_col] = result

    df_sku_main_costs_by_operation = pd.DataFrame(rows)

    # Добавляем к учтенным обоснованиям те, которые учли в К перечислению продавцу
    sales_returns_to_client_reasons = (
        df_sku_costs
        .loc[
            df_sku_costs['id'].isin(sales_returns_to_client_ids),
            'Обоснование для оплаты'
        ]
        .unique()
        .tolist()
    )
    used_cost_reasons.extend(sales_returns_to_client_reasons)

    # ------ Прочие расходы (штрафы/доплаты/прочее) ------
    df_sku_other_costs = df_sku_costs.loc[
        ~df_sku_costs['Обоснование для оплаты'].isin(used_cost_reasons)
    ].copy()

    other_sum_columns = [
        'Штрафы',
        'Доплаты',
        'Прочие удержания\\Выплаты',
    ]

    # ВАЖНО: заполняем NaN в "Обоснование штрафов и доплат" обоснованием для оплаты,
    # чтобы строки без второго уровня не терялись при groupby
    df_sku_other_costs['Обоснование штрафов и доплат'] = (
        df_sku_other_costs['Обоснование штрафов и доплат']
        .fillna(df_sku_other_costs['Обоснование для оплаты'])
    )

    rows = []
    for _, df_group in df_sku_other_costs.groupby(
        ['Обоснование для оплаты', 'Обоснование штрафов и доплат']
    ):
        obsh_pay = df_group['Обоснование для оплаты'].iloc[0]
        obsh_fine = df_group['Обоснование штрафов и доплат'].iloc[0]

        # Если штрафов и доплат нет (пусто/NaN/пробелы) — используем обоснование для оплаты
        col_name = obsh_fine if pd.notna(obsh_fine) and str(obsh_fine).strip() else obsh_pay

        # Если и обоснование пустое (редкий случай) — пропускаем
        if pd.isna(col_name) or str(col_name).strip() == '':
            continue

        sum_by_other_cols = [
            df_group[col].sum()
            for col in other_sum_columns
            if col in df_group.columns and not df_group[col].isnull().all()
        ]
        if not any([s > 0 or s < 0 for s in sum_by_other_cols]):
            continue

        used_col = next(
            (col for col in other_sum_columns
             if col in df_group.columns and not df_group[col].isnull().all()),
            'Итого руб'
        )

        doc_types_not_empty = df_group['Тип документа'].dropna().astype(str).str.strip()
        doc_types_not_empty = doc_types_not_empty[doc_types_not_empty != '']
        has_sales = 'Продажа' in doc_types_not_empty.values
        has_returns = 'Возврат' in doc_types_not_empty.values

        sum_sales = df_group.loc[df_group['Тип документа'] == 'Продажа', used_col].sum()
        sum_returns = df_group.loc[df_group['Тип документа'] == 'Возврат', used_col].sum()
        sum_empty = df_group.loc[
            df_group['Тип документа'].isna()
            | (df_group['Тип документа'].astype(str).str.strip() == ''),
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
                df_group['Тип документа'].isna()
                | (df_group['Тип документа'].astype(str).str.strip() == '')
            ]
            .groupby(groupby_cols)[used_col].sum()
        )

        if has_sales and has_returns:
            result = group_sum_sales.sub(group_sum_returns, fill_value=0).add(group_sum_empty, fill_value=0)
        else:
            result = df_group.groupby(groupby_cols)[used_col].sum()

        df_sku_sizes_costs_result[col_name] = result

    df_sku_other_costs_by_operation = pd.DataFrame(rows)
    # Удаляем временную колонку
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.drop(columns=['tmp_col'])
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.reset_index()

    # Переводим артикул продавца в верхний регистр
    df_sku_sizes_costs_result['Артикул продавца'] = (
        df_sku_sizes_costs_result['Артикул продавца'].str.upper()
    )

    # === Удаляем "нулевые" размеры без продаж и возвратов (они будут учтены позже) ===
    mask_zero_size_no_moves = (
        (df_sku_sizes_costs_result['Размер'] == '0')
        & (df_sku_sizes_costs_result['Продажи шт'] == 0)
        & (df_sku_sizes_costs_result['Возвраты шт'] == 0)
    )

    # Сохраняем эти строки отдельно как "расходы без размеров"
    df_raw_no_size_costs = df_sku_sizes_costs_result.loc[mask_zero_size_no_moves].copy()

    # Удаляем их из основного результата
    df_sku_sizes_costs_result = df_sku_sizes_costs_result.loc[~mask_zero_size_no_moves, :]

    # Объединяем Итоговые суммы по SKU в один датафрейм
    df_sku_costs_by_operation = pd.concat(
        [df_sku_main_costs_by_operation, df_sku_other_costs_by_operation],
        ignore_index=True
    )

    # Конвертируем валюту в рубли (если необходимо)
    df_sku_sizes_costs_result = convert_currency(
        df_sku_sizes_costs_result,
        k_exchange_rate=k_currency_to_rub
    )
    df_sku_costs_by_operation = convert_currency(
        df_sku_costs_by_operation,
        k_exchange_rate=k_currency_to_rub
    )

    result_sku_costs = {
        'df_sku_sizes_costs_result': df_sku_sizes_costs_result,
        'df_sku_costs_by_operation': df_sku_costs_by_operation,
        'sku_costs_ids': sku_costs_ids,
        'df_raw_no_size_costs': df_raw_no_size_costs,  # <- расходы без размеров и без движений
    }
    return result_sku_costs

# Функция выгрузки и расчета расходов по платому хранению
def calc_paid_storage_costs(
        headers,
        df_realization_report_with_sizes,
        date_start,
        date_end
    ):

    # df, в который будем помещать итоговый результат
    df_paid_storage_stats = pd.DataFrame()

    # df, в который будем помещать результаты выгрузки отчетов апи
    df_paid_storage = pd.DataFrame()

    # Формируем недели, из которых был собран отчет о реализации и где были расходы на хранение
    paid_storage_dates = (
        df_realization_report_with_sizes
        .loc[
            df_realization_report_with_sizes['Обоснование для оплаты'].isin(['Хранение (Отчет реализации)']),
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
        # Разбиваем на интервалы по 7 дней
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
    df_sku_sizes_costs_result = result_sku_costs['df_sku_sizes_costs_result'].copy()
    # Получаем df с расходами на пл. хранение
    df_paid_storage_stats = result_paid_storage['df_paid_storage_stats'].copy()
    # Объединяем расходы по SKU из всех отчетов
    df_sku_sizes_costs_union = pd.concat([
        df_sku_sizes_costs_result,
        #df_orders,
        df_paid_storage_stats
    ])
    # Колонки, по которым будет производиться группировка
    groupby_cols = ['Артикул WB', 'Артикул продавца', 'Размер']
    # Считаем суммы
    df_sku_sizes_costs_union = (
        df_sku_sizes_costs_union
        .groupby(groupby_cols)
        .sum()
        .reset_index()
    )
    # Считаем долю продаж в каждом артикуле в каждом размере (изменится, если добавились новые артикулы)
    df_sku_sizes_costs_union['Продажи в артикуле'] = df_sku_sizes_costs_union.groupby('Артикул WB')['Продажи шт'].transform('sum')
    df_sku_sizes_costs_union['Доля продаж в артикуле'] = df_sku_sizes_costs_union['Продажи шт'] / df_sku_sizes_costs_union['Продажи в артикуле']
    df_sku_sizes_costs_union['Количество размеров в артикуле'] = df_sku_sizes_costs_union.groupby('Артикул WB').transform('size')
    # Перемещаем колонки в начало df
    df_sku_sizes_costs_union = move_columns(
        df_sku_sizes_costs_union,
        ['Количество размеров в артикуле', 'Продажи в артикуле', 'Доля продаж в артикуле'],
        position='Продажи шт',
        insert_type='before'
    )

    return df_sku_sizes_costs_union


# Функция выгрузки и расчета расходов по платной приемке
def calc_paid_acceptance_costs(
        headers,
        df_sku_sizes_costs_union,
        df_realization_report_with_sizes,
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
    df_realization_report_ = df_realization_report_with_sizes.copy()

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
            df_sku_sizes_costs_union_ = df_sku_sizes_costs_union.copy()
            # Мерджим с расходами на приемку, чтобы получить количество размеров в одном артикуле
            df_paid_acceptance_sku_sizes = df_paid_acceptance_stats.merge(
                df_sku_sizes_costs_union_[[
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


# Функция получения дат расходов по ВБ / WB Продвижению
def get_promotion_dates(df_realization_report_with_sizes):
    # Регулярка для строк вида:
    # "Оказание услуг «ВБ.Продвижение»"
    # "Оказание услуг «ВБ Продвижение», документ №123"
    # "Оказание услуг «WB Продвижение», документ №123"
    promo_regex = re.compile(
        r'^оказание услуг\s+«(?:вб\.?\s*|wb\s*)продвижение»(?:, документ №\d+)?$',
        flags=re.IGNORECASE
    )

    # Фильтр по столбцу "Обоснование штрафов и доплат"
    mask = (
        df_realization_report_with_sizes['Обоснование штрафов и доплат']
        .astype(str)  # на случай NaN
        .str.strip()
        .str.contains(promo_regex, regex=True, na=False)
    )

    promotion_dates = df_realization_report_with_sizes.loc[
        mask,
        ['Дата начала отчётного периода', 'Дата конца отчётного периода', 'Итого руб']
    ]

    if not promotion_dates.empty:
        # Даты для API рекламы
        date_start_promotion = (
            pd.to_datetime(promotion_dates['Дата начала отчётного периода'].min())
              .strftime('%Y-%m-%d')
        )
        date_end_promotion = (
            pd.to_datetime(promotion_dates['Дата конца отчётного периода'].max())
              .strftime('%Y-%m-%d')
        )
        # Даты для заказов
        date_start_orders = date_start_promotion + 'T00:00:00.000Z'
        date_end_orders = date_end_promotion + 'T23:59:59.000Z'
    else:
        date_start_promotion = None
        date_end_promotion = None
        date_start_orders = None
        date_end_orders = None

    return {
        'date_start_promotion': date_start_promotion,
        'date_end_promotion': date_end_promotion,
        'date_start_orders': date_start_orders,
        'date_end_orders': date_end_orders,
    }

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
def calc_companies_by_sku_and_size(df_companies_stats_processed, df_costs_history, df_sku_sizes_costs_union):
    # Создаем копию для избежания изменений в оригинальном df
    # df_costs_history_ = df_costs_history.copy()
    df_companies_stats_by_sku = df_companies_stats_processed.copy()
    df_sku_sizes_costs_union_ = df_sku_sizes_costs_union.copy()
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
        df_sku_sizes_costs_union_[[
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
    # total_sales = df_sku_sizes_costs_union_['Продажи шт'].sum()
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



# Функция расчета расходов по рекламным кампаниям (v2)
def calc_companies_costs_v2(df_realization_report_with_sizes, df_sku_sizes_costs_union, date_report):
    # Ищем даты, в которых были расходы на ВБ Продвижение
    promotion_dates = get_promotion_dates(df_realization_report_with_sizes)

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
        df_companies_stats = get_company_stats(
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
            df_sku_sizes_costs_union
        )

    # Формируем словарь с результатами
    result_companies_costs = {
        'df_companies_stats': df_companies_stats,
        'df_companies_stats_by_sku': df_companies_stats_by_sku
    }


    return result_companies_costs


# Функция расчета расходов списаний за отзывы
def calc_reviews_costs(df_realization_report_with_sizes, df_sku_sizes_costs_union):
    # Ищем расходы списаний за отзывы
    # Создаём маску для строк по паттерну
    reviews_costs_text_pattern = r'Списание за отзыв.*акция №\d+, товар \d+$'
    # Колонка, в которой ищем списание за отзывы
    column_to_search = 'Обоснование штрафов и доплат'
    mask_reviews_costs = (
        df_realization_report_with_sizes
        [column_to_search]
        .str.contains(
            reviews_costs_text_pattern,
            na=False,
            regex=True
        )
    )
    # Делаем выборку по списаниям за отзывы
    df_reviews_sku_costs = (
        df_realization_report_with_sizes
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
        df_sku_sizes = df_sku_sizes_costs_union.copy()
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
        df_realization_report_with_sizes
        .loc[
            df_realization_report_with_sizes[column_to_search]
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
        df_sku_sizes_costs_union,
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
        df_sku_sizes_costs_union,
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


# Функция расчет расходов, у которых есть артикул, но нет размера
def calc_sku_no_size_costs(df_sku_sizes_costs_all, df_realization_report_with_sizes):

    # Создаем копию для избежания изменений в оригинальном df
    df_sku_no_sizes_costs = df_realization_report_with_sizes.copy()
    # Выбираем тип документа продажа или возврат
    # df_sales_returns = df_sales_returns.loc[df_sales_returns['Тип документа'].isin(['Продажа', 'Возврат']), :]

    # ------ Расчет расходов, которые не разбиты по размерам ------

    # Выбираем расходы, где есть артикул
    df_sku_no_sizes_costs = df_sku_no_sizes_costs.loc[df_sku_no_sizes_costs['Артикул продавца'] != '', :]
    # Выбираем расходы, где нет размера
    df_sku_no_sizes_costs = df_sku_no_sizes_costs.loc[
        (df_sku_no_sizes_costs['Размер'].isnull())
        | (df_sku_no_sizes_costs['Размер'] == '0')
        ,
        :
    ]
    # Формируем id строк с расходами по артикулам
    sku_no_sizes_ids = df_sku_no_sizes_costs['id'].to_list()

    # ------ Расчет расходов без размеров ------

    # 2. Группируем по Артикулу продавца
    groupby_cols_no_size = ['Артикул WB', 'Артикул продавца']

    custom_sum_cols_no_sizes = {
        # Столбец для суммирования: [Обоснования для оплаты, которые учитывает этот столбец]
        'Возмещение издержек по перевозке': [
            'Возмещение издержек по перевозке/по складским операциям с товаром',
            'Возмещение издержек по перемещению и операционной обработке товара'
        ],
    }

    result_frames = []
    # 3. Считаем суммы по расходам без размеров
    for sum_col, reasons_list in custom_sum_cols_no_sizes.items():
        valid_reasons_no_sizes = [r for r in reasons_list if r and str(r).strip() and not pd.isna(r)]
        df_op = df_sku_no_sizes_costs[df_sku_no_sizes_costs['Обоснование для оплаты'].isin(valid_reasons_no_sizes)]
        if df_op.empty:
            continue

        group_sales = df_op[df_op['Тип документа'] == 'Продажа'].groupby(groupby_cols_no_size)[sum_col].sum()
        group_returns = df_op[df_op['Тип документа'] == 'Возврат'].groupby(groupby_cols_no_size)[sum_col].sum()
        group_empty = df_op[
            df_op['Тип документа'].isna() | (df_op['Тип документа'].astype(str).str.strip() == '')
        ].groupby(groupby_cols_no_size)[sum_col].sum()

        # Финальный результат по формуле
        result = group_sales.sub(group_returns, fill_value=0).add(group_empty, fill_value=0).rename(sum_col)
        result = result.reset_index()
        result_frames.append(result)

        # Добавляем к использованным обоснованиям для оплаты новые
        # used_cost_reasons.extend(valid_reasons_no_sizes)

    # 4. Собираем расходы без размеров в итоговый DataFrame
    if result_frames:
        df_sku_no_size_costs_total = result_frames[0]
        for frame in result_frames[1:]:
            df_sku_no_size_costs_total = pd.merge(
                df_sku_no_size_costs_total, frame, on=groupby_cols_no_size, how='outer'
            )
    else:
        df_sku_no_size_costs_total = pd.DataFrame(columns=groupby_cols_no_size)

    # Переводим Артикул продавца в верхний регистр
    df_sku_no_size_costs_total['Артикул продавца'] = df_sku_no_size_costs_total['Артикул продавца'].str.upper()
    # 5. Мерджим с расходами по размерам, чтобы получить количество размеров в артикуле
    df_sizes = (
        df_sku_sizes_costs_all
        .loc[:, ['Артикул продавца', 'Размер', 'Количество размеров в артикуле']]
    )
    df_sku_no_size_costs_total = df_sku_no_size_costs_total.merge(
        df_sizes[['Артикул продавца', 'Размер', 'Количество размеров в артикуле']],
        on='Артикул продавца',
        how='left'
    )
    # 6. Заполняем пропуски в размере
    df_sku_no_size_costs_total = df_sku_no_size_costs_total.fillna({
        'Артикул WB': 0,
        'Размер': '0',
        'Количество размеров в артикуле': 1
    })

    # 7. Распределяем расходы без размеров от количества размеров в артикуле
    for col in custom_sum_cols_no_sizes.keys():
        if col in df_sku_no_size_costs_total.columns:
            df_sku_no_size_costs_total[col] = df_sku_no_size_costs_total[col] / df_sku_no_size_costs_total['Количество размеров в артикуле']

    # 8. Удаляем колонку с количеством размеров
    df_sku_no_size_costs_total = df_sku_no_size_costs_total.drop(columns=['Количество размеров в артикуле'])

    # 9. Переводим артикул и размер в index
    df_sku_no_size_costs_total = (
        df_sku_no_size_costs_total
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'])
        .sum(numeric_only=True)
    )
    # Если в расходах по SKU уже встретились расходы без размеров, то удаляем их
    for col in df_sku_no_size_costs_total.columns:
        if col in df_sku_sizes_costs_all.columns:
            df_sku_sizes_costs_all = df_sku_sizes_costs_all.drop(columns=col)

    # 10. Объединяем с основными расходами по SKU
    df_sku_sizes_costs_final = (
        pd.concat([
            df_sku_sizes_costs_all,
            df_sku_no_size_costs_total.reset_index()
        ])
        # Считаем сумму с группировкой по индексу
       .groupby(['Артикул WB', 'Артикул продавца', 'Размер'], as_index=False)
       .sum(numeric_only=True)
    )

    return df_sku_sizes_costs_final

def calc_sku_no_size_costs_v2(df_sku_sizes_costs_all, result_sku_costs):
    """
    df_sku_sizes_costs_all — результат df_sku_sizes_costs_result из calc_sku_costs_v3
    result_sku_costs — словарь-результат calc_sku_costs_v3, в т.ч. с ключом 'df_raw_no_size_costs'
    """

    # Берём сырые расходы без размеров, которые мы вытащили в первой функции
    df_raw_no_size_costs = result_sku_costs.get('df_raw_no_size_costs')

    # Если нечего распределять — возвращаем исходный df_sku_sizes_costs_all без изменений
    if df_raw_no_size_costs is None or df_raw_no_size_costs.empty:
        return df_sku_sizes_costs_all.copy()

    # Копируем, чтобы не модифицировать исходный df_raw_no_size_costs
    df_raw = df_raw_no_size_costs.copy()

    # Приводим артикул продавца к верхнему регистру для согласованности
    df_raw['Артикул продавца'] = df_raw['Артикул продавца'].str.upper()

    # Группировка без размера
    groupby_cols_no_size = ['Артикул WB', 'Артикул продавца']

    # Определяем, какие числовые столбцы распределять
    numeric_cols = df_raw.select_dtypes(include='number').columns.tolist()
    # Служебные столбцы, которые не нужно распределять
    cols_to_exclude = [
        'Артикул WB',
        'Артикул_Размер',
        'Артикул продавца',
        'Размер',
        'Артикул 1С',
        'Статус',
        'Цвет',
        'Себестоимость',
        'Количество размеров в артикуле',
        'Продажи в артикуле',
        'Доля продаж в артикуле',
        'Продажи шт',
        'Продажи руб с СПП',
        'Продажи руб без СПП',
        'Продажи руб к перечислению',
        'Возвраты шт',
        'Возвраты руб с СПП',
        'Возвраты руб без СПП',
        'Возвраты руб к перечислению',
        'Продажи минус возвраты шт',
        'Продажи минус возвраты руб с СПП',
        'Продажи минус возвраты руб без СПП',
        'Продажи минус возвраты руб к перечислению'
    ]  # можно добавить другие служебные колонки, если нужно
    expense_cols = [c for c in numeric_cols if c not in cols_to_exclude]

    if not expense_cols:
        # Нет числовых расходов, нечего распределять
        return df_sku_sizes_costs_all.copy()

    # Агрегируем расходы по (Артикул WB, Артикул продавца) без учёта размера
    df_grouped = (
        df_raw
        .groupby(groupby_cols_no_size, as_index=False)[expense_cols]
        .sum()
    )

    # Берём из df_sku_sizes_costs_all список размеров и "Количество размеров в артикуле"
    df_sizes = (
        df_sku_sizes_costs_all
        .loc[:, ['Артикул продавца', 'Размер', 'Количество размеров в артикуле']]
        .drop_duplicates()
    )

    # Мерджим, чтобы на каждый (Артикул WB, Артикул продавца) получить все размеры и их количество
    df_merged = df_grouped.merge(
        df_sizes,
        on='Артикул продавца',
        how='left'
    )

    # Если по какому-либо артикулу нет размеров в df_sku_sizes_costs_all —
    # считаем, что у него один "виртуальный" размер '0'
    df_merged['Размер'] = df_merged['Размер'].fillna('0')
    df_merged['Количество размеров в артикуле'] = df_merged['Количество размеров в артикуле'].fillna(1)

    # Делим каждый расход на количество размеров в артикуле
    for col in expense_cols:
        df_merged[col] = df_merged[col] / df_merged['Количество размеров в артикуле']

    # Убираем вспомогательную колонку
    df_merged = df_merged.drop(columns=['Количество размеров в артикуле'])

    # Приводим к формату (Артикул WB, Артикул продавца, Размер) + суммы
    df_no_size_distributed = (
        df_merged
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'], as_index=False)
        .sum(numeric_only=True)
    )

    # Объединяем с исходным df_sku_sizes_costs_all, не изменяя его
    df_combined = pd.concat(
        [df_sku_sizes_costs_all, df_no_size_distributed],
        ignore_index=True
    )

    df_sku_sizes_costs_final = (
        df_combined
        .groupby(['Артикул WB', 'Артикул продавца', 'Размер'], as_index=False)
        .sum(numeric_only=True)
    )

    return df_sku_sizes_costs_final

# Функция расчета расходов, которые не привязаны к артикулам (v3)
def calc_other_costs_v3(result_sku_costs, df_sku_sizes_costs_final):
    sku_costs_ids = result_sku_costs['sku_costs_ids']

    # Копия df с расходами по SKU
    df_sku_and_other_costs = df_sku_sizes_costs_final.copy()

    # --- 1. Берём строки без SKU ---
    df_other_costs = df_realization_report_with_sizes.loc[
        ~df_realization_report_with_sizes['id'].isin(sku_costs_ids), :
    ]
    other_costs_ids = df_realization_report_with_sizes.loc[
        ~df_realization_report_with_sizes['id'].isin(sku_costs_ids), 'id'
    ].to_list()

    # Заготовка пустого df на случай отсутствия прочих расходов
    df_other_costs_by_operation = pd.DataFrame(
        columns=['Тип расходов', 'Продажа', 'Возврат', 'Итого прочие расходы']
    )

    # Флаг: будем ли вообще что‑то распределять
    do_distribute = True

    # --- 2. Если прочих расходов нет вообще ---
    if df_other_costs.empty:
        do_distribute = False
    else:
        # --- 3. Тип расходов ---
        df_other_costs['Тип расходов'] = np.where(
            (df_other_costs['Обоснование штрафов и доплат'] != '') &
            (~df_other_costs['Обоснование штрафов и доплат'].isna()),
            df_other_costs['Обоснование штрафов и доплат'],
            df_other_costs['Обоснование для оплаты']
        )

        # --- 4. Группировка прочих расходов ---
        other_costs_result = []

        for op_type, df_op in df_other_costs.groupby('Тип расходов'):
            doc_types = df_op['Тип документа'].dropna().astype(str).str.strip()
            doc_types = doc_types[doc_types != '']

            has_sales   = 'Продажа' in doc_types.values
            has_returns = 'Возврат' in doc_types.values

            if has_sales and has_returns:
                sum_sales   = df_op.loc[df_op['Тип документа'] == 'Продажа', 'Итого руб'].sum()
                sum_returns = df_op.loc[df_op['Тип документа'] == 'Возврат',  'Итого руб'].sum()
                sum_empty   = df_op.loc[
                    df_op['Тип документа'].isna() |
                    (df_op['Тип документа'].astype(str).str.strip() == ''),
                    'Итого руб'
                ].sum()
                total = sum_sales - sum_returns + sum_empty
                sale  = sum_sales
                ret   = sum_returns
            else:
                total = df_op['Итого руб'].sum()
                sale  = df_op.loc[df_op['Тип документа'] == 'Продажа', 'Итого руб'].sum()
                ret   = df_op.loc[df_op['Тип документа'] == 'Возврат',  'Итого руб'].sum()

            other_costs_result.append({
                'Тип расходов': op_type,
                'Продажа': sale,
                'Возврат': ret,
                'Итого прочие расходы': total
            })

        df_other_costs_by_operation = pd.DataFrame(other_costs_result)

        # Если после группировки ничего не получилось — дальше не распределяем
        if df_other_costs_by_operation.empty:
            do_distribute = False

    # --- 5. Условное исключение некоторых типов расходов ---
    if do_distribute and not df_other_costs_by_operation.empty:
        conditional_exclude_flags = {
            'Платное хранение (отдельный отчет)': [
                'Хранение (Отчет реализации)',
                'Стоимость хранения',
            ],
            'Платная приемка (отдельный отчет)': [
                'Платная приемка (Отчет реализации)',
                'Стоимость платной приемки',
            ],
        }

        exclude_condition_mask = pd.Series(False, index=df_other_costs_by_operation.index)

        for separate_col, realization_types in conditional_exclude_flags.items():
            if separate_col in df_sku_and_other_costs.columns:
                if df_sku_and_other_costs[separate_col].sum() > 0:
                    for realization_type in realization_types:
                        exclude_condition_mask |= (
                            df_other_costs_by_operation['Тип расходов']
                            .astype(str).str.strip().str.lower()
                            == realization_type.lower()
                        )

        # --- 6. Фильтр по regex для исключения ---
        exclude_costs_list = [
            r'списание за отзыв.*акция №\d+, товар \d+$',
            # 'оказание услуг «вб.продвижение»',
            # 'оказание услуг «wb продвижение»',
        ]
        exclude_costs_filter = '|'.join(f"(?:{p})" for p in exclude_costs_list)

        exclude_mask = (
            df_other_costs_by_operation['Тип расходов']
            .astype(str).str.strip().str.lower()
            .str.contains(exclude_costs_filter, regex=True, na=False)
        )

        mask = ~(exclude_mask | exclude_condition_mask)
        df_other_costs_by_operation = df_other_costs_by_operation.loc[mask].reset_index(drop=True)

        # Если после фильтров всё удалили — не распределяем
        if df_other_costs_by_operation.empty:
            do_distribute = False

    # --- 7. Перевод в валюту ---
    if not df_other_costs_by_operation.empty:
        df_other_costs_by_operation = convert_currency(
            df_other_costs_by_operation,
            k_exchange_rate=k_currency_to_rub
        )
    else:
        do_distribute = False

    # --- 8. Распределение прочих расходов по SKU (только если решили распределять) ---
    if do_distribute:
        total_sales = df_sku_and_other_costs['Продажи шт'].sum()

        # Логика: если продаж нет — НЕ распределяем
        if total_sales == 0:
            do_distribute = False
        else:
            for operation_type in df_other_costs_by_operation['Тип расходов']:
                operation_costs = df_other_costs_by_operation.loc[
                    df_other_costs_by_operation['Тип расходов'] == operation_type,
                    'Итого прочие расходы'
                ].values[0]

                df_sku_and_other_costs[operation_type] = (
                    df_sku_and_other_costs['Продажи шт'] / total_sales * operation_costs
                )

    # --- 9. Финальный return (один) ---
    result_other_costs = {
        'df_sku_and_other_costs': df_sku_and_other_costs,
        'df_other_costs_by_operation': df_other_costs_by_operation,
        'other_costs_ids': other_costs_ids
    }
    return result_other_costs


# Функция добавления данных из справочной таблицы
def add_data_from_catalog(result_other_costs):
    # Чтение справочной таблицы
    # catalog = pd.read_excel(
    #     marketplace_dir_path
    #     / "Clients"
    #     / client_name
    #     / "catalog"
    #     / f"Справочная_таблица_{client_name}_WB.xlsx"
    # )
    catalog = get_reference_table()
    # Создаем колонку размера, если её нет в справочной таблице
    if 'Размер' not in catalog.columns:
        catalog['Размер'] = '0'

    # catalog['Размер'] = pd.to_numeric(catalog['Размер'], errors='coerce').fillna(0).astype(int)
    # Создаем остальные колонки, если их нет в справочной таблице
    for col in catalog_finace_svod_columns:
        if col not in catalog.columns:
            catalog[col] = np.nan

    # Заполняем пропуски в размерах
    catalog['Размер'] = catalog['Размер'].fillna('0')

    # Переводим артикул продавца в верхний регистр
    catalog['Артикул продавца'] = catalog['Артикул продавца'].str.upper()

    # Переводим себестоимость в число
    catalog['Себестоимость'] = pd.to_numeric(catalog['Себестоимость'], errors='coerce')

    # В каждой группе артикулов искуственно создаем артикул с размером 0, если такого размера нет
    # Словарь, где для каждого нужного столбца указана агрегация
    if client_name in ['SENS', 'SENS_IP', 'SENS_KZ']:
        agg_dict = {
            'Себестоимость': 'mean',
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


def add_extra_metrics_wb(client_name: str, df: pd.DataFrame) -> pd.DataFrame:
    """
    Дополнительные метрики для клиентов на ВБ.

    В рамках артикула = суммирование по колонке 'Артикул продавца'.

    Ожидаемые колонки:
      - 'Статус'                (для SENS_IP: 'кз' / 'рф', регистр не важен)
      - 'Артикул продавца'
      - 'Продажи минус возвраты руб к перечислению'
      - 'Продажи минус возвраты шт'
      - 'Приход денег на РС'
      - 'Себестоимость'
    """
    if client_name not in ("SENS_IP", "SENS"):
        return df

    df_final = df.copy()

    required_cols = [
        'Статус',
        'Артикул продавца',
        'Продажи минус возвраты руб к перечислению',
        'Продажи минус возвраты шт',
        'Приход денег на РС',
        'Себестоимость',
    ]
    missing = [c for c in required_cols if c not in df_final.columns]
    if missing:
        logger.warning(
            f"Для расчёта дополнительных полей ВБ ({client_name}) не хватает колонок: "
            f"{missing}. Дополнительные расчёты пропущены."
        )
        return df_final

    # Общие агрегаты "в рамках артикула" — одни и те же для всех клиентов
    groupby_cols = ['Артикул продавца']

    grouped_common = (
        df_final
        .groupby(groupby_cols, as_index=False)
        .agg({
            'Продажи минус возвраты руб к перечислению': 'sum',
            'Продажи минус возвраты шт': 'sum',
            'Приход денег на РС': 'sum',
        })
        .rename(columns={
            'Продажи минус возвраты руб к перечислению':
                'Продажи минус возвраты руб к перечислению в рамках артикула',
            'Продажи минус возвраты шт':
                'Продажи минус возвраты шт в рамках артикула',
            'Приход денег на РС':
                'Приход денег на РС в рамках артикула',
        })
    )

    df_final = df_final.merge(grouped_common, on=groupby_cols, how='left')

    # Переставляем для удобства
    df_final = move_columns(
        df=df_final,
        columns_to_move=[
            'Продажи минус возвраты шт в рамках артикула',
            'Продажи минус возвраты руб к перечислению в рамках артикула',
        ],
        position='Продажи минус возвраты шт',
        insert_type='after'
    )
    df_final = move_columns(
        df=df_final,
        columns_to_move=['Приход денег на РС в рамках артикула'],
        position='Приход денег на РС',
        insert_type='after'
    )

    denom = df_final['Продажи минус возвраты шт в рамках артикула']
    sum_sales_rub_art = df_final['Продажи минус возвраты руб к перечислению в рамках артикула']
    sum_income_art = df_final['Приход денег на РС в рамках артикула']

    # --- Вариант для клиента SENS ---
    if client_name == "SENS":
        # 1) Заплатил клиент за 1 ед. артикула
        df_final['Заплатил клиент за 1 ед. артикула'] = np.where(
            denom != 0,
            sum_sales_rub_art / denom,
            np.nan
        )

        # 2) СРЕДНИЙ приход за 1 ед. артикула
        df_final['СРЕДНИЙ приход за 1 ед. артикула'] = np.where(
            denom != 0,
            sum_income_art / denom,
            np.nan
        )

        # 3) Приход за артикул
        df_final['Приход за артикул'] = sum_income_art

        # 4) Средняя маржа в руб за 1 ед
        df_final['Средняя маржа в руб за 1 ед'] = (
            df_final['СРЕДНИЙ приход за 1 ед. артикула'] - df_final['Себестоимость']
        )

        return df_final

    # --- Вариант для клиента SENS_IP (КЗ/РФ, как раньше) ---
    # нормализуем статус
    df_final['Статус_norm'] = df_final['Статус'].astype(str).str.lower()

    mask_kz = df_final['Статус_norm'].str.contains('кз', na=False)
    mask_rf = df_final['Статус_norm'].str.contains('рф', na=False)

    # "Заплатил клиент..." — общий для всех, как договорились с тобой выше
    df_final['Заплатил клиент за 1 ед. артикула'] = np.where(
        denom != 0,
        sum_sales_rub_art / denom,
        np.nan
    )

    # ---------- Блок: Статус содержит "кз" ----------
    if mask_kz.any():
        df_final.loc[mask_kz, 'ДС в КЗ (ВБ)'] = (
            df_final.loc[mask_kz, 'Приход денег на РС'] * 0.75
        )
        df_final.loc[mask_kz, 'ДС в РФ (ВБ)'] = (
            df_final.loc[mask_kz, 'Приход денег на РС'] * 0.25
        )

        sum_ds_kz_art = 0.75 * sum_income_art
        sum_ds_rf_art = 0.25 * sum_income_art

        df_final['ДС в КЗ за 1 ед. артикула (ВБ)'] = np.where(
            (mask_kz) & (denom != 0),
            sum_ds_kz_art / denom,
            np.nan
        )
        df_final['ДС в РФ за 1 ед. артикула (ВБ)'] = np.where(
            (mask_kz) & (denom != 0),
            sum_ds_rf_art / denom,
            np.nan
        )
        df_final['ИТОГО ДС В РФ ЗА АРТИКУЛ (ВБ)'] = np.where(
            mask_kz,
            sum_ds_rf_art,
            np.nan
        )

    # ---------- Блок: Статус содержит "рф" ----------
    if mask_rf.any():
        df_final['СРЕДНИЙ приход за 1 ед. артикула (ВБ)'] = np.where(
            (mask_rf) & (denom != 0),
            sum_income_art / denom,
            np.nan
        )
        df_final['Приход за артикул (ВБ)'] = np.where(
            mask_rf,
            sum_income_art,
            np.nan
        )
        df_final['Средняя маржа в руб за 1 ед (ВБ)'] = np.where(
            mask_rf,
            df_final['СРЕДНИЙ приход за 1 ед. артикула (ВБ)'] - df_final['Себестоимость'],
            np.nan
        )
        df_final['Маржа в руб за артикул (ВБ)'] = np.where(
            mask_rf,
            df_final['Средняя маржа в руб за 1 ед (ВБ)'] * denom,
            np.nan
        )

    df_final = df_final.drop(columns=['Статус_norm'], errors='ignore')

    return df_final

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

    # 1. Базовый список (как был)
    base_promotion_columns = [
        'Расходы по рекламным кампаниям',
        'Списания за отзывы',
        'Витрина Магазина',
        'Оказание услуг «ВБ.Медиа»',
        'Предоставление услуг по подписке «Джем»',
        'Стоимость участия в программе лояльности',
    ]

    # 2. regex для колонок "Оказание услуг ... Продвижение, документ №..."
    # promo_regex_list = ['jopa']
    promo_regex_list = [
        # r'^оказание услуг\s+«(?:вб\.?\s*|wb\s*)продвижение»(?:, документ №\d+)?$',   # продвижение
        r'^витрина\s+магазина(?:, документ №\d+)?$',                                # витрина магазина
        r'^предоставление\s+услуг\s+по\s+подписке\s+«джем»(?:, документ №\d+)?$',   # Джем
        r'^оказание\s+услуг\s+«(?:вб\.?\s*|wb\s*)медиа»(?:, документ №\d+)?$',      # WB Медиа
    ]

    promo_regex = re.compile('|'.join(promo_regex_list), flags=re.IGNORECASE)

    regex_promo_columns = [col for col in df_final_costs.columns if promo_regex.search(col)]    # Все колонки df_final_costs, которые подходят под regex
    regex_promo_columns = [
        col for col in df_final_costs.columns
        if promo_regex.search(col)
    ]

    # 3. Объединяем: явные + найденные по regex
    all_promotion_columns = list(dict.fromkeys(base_promotion_columns + regex_promo_columns))
    # dict.fromkeys сохраняет порядок и выбрасывает дубликаты
    # 4. Фильтруем по фактически существующим колонкам
    existing_promotion_columns = [
        col for col in df_final_costs.columns
        if col in all_promotion_columns
    ]

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
        ] + existing_promotion_columns
    )
    columns_for_final_costs = ([
        'Стоимость логистики',
        'Затраты на рекламу руб',
        'Платное хранение возвратов на ПВЗ более 3 дней',
        'Платное хранение (отдельный отчет)',
        'Платная приемка (отдельный отчет)',
        'Выявленные расхождения в карточке товара после приемки на складе WB',
        'Компенсация скидки по программе лояльности',
        'Занижение фактических габаритов упаковки товара',
        'Сторно. Выявленные расхождения в карточке товара после приемки на складе WB',
        'Подмена товара (с фото)',
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
    for col in ['Маржинальность на ед. товара', 'Маржинальность %']:
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

    # Рассчитываем дополнительные метрики для некоторых клиентов
    df_final_costs = add_extra_metrics_wb(client_name, df_final_costs)

    return df_final_costs, existing_promotion_columns, columns_for_final_costs


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
        'Приход денег на РС в рамках артикула',
        'Заплатил клиент за 1 ед. артикула (ВБ)', 'ДС в КЗ (ВБ)',
        'ДС в РФ (ВБ)', 'ДС в КЗ за 1 ед. артикула (ВБ)',
        'ДС в РФ за 1 ед. артикула (ВБ)', 'ИТОГО ДС В РФ ЗА АРТИКУЛ (ВБ)',
        'СРЕДНИЙ приход за 1 ед. артикула (ВБ)', 'Приход за артикул (ВБ)',
        'Средняя маржа в руб за 1 ед (ВБ)', 'Маржа в руб за артикул (ВБ)'
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


# Функция добавления дополнительных формул в итоговую таблицу для некоторых клиентов
def apply_wb_extra_formulas(final_ws, col_map, row, start_row, end_row, client_name):
    """
    Дополнительные формулы для некоторых клиентов на ВБ.
    В рамках артикула = суммирование по 'Артикул продавца'.
    Используются колонки:
      - Статус, Артикул продавца
      - Продажи минус возвраты руб к перечислению
      - Продажи минус возвраты шт
      - Приход денег на РС
      - Себестоимость
      - Продажи минус возвраты руб к перечислению в рамках артикула
      - Продажи минус возвраты шт в рамках артикула
      - Приход денег на РС в рамках артикула
      - и итоговые колонки с результатами.
    """
    # Работает только для клиентов SENS_IP и SENS
    if client_name not in ('SENS_IP', 'SENS'):
        return

    # --- 1. Агрегаты "в рамках артикула" через СУММЕСЛИМН ---
    base_needed = [
        'Артикул продавца',
        'Приход денег на РС',
        'Продажи минус возвраты шт',
        'Продажи минус возвраты руб к перечислению',
        'Продажи минус возвраты шт в рамках артикула',
        'Продажи минус возвраты руб к перечислению в рамках артикула',
        'Приход денег на РС в рамках артикула',
        'Заплатил клиент за 1 ед. артикула',
    ]
    if all(col in col_map for col in base_needed):
        art_col_letter  = col_map['Артикул продавца']
        inc_col_letter  = col_map['Приход денег на РС']
        qty_col_letter  = col_map['Продажи минус возвраты шт']
        rub_col_letter  = col_map['Продажи минус возвраты руб к перечислению']

        # Абсолютные диапазоны
        rng_art  = f"${art_col_letter}${start_row}:${art_col_letter}${end_row}"
        rng_inc  = f"${inc_col_letter}${start_row}:${inc_col_letter}${end_row}"
        rng_qty  = f"${qty_col_letter}${start_row}:${qty_col_letter}${end_row}"
        rng_rub  = f"${rub_col_letter}${start_row}:${rub_col_letter}${end_row}"

        cell_art = f"{art_col_letter}{row}"

        # Продажи минус возвраты шт в рамках артикула
        cell_sum_qty = f"{col_map['Продажи минус возвраты шт в рамках артикула']}{row}"
        final_ws[cell_sum_qty].value = (
            f'=IFERROR(SUMIFS({rng_qty},{rng_art},{cell_art}),"")'
        )

        # Продажи минус возвраты руб к перечислению в рамках артикула
        cell_sum_rub = f"{col_map['Продажи минус возвраты руб к перечислению в рамках артикула']}{row}"
        final_ws[cell_sum_rub].value = (
            f'=IFERROR(SUMIFS({rng_rub},{rng_art},{cell_art}),"")'
        )

        # Приход денег на РС в рамках артикула
        cell_sum_inc = f"{col_map['Приход денег на РС в рамках артикула']}{row}"
        final_ws[cell_sum_inc].value = (
            f'=IFERROR(SUMIFS({rng_inc},{rng_art},{cell_art}),"")'
        )

        # 1) Заплатил клиент за 1 ед. артикула — одинаково для всех (без учёта статуса)
        cell_paid_common = f"{col_map['Заплатил клиент за 1 ед. артикула']}{row}"
        final_ws[cell_paid_common].value = (
            f'=IFERROR({cell_sum_rub}/{cell_sum_qty},"")'
        )

        # === Дополнительно для клиента SENS ===
        if client_name == 'SENS':
            needed_sens = [
                'СРЕДНИЙ приход за 1 ед. артикула',
                'Приход за артикул',
                'Средняя маржа в руб за 1 ед',
                'Себестоимость',
            ]
            if all(c in col_map for c in needed_sens):
                cell_avg_inc   = f"{col_map['СРЕДНИЙ приход за 1 ед. артикула']}{row}"
                cell_inc_art   = f"{col_map['Приход за артикул']}{row}"
                cell_avg_marg  = f"{col_map['Средняя маржа в руб за 1 ед']}{row}"
                cell_cost      = f"{col_map['Себестоимость']}{row}"

                # 2) СРЕДНИЙ приход за 1 ед. артикула =
                #    Сумма Приход денег на РС в рамках артикула / Сумма Продажи - Возвраты (шт) в рамках артикула
                final_ws[cell_avg_inc].value = (
                    f'=IFERROR({cell_sum_inc}/{cell_sum_qty},"")'
                )

                # 3) Приход за артикул = Сумма Приход денег на РС в рамках артикула
                final_ws[cell_inc_art].value = (
                    f'=IFERROR({cell_sum_inc},"")'
                )

                # 4) Средняя маржа в руб за 1 ед = СРЕДНИЙ приход за 1 ед – Себестоимость
                final_ws[cell_avg_marg].value = (
                    f'=IFERROR({cell_avg_inc}-{cell_cost},"")'
                )

            # Для SENS дальше формулы по статусу не нужны
            return

    # --- 2. Производные формулы по ТЗ (только для SENS_IP: КЗ / РФ) ---
    if client_name != 'SENS_IP':
        return

    needed = [
        'Статус',
        'Себестоимость',
        'Продажи минус возвраты шт в рамках артикула',
        'Продажи минус возвраты руб к перечислению в рамках артикула',
        'Приход денег на РС в рамках артикула',
    ]
    if not all(col in col_map for col in needed):
        return

    cell_status      = f"{col_map['Статус']}{row}"
    cell_cost        = f"{col_map['Себестоимость']}{row}"
    cell_sum_qty_art = f"{col_map['Продажи минус возвраты шт в рамках артикула']}{row}"
    cell_sum_rub_art = f"{col_map['Продажи минус возвраты руб к перечислению в рамках артикула']}{row}"
    cell_sum_inc_art = f"{col_map['Приход денег на РС в рамках артикула']}{row}"

    # === Блок "кз" для SENS_IP ===
    kz_cols = [
        'ДС в КЗ (ВБ)',
        'ДС в КЗ за 1 ед. артикула (ВБ)',
        'ДС в РФ (ВБ)',
        'ДС в РФ за 1 ед. артикула (ВБ)',
        'ИТОГО ДС В РФ ЗА АРТИКУЛ (ВБ)',
    ]
    if all(c in col_map for c in kz_cols):
        cell_ds_kz        = f"{col_map['ДС в КЗ (ВБ)']}{row}"
        cell_ds_kz_unit   = f"{col_map['ДС в КЗ за 1 ед. артикула (ВБ)']}{row}"
        cell_ds_rf        = f"{col_map['ДС в РФ (ВБ)']}{row}"
        cell_ds_rf_unit   = f"{col_map['ДС в РФ за 1 ед. артикула (ВБ)']}{row}"
        cell_ds_rf_total  = f"{col_map['ИТОГО ДС В РФ ЗА АРТИКУЛ (ВБ)']}{row}"

        if 'Приход денег на РС' in col_map:
            cell_income_rs = f"{col_map['Приход денег на РС']}{row}"
            final_ws[cell_ds_kz].value = (
                f'=IF(ISNUMBER(SEARCH("кз",{cell_status})),'
                f'IFERROR({cell_income_rs}*0.75,""),"")'
            )
            final_ws[cell_ds_rf].value = (
                f'=IF(ISNUMBER(SEARCH("кз",{cell_status})),'
                f'IFERROR({cell_income_rs}*0.25,""),"")'
            )
        else:
            final_ws[cell_ds_kz].value = ""
            final_ws[cell_ds_rf].value = ""

        final_ws[cell_ds_kz_unit].value = (
            f'=IF(ISNUMBER(SEARCH("кз",{cell_status})),'
            f'IFERROR(0.75*{cell_sum_inc_art}/{cell_sum_qty_art},""),"")'
        )
        final_ws[cell_ds_rf_unit].value = (
            f'=IF(ISNUMBER(SEARCH("кз",{cell_status})),'
            f'IFERROR(0.25*{cell_sum_inc_art}/{cell_sum_qty_art},""),"")'
        )
        final_ws[cell_ds_rf_total].value = (
            f'=IF(ISNUMBER(SEARCH("кз",{cell_status})),'
            f'IFERROR(0.25*{cell_sum_inc_art},""),"")'
        )

    # === Блок "рф" для SENS_IP ===
    rf_cols = [
        'СРЕДНИЙ приход за 1 ед. артикула (ВБ)',
        'Приход за артикул (ВБ)',
        'Средняя маржа в руб за 1 ед (ВБ)',
        'Маржа в руб за артикул (ВБ)',
    ]
    if all(c in col_map for c in rf_cols):
        cell_avg_inc    = f"{col_map['СРЕДНИЙ приход за 1 ед. артикула (ВБ)']}{row}"
        cell_inc_art    = f"{col_map['Приход за артикул (ВБ)']}{row}"
        cell_avg_margin = f"{col_map['Средняя маржа в руб за 1 ед (ВБ)']}{row}"
        cell_margin_art = f"{col_map['Маржа в руб за артикул (ВБ)']}{row}"

        final_ws[cell_avg_inc].value = (
            f'=IF(ISNUMBER(SEARCH("рф",{cell_status})),'
            f'IFERROR({cell_sum_inc_art}/{cell_sum_qty_art},""),"")'
        )
        final_ws[cell_inc_art].value = (
            f'=IF(ISNUMBER(SEARCH("рф",{cell_status})),'
            f'IFERROR({cell_sum_inc_art},""),"")'
        )
        final_ws[cell_avg_margin].value = (
            f'=IF(ISNUMBER(SEARCH("рф",{cell_status})),'
            f'IFERROR({cell_avg_inc}-{cell_cost},""),"")'
        )
        final_ws[cell_margin_art].value = (
            f'=IF(ISNUMBER(SEARCH("рф",{cell_status})),'
            f'IFERROR({cell_avg_margin}*{cell_sum_qty_art},""),"")'
        )

# Функция форматирования excel
def format_excel(
    current_report_dir,
    date_report,
    client_name,
    df_final_costs,
    existing_promotion_columns,
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
    # existing_promotion_columns = [col for col in df_final_costs.columns if col in promotion_columns]

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
            f'=IFERROR(ABS({cell_marg_rub}/{cell_sales_for_percent}) * 100 *SIGN({cell_marg_rub}), "")'
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
        # pandas-аналог:
        # sales_rub_mean_by_article = df.groupby('Артикул WB')[sales_for_percent].transform('mean')
        # marg_rub_mean_by_article  = df.groupby('Артикул WB')['Маржинальность руб'].transform('mean')
        # Маржа по артикулу % = marg_rub_mean_by_article / sales_rub_mean_by_article * 100
        if 'Артикул WB' in col_map and 'Маржинальность руб' in col_map and sales_for_percent_col in col_map \
           and 'Маржа по артикулу %' in col_map:
            art_col_letter   = col_map['Артикул WB']
            marg_col_letter  = col_map['Маржинальность руб']
            sales_col_letter = col_map[sales_for_percent_col]

            # Абсолютные диапазоны для AVERAGEIFS
            rng_art   = f"${art_col_letter}${start_row}:${art_col_letter}${end_row}"
            rng_marg  = f"${marg_col_letter}${start_row}:${marg_col_letter}${end_row}"
            rng_sales = f"${sales_col_letter}${start_row}:${sales_col_letter}${end_row}"

            cell_art = f"{art_col_letter}{row}"
            cell_margin_article = f"{col_map['Маржа по артикулу %']}{row}"

            formula_margin_article = (
                f'=IFERROR('
                f'AVERAGEIFS({rng_marg},{rng_art},{cell_art}) / '
                f'AVERAGEIFS({rng_sales},{rng_art},{cell_art}) * 100,'
                f'""'
                f')'
            )
            final_ws[cell_margin_article].value = formula_margin_article

        # --- Формула: "Маржа по артикулу %" через AVERAGEIFS ---
        required_for_margin_article = [
            'Артикул продавца',
            'Маржинальность руб',
            sales_for_percent_col,
            'Маржа по артикулу %',
        ]

        if all(col in col_map for col in required_for_margin_article):
            art_col_letter   = col_map['Артикул продавца']
            marg_col_letter  = col_map['Маржинальность руб']
            sales_col_letter = col_map[sales_for_percent_col]

            rng_art   = f"${art_col_letter}${start_row}:${art_col_letter}${end_row}"
            rng_marg  = f"${marg_col_letter}${start_row}:${marg_col_letter}${end_row}"
            rng_sales = f"${sales_col_letter}${start_row}:${sales_col_letter}${end_row}"

            cell_art = f"{art_col_letter}{row}"
            cell_margin_article = f"{col_map['Маржа по артикулу %']}{row}"

            formula_margin_article = (
                f'=IFERROR('
                f'AVERAGEIFS({rng_marg},{rng_art},{cell_art}) / '
                f'AVERAGEIFS({rng_sales},{rng_art},{cell_art}) * 100,'
                f'""'
                f')'
            )
            final_ws[cell_margin_article].value = formula_margin_article

        # --- Дополнительные формулы для некоторых клиентов --- #
        apply_wb_extra_formulas(final_ws, col_map, row, start_row, end_row, client_name)

    # --- Условное форматирование столбцов с маржинальностью ---

    # Список названий столбцов маржинальности
    margin_columns = [
        "Маржинальность руб",
        "Маржинальность %",
        "Маржинальность на ед. товара",
        "Маржа от себестоимости %",
        'Маржа по артикулу %',
        "Приход денег на РС",
        "Приход денег на РС в рамках артикула",
    ]

    # Цвета
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    green_font = Font(color='006100', bold=True)
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    red_font = Font(color='9C0006', bold=True)

    for col_name in margin_columns:
        if col_name in col_map:
            col_letter = col_map[col_name]
            rng = f"{col_letter}{row_start}:{col_letter}{end_row}"

            # Формула для первой строки диапазона (например, A2)
            first_cell = f"{col_letter}{row_start}"
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
    # Переименовываем столбцы
    for old_col, new_col in rename_map.items():
        # Проверяем наличие столбца в col_map
        if old_col in col_map.keys():
            final_ws[f"{col_map[old_col]}1"].value = f"{old_col} ({new_col})"

    # Закрепляем первую строку
    final_ws.freeze_panes = 'A2'

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
        # Лист с итоговой таблицей пропускаем
        if worksheet_name == "Итоговая таблица":
            continue
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
        date_start,
        date_end,
        # date_start = '2026-09-01',
        # date_end='2026-09-30'
    )
    # Создаем директорию для текущего отчета
    current_report_dir, date_report = create_finance_report_dir(client_name, date_start, date_end)
    # Получение отчета о реализации по АПИ
    df_realization_report = upload_realization_report_new(
        headers,
        date_start,
        date_end,
        period='weekly',
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
        filter_dates=True,
        period='weekly',
    )
    # Список текущих товаров по апи
    df_products = getWBProduct(headers, to_save=False)
    # Добавляем к отчету о реализации размеры из списка товаров
    df_realization_report_with_sizes = add_sizes_to_realization_report(
        df_realization_report_date_filtered,
        df_products
    )
    # df_realization_report_with_sizes = get_realization_report(
    #     headers,
    #     date_start,
    #     date_end,
    #     period='weekly',
    #     df_products=df_products
    # )

    # Получаем список дат отчета
    report_dates = get_report_dates(
        df_realization_report_with_sizes
    )

    # Расчет затрат, которые уже разбиты по артикулам
    result_sku_costs = calc_sku_costs_v4(
        df_realization_report_with_sizes
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
        df_realization_report_with_sizes,
        date_start,
        date_end
    )
    # Объединяем расходы, которые уже разбиты по размерам
    df_sku_sizes_costs_union = union_sku_sizes_costs(
        result_sku_costs,
        result_paid_storage
    )
    # Расходы на платную приемку (отдельный отчет)
    result_paid_acceptance = calc_paid_acceptance_costs(
        headers,
        df_sku_sizes_costs_union,
        df_realization_report_with_sizes,
        date_start,
        date_end
    )
    # Считаем расходы на рекламные кампании
    result_companies_costs = calc_companies_costs_v2(
        df_realization_report_with_sizes,
        df_sku_sizes_costs_union,
        date_report
    )
    # Считаем расходы на списания за отзывы на товары
    result_reviews_costs = calc_reviews_costs(
        df_realization_report_with_sizes,
        df_sku_sizes_costs_union,
    )
    # Объединяем расходы на пл. приемку с расходами по SKU и хранению
    # df_sku_paid_storage_acceptance_stats = union_sku_paid_storage_acceptance_costs(df_sku_and_paid_storage_stats, result_paid_acceptance)
    # Объединяем все расходы из отдельных отчетов в один df
    df_sku_sizes_costs_all = merge_all_costs(
        df_sku_sizes_costs_union,
        result_paid_acceptance,
        result_companies_costs,
        result_reviews_costs,
    )
    # Расчет затрат, которые не разбиты по размерам
    df_sku_sizes_costs_final = calc_sku_no_size_costs_v2(
        df_sku_sizes_costs_all,
        result_sku_costs,
    )
    result_other_costs = calc_other_costs_v3(
        result_sku_costs,
        df_sku_sizes_costs_final
    )
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
    df_final_costs, existing_promotion_columns, columns_for_final_costs = calc_final_costs(
        df_all_costs_stats_with_catalog
    )
    # Считаем сводную по затратам
    df_final_costs_by_operation = create_svod_by_operations(df_final_costs)
    # Сохранение результатов в excel
    filepath_finance_report = save_excel(
        current_report_dir,
        date_report,
        # df_realization_report_new_columns,
        df_realization_report_with_sizes,
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
        existing_promotion_columns,
        columns_for_final_costs,
    )
    print(f"\033[32m\033[47mDone calculating Finance Report for Client {client_name}")
    logger.info(f"Finance Report saved in {filepath_finance_report}")



# %%
