
# %% Определение всех функций
import argparse
import json
import re
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests
from loguru import logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

# Файл с настройками и номером клиента
from ozon.scripts.constants import client_name, headers, marketplace_dir_name
# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_dates

parser = argparse.ArgumentParser(
    description="Выгрузка продаж по отчету реализации Ozon",
    usage=(
        "python.exe -m ozon.scripts.create_sales_and_orders_svod [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD --reference_date YYYY-MM-DD] "
        "[-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "    python.exe -m ozon.scripts.create_sales_and_orders_svod 2025-08-27 2025-09-02\n\n"
        "    python.exe -m ozon.scripts.create_sales_and_orders_svod --start_date 2025-08-27 --end_date 2025-09-02 --client 2\n\n"
        "    python.exe -m ozon.scripts.create_sales_and_orders_svod --reference_date 2025-09-01 -c 3"
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

    # Генерируем диапазондат
    df_date_range = pd.DataFrame({
        'date': pd.date_range(start=date_start, end=date_end, freq='MS')
    })
    # Получаем год и месяц
    df_date_range = df_date_range.assign(
        year=df_date_range['date'].dt.year,
        month=df_date_range['date'].dt.month
    )

    return df_date_range




# Функция выгрузки отчета по реализации товаров
def get_products_realization_report(headers, year, month, to_save=False):
    params_finance_report = json.dumps({
    "year": year,
    "month": month
    })
    resp_data_finance_report = requests.post("https://api-seller.ozon.ru/v2/finance/realization", headers=headers, data=params_finance_report).json()
    # result_finance_report_title = pd.DataFrame(resp_data_finance_report['result']['header'], index=[0])
    result_finance_report = pd.DataFrame(resp_data_finance_report['result']['rows'])
    # Получаем данные по товару
    df_products_realization = (
        pd
        .json_normalize(result_finance_report['item'])
        .rename(columns={
            'name': 'Наименование товара',
            'offer_id': 'Артикул',
            'barcode': 'Штрихкод',
            'sku': 'SKU'
        })
    )
    # Словарики для переименовывания
    sales_rename_columns = {
        'amount': 'Сумма',
        'bonus': 'Баллы за скидки',
        'commission': 'Комиссия',
        'compensation': 'Доплата за счёт Ozon',
        'price_per_instance': 'Цена за экземпляр',
        'quantity': 'Количество товара',
        'standard_fee': 'Базовое вознаграждение Ozon',
        'bank_coinvestment': 'Механики лояльности партнеров: зелёные цены',
        'stars': 'Механики лояльности партнеров: звезды',
        'pick_up_point_coinvestment': 'Механики лояльности партнеров: АПВЗ',
        'total': 'Итого к начислению',
    }
    return_rename_columns = sales_rename_columns.copy()
    # Получаем данные по продажам
    df_sales_realization = (
        pd
        .json_normalize(result_finance_report['delivery_commission'])
        .rename(columns=sales_rename_columns)
    )
    # Добавляем префикс к колонкам продаж
    df_sales_realization.columns = [f"{col} (Продажи)" for col in df_sales_realization.columns]

    # Получаем данные по возвратам
    df_returns_realization = (
        pd
        .json_normalize(result_finance_report['return_commission'])
        .rename(columns=return_rename_columns)
    )
    # Добавляем префикс к колонкам возвратов
    df_returns_realization.columns = [f"{col} (Возвраты)" for col in df_returns_realization.columns]

    # Объединяем отчет в один df
    df_realization_report = pd.concat(
        [df_products_realization, df_sales_realization, df_returns_realization],
        axis=1
    )
    # Добавляем год и месяц
    df_realization_report['Год'] = year
    df_realization_report['Месяц'] = month

    return df_realization_report


# Функция выгрузки нескольких отчетов реализации в один
def get_all_realization_report(df_date_range):

    # df, куда будем помещать результат
    df_all_reports = pd.DataFrame()

    # Цикл по каждому году и месяцу
    for idx, row in df_date_range.iterrows():
        year = int(row['year'])
        month = int(row['month'])

        # Выгружаем отчет реализации
        logger.info(f"Выгрузка отчета реализации. Год: {year} Месяц: {month}")
        time.sleep(1)

        df_realization_report = get_products_realization_report(headers, year=year, month=month, to_save=False)
        # Объединяем с предыдущим проходом цикла
        df_all_reports = pd.concat(
            [df_all_reports, df_realization_report],
            ignore_index=True
        )

    return df_all_reports


# Функция вывода отчетов реализации в консоль
def print_realization_reports(df_all_reports):
    # df_to_print = df_all_reports.to_string(index=False)
    # print(df_to_print)

    df_all_reports.to_clipboard(excel=True, index=False)




# %% Вызов всех функций
if __name__ == '__main__':
    # Парсим аргументы командной строки
    dates_cli = cli_dates(parser)

    # Формируем диапазон дат
    df_date_range = generate_date_range(
        date_start=dates_cli['start_date'],
        date_end=dates_cli['end_date'],
        # date_start='2025-10-01',
        # date_end='2025-10-31'
    )

    # Выгружаем отчеты реализации
    df_all_reports = get_all_realization_report(df_date_range)

    # Выводим в консоль
    print_realization_reports(df_all_reports)

    logger.info(
        f"Отчет реализации Озон по клиенту {client_name} в буфере обмена. "
        "Скопируйте его в Excel"
    )

# %%
