
# %% Определение всех функций
import argparse
import json
import os
import warnings
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import requests
# # from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

warnings.simplefilter(action='ignore', category=FutureWarning)
# pd.set_option('future.no_silent_downcasting', True)

# Доп. функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from ozon.scripts.command_args_parser import cli_dates
# Файл с некоторыми константами
from ozon.scripts.constants import (client_name, headers, marketplace_dir_path,
                                    ozon_seller_api_url)
# Функции выгрузки данных по АПИ, которые уже готовы
from ozon.scripts.uploadDataFromOzon import getOrders, get_transactions

parser = argparse.ArgumentParser(
    description="Создание сводного отчета по продажам и заказам на Ozon.",
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

# Создание нужных директорий
def create_dirs():
    # Папка с клиентом
    client_dir = f"{marketplace_dir_path}/Clients/{client_name}/"
    # Список директорий для создания
    dir_names = ['SalesOrdersSvod']
    for dir_name in dir_names:
        dir_path = os.path.join(client_dir, dir_name)
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)


# Функция создания диапазона дат
# GPT START ----
def generate_date_range(
        reference_date: str = None,
        start_date: str = None,
        end_date: str = None
    ):
    # Проверка на несовместимость параметров:
    # либо используем reference_date, либо вручную заданный диапазон
    if reference_date and (start_date or end_date):
        raise ValueError("Нельзя одновременно задавать reference_date и start_date/end_date.")

    # Если оба manual-диапазона заданы
    if start_date and end_date:
        # Парсим строки дат в объекты datetime
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    elif start_date or end_date:
        # Один из диапазонов не указан — ошибка
        raise ValueError("Нужно задать одновременно start_date и end_date.")
    else:
        # Ни диапазон, ни reference_date не заданы явно
        #  — значит работаем с reference_date (по умолчанию — сегодня)
        if reference_date is None:
            ref_dt = datetime.now()
        else:
            # Парсим reference_date из строки
            ref_dt = datetime.strptime(reference_date, "%Y-%m-%d")

        # Если дата — первое число месяца
        if ref_dt.day == 1:
            # Берём последний день прошлого месяца
            prev_month_last_day = ref_dt - timedelta(days=1)
            # Начало диапазона — первое число прошлого месяца
            start_dt = prev_month_last_day.replace(day=1)
            # Конец диапазона — последний день прошлого месяца
            end_dt = prev_month_last_day
        else:
            # Если не первое число месяца:
            # Начало диапазона — первое число текущего месяца
            start_dt = ref_dt.replace(day=1)
            # Конец диапазона — предыдущие сутки относительно reference_date
            end_dt = ref_dt - timedelta(days=1)

    # Форматы для строк и для ISO-даты по времени суток
    start_date_str = start_dt.strftime('%Y-%m-%d')
    end_date_str = end_dt.strftime('%Y-%m-%d')
    start_date_iso = start_dt.strftime('%Y-%m-%dT00:00:00Z')
    end_date_iso = end_dt.strftime('%Y-%m-%dT23:59:59Z')

    # Собираем результат в DataFrame
    df_dates = pd.DataFrame([{
        'date_start': start_date_str,       # начало диапазона (дд.мм.гггг)
        'date_end': end_date_str,           # конец диапазона (дд.мм.гггг)
        'datetime_start': start_date_iso,   # начало ISO-датой с 00:00:00
        'datetime_end': end_date_iso        # конец ISO-датой с 23:59:59
    }])

    return start_date_iso, end_date_iso
# GPT END ----


# Функция создания датафрейма с диапазоном дат
def generate_date_range_df(date_start, date_end):
    # Генерируем диапазон дат от начальной до конечной даты с интервалом в один день
    date_range_df = pd.DataFrame({'date': pd.date_range(date_start, date_end, freq='d')})
    # Переводим в формат даты, чтобы убрать время
    date_range_df['date'] = date_range_df['date'].dt.date

    return date_range_df

# Функция выгрузки списка транзакций
def getTransactionReport(headers, date_start, date_end):
    """
    Выгружает транзакции по API Ozon с разбивкой по месяцам и логированием.
    """
    dt_start = pd.to_datetime(date_start)
    dt_end = pd.to_datetime(date_end)
    period_days = (dt_end - dt_start).days

    def generate_monthly_intervals(date_start: str, date_end: str):
        """Генерирует месячные интервалы, если период больше месяца"""
        dt_start = pd.to_datetime(date_start)
        dt_end = pd.to_datetime(date_end)
        intervals = []

        current_start = dt_start
        while current_start <= dt_end:
            next_month = (current_start + pd.DateOffset(months=1)).replace(day=1)
            current_end = min(next_month - pd.Timedelta(seconds=1), dt_end)
            intervals.append({
                'date_start': current_start.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
                'date_end': current_end.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
                'dt_start': current_start,
                'dt_end': current_end
            })
            current_start = next_month

        return pd.DataFrame(intervals)

    if period_days <= 29:
        # Один интервал на весь указанный диапазон
        date_range_df = pd.DataFrame([{
            'date_start': dt_start.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'date_end': dt_end.strftime('%Y-%m-%dT%H:%M:%S.000Z'),
            'dt_start': dt_start,
            'dt_end': dt_end
        }])
    else:
        # Дробим на месяцы
        date_range_df = generate_monthly_intervals(date_start, date_end)

    logger.info(f"Выгрузка транзакций с {date_start} по {date_end} начата.")

    df_transaction_list = pd.DataFrame()

    for idx, interval in date_range_df.iterrows():
        logger.info(f"Обработка периода: {interval['dt_start']} - {interval['dt_end']}")
        page = 1
        total_pages = 1  # Пажинация: по умолчанию хотя бы одну страницу пробуем
        while page <= total_pages:
            payload = {
                "filter": {
                    "date": {
                        "from": interval['date_start'],
                        "to": interval['date_end']
                    },
                    "operation_type": [],
                    "posting_number": "",
                    "transaction_type": "all"
                },
                "page": page,
                "page_size": 1000
            }
            try:
                resp = requests.post(
                    f"{ozon_seller_api_url}/v3/finance/transaction/list",
                    headers=headers,
                    data=json.dumps(payload),
                    timeout=60
                )
                resp.raise_for_status()
                data = resp.json()
                # Получение страниц
                result = data.get('result', {})
                total_pages = result.get('page_count', 0)
                operations = result.get('operations', [])
                if not operations:
                    logger.info(f"Нет транзакций за этот период, страница {page}.")
                tmp_df = pd.DataFrame(operations)
                if not tmp_df.empty:
                    tmp_df = tmp_df.assign(
                        dt_start=interval['dt_start'],
                        dt_end=interval['dt_end']
                    )
                    df_transaction_list = pd.concat([df_transaction_list, tmp_df], ignore_index=True)
                logger.info(f"Страница {page} из {total_pages} обработана.")
            except requests.RequestException as ex:
                logger.error(f"Ошибка запроса API на странице {page}: {ex}")
            except Exception as ex:
                logger.error(f"Общая ошибка на странице {page}: {ex}")
            page += 1

    if not df_transaction_list.empty:
        df_transaction_list['items_amount'] = df_transaction_list['items'].apply(lambda x: len(x) if isinstance(x, list) else 0)
        df_transaction_list['services_amount'] = df_transaction_list['services'].apply(lambda x: len(x) if isinstance(x, list) else 0)
        logger.info(f"Выгрузка транзакций завершена. Найдено {df_transaction_list.shape[0]} транзакций.")
    else:
        logger.info("Транзакции не найдены по указанному периоду.")

    return df_transaction_list

# Функция выгрузки отчета о заказах
def get_orders(headers, date_start, date_end):
    # Выгружаем отчет об отправлениях fbo и fbs
    df_orders_fbo = getOrders(headers, date_start, date_end, delivery_schema='fbo', to_save=False)
    df_orders_fbs = getOrders(headers, date_start, date_end, delivery_schema='fbs', to_save=False)
    # Убираем лишние колонки из отчета fbo
    df_orders_fbo = df_orders_fbo.loc[:, ~df_orders_fbo.columns.isin(['Объемный вес товаров, кг'])]
    # Объединяем два отчета в один
    df_orders = pd.concat([df_orders_fbo, df_orders_fbs])

    return df_orders

# Функция объединения заказов fbo и fbs
def union_fbo_and_fbs_orders(df_postings_fbo, df_postings_fbs):
    df_postings_all = pd.concat([df_postings_fbo, df_postings_fbs])
    return df_postings_all

# Функция расчета продаж
def calc_sales(df_transactions: pd.DataFrame, date_range_df: pd.DataFrame) -> pd.DataFrame:
    """
    Рассчитывает продажи по дням на основе нового списка транзакций.

    На вход подается df_transactions в новом формате:
      - 'Дата'         — дата операции (или колонка, которую можно привести к дате)
      - 'Тип начисления' ∈ {
            'Продажи руб', 'Продажи шт',
            'Программы партнеров', 'Баллы за скидки',
            'Возвраты руб', 'Возвраты шт',
            ...
        }
      - 'Итого'        — сумма/количество по транзакции

    Условия:
      - Возвраты в рублях уже приходят ОТРИЦАТЕЛЬНЫМИ в 'Итого'.
      - Возвраты в штуках приходят ПОЛОЖИТЕЛЬНЫМИ в 'Итого'.

    В результате считаем ЧИСТЫЕ продажи:
      - 'Продажи руб' = (Продажи руб + Программы партнёров + Баллы за скидки + Возвраты руб)
        (возвраты руб уже отрицательные, просто суммируем).
      - 'Продажи шт' = Продажи шт - Возвраты шт
        (возвраты шт положительные, вычитаем).
    """

    date_column = "Дата"
    type_column = "Тип начисления"
    value_column = "Итого"

    if df_transactions is None or df_transactions.empty:
        logger.info("calc_sales: df_transactions пуст, возвращаем нули по date_range_df.")
        sales = date_range_df.copy()
        sales["Продажи руб"] = 0
        sales["Продажи шт"] = 0
        return sales

    sales_df = df_transactions.copy()

    # Переводим колонку с датой в timestamp (если ещё не datetime)
    sales_df[date_column] = pd.to_datetime(sales_df[date_column], errors="coerce")
    # Оставляем только дату
    sales_df[date_column] = sales_df[date_column].dt.date

    # Типы начислений, которые учитываем
    money_types_positive = ["Продажи руб", "Программы партнеров", "Баллы за скидки"]
    money_types_negative = ["Возвраты руб"]     # уже отрицательные в Итого
    qty_types_positive = ["Продажи шт"]
    qty_types_negative = ["Возвраты шт"]        # приходят положительными, надо вычесть

    filtered = sales_df.loc[
        sales_df[type_column].isin(
            money_types_positive + money_types_negative +
            qty_types_positive + qty_types_negative
        ),
        [date_column, type_column, value_column],
    ].copy()

    if filtered.empty:
        logger.info("calc_sales: после фильтрации по типам продаж данных нет, возвращаем нули.")
        sales = date_range_df.copy()
        sales["Продажи руб"] = 0
        sales["Продажи шт"] = 0
        return sales

    # Разносим по отдельным колонкам
    # Денежная часть: продажи и «доп. рублёвые» (баллы, программы) идут в плюс,
    # возвраты руб уже отрицательные в 'Итого', просто добавляем их как есть.
    filtered["Продажи_руб_плюс"] = np.where(
        filtered[type_column].isin(money_types_positive),
        filtered[value_column],
        0,
    )
    filtered["Возвраты_руб_сырые"] = np.where(
        filtered[type_column].isin(money_types_negative),
        filtered[value_column],
        0,
    )

    # Штучная часть: продажи шт → плюс, возвраты шт → минус (т.к. приходят положительными)
    filtered["Продажи_шт_плюс"] = np.where(
        filtered[type_column].isin(qty_types_positive),
        filtered[value_column],
        0,
    )
    filtered["Возвраты_шт_плюс"] = np.where(
        filtered[type_column].isin(qty_types_negative),
        filtered[value_column],
        0,
    )

    # Группируем по дате и считаем суммы
    agg = (
        filtered
        .groupby(date_column, as_index=False)
        .agg(
            Продажи_руб_плюс=("Продажи_руб_плюс", "sum"),
            Возвраты_руб_сырые=("Возвраты_руб_сырые", "sum"),
            Продажи_шт_плюс=("Продажи_шт_плюс", "sum"),
            Возвраты_шт_плюс=("Возвраты_шт_плюс", "sum"),
        )
    )

    # Чистые продажи:
    # - рубли: продажи плюс + возвраты (возвраты уже отрицательные) → просто сумма
    agg["Продажи руб"] = agg["Продажи_руб_плюс"] + agg["Возвраты_руб_сырые"]

    # - штуки: возвраты шт приходят положительными, надо вычесть
    agg["Продажи шт"] = agg["Продажи_шт_плюс"] - agg["Возвраты_шт_плюс"]

    sales = agg[[date_column, "Продажи руб", "Продажи шт"]].rename(
        columns={date_column: "date"}
    )

    # Мерджим с диапазоном дат, т.к. бывает, что не все дни присутствуют в заказах
    sales = sales.merge(
        date_range_df,
        on="date",
        how="outer"
    )

    # Заполняем пропуски после мерджа
    for col in ["Продажи руб", "Продажи шт"]:
        if col not in sales.columns:
            sales[col] = 0
        else:
            sales[col] = sales[col].fillna(0)

    return sales

# Функция расчета заказов
def calc_orders(df_orders_all, date_range_df):
    # Создаем копию для избежания изменений в оригинальном df
    df_orders_all_ = df_orders_all.copy()
    # Колонка, по котрой берем дату для расчета продаж
    date_column = 'Принят в обработку'
    # Колонка, по которой считаем итоговую сумму
    total_column = 'Сумма отправления'
    # Переводим колонку с датой в timestamp
    df_orders_all_[date_column] = pd.to_datetime(df_orders_all_[date_column])
    # Оставляем только дату
    df_orders_all_[date_column] = df_orders_all_[date_column].dt.date
    # Переименовываем колонку для мерджа
    df_orders_all_ = df_orders_all_.rename(columns={
        date_column: 'date',
        total_column: 'Заказы руб'
        })
    # Считаем заказы в штуках и рублях
    orders = (df_orders_all_
              .groupby(['date'])
              .agg(**{
                  'Заказы руб': ('Заказы руб', 'sum'),
                  'Заказы шт': ('Количество', 'sum')
              })
              .reset_index()
    )

    # Мерджим с диапазоном дат, т.к. бывает, что не все дни присутствуют в заказах
    orders = orders.merge(
        date_range_df,
        on='date',
        how='outer'
    )
    # Заполняем пропуски после мерджа
    for col in ['Заказы руб', 'Заказы шт']:
        orders[col] = orders[col].fillna(0)

    return orders

# Объединяем заказы и продажи в одну таблицу
def union_orders_and_sales(orders, sales):
    # Объединяем в один df
    df_orders_and_sales = pd.concat([
        orders.set_index(orders['date']),
        sales.set_index(sales['date'])
    ], axis=1)
    # Удаляем ненужную колонку
    df_orders_and_sales = df_orders_and_sales.drop(columns=['date'])
    # Достаем дату из индекса
    df_orders_and_sales = df_orders_and_sales.reset_index()
    # Форматируем дату в нужный формат
    df_orders_and_sales['Дата'] = df_orders_and_sales['date'].apply(lambda x: x.strftime('%d.%m.%Y'))
    # Разворачиваем в широкий вид
    df_orders_and_sales_wide = pd.pivot_table(
        df_orders_and_sales,
        values=['Заказы руб', 'Заказы шт', 'Продажи руб', 'Продажи шт'],
        # index=['Заказы руб', 'Заказы шт', 'Продажи руб', 'Продажи шт'],
        columns=['date'],
        aggfunc='sum'
    )
    # Делаем нужный формат даты у колонок
    df_orders_and_sales_wide.columns = pd.to_datetime(df_orders_and_sales_wide.columns).strftime('%d.%m.%Y')

    return df_orders_and_sales_wide


# Функция добавления плана и факта продаж
def add_plan(df_orders_and_sales_wide,
             client_name,
             plan_orders_rub = 0,
             plan_orders_amount = 0,
             plan_sales_rub = 0,
             plan_sales_amount = 0):
    # Создаем копию для избежания изменений в оригинальном df
    df_plan = df_orders_and_sales_wide.copy()
    # Считаем итоговые суммы
    df_plan['Факт'] = df_plan.sum(axis=1)
    # Добавляем План
    df_plan['План'] = np.nan
    df_plan.loc['Заказы руб', 'План'] = plan_orders_rub
    df_plan.loc['Заказы шт', 'План'] = plan_orders_amount
    df_plan.loc['Продажи руб', 'План'] = plan_sales_rub
    df_plan.loc['Продажи шт', 'План'] = plan_sales_amount
    # Считаем отклонение плана от факта
    df_plan['Факт/План, %'] = df_plan['Факт'].divide(df_plan['План'].replace({0: np.nan})) * 100
    # Заменяем inf, если получилось деление на ноль
    df_plan['Факт/План, %'] = df_plan['Факт/План, %'].replace([-np.inf, np.inf], np.nan)
    # Добавляем имя магазина
    # df_plan['Магазин'] = client_name
    # # Добавление пустой строки в начало DataFrame
    # empty_row = pd.DataFrame([[np.nan] * df_plan.shape[1]], columns=df_plan.columns, index=['Магазин'])
    # df = pd.concat([empty_row, df_plan])
    # # Перемещаем колонки в начало df
    df_plan = move_columns(df_plan, ['План', 'Факт', 'Факт/План, %'], 0)
    # Убираем имя у индекса
    df_plan.index.name = None

    return df_plan


# Функция сохранения отчета в excel
def save_excel(df_plan, date_start, date_end, date_report=str(date.today())):
    # Создаем df с именем клиента
    shop_name = pd.DataFrame({'Магазин': [client_name]})
    # Создаем df с диапазоном формирования отчета
    df_dates = pd.DataFrame({
        'Начальная дата': [date_start],
        'Конечная дата': [date_end]
    })
    # Даты для имени файла
    date_start_file = pd.to_datetime(date_start).strftime('%d-%m-%Y')
    date_end_file = pd.to_datetime(date_end).strftime('%d-%m-%Y')

    # Имя файла для сохранения отчета
    report_file_name = (
        marketplace_dir_path
        / "Clients"
        / client_name
        / "SalesOrdersSvod"
        / f"{date_start_file}_{date_end_file}_Свод_заказы_продажи_Ozon_{client_name}.xlsx"
    )

    # GPT START ----
    # Параметры для записи датафреймов
    dataframes = [
        (shop_name, 2, 2, False),  # Начало в строке 1, колонке 1, не записывать индекс датафрейма
        (df_plan, 6, 2, True),
        (df_dates, 12, 2, False),
    ]

    # Запись датафреймов на один лист Excel
    with pd.ExcelWriter(report_file_name, engine='openpyxl') as writer:
        ranges = []  # Здесь будем сохранять диапазоны (начало и конец) для границ
        for df, start_row, start_col, save_index in dataframes:
            # Если сохраняется индекс, добавляется 1 столбец
            end_row = start_row + len(df)  # Вычисляем конечную строку
            end_col = start_col + len(df.columns) - 1 + (1 if save_index else 0)  # Конечный столбец с учётом индекса
            ranges.append((start_row, end_row, start_col, end_col, save_index))  # Добавляем "save_index"

            # Запись датафрейма с учётом параметра индекса
            df.to_excel(writer, index=save_index, sheet_name=client_name, startrow=start_row - 1, startcol=start_col - 1)


    # Форматирование Excel
    logger.info(f"Formatting Sales svod for client {client_name}")

    wb = load_workbook(report_file_name)
    ws = wb[client_name]

    # Определяем стиль границ
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Применяем границы только к диапазону данных таблиц (без индекса)
    for start_row, end_row, start_col, end_col, save_index in ranges:
        for row in ws.iter_rows(
            min_row=start_row,
            max_row=end_row,
            min_col=start_col + (1 if save_index else 0),  # Сдвигаем начало диапазона вправо, если индекс записан
            max_col=end_col
        ):
            for cell in row:
                cell.border = thin_border
                # Выравнивание текста по центру
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Применение числового форматирования с разделением разрядов (через пробелы)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):  # Проверяем, является ли значение числом
                cell.number_format = '#,##0'  # Формат с пробелами и двумя знаками после запятой


    # Автоподбор ширины столбцов
    for col in ws.columns:
        max_length = 0
        column = col[0].column  # Номер текущего столбца
        column_letter = get_column_letter(column)  # Буква столбца
        for cell in col:
            try:
                if cell.value:  # Проверка на наличие данных
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = max_length + 2

    # Сохранение обновленного файла
    wb.save(report_file_name)

    return report_file_name


# %% Вызов всех функций
if __name__ == '__main__':
    # Парсим аргументы командной строки
    dates_cli = cli_dates(parser)

    # Формируем диапазон дат
    date_start, date_end = generate_date_range(
        start_date=dates_cli['start_date'],
        end_date=dates_cli['end_date'],
        # start_date='2025-08-27',
        # end_date='2025-09-02'
    )
    # Планы продаж и заказов в штуках и рублях
    plan_orders_rub = 0
    plan_orders_amount = 0
    plan_sales_rub = 0
    plan_sales_amount = 0

    # Генерируем диапазон дат от начальной до конечной даты с интервалом в один день
    date_range_df = generate_date_range_df(date_start, date_end)

    logger.info(
        f"Calculating Sales and Orders svod for client {client_name} "
        f"for dates {date_start} - {date_end}"
    )
    # Выгружаем заказы FBO и FBS
    df_orders_all = get_orders(headers, date_start, date_end)
    # Выгружаем список транзакций
    df_transactions_list = get_transactions(headers, date_start, date_end, to_save=False)

    # Считаем заказы в штуках и рублях
    orders = calc_orders(df_orders_all, date_range_df)
    # Считаем продажи в штуках и рублях
    sales = calc_sales(df_transactions_list, date_range_df)
    # Объединяем продажи и заказы в один df
    df_orders_and_sales_wide = union_orders_and_sales(orders, sales)
    # Добавляем планы продаж и считаем отклонение плана от факта
    df_plan = add_plan(
        df_orders_and_sales_wide,
        client_name,
        plan_orders_rub,
        plan_orders_amount,
        plan_sales_rub,
        plan_sales_amount
    )
    # Создаем нужную директорию
    create_dirs()
    # Сохраняем Excel
    report_file_name = save_excel(df_plan, date_start, date_end)
    # Форматируем Excel
    # format_excel()
    logger.info(f"Finished calculating Sales svod for client {client_name}")
    logger.info(f"Report saved in: {report_file_name}")

# %%
