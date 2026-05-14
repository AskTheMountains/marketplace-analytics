# %% Определение функций
import argparse
import copy
from math import prod
import time
import re
from datetime import date, datetime, timedelta
from collections import defaultdict, deque
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.options.mode.chained_assignment = None  # default='warn'

# Некоторые вспомогательные функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from wb.scripts.command_args_parser import cli_companies
# Файл с настройками и номером клиента
from wb.scripts.constants import (
    client_name, headers, article_replacement, PROMOTION_COMPANIES, COMPANY_STATS_COL_FORMAT
)
# Функции выгрузки данных
from wb.scripts.uploadDataFromWB import getOrdersWB, getWBProduct
# Доп. функции
from wb.scripts.wb_generic_functions import (
    WbRateLimiter,
    get_sales_funnel
)

parser = argparse.ArgumentParser(
    description="Create promotion report (WB)",
    usage=(
        "python.exe -m wb.scripts.create_campaigns_report [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD --reference_date YYYY-MM-DD] "
        "[-c N | --client N] "
        "[--promotion_companies \"ID1, ID2, ...\"]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report 2024-11-13 2024-11-20\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --start_date 2024-11-13 --end_date 2024-11-20 --client 2\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --reference_date 2024-11-21 -c 3\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --promotion_companies \"13687629, 14962180, 16997671\"\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --promotion_companies \"13687629 14962180 16997671\"\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --promotion_companies \"13687629;14962180;16997671\"\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --promotion_companies \"13687629\\n14962180\\n16997671\"\n\n"
        "Можно комбинировать любые разделители — запятая, пробел, точка с запятой, перевод строки.\n"
    ),
    formatter_class=ParagraphRichHelpFormatter
)

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent


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


# Создание директории для выгрузки статистики РК за конкретные даты
def create_report_path(date_campaign_report, client_name):
    # Задаем путь к директории
    # upload_path = (
    #     f"Clients/{client_name}/CampaignsReport/{str(date.today())}/"
    #     f"{dt_start_dir}-{dt_end_dir}_API/"
    # )
    report_path = (
        BASE_DIR
        / "Clients"
        / client_name
        / "CampaignsReport"
        / date_campaign_report
    )
    if not report_path.exists():
        report_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"Creating Companies Upload directory:{report_path}")

    return report_path


# GPT START ----
def generate_dates(
        reference_date: str = None,
        start_date: str = None,
        end_date: str = None
    ) -> pd.DataFrame:
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
        if today.weekday() == 0:
            # Если сегодня понедельник (weekday() == 0)
            # Понедельник предыдущей недели:
            start_dt = today - timedelta(days=7)
            start_dt = start_dt - timedelta(days=start_dt.weekday())
            # Воскресенье предыдущей недели:
            end_dt = start_dt + timedelta(days=6)
        else:
            # Если не понедельник:
            # Понедельник текущей недели:
            start_dt = today - timedelta(days=today.weekday())
            # Вчерашний день:
            end_dt = today - timedelta(days=1)

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
def get_start_end_date(df_date_range, type_dates='companies'):
    # Если задан тип дат - даты для заказов, то берем даты со временм
    if type_dates == 'orders':
        date_start, date_end = df_date_range.loc[0, ['datetime_start', 'datetime_end']]
    # В остальных случаях берем даты без времени
    else:
        date_start, date_end = df_date_range.loc[0, ['date_start', 'date_end']]

    return date_start, date_end

# Функция создания списка РК с датами
def create_input_companies(
        client_name,
        df_date_range,
        cli_companies_parsed=[],
    ):
    # Если передан список из CLI, используем его
    if cli_companies_parsed:
        input_companies = [[int(company)] for company in cli_companies_parsed]
    else:
        # Берём по client_name из словаря
        companies = PROMOTION_COMPANIES.get(client_name, [])
        if not companies:
            raise ValueError(f"Список кампаний пуст для клиента {client_name}!")
        input_companies = [[int(company)] for company in companies]

    # Получаем даты для отчёта РК
    date_start_performance, date_end_performance = get_start_end_date(df_date_range, type_dates='companies')
    # Добавляем даты к каждой кампании
    for company in input_companies:
        company.extend([date_start_performance, date_end_performance])

    return input_companies


# Функция создания датафрейма со списком кампаний и датами их выгрузки
def create_companies_upload_df(
        input_companies,
):
    # Создаем df из словаря
    df_input_companies = (
        pd.DataFrame(
            input_companies,
            columns=['ID кампании', 'date_start', 'date_end']
        )
        # .reset_index()
        # .rename(columns={
        #     'index': 'ID кампании'
        # })
    )
    # Удаляем дубликаты
    df_input_companies = (
        df_input_companies
        .loc[~df_input_companies['ID кампании'].isin(['']), :]
        .drop_duplicates(subset=['ID кампании', 'date_start', 'date_end'])
    )
    # Присваиваем каждой кампании уникальный ID по номеру кампании и дате
    df_input_companies['unique_company_id'] = (
        df_input_companies
        .groupby(['ID кампании', 'date_start', 'date_end'], as_index=False)
        .ngroup() + 1
    )
    # Сортировка
    df_input_companies = df_input_companies.sort_values(by=['unique_company_id'], ignore_index=True)
    # Перемещаем колонку в начало df
    df_input_companies = move_columns(
        df_input_companies,
        columns_to_move=['unique_company_id'],
        position=0
    )

    return df_input_companies

# Функция обработки списка товаров
def process_product_list(df_products):
    # Создаем копию для избежания изменений в оригинальном df
    df_products_processed = df_products.copy()
    # Переименовываем колонки
    df_products_processed = df_products_processed.rename(columns={
        'nmID': 'Артикул WB',
    })

    return df_products_processed


# Функция получения списка кампаний
def get_company_list(
     headers
):
    # Делаем запрос для получения списка кампаний
    resp_data_company_list = requests.get("https://advert-api.wildberries.ru/adv/v1/promotion/count", headers=headers).json()
    # Переводим в df
    df_company_list_stats = pd.DataFrame(resp_data_company_list['adverts'])
    # Распаковываем список кампаний
    df_company_list_unpacked = df_company_list_stats.explode(column='advert_list').reset_index()
    # Получаем id кампаний
    df_company_ids = pd.json_normalize(df_company_list_unpacked['advert_list'])
    # Объединяем с исходным df
    df_company_list = pd.concat([df_company_ids, df_company_list_unpacked], axis=1)
    # Переводим колонку с датой в tiemstamp
    df_company_list['changeTime'] = pd.to_datetime(df_company_list['changeTime'], format='mixed').dt.tz_localize(None)

    return df_company_list

# Функция получения "сырых" данных по доп. информации по кампаниям из WB API
def fetch_companies_info_raw(headers, df_company_list_filtered, statuses='-1,4,7,8,9,11', payment_type=None):
    """
    Выполняет запросы к API WB /api/advert/v2/adverts для получения данных по кампаниям.

    :param headers: dict с обязательным Authorization
    :param df_company_list_filtered: pd.DataFrame с колонкой 'ID кампании', содержащий кампании для запроса
    :param statuses: строка со статусами кампаний для фильтра (по умолчанию все актуальные)
    :param payment_type: опционально, 'cpm' или 'cpc'
    :return: список словарей с данными по кампаниям (синоним json ответа 'adverts')
    """
    if df_company_list_filtered.empty:
        logger.warning("Нет кампаний для запроса.")
        return []

    # Разбиваем на чанки по 50 кампаний
    step = 50
    df = df_company_list_filtered.reset_index(drop=True)
    df['chunk'] = df.index // step

    REQUEST_LIMIT = 5       # Максимум запросов в секунду
    REQUEST_INTERVAL = 0.2  # Пауза 200 мс
    TIME_WINDOW = 1.0       # Временное окно для лимита в секундах
    request_times = deque()

    all_adverts_info = []
    session = requests.Session()
    base_url = "https://advert-api.wildberries.ru/api/advert/v2/adverts"

    for chunk_id in df['chunk'].unique():
        now = time.time()
        while request_times and now - request_times[0] > TIME_WINDOW:
            request_times.popleft()
        if len(request_times) >= REQUEST_LIMIT:
            time_to_wait = TIME_WINDOW - (now - request_times[0])
            logger.info(f"Достигнут лимит запросов API, ждем {round(time_to_wait,3)} сек")
            time.sleep(max(time_to_wait, REQUEST_INTERVAL))
        else:
            time.sleep(REQUEST_INTERVAL)
        request_times.append(time.time())

        df_chunk = df[df['chunk'] == chunk_id]
        ids_str = ",".join(str(x) for x in df_chunk['ID кампании'])

        params = {
            "ids": ids_str,
            "statuses": statuses
        }
        if payment_type:
            params["payment_type"] = payment_type

        try:
            resp = session.get(base_url, headers=headers, params=params, timeout=30)
            resp.raise_for_status()
            result = resp.json()
            adverts = result.get("adverts", [])
            all_adverts_info.extend(adverts)
            logger.info(f"Чанк {chunk_id+1}: получено {len(adverts)} кампаний")
        except Exception as e:
            logger.warning(f"Ошибка запроса по кампаниям [{ids_str}]: {e}")

    return all_adverts_info

# Функция обработки "сырых" данных доп. информации по кампаниям в DataFrame
def process_companies_info_to_df(all_adverts_info):
    """
    Преобразует "сырые" данные по кампаниям из WB API в плоский DataFrame.

    :param all_adverts_info: список словарей с данными по кампаниям (ключ 'adverts' из API)
    :return: pd.DataFrame с распарсенными и обработанными данными
    """
    if not all_adverts_info:
        return pd.DataFrame()

    df_adverts_info = pd.json_normalize(
        all_adverts_info,
        record_path=['nm_settings'],
        meta=[
            'id',
            ['settings', 'name'],
            ['settings', 'payment_type'],
            'status',
            ['timestamps', 'created'],
            ['timestamps', 'started'],
            ['timestamps', 'updated'],
            ['timestamps', 'deleted']
        ],
        errors='ignore'
    )

    # Переименование столбцов для удобства
    df_adverts_info = df_adverts_info.rename(columns={
        'id': 'ID кампании',
        'nm_id': 'Артикул WB',
        'subject.name': 'Категория',
        'bids_kopecks.recommendations': 'Ставка в рекомендациях (коп.)',
        'bids_kopecks.search': 'Ставка в поиске (коп.)',
        'status': 'Статус кампании',
        'settings.name': 'Название кампании',
        'settings.payment_type': 'Тип ставки',
        'timestamps.created': 'Дата создания',
        'timestamps.started': 'Дата запуска',
        'timestamps.updated': 'Дата последнего изменения',
        'timestamps.deleted': 'Дата удаления'
    })
    # Приведение дат к datetime без таймзоны
    date_cols = ['Дата создания', 'Дата запуска', 'Дата последнего изменения', 'Дата удаления']
    for col in date_cols:
        if col in df_adverts_info.columns:
            df_adverts_info[col] = pd.to_datetime(df_adverts_info[col], errors='coerce').dt.tz_localize(None)
    # Добавляем сегодняшнюю дату как дату выгрузки
    df_adverts_info['Дата'] = pd.Timestamp.now().normalize()
    # Перевод ставок из копеек в рубли
    df_adverts_info =  df_adverts_info.assign(**{
            'Ставка в рекомендациях (руб.)': df_adverts_info['Ставка в рекомендациях (коп.)'] / 100,
            'Ставка в поиске (руб.)': df_adverts_info['Ставка в поиске (коп.)'] / 100,
    })
    # Удаляем колонки в копейках
    df_adverts_info = df_adverts_info.drop(columns=[
        'Ставка в рекомендациях (коп.)',
        'Ставка в поиске (коп.)'
    ])
    # Отброс пустых артикулов
    df_adverts_info = df_adverts_info.dropna(subset=['Артикул WB'])
    # Приведение типов
    df_adverts_info['ID кампании'] = df_adverts_info['ID кампании'].astype(int)
    df_adverts_info['Артикул WB'] = df_adverts_info['Артикул WB'].astype(int)
    # Переводим Тип ставки в верхний регистр
    df_adverts_info['Тип ставки'] = df_adverts_info['Тип ставки'].str.upper()
    # Заменяем коды статусов на текстовые значения
    status_mapping = {
        -1: 'Удалена',
        4: 'Готова к запуску',
        7: 'Завершена',
        8: 'Отменена',
        9: 'Активна',
        11: 'На паузе'
    }
    df_adverts_info['Статус кампании'] = df_adverts_info['Статус кампании'].map(status_mapping)

    # Выбираем нужные колонки
    df_adverts_info = df_adverts_info[[
        'ID кампании',
        'Название кампании',
        'Статус кампании',
        'Тип ставки',
        'Артикул WB',
        'Ставка в рекомендациях (руб.)',
        'Ставка в поиске (руб.)'
    ]]

    df_adverts_info = df_adverts_info.reset_index(drop=True)

    return df_adverts_info

# Функция получения информации о кампаниях с помощью двух предыдущих функций
def get_companies_info(headers, df_company_list, statuses='-1,4,7,8,9,11', payment_type=None):
    # Фильтруем по списку кампаний из входного df
    df_company_list_filtered = df_company_list.loc[df_company_list['ID кампании'].isin(df_input_companies['ID кампании']), :]
    # Получаем "сырые" данные по доп. информации по кампаниям из WB API
    all_adverts_info = fetch_companies_info_raw(headers, df_company_list_filtered, statuses, payment_type)
    # Обрабатываем "сырые" данные доп. информации по кампаниям в DataFrame
    df_companies_info = process_companies_info_to_df(all_adverts_info)

    return df_companies_info


def split_into_intervals(start_date, end_date, max_days=31):
    """Разбивает интервал на интервалы не длиннее max_days."""
    intervals = []
    current_start = start_date
    while current_start <= end_date:
        current_end = min(current_start + timedelta(days=max_days - 1), end_date)
        intervals.append({
            'begin': current_start.strftime('%Y-%m-%d'),
            'end': current_end.strftime('%Y-%m-%d')
        })
        current_start = current_end + timedelta(days=1)
    return intervals

# Функция получения "сырых" данных по статистике кампаний из WB API с учетом разбиения на интервалы и чанки
def get_company_stats_raw(
        headers: dict,
        df_input_companies: pd.DataFrame,
        base_url: str = "https://advert-api.wildberries.ru/adv/v3/fullstats",
        max_days: int = 31,
        step: int = 100,
        period_seconds: int = 60,
        period_limit: int = 3,
        max_burst: int = 1,
        max_retries: int = 3,
        batch_size_companies: int = 50
) -> pd.DataFrame:

    if df_input_companies.empty:
        logger.warning("Входной датафрейм с кампаниями пуст. Возвращаю пустой результат.")
        return pd.DataFrame()

    df_companies = df_input_companies.copy()
    total_companies = len(df_companies['ID кампании'].unique())
    logger.info(f"Всего уникальных кампаний в запросе: {total_companies}")

    limiter = WbRateLimiter(
        period_seconds=period_seconds,
        period_limit=period_limit,
        max_burst=max_burst
    )

    session = requests.Session()
    session.headers.update(headers)

    unique_ids = df_companies['ID кампании'].unique()
    num_batches = (len(unique_ids) + batch_size_companies - 1) // batch_size_companies
    logger.info(
        f"Кампании будут обработаны батчами по {batch_size_companies} шт. "
        f"Всего батчей: {num_batches}"
    )

    all_stats = pd.DataFrame()

    for batch_idx in range(num_batches):
        batch_start = batch_idx * batch_size_companies
        batch_ids = unique_ids[batch_start:batch_start + batch_size_companies]

        logger.info(
            f"Обработка батча {batch_idx + 1}/{num_batches}."
        )

        df_batch = df_companies[df_companies['ID кампании'].isin(batch_ids)].copy()

        df_batch['intervals'] = df_batch.apply(
            lambda row: split_into_intervals(
                pd.to_datetime(row['date_start']),
                pd.to_datetime(row['date_end']),
                max_days=max_days
            ),
            axis=1
        )

        df_batch_full = (
            df_batch
            .explode('intervals')
            .reset_index(drop=True)
        )
        df_batch_full['interval_begin'] = df_batch_full['intervals'].apply(lambda x: x['begin'])
        df_batch_full['interval_end'] = df_batch_full['intervals'].apply(lambda x: x['end'])

        unique_intervals = df_batch_full[['interval_begin', 'interval_end']].drop_duplicates()
        logger.info(
            f"Для батча {batch_idx + 1} сформировано {len(unique_intervals)} уникальных интервалов дат."
        )

        for interval_idx, interval_row in unique_intervals.iterrows():
            begin = interval_row['interval_begin']
            end = interval_row['interval_end']

            company_ids = df_batch_full[
                (df_batch_full['interval_begin'] == begin) &
                (df_batch_full['interval_end'] == end)
            ]['ID кампании'].unique()

            num_chunks = (len(company_ids) + step - 1) // step
            logger.info(
                f"Батч {batch_idx + 1}: интервал {begin} - {end}, "
                f"кампаний в интервале: {len(company_ids)}, чанков: {num_chunks}"
            )

            for chunk_idx in range(num_chunks):
                i = chunk_idx * step
                chunk = company_ids[i:i + step]
                if len(chunk) == 0:
                    continue

                params_companies_stats = {
                    'ids': ','.join(str(cid) for cid in chunk),
                    'beginDate': begin,
                    'endDate': end
                }

                logger.info(
                    f"Батч {batch_idx + 1}/{num_batches}, интервал {begin} - {end}, "
                    f"запрос по кампаниям: {list(map(int, chunk))}"
                )

                # Определяем, является ли этот запрос последним вообще
                last_batch = (batch_idx == num_batches - 1)
                last_interval = (interval_idx == len(unique_intervals) - 1)
                last_chunk = (chunk_idx == num_chunks - 1)
                is_last_request = last_batch and last_interval and last_chunk

                for attempt in range(1, max_retries + 1):
                    resp, ok = limiter.call(
                        session.get,
                        url=base_url,
                        params=params_companies_stats,
                        skip_sleep=is_last_request,   # <- не спим после последнего запроса
                    )
                    if not ok or resp is None:
                        logger.error(
                            f"Ошибка соединения/лимита при получении статистики "
                            f"(попытка {attempt}/{max_retries})."
                        )
                        if attempt == max_retries:
                            logger.error("Достигнут максимум попыток. Пропускаю этот чанк.")
                        continue

                    status = resp.status_code
                    if status == 200:
                        try:
                            resp_json = resp.json()
                        except Exception as e:
                            logger.error(f"Не удалось распарсить JSON ответа: {e}")
                            break

                        tmp_df = pd.DataFrame(resp_json)
                        if not tmp_df.empty:
                            tmp_df['request_start_date_chunk'] = begin
                            tmp_df['request_end_date_chunk'] = end
                            all_stats = pd.concat([all_stats, tmp_df], ignore_index=True)
                            logger.info(
                                f"Успешно получена статистика для {len(chunk)} кампаний "
                                f"за период {begin} - {end}."
                            )
                        else:
                            logger.info(
                                f"Ответ для кампаний {list(map(int, chunk))} "
                                f"за период {begin} - {end} пуст."
                            )
                        break  # успешный запрос — выходим из цикла попыток

                    else:
                        try:
                            err_json = resp.json()
                        except Exception:
                            err_json = resp.text
                        logger.error(
                            f"Ошибка ответа API (HTTP {status}), попытка {attempt}/{max_retries}. "
                            f"Тело ответа: {err_json}"
                        )
                        if attempt == max_retries:
                            logger.error("Достигнут максимум попыток. Пропускаю этот чанк.")
                            break
                        time.sleep(1)
                        continue

    if all_stats.empty:
        logger.warning("Не удалось получить ни одной записи статистики. Возвращаю пустой DataFrame.")
        return all_stats

    all_stats = all_stats.sort_values(by='advertId', ignore_index=True)
    logger.info(f"Итоговый размер сырых данных со статистикой: {all_stats.shape}")
    return all_stats

# Функция обработки статистики кампаний
def process_company_stats(
        df_raw: pd.DataFrame,
        df_input_companies: pd.DataFrame
) -> pd.DataFrame:
    """
    Обрабатывает сырую статистику кампаний:
      - добавляет интервалы запроса,
      - распаковывает days/apps/nms,
      - группирует по кампании и артикулу,
      - считает метрики,
      - фильтрует по затратам > 0.
    """

    if df_raw.empty:
        logger.warning("df_raw is empty in process_company_stats.")
        return pd.DataFrame()

    # Добавляем в df статистики кампаний исходные интервалы запроса
    df_request_interval = (
        df_input_companies
        .rename(columns={
            'ID кампании': 'advertId',
            'date_start': 'Начальная дата запроса статистики РК',
            'date_end': 'Конечная дата запроса статистики РК'
        })
        .loc[:, [
            'advertId',
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК'
        ]]
        .drop_duplicates()
    )

    df_stats = (
        df_raw
        .merge(df_request_interval, on='advertId', how='left')
    )

    # Колонки, которые будем держать "в голове" дальше
    companies_columns = [
        'advertId',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'request_start_date_chunk', 'request_end_date_chunk',
    ]

    df_stats = move_columns(df_stats, companies_columns, position=0)

    # Распаковываем по дням
    df_unpacked_days = df_stats.explode(column='days', ignore_index=True)
    df_by_days = pd.concat(
        [
            df_unpacked_days[companies_columns],
            pd.json_normalize(df_unpacked_days['days'])
        ],
        axis=1
    )

    # Распаковываем по платформам
    df_unpacked_apps = df_by_days.explode(column='apps', ignore_index=True)
    df_by_apps = pd.concat(
        [
            df_unpacked_apps[companies_columns],
            pd.json_normalize(df_unpacked_apps['apps'])
        ],
        axis=1
    )

    # Распаковываем по артикулам
    df_unpacked_nms = df_by_apps.explode(column='nms', ignore_index=True)
    df_by_nms = pd.concat(
        [
            df_unpacked_nms[companies_columns],
            pd.json_normalize(df_unpacked_nms['nms'])
        ],
        axis=1
    )

    # Группировка по кампании и артикулу
    groupby_cols = [
        'advertId',
        'Начальная дата запроса статистики РК',
        'Конечная дата запроса статистики РК',
        'nmId',
        'name'
    ]

    df_agg = (
        df_by_nms
        .loc[:, groupby_cols + ['views', 'clicks', 'sum', 'atbs', 'orders', 'shks', 'sum_price']]
        .groupby(groupby_cols, as_index=False)
        .sum()
    )

    # Метрики (с защитой от деления на ноль)
    df_agg = df_agg.assign(
        ctr=lambda df: np.where(df['views'] > 0, df['clicks'] / df['views'] * 100, 0.0),
        cr=lambda df: np.where(df['clicks'] > 0, df['shks'] / df['clicks'] * 100, 0.0),
        cpc=lambda df: np.where(df['clicks'] > 0, df['sum'] / df['clicks'], 0.0),
        cpm=lambda df: np.where(df['views'] > 0, df['sum'] / df['views'] * 1000, 0.0),
    )

    # Переименование колонок
    df_agg = df_agg.rename(columns={
        'advertId': 'ID кампании',
        'nmId': 'Артикул WB',
        'name': 'Наименование товара',
        'views': 'Показы',
        'clicks': 'Клики (РК)',
        'sum': 'Затраты, руб.',
        'atbs': 'Корзина (РК)',
        'orders': 'Заказы, шт. (РК)',
        'sum_price': 'Заказы, руб. (РК)',
        'shks': 'Заказанные товары, шт. (РК)',
        'ctr': 'CTR, %',
        'cr': 'CR, %',
        'cpc': 'CPC',
        'cpm': 'CPM',
    })

    # Уникальный ID кампании по периоду
    df_agg['unique_company_id'] = (
        df_agg
        .groupby(
            ['ID кампании',
             'Начальная дата запроса статистики РК',
             'Конечная дата запроса статистики РК'],
            as_index=False
        )
        .ngroup() + 1
    )

    # Порядок колонок
    df_agg = df_agg.loc[:, [
        'unique_company_id',
        'ID кампании',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'Артикул WB', 'Наименование товара',
        'Показы', 'Клики (РК)',
        'CTR, %', 'CR, %', 'CPC', 'CPM',
        'Затраты, руб.',
        'Корзина (РК)',
        'Заказы, шт. (РК)',
        # 'Заказанные товары, шт. (РК)',
        'Заказы, руб. (РК)',
    ]]

    # Фильтр по затратам
    df_agg = (
        df_agg
        .loc[df_agg['Затраты, руб.'] > 0, :]
        .reset_index(drop=True)
    )

    return df_agg

# Функция добавления доп. информации по кампаниям к статистике кампаний
def add_companies_info_to_stats(
        df_companies_stats,
        df_companies_info
):
    # Выбираем нужные колонки из доп. информации по кампаниям
    companies_info_columns = [
        'ID кампании', 'Название кампании', 'Статус кампании',
        'Тип ставки',
        'Артикул WB',
        'Ставка в рекомендациях (руб.)',
        'Ставка в поиске (руб.)'
    ]
    missing_columns = [col for col in companies_info_columns if col not in df_companies_info.columns]
    if missing_columns:
        logger.warning(
            f"В df_companies_info отсутствуют колонки: {missing_columns}. "
        )
        df_companies_info[missing_columns] = np.nan

    # Мерджим статистику кампаний с доп. информацией по кампаниям
    df_stats_with_info = (
        df_companies_stats
        .merge(
            df_companies_info,
            how='left',
            on=['ID кампании', 'Артикул WB']
        )
    )

    return df_stats_with_info


# Функция добавления информации о товарах к статистике кампаний
def add_product_info_to_stats(
        df_companies_stats_with_info,
        df_products_processed
):
    # Выбираем нужные колонки из списка товаров
    product_info_columns = ['Артикул WB', 'Артикул продавца']
    missing_columns = [col for col in product_info_columns if col not in df_products_processed.columns]
    if missing_columns:
        logger.warning(
            f"В df_products_processed отсутствуют колонки: {missing_columns}. "
        )
        df_products_processed[missing_columns] = np.nan
    df_products_processed = df_products_processed[product_info_columns]

    # Мерджим статистику кампаний с информацией о товарах
    df_stats_with_product_info = (
        df_companies_stats_with_info
        .merge(
            df_products_processed.drop_duplicates(subset=['Артикул WB']),  # удаляем дубликаты по Артикулу WB
            how='left',
            on='Артикул WB'
        )
    )

    return df_stats_with_product_info


def get_keywords_stats(
    headers,
    df_input_companies,
    df_companies_info,
    base_url: str = "https://advert-api.wildberries.ru/adv/v0/normquery/stats",
    period_seconds: int = 60,
    period_limit: int = 10,
    max_burst: int = 20,
    max_items_per_request: int = 100,
    max_retries: int = 3,
):
    """
    Выгружает статистику поисковых кластеров (normquery stats) по WB API:
      POST https://advert-api.wildberries.ru/adv/v0/normquery/stats

    Вход:
      df_input_companies:
        - 'ID кампании'
        - 'date_start'
        - 'date_end'

      df_companies_info:
        - 'ID кампании'
        - 'Артикул WB'

    Особенности:
      - В одном запросе в "items" может быть до max_items_per_request (по умолчанию 100)
        связок (advert_id, nm_id) из РАЗНЫХ кампаний и артикулов.
      - Лимиты метода задаются через аргументы:
          period_seconds (по умолчанию 60),
          period_limit   (по умолчанию 10),
          max_burst      (по умолчанию 20).
    """

    if df_input_companies.empty:
        logger.warning("df_input_companies пуст. Возвращаю пустой DataFrame.")
        return pd.DataFrame()

    if df_companies_info.empty:
        logger.warning("df_companies_info пуст (нет артикулов по кампаниям). Возвращаю пустой DataFrame.")
        return pd.DataFrame()

    # Инициализируем лимитер с заданными параметрами
    limiter = WbRateLimiter(
        period_seconds=period_seconds,
        period_limit=period_limit,
        max_burst=max_burst,
    )

    session = requests.Session()
    session.headers.update(headers)

    all_rows = []

    # Приводим даты к datetime
    df_companies = df_input_companies.copy()
    df_companies['date_start'] = pd.to_datetime(df_companies['date_start'])
    df_companies['date_end'] = pd.to_datetime(df_companies['date_end'])

    # Оставляем только нужные колонки
    df_companies = df_companies[['ID кампании', 'date_start', 'date_end']].drop_duplicates()

    # Присоединяем информацию об артикулах к компаниям
    df_nms = df_companies_info[['ID кампании', 'Артикул WB']].dropna()
    if df_nms.empty:
        logger.warning("В df_companies_info нет ни одной строки с 'Артикул WB'. Возвращаю пустой DataFrame.")
        return pd.DataFrame()

    pairs = (
        df_nms
        .merge(df_companies, on='ID кампании', how='inner')
        .rename(columns={'Артикул WB': 'nm_id'})
    )

    if pairs.empty:
        logger.warning("После объединения кампаний и артикулов не осталось ни одной связки. Возвращаю пустой DataFrame.")
        return pd.DataFrame()

    logger.info(
        f"Всего связок (кампания + артикул) для запроса normquery/stats: {len(pairs)}"
    )

    # Собираем все связки в список
    items_all = []
    for _, r in pairs.iterrows():
        items_all.append({
            "advert_id": int(r['ID кампании']),
            "nm_id": int(r['nm_id']),
            "from": r['date_start'].strftime('%Y-%m-%d'),
            "to": r['date_end'].strftime('%Y-%m-%d')
        })

    # Разбиваем общий список связок на чанки по max_items_per_request
    num_chunks = (len(items_all) + max_items_per_request - 1) // max_items_per_request
    logger.info(
        f"Связки будут отправлены чанками по {max_items_per_request} шт. "
        f"Всего чанков: {num_chunks}"
    )

    for chunk_idx in range(num_chunks):
        start_i = chunk_idx * max_items_per_request
        end_i = start_i + max_items_per_request
        chunk_items = items_all[start_i:end_i]

        if not chunk_items:
            continue

        # Определяем общий период для чанка
        from_dates = [pd.to_datetime(x['from']) for x in chunk_items]
        to_dates = [pd.to_datetime(x['to']) for x in chunk_items]

        from_date = min(from_dates).strftime('%Y-%m-%d')
        to_date = max(to_dates).strftime('%Y-%m-%d')

        # Формируем items для запроса (без дат, только advert_id и nm_id)
        items_payload = [
            {"advert_id": it["advert_id"], "nm_id": it["nm_id"]}
            for it in chunk_items
        ]

        payload = {
            "from": from_date,
            "to": to_date,
            "items": items_payload
        }

        logger.info(
            f"Чанк {chunk_idx + 1}/{num_chunks}: "
            f"период {from_date} - {to_date}, "
            f"связок (advert_id, nm_id) в items: {len(items_payload)}"
        )

        # Вызов API с учётом лимита
        resp, ok = limiter.call(
            session.post,
            url=base_url,
            json=payload,
            max_retries=max_retries,
        )
        if not ok or resp is None:
            logger.error(
                f"Не удалось получить статистику normquery для чанка {chunk_idx + 1}/{num_chunks}, "
                f"период {from_date}-{to_date}."
            )
            continue

        if resp.status_code != 200:
            logger.error(
                f"Ошибка ответа normquery/stats: HTTP {resp.status_code} | {resp.text}"
            )
            continue

        try:
            data = resp.json()
        except Exception as e:
            logger.error(
                f"Ошибка парсинга JSON для чанка {chunk_idx + 1}/{num_chunks}, "
                f"период {from_date}-{to_date}: {e}"
            )
            continue

        stats_list = data.get("stats", [])
        if not isinstance(stats_list, list) or not stats_list:
            logger.info(
                f"Пустой список 'stats' для чанка {chunk_idx + 1}/{num_chunks}, "
                f"период {from_date}-{to_date}."
            )
            continue

        # Разбира��м ответ по каждому (advert_id, nm_id)
        for item_stats in stats_list:
            item_advert_id = item_stats.get("advert_id")
            item_nm_id = item_stats.get("nm_id")
            stats_by_queries = item_stats.get("stats", [])

            if not stats_by_queries:
                continue

            for s in stats_by_queries:
                row_out = {
                    "ID кампании": item_advert_id,
                    "Артикул WB": item_nm_id,
                    "from": from_date,
                    "to": to_date,
                    "norm_query": s.get("norm_query"),
                    "atbs": s.get("atbs"),
                    "avg_pos": s.get("avg_pos"),
                    "clicks": s.get("clicks"),
                    "cpc": s.get("cpc"),
                    "cpm": s.get("cpm"),
                    "ctr": s.get("ctr"),
                    "orders": s.get("orders"),
                    "views": s.get("views"),
                    "spend": s.get("spend"),
                }
                all_rows.append(row_out)

    if not all_rows:
        logger.warning("Не удалось получить ни одной записи статистики normquery. Возвращаю пустой DataFrame.")
        return pd.DataFrame()

    df_result = pd.DataFrame(all_rows)

    # Приведение типов (по возможности)
    for col in ["ctr", "cpc", "cpm", "avg_pos"]:
        if col in df_result.columns:
            df_result[col] = pd.to_numeric(df_result[col], errors="coerce")

    # Переименования под более «человеческий» вид
    df_result = df_result.rename(columns={
        "norm_query": "Поисковый кластер",
        "atbs": "Добавления в корзину",
        "avg_pos": "Средняя позиция",
        "clicks": "Клики",
        "cpc": "CPC",
        "cpm": "CPM",
        "ctr": "CTR, %",
        "orders": "Заказы",
        "views": "Просмотры",
        "spend": "Затраты, руб.",
        "from": "Дата начала интервала",
        "to": "Дата конца интервала",
    })

    df_result = df_result.loc[
        :,
        [
            "ID кампании",
            "Артикул WB",
            "Дата начала интервала",
            "Дата конца интервала",
            "Поисковый кластер",
            "Добавления в корзину",
            "Заказы",
            "Просмотры",
            "Клики",
            "Затраты, руб.",
            "CTR, %",
            "CPC",
            "CPM",
            "Средняя позиция",
        ]
    ]

    df_result = df_result.sort_values(
        by=["ID кампании", "Артикул WB", "Затраты, руб.", "Заказы", "Клики"],
        ascending=[True, True, False, False, False],
        ignore_index=True
    )

    return df_result

# Функция определения минимальной и максимальной даты в кампаниях
def get_min_max_dates(df_input_companies):
    # Создаем копию для избежания изменений в оригинальном df
    df_input_companies_ = df_input_companies.copy()
    # Переводим колонки в datetime
    df_input_companies_['datetime_start'] = pd.to_datetime(df_input_companies_['date_start'])
    df_input_companies_['datetime_end'] = pd.to_datetime(df_input_companies_['date_end'])
    # Находим мин. и макс. дату
    # GPT START----
    # Создаем столбец с минимумом по строке
    df_input_companies_['row_min'] = df_input_companies_[['datetime_start', 'datetime_end']].min(axis=1)
    df_input_companies_['row_max'] = df_input_companies_[['datetime_start', 'datetime_end']].max(axis=1)

    # Индекс общего минимума среди этих двух столбцов (по всем строкам)
    idx_min = df_input_companies_['row_min'].idxmin()
    idx_max = df_input_companies_['row_max'].idxmax()
    # Получаем label для строки с общим минимумом
    min_date = df_input_companies_.loc[idx_min, 'date_start']
    max_date = df_input_companies_.loc[idx_max, 'date_end']
    # Переводим в нужный формат
    min_date = min_date + 'T00:00:00'
    max_date = max_date + 'T23:59:59'
    # GPT END----

    return min_date, max_date


# Функция выгрузки заказов
def get_orders(headers, date_start, date_end):
    # Выгружаем заказы
    df_orders = getOrdersWB(headers, date_start, date_end, to_save=False)
    # Заменяем некоторые артикулы по маппингу замен
    if article_replacement and 'Артикул продавца' in df_orders.columns:
        df_orders['Артикул продавца'] = df_orders['Артикул продавца'].replace(article_replacement)
    # Переводим колонку со временем заказа в timestamp
    df_orders['datetime_orders'] = pd.to_datetime(df_orders['date'])
    return df_orders


# Функция расчета заказов для отдельной рекламной кампании с фильтром по дате
def calc_orders_for_company(
        datetime_start,
        datetime_end,
        # begin_date_company,
        # end_date_company,
        # unique_company_id,
        # company_number,
        df_orders,
        tmp_df_company,
):
    # Создаем копию для избежания изменений в оригинальном df
    # df_orders_for_company = df_orders.copy()
    # Фильтруем заказы по датам кампании
    df_orders_date_filtered = (
        df_orders
        .loc[
            df_orders['datetime_orders'].between(datetime_start, datetime_end),
              :
            ]
            .copy()
    )
    # Считаем заказы
    df_orders_for_company = (
        df_orders_date_filtered
        .groupby(['Артикул WB'], as_index=False)
        .agg(**{
            'Заказы артикула, шт. (Заказы)' : ('Артикул WB', 'count'),
            'Заказы артикула, руб. (Заказы)': ('priceWithDisc', 'sum'),
        })
    )

    # Добавляем ID кампании к заказам
    # df_orders_for_company = df_orders_for_company.assign(**{
    #     'unique_company_id': unique_company_id,
    #     'ID кампании': company_number,
    #     'begin': begin_date_company,
    #     'end': end_date_company,

    # })
    # Колонки для заполнения пропусков
    columns_to_fillna = ['Заказы артикула, шт. (Заказы)', 'Заказы артикула, руб. (Заказы)']
    # Мерджим заказы с компанией
    tmp_df_company_with_orders = (
        tmp_df_company
        .loc[:, ~tmp_df_company.columns.isin(['unique_company_id'])]
        .merge(
            df_orders_for_company,
            how='left',
            on='Артикул WB'
        )
        # Заполняем пропуски в колонках заказов
        .fillna({col: 0 for col in columns_to_fillna})
    )

    return tmp_df_company_with_orders

# Функция получения ассоциированных артикулов БЕЗ артикулов кампании
def get_associated_products(
        company_sku_list,
        df_products_processed
):
    # Удаляем дубликаты списка товаров, поскольку он распакован по размерам
    df_products_no_duplicates = (
        df_products_processed
        .loc[:, ['Артикул WB', 'imtID', 'Артикул продавца', 'Наименование товара']]
        .drop_duplicates(ignore_index=True)
    )
    # Получаем список связанных артикулов артикула РК из списка товаров
    associated_skus = (
        df_products_processed
        .loc[df_products_processed['Артикул WB'].isin(company_sku_list), 'imtID']
        .drop_duplicates()
        .to_list()
    )
    # Получаем df соответствия артикулов РК и их связанных артикулов
    df_associated_sku_all = (
        df_products_processed
        # Делаем выборку по ассоциированным артикулам
        .loc[
            df_products_processed['imtID'].isin(associated_skus),
            ['Артикул WB', 'imtID', 'Артикул продавца', 'Наименование товара']
        ]
        # Удаляем дубликаты по артикулу
        .drop_duplicates(ignore_index=True)
        # .pipe(lambda df:
        #       # Получаем
        #       df.loc[df_associated_skus['nmID'].isin(company_sku_list), :]
        # )
    )
    # Получаем df артикула РК и ID общего артикула
    df_company_sku = (
        df_associated_sku_all
        .loc[df_associated_sku_all['Артикул WB'].isin(company_sku_list), ['Артикул WB', 'imtID']]
        # .groupby(['imtID'], as_index=False)
        # .agg({
        #     'Артикул WB': lambda x: ';'.join(map(str, x))
        # })
        .rename(columns={
            'Артикул WB': 'Артикул РК'
        })
    )
    # Получаем df связанных артикулов НЕ ВКЛЮЧАЯ артикулы РК
    df_associated_sku_company = (
        # Добавляем в df ассоциированных артикулов артикул РК
        # через merge по ID связанного артикула
        df_associated_sku_all
        .merge(
            df_company_sku,
            how='inner',
            on='imtID'
        )
        # Удаляем артикулы из РК
        .pipe(lambda df:
              df.loc[
                  ~df['Артикул WB'].isin(company_sku_list),
                  ~df.columns.isin(['imtID'])
                ]
        )
        .rename(columns={
            'Артикул WB': 'Ассоциированный артикул'
        })
    )

    # Один пайплайн (нужно тестить)
    # df_associated_sku_company = (
    #     df_products_no_duplicates
    #     # Делаем выборку по артикулам РК
    #     .loc[df_products_no_duplicates['nmID'].isin(company_sku_list), ['nmID', 'imtID']]
    #     # Удаляем дубликаты по артикулу
    #     # .drop_duplicates(ignore_index=True)
    #     # Переименовываем колонку артикулом РК
    #     .rename(columns={
    #         'nmID': 'Артикул РК'
    #     })
    #     .merge(
    #         df_products_no_duplicates[['nmID', 'imtID']],
    #         how='inner',
    #         on='imtID'
    #     )
    #     # Удаляем дубликаты по артикулу (поскольку список товаров распакован по размерам)
    #     # .drop_duplicates(ignore_index=True)
    #     # Удаляем артикулы из РК
    #     .pipe(lambda df:
    #           df.loc[
    #               ~df['nmID'].isin(company_sku_list),
    #               ~df.columns.isin(['imtID'])
    #             ]
    #     )
    #     # Переименовываем колонку с артикулами, которые пришли из списка товаров
    #     .rename(columns={
    #         'nmID': 'Ассоциированный артикул'
    #     })
    # )

    return df_associated_sku_company

# Функция расчета заказов по ассоциированным артикулам
def calc_associated_orders(
        datetime_start,
        datetime_end,
        company_sku_list,
        # begin_date_company,
        # end_date_company,
        # unique_company_id,
        company_number,
        df_products_processed,
        df_orders,
        tmp_df_company,
):
    # Получаем список ассоциированных артикулов кампании
    # БЕЗ самого артикула РК
    df_associated_sku_company = get_associated_products(
        company_sku_list,
        df_products_processed
    )
    # Получаем список для фильтрации заказов
    associated_sku_list = df_associated_sku_company['Ассоциированный артикул'].unique().tolist()
    # Создаем копию для избежания изменений в оригинальном df
    # df_orders_for_company = df_orders.copy()
    # Фильтруем заказы по датам кампании
    df_orders_date_filtered = (
        df_orders
        .loc[
            df_orders['datetime_orders'].between(datetime_start, datetime_end),
              :
            ]
            .copy()
    )
    # Фильтруем заказы по ассоциированным артикулам
    df_orders_sku_date_filtered = (
        df_orders_date_filtered
        .loc[
            df_orders_date_filtered['Артикул WB'].isin(associated_sku_list),
            :
        ]
    )
    # Считаем заказы
    df_orders_for_company = (
        df_orders_sku_date_filtered
        .groupby(['Артикул WB'], as_index=False)
        .agg(**{
            'Заказы артикула, шт. (Заказы)' : ('Артикул WB', 'count'),
            'Заказы артикула, руб. (Заказы)': ('priceWithDisc', 'sum'),
        })
        .rename(columns={
            'Артикул WB': 'Ассоциированный артикул'
        })
    )
    # Добавляем заказы к списку ассоицированнных артикулов
    df_associated_sku_with_orders = df_associated_sku_company.merge(
        df_orders_for_company,
        on='Ассоциированный артикул',
        how='left'
    )

    # Заполняем пропуски
    df_associated_sku_with_orders = df_associated_sku_with_orders.fillna(0)

    # Добавляем ID кампании к заказам
    df_associated_sku_with_orders = df_associated_sku_with_orders.assign(**{
        # 'unique_company_id': unique_company_id,
        'ID кампании': company_number,
        'Начальная дата запроса статистики РК': datetime_start,
        'Конечная дата запроса статистики РК': datetime_end,
    })
    # Делаем порядок колонок
    df_associated_sku_with_orders = df_associated_sku_with_orders.loc[
        :,
        [
            'ID кампании',
            'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
            'Ассоциированный артикул', 'Артикул РК',
            'Артикул продавца', 'Наименование товара',
            'Заказы артикула, шт. (Заказы)',
            'Заказы артикула, руб. (Заказы)',
        ]
    ]

    return df_associated_sku_with_orders

# Функция расчета заказов для каждой из рекламных кампаний
def calc_orders_for_companies(
        df_input_companies,
        df_orders,
        df_companies_stats_with_products,
):
    # df, куда будем помещать итоговый результат
    df_companies_stats_with_orders = pd.DataFrame()
    df_associated_orders = pd.DataFrame()
    # Цикл по каждой кампании
    for unique_company_id in df_input_companies['unique_company_id'].unique():
        # Временный df с данными текущей кампании
        tmp_df_company = (
            df_companies_stats_with_products
            .loc[df_companies_stats_with_products['unique_company_id'] == unique_company_id, :]
            # .drop_duplicates(subset=['ID кампании'])
        )
        # Если вернулась статистика по данной РК, то начинаем расчет статистики
        if not tmp_df_company.empty:
            # Получаем ID кампании
            company_number = tmp_df_company['ID кампании'].iloc[0]
            # Получаем даты начала и окончания кампании (для фильтра заказов)
            datetime_start_orders = (
                pd.to_datetime(tmp_df_company['Начальная дата запроса статистики РК'])
                .dt.tz_localize(None)
                .iloc[0]
            )
            datetime_end_orders = (
                pd.to_datetime(tmp_df_company['Конечная дата запроса статистики РК'])
                .dt.tz_localize(None)
                # У конечной даты заменяем время на 23:59:59
                .apply(lambda x: x.replace(hour=23, minute=59, second=59))
                .iloc[0]
            )
            # Эти даты нужны для мерджа
            begin_date_company = tmp_df_company['Начальная дата запроса статистики РК'].iloc[0]
            end_date_company = tmp_df_company['Конечная дата запроса статистики РК'].iloc[0]
            # Получаем список исходных SKU кампании
            company_sku_list = tmp_df_company['Артикул WB'].to_list()
            # Считаем заказы с фильтром по данным датам
            # и добавляем данные по заказам к кампании (расчет + merge в одной функции)
            tmp_df_company_with_orders = calc_orders_for_company(
                datetime_start_orders,
                datetime_end_orders,
                # begin_date_company,
                # end_date_company,
                # unique_company_id,
                # company_number,
                df_orders,
                tmp_df_company
            )
            # Считаем заказы по ассоциированным артикулам
            tmp_df_associated_orders = calc_associated_orders(
                datetime_start_orders,
                datetime_end_orders,
                company_sku_list,
                company_number,
                df_products_processed,
                df_orders,
                tmp_df_company
            )
        # Если статистика по РК не вернулась, то заказы по ней не считаем
        else:
            company_with_no_stats = df_input_companies.loc[
                df_input_companies['unique_company_id'] == unique_company_id,
                ['ID кампании', 'date_start', 'date_end']
            ]
            logger.warning(f"No company stats for company\n {company_with_no_stats}")
            tmp_df_company_with_orders = pd.DataFrame()
            tmp_df_associated_orders = pd.DataFrame()
        # Объединяем с предыдущей кампанией
        df_companies_stats_with_orders = pd.concat([
                df_companies_stats_with_orders,
                tmp_df_company_with_orders
                ],
            ignore_index=True
        )
        df_associated_orders = pd.concat([
                df_associated_orders,
                tmp_df_associated_orders
                ],
            ignore_index=True
        )

    # Формируем словарь с результатами
    result_companies_orders = {
        'Статистика РК': df_companies_stats_with_orders,
        'Ассоциированные заказы': df_associated_orders
    }

    return df_companies_stats_with_orders, df_associated_orders


# Функция добавления данных из воронки продаж
def add_data_from_sales_funnel(df_companies_stats_with_orders):
    # Из статистики кампаний формируем список уникальных артикулов и дат
    results = []
    groups = (
        df_companies_stats_with_orders
        .groupby([
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК'
        ])
    )
    # Цикл по каждой группе
    for (date_from, date_to), group in groups:
        # Выбираем артикулы в группе
        nm_ids = group['Артикул WB'].tolist()
        # Разбиваем список артикулов на чанки по 1000
        for i in range(0, len(nm_ids), 1000):
            nm_ids_chunk = nm_ids[i:i+1000]
            # Выгружаем данные воронки продаж
            tmp_df_sales_funnel = get_sales_funnel(
                headers=headers,
                nm_ids=nm_ids_chunk,
                date_from=date_from,
                date_to=date_to,
                limit=50
            )
            # Добавляем даты из группы к воронке продаж
            tmp_df_sales_funnel['Начальная дата запроса статистики РК'] = date_from
            tmp_df_sales_funnel['Конечная дата запроса статистики РК'] = date_to
            # Добавляем результаты в общий список
            results.append(tmp_df_sales_funnel)

    # Объединяем результаты в один DataFrame
    df_sales_funnel = pd.concat(results, ignore_index=True)
    # Выбираем нужные колонки из воронки продаж
    df_sales_funnel = (
        df_sales_funnel
        .rename(columns={
            'Количество переходов в карточку товара': 'Клики (общие)',
            'Положили в корзину, шт.': 'Корзина (общая)'
        })
        .loc[:, [
            'Артикул WB',
            'Клики (общие)',
            'Корзина (общая)',
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК',
        ]]
    )

    # Объединяем данные со статистикой РК
    df_companies_stats_with_sales_funnel = df_companies_stats_with_orders.merge(
        df_sales_funnel,
        on=[
            'Артикул WB',
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК',
        ],
        how='left'
    )
    # Перемещаем колонки
    df_companies_stats_with_sales_funnel = move_columns(
        df_companies_stats_with_sales_funnel,
        'Клики (общие)',
        'Клики (РК)',
    )
    df_companies_stats_with_sales_funnel = move_columns(
        df_companies_stats_with_sales_funnel,
        'Корзина (общая)',
        'Корзина (РК)',
    )

    return df_companies_stats_with_sales_funnel

# Функция расчета итоговых колонок в статистике РК
def calc_final_orders(df_companies_stats_with_sales_funnel):
    # Достаем df со статистикой РК из словаря
    df_companies_stats_final_orders = df_companies_stats_with_sales_funnel.copy()
    # Считаем затраты артикула РК
    df_companies_stats_final_orders['Затраты артикула, руб.'] = (
        df_companies_stats_final_orders
        .groupby(['Артикул WB'])['Затраты, руб.']
        .transform('sum')
    )
    # Считаем ДРР
    df_companies_stats_final_orders['ДРР'] = np.where(
        df_companies_stats_final_orders['Заказы артикула, руб. (Заказы)'] > 0,
        (
            df_companies_stats_final_orders['Затраты, руб.']
            / df_companies_stats_final_orders['Заказы артикула, руб. (Заказы)']
            * 100
        ),
        np.nan
    )
    # Считаем ДРР артикула
    df_companies_stats_final_orders['ДРР артикула'] = np.where(
        df_companies_stats_final_orders['Заказы артикула, руб. (Заказы)'] > 0,
        (
            df_companies_stats_final_orders['Затраты артикула, руб.']
            / df_companies_stats_final_orders['Заказы артикула, руб. (Заказы)']
            * 100
        ),
        np.nan
    )

    return df_companies_stats_final_orders


# Функция создания df для записи в Excel
def create_final_df_for_excel(df_final_orders):
    # Создаем копию для избежания изменений в оригинальном df
    df_excel_companies_report = df_final_orders.copy()
    # Переименовываем колонки
    df_excel_companies_report = df_excel_companies_report.rename(columns={
        'Наименование товара': 'Название товара',
    })
    # Формируем список колонок для итогового df
    columns_for_excel = [
        'ID кампании',
        'Артикул WB',
        'Артикул продавца',
        'Тип ставки',
        'Ставка в рекомендациях (руб.)',
        'Ставка в поиске (руб.)',
        'Название товара',
        'Показы',
        'Клики (РК)',
        'Клики (общие)',
        'CTR, %',
        'CR, %',
        'CPC',
        'CPM',
        'Затраты, руб.',
        'Затраты артикула, руб.',
        'Корзина (РК)',
        'Корзина (общая)',
        'Заказы, шт. (РК)',
        'Заказы, руб. (РК)',
        'Заказы артикула, шт. (Заказы)',
        'Заказы артикула, руб. (Заказы)',
        'ДРР',
        'ДРР артикула'
    ]
    missing_columns = [col for col in columns_for_excel if col not in df_excel_companies_report.columns]
    if missing_columns:
        logger.warning(
            f"В финальном df для Excel отсутствуют колонки: {missing_columns}. "
            f"Итоговый df для Excel будет содержать NaN в этих колонках."
        )
        df_excel_companies_report[missing_columns] = np.nan
    # Выбираем нужные колонки и переименовываем их для итогового df
    df_excel_companies_report = df_excel_companies_report[columns_for_excel]

    # Заменяем 0 и 0.0 на NaN во всех колонках
    # df_final_for_excel = df_final_for_excel.replace({0: np.nan, 0.0: np.nan})


    return df_excel_companies_report



# Функция расчета показателей по ключевым словам
def calc_keyword_stats(df_companies_keywords_stats):
    # Создаем копию для избежания изменений в оригинальном df
    df_keywords_stats = df_companies_keywords_stats.copy()
    # Считаем суммарные показатели
    df_keywords_stats = (
        df_keywords_stats
        .groupby(['ID кампании', 'Поисковый кластер'], as_index=False)
        # Считаем суммы по кампании
        .agg(**{
            'Затраты': ('Затраты, руб.', 'sum'),
            'Клики': ('Клики', 'sum'),
            'Просмотры': ('Просмотры', 'sum')
        })
        # Считаем доп. столбцы
        .assign(
            CTR=lambda x: (
                (x['Клики'] / x['Просмотры'] * 100)
                .replace([np.inf, -np.inf], np.nan)
                .fillna(0)
                .round(2)
            )
        )
        # Сортировка
        .sort_values(by=['ID кампании', 'Затраты'], ascending=[True, False])
    )

    return df_keywords_stats

# Функция формирования словаря с итоговыми результатами
def create_result_dictionary(
        df_date_range,
        df_final_orders,
        df_associated_orders,
        df_companies_keywords_stats,
        df_keywords_stats,
):
    # # Убираем лишние колонки из статистики РК
    # df_final_orders_ = (
    #     df_final_orders.drop(columns=[
    #         'Начальная дата запроса статистики РК',
    #         'Конечная дата запроса статистики РК'
    #     ])
    # )

    # Получаем мин. и макс. дату для записи на лист
    date_start_excel = (
        pd.to_datetime(df_date_range['date_start'])
        .min()
        .strftime('%d.%m')
    )
    date_end_excel = (
        pd.to_datetime(df_date_range['date_end'])
        .max()
        .strftime('%d.%m')
    )

    # Добавляем статистику по ключевым словам в словарь с кампаниями
    result_companies_report = {
        f'{client_name} {date_start_excel}-{date_end_excel}': df_final_orders,
        'Ассоциированные заказы': df_associated_orders,
        'Ключевые слова (свод)': df_keywords_stats,
        # 'Ключевые слова': df_companies_keywords_stats
    }
    # В каждом df заменяем 0 и 0.0 на NaN
    for key in result_companies_report:
        result_companies_report[key] = result_companies_report[key].replace({0: np.nan, 0.0: np.nan})

    return result_companies_report


def format_all_sheets_common_wb(wb, exclude_columns=None):
    """
    Общее форматирование для всех листов книги (WB):
    - границы
    - заголовки: центр, перенос, жирный, общий цвет (RGB 204,192,218)
    - данные: вертикальное выравнивание по центру, без переноса (по умолчанию)
    - числовой формат (целое / с двумя знаками) для чисел, кроме исключённых колонок
    - автоширина столбцов с небольшим запасом
    - автофильтр по всей таблице
    - дополнительное форматирование столбцов по COMPANY_STATS_COL_FORMAT:
        * поиск по startswith(ключ)
        * для совпавших — выставляем width и alignment из словаря
        * для остальных — остаётся автоподбор ширины и базовое выравнивание
    """

    if exclude_columns is None:
        exclude_columns = [
            "Артикул",
            "Артикул WB",
            "Артикул продавца",
            "ID кампании",
            "Ассоциированный артикул",
            "Артикул РК"
        ]

    exclude_columns_lower = {col.lower() for col in exclude_columns}

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

    # Базовое выравнивание данных
    data_alignment_default = Alignment(
        vertical='center',
        wrap_text=False
    )

    # Цвет заголовка (RGB 204, 192, 218)
    HEADER_COLOR = "CCC0DA"

    header_fill_common = PatternFill(
        start_color=HEADER_COLOR,
        end_color=HEADER_COLOR,
        fill_type="solid"
    )

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

        # Высота первой строки (как у Ozon)
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
                    cell.fill = header_fill_common
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

        # Автофильтр по всей таблице
        ws.auto_filter.ref = ws.dimensions

        # Дополнительное форматирование по COMPANY_STATS_COL_FORMAT
        # Для каждой колонки листа ищем первый key, который является префиксом заголовка
        for col_idx, header in enumerate(headers, start=1):
            if not isinstance(header, str):
                continue

            fmt = None
            for key, cfg in COMPANY_STATS_COL_FORMAT.items():
                if header.startswith(key):
                    fmt = cfg
                    break

            if fmt is None:
                continue  # для этой колонки остаётся автофит и базовое выравнивание

            width = fmt.get('width')
            align_kwargs = fmt.get('alignment', {})
            alignment = Alignment(**align_kwargs)

            col_letter = get_column_letter(col_idx)

            # переопределяем ширину
            if width is not None:
                ws.column_dimensions[col_letter].width = width

            # выравнивание только данных (со 2-й строки)
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.alignment = alignment


def format_wb_client_sheet(
        wb,
        sheet_name: str,
        df_client: pd.DataFrame,
        percent_cols=None,
):
    """
    Форматирование листа сводного отчёта по клиенту (WB):
    - делим значения ДРР на 100
    - задаём формат процентов
    - условное форматирование (градиент) по % столбцам
    """
    if sheet_name not in wb.sheetnames:
        return

    if percent_cols is None:
        percent_cols = ['ДРР', 'ДРР артикула']

    ws = wb[sheet_name]

    # --- 1. Процентный формат ---
    for col_name in percent_cols:
        if col_name not in df_client.columns:
            continue

        col_idx = list(df_client.columns).index(col_name) + 1  # Excel: 1-based
        for row in range(2, ws.max_row + 1):  # пропускаем заголовок
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                # если в df уже проценты (15.3 → 15.3%), делим на 100
                cell.value = cell.value * 0.01
            cell.number_format = '0.00%'

    # --- 2. Градиентное условное форматирование по тем же столбцам ---
    for col_name in percent_cols:
        if col_name not in df_client.columns:
            continue

        col_idx = list(df_client.columns).index(col_name) + 1
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        cell_range = f"{col_letter}2:{col_letter}{ws.max_row}"

        rule = ColorScaleRule(
            start_type='min', start_color='63BE7B',           # зелёный
            mid_type='percentile', mid_value=50, mid_color='FFEB84',  # жёлтый
            end_type='max', end_color='E06967',               # красный
        )
        ws.conditional_formatting.add(cell_range, rule)


def save_and_format_excel(
        report_path: Path,
        date_campaign_report: str,
        client_name: str,
        result_companies_report: dict,
        df_date_range: pd.DataFrame,
):
    """
    Сохранение и форматирование отчёта РК по WB.

    result_companies_report: dict[sheet_name -> DataFrame]
    """
    # Имя файла для сохранения
    file_name = report_path / f"{date_campaign_report}_Отчет_РК_{client_name}_WB.xlsx"

    report_path.mkdir(parents=True, exist_ok=True)

    # Создаём книгу и записываем все dfs по листам
    with pd.ExcelWriter(file_name) as w:
        for sheet_name, df in result_companies_report.items():
            df.to_excel(w, index=False, sheet_name=sheet_name)

    # Открываем и форматируем
    wb = openpyxl.load_workbook(file_name)

    # 1) Общее форматирование всех листов (границы, заголовки, числа, ширина, формат по маппингу)
    format_all_sheets_common_wb(wb)

    # 2) Форматирование листа по клиенту (ДРР и ДРР артикула)
    date_start_excel = (
        pd.to_datetime(df_date_range['date_start'])
        .min()
        .strftime('%d.%m')
    )
    date_end_excel = (
        pd.to_datetime(df_date_range['date_end'])
        .max()
        .strftime('%d.%m')
    )
    sheet_name_client = f"{client_name} {date_start_excel}-{date_end_excel}"

    format_wb_client_sheet(
        wb,
        sheet_name=sheet_name_client,
        df_client=result_companies_report.get(sheet_name_client, pd.DataFrame()),
        percent_cols=['ДРР', 'ДРР артикула']
    )

    wb.save(file_name)
    return file_name


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


    # Генерируем диапазон дат
    df_date_range = generate_dates(
        start_date=date_start_cli,
        end_date=date_end_cli,
        # start_date='2026-03-01',
        # end_date='2026-03-31'
    )
    # Список кампаний и даты для каждой кампании,
    # за которые нужно сформировать отчет
    input_companies = create_input_companies(client_name, df_date_range, cli_companies_parsed)
    # Дата формирования отчета
    date_campaign_report = str(date.today())
    # Создаем df с нужным списком кампаний и датами
    df_input_companies = create_companies_upload_df(input_companies)

    # Получаем список кампаний
    df_company_list = get_company_list(headers)
    # Получаем доп. информацию по кампаниям
    df_companies_info = get_companies_info(headers, df_input_companies)


    # Получаем статистику кампаний
    df_companies_stats_raw = get_company_stats_raw(headers, df_input_companies)

    # Получаем статистику по ключевым фразам в поиске
    df_companies_keywords_stats = get_keywords_stats(
        headers,
        df_input_companies,
        df_companies_info,
    )

    # Получаем список товаров
    df_products = getWBProduct(headers, to_save=False)
    # Обрабатываем список товаров
    df_products_processed = process_product_list(df_products)

    # Находим мин. и макс. дату для получения заказов
    date_start_orders, date_end_orders = get_min_max_dates(df_input_companies)
    # Выгружаем заказы
    df_orders = get_orders(headers, date_start_orders, date_end_orders)

    # Обрабатываем статистику кампаний
    df_companies_stats = process_company_stats(df_companies_stats_raw, df_input_companies)
    # Добавляем к статистике кампаний информацию по SKU из кампании
    df_companies_stats_with_info = add_companies_info_to_stats(df_companies_stats, df_companies_info)


    # Добавляем данные по товарам в статистику кампаний
    df_companies_stats_with_products = add_product_info_to_stats(
        df_companies_stats_with_info,
        df_products_processed
    )

    # Считаем заказы для каждой кампании
    df_companies_stats_with_orders, df_associated_orders = calc_orders_for_companies(
        df_input_companies,
        df_orders,
        df_companies_stats_with_products
    )
    # Добавляем данные из воронки продаж
    df_companies_stats_with_sales_funnel = add_data_from_sales_funnel(
        df_companies_stats_with_orders
    )
    # Получаем историю затрат (для корректных расходов на кампании)
    # df_costs_history_stats = get_costs_history(headers, date_start, date_end)
    # Считаем расходы на номенклатуры в кампании
    # df_companies_sku_stats = calc_costs_by_sku(result_companies_info, df_costs_history_stats)

    # Считаем ДРР и итоговые заказы
    df_final_orders = calc_final_orders(df_companies_stats_with_sales_funnel)
    # Создаем df для записи в Excel
    df_final_orders_for_excel = create_final_df_for_excel(df_final_orders)

     # Считаем статистику по ключевым словам с группировкой
    df_keywords_stats = calc_keyword_stats(df_companies_keywords_stats)

    # Создаем итоговый словарь с результатами
    result_companies_report = create_result_dictionary(
        df_date_range,
        df_final_orders_for_excel,
        df_associated_orders,
        df_companies_keywords_stats,
        df_keywords_stats,
    )
    # Создаем директорию, куда будет сохранен отчет
    report_path = create_report_path(date_campaign_report, client_name)
    # Сохраняем и форматируем файл с отчетом
    filepath_report = save_and_format_excel(
        report_path,
        date_campaign_report,
        client_name,
        result_companies_report,
        df_date_range
    )
    logger.info(f"Done creating campaigns report for client {client_name}")
    logger.info(f"Campaigns report saved in {filepath_report}")

# %%
