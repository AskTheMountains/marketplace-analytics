# %% Определение функций
import argparse
import copy
import time
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

pd.options.mode.chained_assignment = None  # default='warn'

# Некоторые вспомогательные функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from wb.scripts.command_args_parser import cli_companies
# Файл с настройками и номером клиента
from wb.scripts.constants import client_name, headers, PROMOTION_COMPANIES
# Функции выгрузки данных
from wb.scripts.uploadDataFromWB import getOrdersWB, getWBProduct
# Доп. функции
from wb.scripts.wb_generic_functions import (
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
            columns=['Номер кампании', 'date_start', 'date_end']
        )
        # .reset_index()
        # .rename(columns={
        #     'index': 'Номер кампании'
        # })
    )
    # Удаляем дубликаты
    df_input_companies = (
        df_input_companies
        .loc[~df_input_companies['Номер кампании'].isin(['']), :]
        .drop_duplicates(subset=['Номер кампании', 'date_start', 'date_end'])
    )
    # Присваиваем каждой кампании уникальный ID по номеру кампании и дате
    df_input_companies['unique_company_id'] = (
        df_input_companies
        .groupby(['Номер кампании', 'date_start', 'date_end'], as_index=False)
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


# Функция получения информации о кампаниях
def get_companies_info(headers, df_input_companies):
    # Создаем копию для избежания изменений в оригинальном df
    df_company_list_ = df_input_companies.copy()
    # Разбиваем df с кампаниями на интервалы по 50 кампаний
    step = 50
    df_company_list_['id'] = np.arange(0, len(df_company_list_))
    df_company_list_['chunks'] = df_company_list_['id'].apply(lambda x: int(x/step) + 1)
    # df для каждого из типов кампаний
    df_companies_6 = pd.DataFrame()
    df_companies_8 = pd.DataFrame()
    df_companies_9 = pd.DataFrame()
    logger.info("Uploading companies info")
    # Цикл по каждому интервалу
    for chunk in df_company_list_['chunks'].unique():
        # Получаем список кампаний
        df_company_list_chunk = df_company_list_.loc[df_company_list_['chunks'] == chunk, :]
        company_list = df_company_list_chunk['Номер кампании'].to_list()
        # Переводим номера кампаний в int для запроса
        # company_list_int = [int(elem) for elem in company_list]
        # Параметры запроса
        # params_companies_info = {company_list}
        resp_data_companies_info = requests.post("https://advert-api.wildberries.ru/adv/v1/promotion/adverts", headers=headers, json=company_list).json()
        # Переводим в df
        tmp_df_companies_info = pd.DataFrame(resp_data_companies_info)
        # Переводим колонки с датой в timestamp
        for col in ['endTime', 'createTime', 'changeTime', 'startTime']:
            tmp_df_companies_info[col] = pd.to_datetime(tmp_df_companies_info[col], format='mixed')
            # Убираем нахуй таймзону
            tmp_df_companies_info[col] = tmp_df_companies_info[col].dt.tz_localize(None)
        # Делаем выборку на каждый из типов кампаний
        tmp_df_companies_6 = tmp_df_companies_info.loc[~tmp_df_companies_info['type'].isin([8, 9]), :]
        tmp_df_companies_8 = tmp_df_companies_info.loc[tmp_df_companies_info['type'] == 8, :]
        tmp_df_companies_9 = tmp_df_companies_info.loc[tmp_df_companies_info['type'] == 9, :]
        # Объединяем с предыдущей итерацией цикла
        df_companies_6 = pd.concat([df_companies_6, tmp_df_companies_6])
        df_companies_8 = pd.concat([df_companies_8, tmp_df_companies_8])
        df_companies_9 = pd.concat([df_companies_9, tmp_df_companies_9])

    # Обработка данных кампаний со статусами, отличными от 8 и 9
    if not df_companies_6.empty:
        # Делаем reset_index после loc
        df_companies_6 = df_companies_6.reset_index(drop=True)
        # Убираем лишние колонки
        df_companies_6 = df_companies_6.loc[:, ~df_companies_6.columns.isin(['nmCPM', 'unitedParams', 'autoParams'])]
        # Переводим колонки с датой в timestamp
        # for col in ['endTime', 'createTime', 'changeTime', 'startTime']:
        #     df_companies_6[col] = pd.to_datetime(df_companies_6[col], format='mixed')
        # Получаем параметры кампании из списка
        df_companies_6_params = (
            df_companies_6
            .loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'params']]
            .explode(column='params')
            .reset_index(drop=True)
        )
        # Получаем параметры кампании в виде столбцов df
        df_companies_6_params = pd.concat([
            df_companies_6_params,
            pd.json_normalize(df_companies_6_params['params'])
        ], axis=1)
        # Получаем список номенклатур кампании
        df_companies_6_nms = df_companies_6_params.explode(column='nms').reset_index(drop=True)
        # Получаем список номенклатур кампании в виде столбцов df
        df_companies_6_nms = pd.concat([df_companies_6_nms, pd.json_normalize(df_companies_6_nms['nms'])], axis=1)
        # Убираем лишние колонки
        df_companies_6_nms = df_companies_6_nms.loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'nm']]
        # Переименовываем колонку чтобы дальше использовать concat
        df_companies_6_nms = df_companies_6_nms.rename(columns={'nm': 'nms'})
    else:
        df_companies_6_nms = pd.DataFrame()

    # Обработка данных кампаний со статусом 8
    if not df_companies_8.empty:
        # Делаем reset_index после loc
        df_companies_8 = df_companies_8.reset_index(drop=True)
        # Убираем лишние колонки
        df_companies_8 = df_companies_8.loc[:, ~df_companies_8.columns.isin(['nmCPM', 'unitedParams'])]
        # Переводим колонки с датой в timestamp
        # for col in ['endTime', 'createTime', 'changeTime', 'startTime']:
        #     df_companies_8[col] = pd.to_datetime(df_companies_8[col], format='mixed')
        # Получаем параметры кампании из списка
        df_companies_8_params = (
            df_companies_8
            .loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'autoParams']]
        )
        # Получаем параметры кампании в виде столбцов df
        df_companies_8_params = pd.concat([df_companies_8_params, pd.json_normalize(df_companies_8_params['autoParams'])], axis=1)
        # Получаем список номенклатур
        df_companies_8_nms = df_companies_8_params.explode(column='nms')
        # Убираем лишние колонки
        df_companies_8_nms = df_companies_8_nms.loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'nms']]
    else:
        df_companies_8_nms = pd.DataFrame()

    # Обработка данных кампаний со статусами, отличными от 8 и 9
    if not df_companies_9.empty:
        # Делаем reset_index после loc
        df_companies_9 = df_companies_9.reset_index(drop=True)
        # Убираем лишние колонки
        df_companies_9 = df_companies_9.loc[:, ~df_companies_9.columns.isin(['nmCPM', 'autoParams'])]
        # Переводим колонки с датой в timestamp
        # for col in ['endTime', 'createTime', 'changeTime', 'startTime']:
        #     df_companies_9[col] = pd.to_datetime(df_companies_9[col], format='mixed')
        # Получаем параметры кампании из списка
        df_companies_9_params = (
            df_companies_9
            .loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'unitedParams']]
            .explode(column='unitedParams')
            .reset_index(drop=True)
        )
        # Получаем параметры кампании в виде столбцов df
        df_companies_9_params = pd.concat([
            df_companies_9_params,
            pd.json_normalize(df_companies_9_params['unitedParams'])
        ], axis=1)
        # Получаем список номенклатур кампании
        df_companies_9_nms = df_companies_9_params.explode(column='nms').reset_index(drop=True)
        # Получаем список номенклатур кампании в виде столбцов df
        # df_companies_9_nms = pd.concat([df_companies_9_nms, pd.json_normalize(df_companies_9_nms['nms'])], axis=1)
        # Убираем лишние колонки
        df_companies_9_nms = df_companies_9_nms.loc[:, ['endTime', 'createTime', 'changeTime', 'startTime', 'advertId', 'nms']]
    else:
        df_companies_9_nms = pd.DataFrame()

    # Формируем итоговый список номенклатур по всем кампаниям
    df_companies_nms_all = pd.concat([df_companies_6_nms, df_companies_8_nms, df_companies_9_nms])
    # Переименовываем некоторые колонки для удобства
    df_companies_nms_all = df_companies_nms_all.rename(columns={
        'nms': 'Артикул WB',
        'advertId': 'Номер кампании',
    })
    # Получаем количество артикулов в кампании
    df_companies_nms_all['Количество артикулов в кампании'] = df_companies_nms_all.groupby('Номер кампании').transform('size')
    # Перемещаем колонки в начало для удобства
    df_companies_nms_all = move_columns(
        df_companies_nms_all,
        ['Номер кампании', 'Артикул WB', 'Количество артикулов в кампании'],
        position=0
    )
    # Формируем словарь с результатами
    result_companies_info = {
        'df_companies_list': df_company_list,
        'df_companies_nms_all': df_companies_nms_all
    }

    return df_companies_nms_all


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


def get_company_stats_raw(
        headers: dict,
        df_input_companies: pd.DataFrame,
        base_url: str = "https://advert-api.wildberries.ru/adv/v3/fullstats",
        max_days: int = 31,
        step: int = 100,
        period_seconds: int = 60,
        period_limit: int = 3,
        min_interval_seconds: float = 20.0,
        max_retries: int = 3,
        batch_size_companies: int = 50
) -> pd.DataFrame:
    """
    Получает сырую статистику кампаний по API WB в виде DataFrame.

    Делает:
      - разбиение списка кампаний на батчи по batch_size_companies;
      - разбиение дат на интервалы не длиннее max_days;
      - вызовы API с учётом rate limit (1 мин / 3 запроса / ~20 сек);
      - сбор ответов в один DataFrame.

    Ожидаемые колонки в df_input_companies:
      - 'Номер кампании'
      - 'date_start'
      - 'date_end'
    """

    if df_input_companies.empty:
        logger.warning("Входной датафрейм с кампаниями пуст. Возвращаю пустой результат.")
        return pd.DataFrame()

    df_companies = df_input_companies.copy()

    total_companies = len(df_companies['Номер кампании'].unique())
    logger.info(f"Всего уникальных кампаний в запросе: {total_companies}")

    # Разбиваем входные кампании на батчи по batch_size_companies
    unique_ids = df_companies['Номер кампании'].unique()
    num_batches = (len(unique_ids) + batch_size_companies - 1) // batch_size_companies

    logger.info(
        f"Кампании будут обработаны батчами по {batch_size_companies} шт. "
        f"Всего батчей: {num_batches}"
    )

    all_stats = pd.DataFrame()

    # Глобальный контроль лимита по времени
    request_counter = 0
    period_start_ts = time.time()

    for batch_idx in range(num_batches):
        batch_start = batch_idx * batch_size_companies
        batch_ids = unique_ids[batch_start:batch_start + batch_size_companies]

        logger.info(
            f"Обработка батча {batch_idx + 1}/{num_batches}. "
            # f"Кампании: {list(map(int, batch_ids))}"
        )

        df_batch = df_companies[df_companies['Номер кампании'].isin(batch_ids)].copy()

        # Формируем интервалы не более max_days
        df_batch['intervals'] = df_batch.apply(
            lambda row: split_into_intervals(
                pd.to_datetime(row['date_start']),
                pd.to_datetime(row['date_end']),
                max_days=max_days
            ),
            axis=1
        )

        # Распаковываем интервалы
        df_batch_full = (
            df_batch
            .explode('intervals')
            .reset_index(drop=True)
        )
        df_batch_full['interval_begin'] = df_batch_full['intervals'].apply(lambda x: x['begin'])
        df_batch_full['interval_end'] = df_batch_full['intervals'].apply(lambda x: x['end'])

        # Уникальные интервалы дат в батче
        unique_intervals = df_batch_full[['interval_begin', 'interval_end']].drop_duplicates()

        logger.info(
            f"Для батча {batch_idx + 1} сформировано {len(unique_intervals)} уникальных интервалов дат."
        )

        # Цикл по интервалам дат в данном батче
        for interval_idx, interval_row in unique_intervals.iterrows():
            begin = interval_row['interval_begin']
            end = interval_row['interval_end']

            # Кампании батча, попадающие в этот интервал
            company_ids = df_batch_full[
                (df_batch_full['interval_begin'] == begin) &
                (df_batch_full['interval_end'] == end)
            ]['Номер кампании'].unique()

            # Разбиваем кампании на чанки по step (лимит 100 кампаний в одном запросе WB)
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

                # === Контроль лимита запросов ===
                now = time.time()
                # Если прошло больше period_seconds с начала окна — сбрасываем счётчик
                if now - period_start_ts >= period_seconds:
                    logger.info("Начинается новое минутное окно лимита. Сбрасываю счётчик запросов.")
                    period_start_ts = now
                    request_counter = 0

                # Если уже достигли лимита запросов в текущем окне — ждём до конца минуты
                if request_counter >= period_limit:
                    sleep_time = period_seconds - (now - period_start_ts)
                    if sleep_time > 0:
                        logger.info(
                            f"Достигнут лимит {period_limit} запросов за {period_seconds} секунд. "
                            f"Пауза {sleep_time:.1f} сек."
                        )
                        time.sleep(sleep_time)
                    period_start_ts = time.time()
                    request_counter = 0

                # Повторные попытки при 429 и временных ошибках
                for attempt in range(1, max_retries + 1):
                    try:
                        resp = requests.get(
                            base_url,
                            headers=headers,
                            params=params_companies_stats
                        )
                    except Exception as e:
                        logger.error(
                            f"Ошибка при выполнении запроса (попытка {attempt}/{max_retries}): {e}"
                        )
                        if attempt == max_retries:
                            logger.error("Достигнут максимум попыток. Пропускаю этот чанк.")
                            break
                        time.sleep(5)
                        continue

                    status = resp.status_code
                    rl_remaining = resp.headers.get("X-Ratelimit-Remaining")

                    if status == 200:
                        request_counter += 1

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
                                f"Ответ для кампаний {list(map(int, chunk))} за период {begin} - {end} пуст."
                            )

                        # Если не последний запрос — выдерживаем интервал
                        last_batch = (batch_idx == num_batches - 1)
                        last_interval = (interval_idx == len(unique_intervals) - 1)
                        last_chunk = (chunk_idx == num_chunks - 1)
                        is_last_request = last_batch and last_interval and last_chunk

                        if not is_last_request:
                            if rl_remaining is not None and rl_remaining == '0':
                                logger.info(
                                    "Заголовок X-Ratelimit-Remaining = 0. "
                                    f"Делаю паузу {min_interval_seconds} сек."
                                )
                                time.sleep(min_interval_seconds)
                            else:
                                logger.info(
                                    f"Ожидание {min_interval_seconds} сек перед следующим запросом."
                                )
                                time.sleep(min_interval_seconds)

                        break  # выходим из цикла попыток для этого чанка

                    elif status == 429:
                        logger.info(
                            f"Ответ 429 Too Many Requests (попытка {attempt}/{max_retries}). "
                            "Пауза 60 сек и повтор запроса."
                        )
                        time.sleep(60)
                        if attempt == max_retries:
                            logger.error("Достигнут максимум попыток после 429. Пропускаю этот чанк.")
                        continue

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
                        time.sleep(5)
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
            'Номер кампании': 'advertId',
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
        'advertId': 'Номер кампании',
        'nmId': 'Артикул WB',
        'name': 'Наименование товара',
        'views': 'Показы',
        'clicks': 'Клики РК',
        'sum': 'Затраты руб',
        'atbs': 'Корзина РК',
        'orders': 'Заказы шт (из отчета РК)',
        'sum_price': 'Заказы руб (из отчета РК)',
        'shks': 'Заказанные товары шт (из отчета РК)',
        'ctr': 'CTR, %',
        'cr': 'CR, %',
        'cpc': 'CPC',
        'cpm': 'CPM',
    })

    # Уникальный ID кампании по периоду
    df_agg['unique_company_id'] = (
        df_agg
        .groupby(
            ['Номер кампании',
             'Начальная дата запроса статистики РК',
             'Конечная дата запроса статистики РК'],
            as_index=False
        )
        .ngroup() + 1
    )

    # Порядок колонок
    df_agg = df_agg.loc[:, [
        'unique_company_id',
        'Номер кампании',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'Артикул WB', 'Наименование товара',
        'Показы', 'Клики РК',
        'CTR, %', 'CR, %', 'CPC', 'CPM',
        'Затраты руб',
        'Корзина РК',
        'Заказы шт (из отчета РК)',
        'Заказанные товары шт (из отчета РК)',
        'Заказы руб (из отчета РК)',
    ]]

    # Фильтр по затратам
    df_agg = (
        df_agg
        .loc[df_agg['Затраты руб'] > 0, :]
        .reset_index(drop=True)
    )

    return df_agg


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
    # Выгружаем заказа
    df_orders = getOrdersWB(headers, date_start, date_end, to_save=False)
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
            'Заказы артикула шт (из отчета по заказам)' : ('Артикул WB', 'count'),
            'Заказы артикула руб (из отчета по заказам)': ('priceWithDisc', 'sum'),
        })
    )

    # Добавляем номер кампании к заказам
    # df_orders_for_company = df_orders_for_company.assign(**{
    #     'unique_company_id': unique_company_id,
    #     'Номер кампании': company_number,
    #     'begin': begin_date_company,
    #     'end': end_date_company,

    # })
    # Колонки для заполнения пропусков
    columns_to_fillna = ['Заказы артикула шт (из отчета по заказам)', 'Заказы артикула руб (из отчета по заказам)']
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
        .loc[:, ['nmID', 'imtID', 'Артикул продавца', 'Наименование товара']]
        .drop_duplicates(ignore_index=True)
    )
    # Получаем список связанных артикулов артикула РК из списка товаров
    associated_skus = (
        df_products_processed
        .loc[df_products_processed['nmID'].isin(company_sku_list), 'imtID']
        .drop_duplicates()
        .to_list()
    )
    # Получаем df соответствия артикулов РК и их связанных артикулов
    df_associated_sku_all = (
        df_products_processed
        # Делаем выборку по ассоциированным артикулам
        .loc[
            df_products_processed['imtID'].isin(associated_skus),
            ['nmID', 'imtID', 'Артикул продавца', 'Наименование товара']
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
        .loc[df_associated_sku_all['nmID'].isin(company_sku_list), ['nmID', 'imtID']]
        # .groupby(['imtID'], as_index=False)
        # .agg({
        #     'nmID': lambda x: ';'.join(map(str, x))
        # })
        .rename(columns={
            'nmID': 'Артикул РК'
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
                  ~df['nmID'].isin(company_sku_list),
                  ~df.columns.isin(['imtID'])
                ]
        )
        .rename(columns={
            'nmID': 'Ассоциированный артикул'
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
            'Заказы артикула шт (из отчета по заказам)' : ('Артикул WB', 'count'),
            'Заказы артикула руб (из отчета по заказам)': ('priceWithDisc', 'sum'),
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

    # Добавляем номер кампании к заказам
    df_associated_sku_with_orders = df_associated_sku_with_orders.assign(**{
        # 'unique_company_id': unique_company_id,
        'Номер кампании': company_number,
        'Начальная дата запроса статистики РК': datetime_start,
        'Конечная дата запроса статистики РК': datetime_end,
    })
    # Делаем порядок колонок
    df_associated_sku_with_orders = df_associated_sku_with_orders.loc[
        :,
        [
            'Номер кампании',
            'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
            'Ассоциированный артикул', 'Артикул РК',
            'Артикул продавца', 'Наименование товара',
            'Заказы артикула шт (из отчета по заказам)',
            'Заказы артикула руб (из отчета по заказам)',
        ]
    ]

    return df_associated_sku_with_orders

# Функция расчета заказов для каждой из рекламных кампаний
def calc_orders_for_companies(
        df_input_companies,
        df_orders,
        df_companies_stats
):
    # df, куда будем помещать итоговый результат
    df_companies_stats_with_orders = pd.DataFrame()
    df_associated_orders = pd.DataFrame()
    # Цикл по каждой кампании
    for unique_company_id in df_input_companies['unique_company_id'].unique():
        # Временный df с данными текущей кампании
        tmp_df_company = (
            df_companies_stats
            .loc[df_companies_stats['unique_company_id'] == unique_company_id, :]
            # .drop_duplicates(subset=['Номер кампании'])
        )
        # Если вернулась статистика по данной РК, то начинаем расчет статистики
        if not tmp_df_company.empty:
            # Получаем номер кампании
            company_number = tmp_df_company['Номер кампании'].iloc[0]
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
                ['Номер кампании', 'date_start', 'date_end']
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
            'Количество переходов в карточку товара': 'Клики общие',
            'Положили в корзину, шт.': 'Корзина общая'
        })
        .loc[:, [
            'Артикул WB',
            'Клики общие',
            'Корзина общая',
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
        'Клики общие',
        'Клики РК',
    )
    df_companies_stats_with_sales_funnel = move_columns(
        df_companies_stats_with_sales_funnel,
        'Корзина общая',
        'Корзина РК',
    )

    return df_companies_stats_with_sales_funnel

# Функция расчета итоговых колонок в статистике РК
def calc_final_orders(df_companies_stats_with_sales_funnel):
    # Достаем df со статистикой РК из словаря
    df_companies_stats_final_orders = df_companies_stats_with_sales_funnel.copy()
    # Считаем ДРР
    df_companies_stats_final_orders['ДРР, %'] = np.where(
        df_companies_stats_final_orders['Заказы артикула руб (из отчета по заказам)'] > 0,
        (
            df_companies_stats_final_orders['Затраты руб']
            / df_companies_stats_final_orders['Заказы артикула руб (из отчета по заказам)']
            * 100
        ),
        np.nan
    )

    return df_companies_stats_final_orders


# Функция выгрузки статистики по ключевым словам
def get_keywords_stats(headers, df_input_companies):


    # df, куда будем помещать итоговый результат
    df_keyword_company_stats_api = pd.DataFrame()

    # Делаем цикл по каждой кампании
    for idx, row in df_input_companies.iterrows():
        # Получаем номер кампании
        company_number = row['Номер кампании']
        # Получаем даты, за которые нужно получить статистику
        date_start_company = pd.to_datetime(row['date_start'])
        date_end_company = pd.to_datetime(row['date_end'])
        # Разбиваем диапазон дат кампании на интервалы по 7 дней
        date_range = []
        # Генерируем даты с разницей не более 8 дней и без пересечений
        current_date = date_start_company
        while current_date <= date_end_company:
            # Устанавливаем конечную дату как текущую дату плюс 7 дней или дату окончания, если она раньше
            new_end_date = min(current_date + pd.Timedelta(days=6), date_end_company)
            date_range.append({
                'advertId': company_number,
                'date_from': current_date,
                'date_to': new_end_date
            })
            # Переходим к следующей начальной дате (добавляем 8 дней)
            current_date = new_end_date + pd.Timedelta(days=1)
        # Создаем DataFrame
        df_date_range = pd.DataFrame(date_range)
        # Создаем даты для запроса
        df_date_range = df_date_range.assign(
            date_from_request=lambda df: df['date_from'].dt.strftime('%Y-%m-%d'),
            date_to_request=lambda df: df['date_to'].dt.strftime('%Y-%m-%d'),

        )

        # Считаем, сколько получилось интервалов
        total_intervals = df_date_range.shape[0]
        # Сколько выгружено интервалов
        uploaded_intervals = 0

        # Цикл по каждому интервалу
        for idx, row in df_date_range.iterrows():
            logger.info(
                f"\n"
                f"Uploading keyword statistics: \n"
                f"advertID: {company_number}: \n"
                f"dates: {row['date_from_request']} - {row['date_to_request']}"
            )
            # Формируем параметры запроса
            params_keyword_companies = {
                'advert_id': company_number,
                'from': row['date_from_request'],
                'to': row['date_to_request']
            }
            # Делаем запрос
            resp_data_keyword_companies = requests.get(
                'https://advert-api.wildberries.ru/adv/v0/stats/keywords',
                headers=headers,
                params=params_keyword_companies
            )
            # Если запрос был успешным, создаем df из данных апи
            if resp_data_keyword_companies.status_code == 200:
                resp_data_keyword_companies = resp_data_keyword_companies.json()
                # Переводим в df
                tmp_df_keyword_companies = pd.DataFrame(resp_data_keyword_companies)
                # Добавляем номер кампании
                tmp_df_keyword_companies['Номер кампании'] = company_number

                # Объединяем с предыдущим проходом цикла\
                df_keyword_company_stats_api = pd.concat([
                    tmp_df_keyword_companies,
                    df_keyword_company_stats_api,
                ], ignore_index=True)

            # Если была ошибка, выводим текст ошибки
            else:
                logger.warning(f'Ошибка запроса: {resp_data_keyword_companies.text}')

            # Делаем паузу между запросами
            time.sleep(0.25)

    if not df_keyword_company_stats_api.empty:
        # Обработка данных кампаний поиска по ключевым фразам
        # Достаем данные из словарей
        df_companies_keywords_normalized = (
            pd.concat([
                df_keyword_company_stats_api,
                pd.json_normalize(df_keyword_company_stats_api['keywords'])
            ], axis=1)
        )
        # Достаем данные из списков
        df_companies_keywords_unpacked = df_companies_keywords_normalized.explode(['stats'], ignore_index=True)
        # Достаем данные из словарей списков
        df_companies_keywords_stats = (
            pd.concat([
                df_companies_keywords_unpacked,
                pd.json_normalize(df_companies_keywords_unpacked['stats']),
            ], axis=1)
            # Убираем колонку со словарями
            .pipe(
                lambda df:
                    df.loc[:, ~df.columns.isin(['stats'])]
            )
            # Переименовываем колонки
            .rename(columns={
                'date': 'Дата',
                'keyword': 'Ключевая фраза',
                'views': 'Просмотры',
                'clicks': 'Клики',
                'ctr': 'CTR',
                'sum': 'Затраты'
            })
        )
        # Переводим колонку с датой в datetime
        df_companies_keywords_stats['Дата'] = pd.to_datetime(df_companies_keywords_stats['Дата'])
        # Сортируем по затратам
        df_companies_keywords_stats = (
            df_companies_keywords_stats
            .sort_values(
                by=['Номер кампании', 'Затраты'],
                ascending=False,
                ignore_index=True
            )
        )
        # Делаем удобный порядок колонок
        df_companies_keywords_stats = df_companies_keywords_stats.loc[
            :,
            [
                'Номер кампании',
                'Дата',
                'Ключевая фраза',
                'Затраты',
                'Клики',
                'Просмотры',
                'CTR'
            ]
        ]
    else:
        df_companies_keywords_stats = pd.DataFrame()

    return df_companies_keywords_stats


# Функция расчета показателей по ключевым словам
def calc_keyword_stats(df_companies_keywords_stats):
    # Создаем копию для избежания изменений в оригинальном df
    df_keywords_stats = df_companies_keywords_stats.copy()
    # Считаем суммарные показатели
    df_keywords_stats = (
        df_keywords_stats
        .groupby(['Номер кампании', 'Ключевая фраза'], as_index=False)
        # Считаем суммы по кампании
        .agg(**{
            'Затраты': ('Затраты', 'sum'),
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
        .sort_values(by=['Номер кампании', 'Затраты'], ascending=[True, False])
    )

    return df_keywords_stats

# Функция формирования словаря с итоговыми результатами
def create_result_dictionary(
        df_final_orders,
        df_associated_orders,
        df_companies_keywords_stats,
        df_keywords_stats,
):
    # Убираем лишние колонки из статистики РК
    df_final_orders_ = (
        df_final_orders.drop(columns=[
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК'
        ])
    )
    # Добавляем статистику по ключевым словам в словарь с кампаниями
    result_companies_report = {
        'Статистика РК': df_final_orders_,
        'Ассоциированные заказы': df_associated_orders,
        'Ключевые слова (свод)': df_keywords_stats,
        # 'Ключевые слова': df_companies_keywords_stats
    }

    return result_companies_report


# Функция сохранения excel
def save_and_format_excel(
        report_path,
        date_campaign_report,
        client_name,
        result_companies_report
):
    logger.info(f"Formatting campaigns report for client {client_name}")

    # Имя файла для сохранения
    file_name = (
        report_path / f"{date_campaign_report}_Отчет_РК_{client_name}_WB.xlsx"
    )

    # Создаем книгу excel и сохраняем в неё df из словаря с результатами
    with pd.ExcelWriter(file_name) as w:
        for sheet_name, df in result_companies_report.items():
            df.to_excel(w, index=False, sheet_name=sheet_name)

    # Форматирование файла с отчетом
    # GPT START----
    wb = openpyxl.load_workbook(file_name)

    # Создаем границы для ячеек
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)


    # 1. Задаем список названий столбцов, которые НЕ надо форматировать как числа
    exclude_format_headers = [
     'Номер кампании', 'Артикул', 'Ассоциированный артикул', 'Артикул РК', 'Артикул WB',
    ]
    # Цикл по каждому листу
    for ws in wb.worksheets:
        # 2. Сохраняем названия колонок для дальнейшей фильтрации
        headers = [cell.value for cell in ws[1]]

        # 3. Определяем индексы столбцов, которые НЕ нужно форматировать как числа
        exclude_format_idx = [
            idx + 1  # openpyxl считает колонки с 1
            for idx, header in enumerate(headers)
            if header in exclude_format_headers
        ]

        # 4. Обходим все ячейки данных (начиная со 2-й строки)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for col_idx, cell in enumerate(row, start=1):
                # 5. Пропускаем столбцы, для которых не нужно форматирование чисел
                if col_idx in exclude_format_idx:
                    continue
                # 6. Форматируем целые числа: разделитель тысяч, без дробной части
                if isinstance(cell.value, int):
                    cell.number_format = '#,##0'
                # 7. Форматируем дробные числа:
                #    - Если значение "целое" по факту (например, 10000.0), показываем без знаков после запятой
                #    - Иначе — 2 знака после запятой, разделитель тысяч
                elif isinstance(cell.value, float):
                    if cell.value % 1 == 0:
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = '#,##0.00'

        # Границы и автоподбор ширины
        column_widths = {}
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell_value_len = len(str(cell.value)) if cell.value is not None else 0
                column_widths[cell.column] = max(column_widths.get(cell.column, 0), cell_value_len)

        for col_idx, width in column_widths.items():
            # Для строк заголовков добавляем запас под фильтр
            header_len = len(str(ws.cell(row=1, column=col_idx).value)) if ws.cell(row=1, column=col_idx).value else 0
            # Даём больше места, если ширина определилась по заголовку
            reserve = 5 if width == header_len else 2
            ws.column_dimensions[get_column_letter(col_idx)].width = width + reserve

        # Высота загловков
        ws.row_dimensions[1].height = 25
        # Фильтр на первый ряд (заголовки)
        ws.auto_filter.ref = ws.dimensions

        # Перенос текста в заголовках (первая строка)
        for cell in ws[1]:
            cell.alignment = Alignment(
                vertical='center',
                horizontal='left',
                wrap_text=True
            )

    wb.save(file_name)

    return file_name

    # GPT END----


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
    # Получаем список номенклатур кампании
    # df_companies_nms = get_companies_info(headers, df_input_companies)
    # Получаем статистику кампаний
    df_companies_stats_raw = get_company_stats_raw(headers, df_input_companies)
    # Обрабатываем статистику кампаний
    df_companies_stats = process_company_stats(df_companies_stats_raw, df_input_companies)

    # Получаем список товаров
    df_products = getWBProduct(headers, to_save=False)
    # Обрабатываем список товаров
    df_products_processed = process_product_list(df_products)

    # Находим мин. и макс. дату для получения заказов
    date_start_orders, date_end_orders = get_min_max_dates(df_input_companies)
    # Выгружаем заказы
    df_orders = get_orders(headers, date_start_orders, date_end_orders)

    # Считаем заказы для каждой кампании
    df_companies_stats_with_orders, df_associated_orders = calc_orders_for_companies(
        df_input_companies,
        df_orders,
        df_companies_stats
    )
    # Добавляем данные из воронки продаж
    df_companies_stats_with_sales_funnel = add_data_from_sales_funnel(
        df_companies_stats_with_orders
    )
    # Получаем статистику по ключевым фразам в поиске
    df_companies_keywords_stats = get_keywords_stats(
        headers,
        df_input_companies
    )
    # Считаем статистику по ключевым словам с группировкой
    df_keywords_stats = calc_keyword_stats(df_companies_keywords_stats)
    # Получаем историю затрат (для корректных расходов на кампании)
    # df_costs_history_stats = get_costs_history(headers, date_start, date_end)
    # Считаем расходы на номенклатуры в кампании
    # df_companies_sku_stats = calc_costs_by_sku(result_companies_info, df_costs_history_stats)

    # Считаем ДРР и итоговые заказы
    df_final_orders = calc_final_orders(df_companies_stats_with_sales_funnel)
    # Создаем итоговый словарь с результатами
    result_companies_report = create_result_dictionary(
        df_final_orders,
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
        result_companies_report
    )
    logger.info(f"Done creating campaigns report for client {client_name}")
    logger.info(f"Campaigns report saved in {filepath_report}")

# %%
