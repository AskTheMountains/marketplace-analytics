# %% Определение функций
import ast
import time
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger

pd.options.mode.chained_assignment = None  # default='warn'

# Файл с настройками и номером клиента
# Функции выгрузки данных
# Некоторые вспомогательные функции
from generic_functions import move_columns

from wb.scripts.constants import (
    realization_report_column_map
)

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent


# Функция приведения размера к строковому типу с нормализацией чисел (векторная)
def normalize_size_series(series):
    """
    Векторно приводит размер к строковому типу.
    Целые числа (36.0, '36.0', 42) -> '36', '42'; дробные (43.5) -> '43.5'.
    Пустые/NaN значения -> np.nan; нечисловые размеры ('M', 'S') остаются как есть.
    """
    s = series.copy()
    # None/NaN -> np.nan (пустые размеры будут представлены как np.nan/pd.isna)
    s = s.where(~s.isna(), np.nan)
    # в строку и убираем лишние пробелы
    s = s.astype(str).str.strip()
    # пустые строки и строковые представления пустот (после astype(str) NaN -> 'nan') -> np.nan
    s = s.replace(['', 'nan', 'None', '<NA>', 'NaT'], np.nan)
    # парсим числа (нечисловые значения и NaN -> NaN, остаются строками как есть)
    num = pd.to_numeric(s, errors='coerce')
    # маска: где удалось распарсить число
    is_num = num.notna()
    # целые числа: 36.0 -> True, 43.5 -> False
    is_integer = is_num & (num == np.floor(num))
    out = s.copy()
    if is_integer.any():
        out.loc[is_integer] = num.loc[is_integer].astype('int64').astype(str)
    fractional = is_num & ~is_integer
    if fractional.any():
        out.loc[fractional] = num.loc[fractional].astype(str)
    return out


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
        headers,
        df_input_companies: pd.DataFrame,
        max_days: int = 31,
        step: int = 100,                 # лимит WB: макс. кампаний в запросе
        period_seconds: int = 60,        # окно лимита (сек)
        period_limit: int = 3,           # макс. запросов за окно
        min_interval_seconds: float = 20.0,  # интервал между запросами
        max_retries: int = 3,
        batch_size_companies: int = 50,  # батчи кампаний по 50
        base_url: str = "https://advert-api.wildberries.ru/adv/v3/fullstats"
) -> pd.DataFrame:
    """
    Получение сырой статистики РК (WB adv/v3/fullstats) с учётом:
      - лимита API: 1 мин / 3 запроса / ~20 сек;
      - разбиения дат по max_days;
      - разбиения кампаний на батчи по batch_size_companies.

    Ожидаемые колонки в df_input_companies:
      - 'Номер кампании'
      - 'date_start'
      - 'date_end'

    Возвращает DataFrame в «сыром» виде, как у WB, + поля:
      - 'request_start_date_chunk'
      - 'request_end_date_chunk'
    """

    if df_input_companies.empty:
        logger.warning("Входной датафрейм с кампаниями пуст. Возвращаю пустой результат.")
        return pd.DataFrame()

    df_companies = df_input_companies.copy()

    total_companies = len(df_companies['Номер кампании'].unique())
    logger.info(f"Всего уникальных кампаний в запросе: {total_companies}")

    # Разбиваем список кампаний на батчи по 50
    unique_ids = df_companies['Номер кампании'].unique()
    num_batches = (len(unique_ids) + batch_size_companies - 1) // batch_size_companies
    logger.info(
        f"Кампании будут обработаны батчами по {batch_size_companies} шт. "
        f"Всего батчей: {num_batches}"
    )

    all_stats = pd.DataFrame()

    # Контроль лимита WB
    request_counter = 0
    period_start_ts = time.time()

    for batch_idx in range(num_batches):
        batch_start = batch_idx * batch_size_companies
        batch_ids = unique_ids[batch_start:batch_start + batch_size_companies]

        logger.info(
            f"Обработка батча {batch_idx + 1}/{num_batches}. "
            f"Кампании: {list(map(int, batch_ids))}"
        )

        df_batch = df_companies[df_companies['Номер кампании'].isin(batch_ids)].copy()

        # Формируем интервалы не более max_days для каждой кампании в батче
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

        # Уникальные интервалы дат внутри батча
        unique_intervals = df_batch_full[['interval_begin', 'interval_end']].drop_duplicates()
        logger.info(
            f"Для батча {batch_idx + 1} сформировано {len(unique_intervals)} уникальных интервалов дат."
        )

        # Цикл по интервалам дат
        for interval_idx, interval_row in unique_intervals.iterrows():
            begin = interval_row['interval_begin']
            end = interval_row['interval_end']

            # Кампании этого батча, попадающие в данный интервал дат
            company_ids = df_batch_full[
                (df_batch_full['interval_begin'] == begin) &
                (df_batch_full['interval_end'] == end)
            ]['Номер кампании'].unique()

            # Разбиваем этот список кампаний на чанки по step (макс. 100 кампаний/запрос)
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
                # если прошла минута - новое окно лимита
                if now - period_start_ts >= period_seconds:
                    logger.info("Начинается новое минутное окно лимита. Сбрасываю счётчик запросов.")
                    period_start_ts = now
                    request_counter = 0

                # если упёрлись в лимит 3 запроса за минуту — ждём до конца окна
                if request_counter >= period_limit:
                    sleep_time = period_seconds - (now - period_start_ts)
                    if sleep_time > 0:
                        logger.warning(
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

                        # Проверяем, последний ли это запрос в целом
                        last_batch = (batch_idx == num_batches - 1)
                        last_interval = (interval_idx == len(unique_intervals) - 1)
                        last_chunk = (chunk_idx == num_chunks - 1)
                        is_last_request = last_batch and last_interval and last_chunk

                        if not is_last_request:
                            # если WB отдал X-Ratelimit-Remaining = 0 — подстрахуемся
                            if rl_remaining is not None and rl_remaining == '0':
                                logger.warning(
                                    "Заголовок X-Ratelimit-Remaining = 0. "
                                    f"Делаю паузу {min_interval_seconds} сек."
                                )
                                time.sleep(min_interval_seconds)
                            else:
                                logger.info(
                                    f"Ожидание {min_interval_seconds} сек перед следующим запросом."
                                )
                                time.sleep(min_interval_seconds)

                        break  # успешно, выходим из цикла попыток

                    elif status == 429:
                        logger.warning(
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

# Функция обработки статистики рекламных кампаний
def process_company_stats(
        df_raw: pd.DataFrame,
        df_input_companies: pd.DataFrame,
        groupby_cols = ['advertId', 'nmId', 'name']
) -> pd.DataFrame:
    """
    Обработка сырой статистики РК:
      - добавляет интервалы запроса,
      - распаковывает days/apps/nms,
      - группирует по groupby_cols,
      - считает метрики,
      - фильтрует по затратам > 0.
    """

    if df_raw.empty:
        logger.warning("Сырая статистика пуста, возвращаю пустой DataFrame.")
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

    companies_columns = [
        'advertId',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'request_start_date_chunk', 'request_end_date_chunk',
    ]

    df_stats = move_columns(df_stats, companies_columns, position=0)

    # Распаковываем по дням
    df_companies_stats_unpacked = df_stats.explode(column='days', ignore_index=True)
    df_companies_stats_by_days = pd.concat(
        [
            df_companies_stats_unpacked[companies_columns],
            pd.json_normalize(df_companies_stats_unpacked['days']),
        ],
        axis=1
    )

    # Распаковываем по платформам
    df_by_days_unpacked = df_companies_stats_by_days.explode(column='apps', ignore_index=True)
    df_by_apps = pd.concat(
        [
            df_by_days_unpacked[companies_columns],
            pd.json_normalize(df_by_days_unpacked['apps']),
        ],
        axis=1
    )

    # Распаковываем по артикулам
    df_by_apps_unpacked = df_by_apps.explode(column='nms', ignore_index=True)
    df_by_nms = pd.concat(
        [
            df_by_apps_unpacked[companies_columns],
            pd.json_normalize(df_by_apps_unpacked['nms']),
        ],
        axis=1
    )

    # Группировка по кампании/артикулам и т.д.
    df_companies_stats_nms = (
        df_by_nms
        .loc[:, groupby_cols + ['views', 'clicks', 'sum', 'atbs', 'orders', 'shks', 'sum_price']]
        .groupby(groupby_cols)
        .sum()
        .reset_index()
    )

    # Метрики с защитой от деления на ноль
    df_companies_stats_nms = df_companies_stats_nms.assign(
        ctr=lambda df: np.where(df['views'] > 0, df['clicks'] / df['views'] * 100, 0.0),
        cr=lambda df: np.where(df['clicks'] > 0, df['shks'] / df['clicks'] * 100, 0.0),
        cpc=lambda df: np.where(df['clicks'] > 0, df['sum'] / df['clicks'], 0.0),
        cpm=lambda df: np.where(df['views'] > 0, df['sum'] / df['views'] * 1000, 0.0),
    )

    # Переименование колонок — как у вас
    df_companies_stats_nms = df_companies_stats_nms.rename(columns={
        'advertId': 'Номер кампании',
        'nmId': 'Артикул WB',
        'name': 'Наименование товара',
        'views': 'Показы',
        'clicks': 'Клики',
        'sum': 'Затраты руб',
        'atbs': 'Добавлений в корзину',
        'orders': 'Заказы по рекламе шт (из отчета РК)',
        'sum_price': 'Заказы по рекламе руб (из отчета РК)',
        'shks': 'Заказанные товары шт (из отчета РК)',
        'ctr': 'CTR, %',
        'cr': 'CR, %',
        'cpc': 'CPC',
        'cpm': 'CPM',
    })

    # Фильтруем только по тем артикулам, по которым были расходы
    df_companies_stats_nms = (
        df_companies_stats_nms
        .loc[df_companies_stats_nms['Затраты руб'] > 0, :]
        .reset_index(drop=True)
    )

    return df_companies_stats_nms

def get_company_stats(
        headers,
        df_input_companies,
        groupby_cols = ['advertId','nmId', 'name']
):
    df_raw = get_company_stats_raw(
        headers=headers,
        df_input_companies=df_input_companies
    )
    df_final = process_company_stats(
        df_raw=df_raw,
        df_input_companies=df_input_companies,
        groupby_cols=groupby_cols
    )
    return df_final


# --- Выгрузка отчета реализации --- #
## Вспомогательные функции
def wildberries_get_report_detail(
    headers, date_from, date_to, rrdid=0, limit=100000,
    last_call_time=None, max_retries=5, period='weekly',
):
    """
    Запрос за один кусок отчёта с лимитом 1 запрос в минуту.
    Перед запросом делается пауза, чтобы выдержать не менее 60 секунд
    с момента последнего успешного запроса, если last_call_time передан.
    При ошибке 429 (Too Many Requests) ожидаем время из заголовка
    X-Ratelimit-Retry или по умолчанию 60 секунд и повторяем запрос
    (не более max_retries раз).
    Аргументы:
      headers (dict): заголовки с авторизацией API.
      date_from (str): дата начала периода 'YYYY-MM-DD'.
      date_to (str): дата конца периода 'YYYY-MM-DD'.
      rrdid (int): смещение пагинации.
      limit (int): лимит строк в ответе (максимум 100000).
      last_call_time (float): время последнего успешного запроса (секунды UNIX).
      max_retries (int): максимальное число повторных попыток при 429.
      period (str): тип периода отчета - недельный или по дням ('weekly' или 'daily').
    Возвращает:
      tuple: (response_data, время успешного запроса).
    """

    url = 'https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed'
    params = {
        'dateFrom': date_from,
        'dateTo': date_to,
        'rrdid': rrdid,
        'limit': limit,
        'period': period,
    }

    attempt = 1  # счётчик попыток (включая первую)
    while attempt <= max_retries:
        # Если известна метка времени последнего успешного запроса, ждём, чтобы
        # прошло как минимум 60 секунд, иначе спим разницу
        if last_call_time is not None:
            elapsed = time.time() - last_call_time
            wait = 60 - elapsed
            if wait > 0:
                if logger:
                    logger.info(f'Пауза {int(wait)} сек перед запросом (лимит 1 в минуту)')
                time.sleep(wait)

        try:
            response = requests.post(url, headers=headers, json=params)
        except Exception as e:
            if logger:
                logger.error(
                    f"Ошибка соединения при запросе отчёта WB "
                    f"(попытка {attempt}/{max_retries}): {e}"
                )
            # при ошибке соединения не знаем точное время вызова, но фиксируем попытку
            if attempt == max_retries:
                # после максимальной попытки пробрасываем исключение
                raise
            attempt += 1
            continue

        # Запоминаем время текущего запроса (успешного или нет) для контроля пауз
        last_call_time = time.time()

        if response.status_code == 200:
            response_data = response.json()
            response.raise_for_status()
            return response_data, last_call_time

        elif response.status_code == 429:
            # Превышен лимит запросов — читаем сколько ждать из заголовка X-Ratelimit-Retry
            retry_after = response.headers.get('X-Ratelimit-Retry')
            wait_time = int(retry_after) if retry_after and retry_after.isdigit() else 60
            if logger:
                logger.warning(
                    f'Получен 429 Too Many Requests. '
                    f'Ждём {wait_time} сек перед повтором '
                    f'(попытка {attempt}/{max_retries}).'
                )
            time.sleep(wait_time)
            if attempt == max_retries:
                if logger:
                    logger.error('Превышено максимальное число повторов после 429')
                response.raise_for_status()
            attempt += 1
            continue

        elif response.status_code == 204:
            if logger:
                logger.warning(f'Пустой отчет за период: {date_from} - {date_to}')
            return [], last_call_time

        else:
            if logger:
                logger.error(
                    f"Ошибка получения отчета реализации (код {response.status_code}), "
                    f"попытка {attempt}/{max_retries}, ответ:\n{response.text}"
                )
            # для других кодов ошибок тоже ограничиваемся max_retries
            if attempt == max_retries:
                response.raise_for_status()
            attempt += 1
            # можно сделать небольшую паузу между повторами, если нужно:
            # time.sleep(5)
            continue

    # На всякий случай (по логике сюда не должны дойти, т.к. raise внутри цикла)
    if logger:
        logger.error("Не удалось получить отчёт WB после всех попыток.")
    raise RuntimeError("wildberries_get_report_detail: превышено число попыток запроса")

## Основная логика (обработка недель, rrdId и бинарный сплит)
def upload_realization_report(headers, date_start, date_end,  period='weekly', verbose=False,):
    """
    Выгружает финансовый отчет Wildberries по частям (по недельным окнам) с контролем лимита
    1 запрос в минуту.

    Аргументы:
      headers (dict): Заголовки запроса с токеном.
      date_start (str): Дата начала ('YYYY-MM-DD').
      date_end (str): Дата конца ('YYYY-MM-DD').
      period (str): тип периода отчета - недельный или по дням ('weekly' или 'daily').
      verbose (bool): Выводить ли подробный лог.

    Возвращает:
      pd.DataFrame со всеми уникальными строками за указанный период.
    """

    all_rows = []             # Список для сбора всех строк отчета
    dedup_key_set = set()     # Для дедупликации по (realizationreport_id, rrdId)

    cursor_start = datetime.strptime(date_start, "%Y-%m-%d")  # Текущая начальная дата окна
    end_date = datetime.strptime(date_end, "%Y-%m-%d")       # Конечная дата всего периода

    last_call_time = None  # Время последнего успешного запроса (для контроля rate limit)

    # Проходим по всему периоду по недельным окнам (максимум 7 дней)
    while cursor_start <= end_date:
        # Определяем дату конца текущего окна: максимум неделя или до конца периода
        window_end = min(cursor_start + timedelta(days=6), end_date)

        if verbose:
            logger.info(f"\nОкно: {cursor_start.strftime('%Y-%m-%d')} - {window_end.strftime('%Y-%m-%d')}")

        temp_from = cursor_start
        temp_to = window_end

        # Цикл дробления окна, если данных слишком много
        while True:
            rrdid = 0                # Старт пагинации с rrdid=0
            rows_in_window = []      # Собираем строки только за это окно
            last_max_date_to = None  # Максимальная дата для контроля дробления

            # Цикл пагинации по rrdid
            while True:
                # Вызываем функцию выгрузки с передачей last_call_time для контроля пауз
                json_resp, last_call_time = wildberries_get_report_detail(
                    headers,
                    temp_from.strftime("%Y-%m-%d"),
                    temp_to.strftime("%Y-%m-%d"),
                    rrdid=rrdid,
                    last_call_time=last_call_time,
                    period=period
                    )

                rows = json_resp
                if not rows:
                    # Нет данных — выходим из пагинации для этого окна
                    break

                # Дедупликация по ключу (realizationreport_id, rrdId)
                for row in rows:
                    key = (row.get('realizationreport_id'), row.get('rrdId'))
                    if key not in dedup_key_set:
                        dedup_key_set.add(key)
                        all_rows.append(row)
                        rows_in_window.append(row)

                # Обновляем rrdid для следующего запроса по пагинации
                max_rrd = max(row['rrdId'] for row in rows)

                # Запоминаем максимальную дату из выгруженных данных
                last_max_date_to = max(pd.to_datetime(row['dateTo']) for row in rows)

                if verbose:
                    logger.info(f"  rrdid={rrdid}: получено {len(rows)} строк; max_rrd={max_rrd}")

                # Если получили меньше лимита, считаем окно выгруженным
                if len(rows) < 100000:
                    break

                # Иначе продолжаем пагинацию с новым rrdid
                rrdid = max_rrd

            # Если данные за окно меньше лимита — окнo полностью выгружено, переходим дальше
            if len(rows_in_window) < 100000:
                break

            # Если много данных и окно не "сдвинулось" по дате — дробим окно
            date_progress = (last_max_date_to - temp_from).days if last_max_date_to else 0
            window_days = (temp_to - temp_from).days

            if verbose:
                logger.info(
                    f"\n!!! Нужно дробление окна: внутри окна {len(rows_in_window)} строк,\n"
                    f"last_max_date_to={last_max_date_to}, date_progress={date_progress}, window_days={window_days}"
                )

            # Бинарное дробление окна — уменьшаем правую границу вдвое
            new_window_days = window_days // 2
            if new_window_days == 0:
                # Нельзя уменьшить окно — прерываем дробление
                break

            temp_to = temp_from + timedelta(days=new_window_days)

        # Переходим к следующему недельному окну
        cursor_start = window_end + timedelta(days=1)

    # Формируем итоговый DataFrame
    df = pd.DataFrame(all_rows)
    if logger:
        logger.info(f"Всего строк в отчете: {df.shape[0]}")

    return df
# --- Конец выгрузки отчета реализации --- #

# Функция переименования столбцов отчета по реализации на русский
def rename_report_columns(df_realization_report):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_ = df_realization_report.copy()
    # Считывание файла с соответствием колонки из апи её описанию из документации
    api_report_columns = pd.DataFrame(
        list(realization_report_column_map.items()),
        columns=['api_name', 'ru_name']
    )
    # Создаем df с колонками из отчета по апи
    df_realization_report_columns = pd.DataFrame({'api_name': df_realization_report_.columns})
    # Мерджим с колонками отчета, который выгрузили по апи
    all_columns = df_realization_report_columns.merge(
        api_report_columns,
        how='outer',
        on='api_name',
        indicator=True
    )
    # Находим общие колонки
    columns_to_rename = all_columns.loc[all_columns['_merge'] == 'both', ['api_name', 'ru_name']]
    # Переименовываем колонки
    # df_realization_report_ = df_realization_report_.rename(columns=columns_to_rename.set_index('api_name')['ru_name'])
    df_realization_report_ = df_realization_report_.rename(columns=realization_report_column_map)
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
        'Компенсация скидки по программе лояльности',
        'Разовое изменение срока перечисления денежных средств',
    ]
    for col in variable_api_columns:
        if col not in df_realization_report_new_columns.columns:
            df_realization_report_new_columns[col] = 0
    # Переименование некоторых колонок
    df_realization_report_new_columns = df_realization_report_new_columns.rename(columns={
        'Стоимость платной приемки': 'Стоимость платной приёмки'
    })
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

    # Переводим некоторые колонки в числа
    col_to_numbers = [
        'Цена розничная',
        'Сумма продаж (возвратов)',
        'Согласованная скидка',
        'Процент комиссии',
        'Цена розничная с учетом согласованной скидки',
        'Стоимость логистики',
        'Скидка постоянного покупателя',
        'Размер кВВ без НДС, % базовый',
        'Итоговый кВВ без НДС, %', 'Размер снижения кВВ из-за рейтинга',
        'Размер снижения кВВ из-за акции',
        'Вознаграждение с продаж до вычета услуг поверенного, без НДС',
        'К перечислению продавцу за реализованный товар',
        'Возмещение за выдачу и возврат товаров на ПВЗ',
        'Возмещение издержек по эквайрингу',
        'Размер комиссии за эквайринг без НДС, %',
        'Вознаграждение WB без НДС',
        'НДС с вознаграждения WB',
        'Штрафы',
        'Доплаты',
        'Возмещение издержек по перевозке',
        'Стоимость хранения',
        'Прочие удержания/выплаты',
        'Стоимость платной приёмки',
        'Сумма, удержанная за начисленные баллы программы лояльности',
        'Компенсация скидки по программе лояльности',
        'Стоимость участия в программе лояльности',
        'Размер скидки лояльности от продавца, %',
        'Разовое изменение срока перечисления денежных средств',
    ]
    existing_col_to_numbers = [col for col in col_to_numbers if col in df_realization_report_new_columns.columns]
    df_realization_report_new_columns[existing_col_to_numbers] = df_realization_report_new_columns[existing_col_to_numbers].apply(pd.to_numeric, errors='coerce')

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
            'Баркод',
            'Штрих-код',
            'Тип документа',
            'Обоснование для оплаты',
            'Обоснование штрафов и доплат',
            'Цена розничная',
            'Цена розничная с учетом согласованной скидки',
            'Сумма продаж (возвратов)',
            'К перечислению продавцу за реализованный товар',
            'Вознаграждение WB без НДС',
            'НДС с вознаграждения WB',
            'Размер кВВ без НДС, % базовый',
            'Итоговый кВВ без НДС, %',
            'Размер снижения кВВ из-за рейтинга',
            'Размер снижения кВВ из-за акции',
            'Стоимость логистики',
            'Тип платежа за Эквайринг/Комиссии за организацию платежей',
            'Возмещение издержек по эквайрингу',
            'Возмещение за выдачу и возврат товаров на ПВЗ',
            'Возмещение издержек по перевозке',
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

    # Приводим размер к строковому типу с нормализацией чисел (36.0 -> '36', 43.5 -> '43.5')
    df_realization_report_new_columns['Размер'] = normalize_size_series(df_realization_report_new_columns['Размер'])

    # Считаем НДС в долях
    df_realization_report_new_columns['НДС'] = (
        df_realization_report_new_columns['НДС с вознаграждения WB'] / df_realization_report_new_columns['Вознаграждение WB без НДС']
    )

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
        'Доплаты',
        'Разовое изменение срока перечисления денежных средств',
    ]

    # Итоговая колонка с прочими расходами
    df_realization_report_new_columns['Итого руб'] = df_realization_report_new_columns.loc[:, other_costs_columns].sum(axis=1)

    return df_realization_report_new_columns

# Функция исключения переходных недель из отчета
def filter_report_dates(date_start, date_end, df_realization_report_new_columns, period='daily', filter_dates=False):
    if period not in ['daily', 'weekly']:
        logger.error(f"Некорректное значение периода для фильтрации: {period}. Допустимые значения: 'daily' или 'weekly'.")
        raise ValueError(f"Некорректное значение периода для фильтрации: {period}. Допустимые значения: 'daily' или 'weekly'.")
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_date_filtered = df_realization_report_new_columns.copy()
    # Если указан флаг фильтра по датам, то делаем фильтр по датам
    if filter_dates:
        # Переводим даты выгрузки в Timestamp
        dt_start = pd.to_datetime(date_start, format='mixed')
        dt_end = pd.to_datetime(date_end, format='mixed')

        if period == 'weekly':
            logger.info("Фильтруем отчет реализации по переходным неделям")
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

        elif period == 'daily':
            logger.info("Фильтруем отчет реализации по дате продажи в пределах указанного диапазона")
            # Делаем фильтр по дате продажи в пределах указанного диапазона
            df_realization_report_date_filtered = (
                df_realization_report_date_filtered
                .loc[
                    (df_realization_report_date_filtered['Дата продажи'] >= dt_start)
                    & (df_realization_report_date_filtered['Дата продажи'] <= dt_end),
                    :
                ]
                .reset_index(drop=True)
            )



    return df_realization_report_date_filtered

# Функция добавления размера товара в отчет реализации
def add_sizes_to_realization_report(df_realization_report_date_filtered, df_products):
    # Создаем копию для избежания изменений в оригинальном df
    df_realization_report_with_sizes = df_realization_report_date_filtered.copy()
    df_products_barcodes = df_products.copy()

    # Преобразуем sku_list из строки вида "['1234', '5678']" в настоящий список
    df_products_barcodes['sku_list'] = df_products_barcodes['sku_list'].apply(
        lambda x: ast.literal_eval(x) if isinstance(x, str) else x
    )

    # Распаковываем список баркодов
    df_products_barcodes = (
        df_products_barcodes
        .explode('sku_list')
        .rename(columns={
            'sku_list': 'Баркод',
            'Размер': 'Размер (список товаров)'
        })
        .reset_index(drop=True)
    )
    # Приводим размер из списка товаров к строковому типу с нормализацией чисел (36.0 -> '36', 43.5 -> '43.5')
    df_products_barcodes['Размер (список товаров)'] = (
        normalize_size_series(df_products_barcodes['Размер (список товаров)'])
    )

    # Переводим баркод в строку
    df_products_barcodes['Баркод'] = (
        df_products_barcodes['Баркод']
        .astype(str)
    )

    df_realization_report_with_sizes['Баркод'] = (
        pd.to_numeric(
            df_realization_report_with_sizes['Баркод'], errors='coerce'
        )
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # Объединяем df с размерами и баркодами с отчетом реализации по баркоду
    df_realization_report_with_sizes = (
        df_realization_report_with_sizes
        .merge(
            df_products_barcodes[['Баркод', 'Размер (список товаров)']],
            on='Баркод',
            how='left',
            indicator='Наличие в списке товаров с размерами'
        )
        .reset_index(drop=True)
    )

    # Там, где в отчете реализации размер 0 или пропуск, берем размер из списка товаров, если он там есть
    mask_sizes = (
        (df_realization_report_with_sizes['Размер'].isna() |
         (df_realization_report_with_sizes['Размер'] == '0'))
        & (df_realization_report_with_sizes['Наличие в списке товаров с размерами'] == 'both')
    )

    # Логируем, сколько строк будет дополнено
    filled_count = mask_sizes.sum()
    total_rows = len(df_realization_report_with_sizes)
    percentage = (filled_count / total_rows * 100) if total_rows else 0
    logger.info(
        "Размеры дополнены из списка товаров для {} строк ({:.2f}% от {})",
        filled_count, percentage, total_rows
    )

    df_realization_report_with_sizes['Размер дополненный'] = np.where(
        mask_sizes,
        df_realization_report_with_sizes['Размер (список товаров)'],
        df_realization_report_with_sizes['Размер']
    )

    # Удаляем вспомогательные колонки
    df_realization_report_with_sizes = df_realization_report_with_sizes.drop(
        columns=[
            'Размер (список товаров)',
            'Наличие в списке товаров с размерами'
        ]
    )

    # Переименовываем дополненную колонку с размерами
    df_realization_report_with_sizes = df_realization_report_with_sizes.rename(columns={
        'Размер': 'Размер (исходный)',
        'Размер дополненный': 'Размер'
    })

    # Перемещаем колонки в начало df
    df_realization_report_with_sizes = move_columns(
        df_realization_report_with_sizes,
        ['Размер (исходный)', 'Размер'],
        position='Артикул продавца'
    )

    return df_realization_report_with_sizes

# Функция выгрузки отчета о реализации
def get_realization_report(headers, date_start, date_end, period='daily', verbose=True, df_products=None):
    # Переводим даты в нужный формат
    date_start_ = datetime.fromisoformat(date_start).strftime('%Y-%m-%d')
    date_end_ = datetime.fromisoformat(date_end).strftime('%Y-%m-%d')
    # Получение отчета о реализации по АПИ
    df_realization_report = upload_realization_report(headers, date_start_, date_end_, period=period, verbose=verbose)
    # Переименование колонок отчета о реализации в русские названия
    df_realization_report_renamed = rename_report_columns(df_realization_report)
    # Расчет некоторых дополнительных колонок, которые будут участвовать в расчетах
    df_realization_report_new_columns = add_new_columns(df_realization_report_renamed)
    # Фильтр по переходным неделям
    df_realization_report_date_filtered = filter_report_dates(
        date_start_, date_end_,
        df_realization_report_new_columns,
        filter_dates=True,
        period=period
    )
    # Выгружаем список товаров WB, если не передан
    if df_products is None:
        from wb.scripts.uploadDataFromWB import get_wb_product
        df_products = get_wb_product(headers, to_save=False)
    # Добавляем размеры в отчет реализации
    df_realization_report_with_sizes = add_sizes_to_realization_report(
        df_realization_report_date_filtered,
        df_products
    )

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


# Функция перевода столбцов в другую валюту
def convert_currency(
    df,
    columns_to_convert=None,
    custom_columns_to_ignore=None,
    k_exchange_rate=1.0,
    reverse=False
):
    """
    Переводит значения указанных столбцов между валютой и рублями по единому курсу.

    Параметры:
    ----------
    df : pd.DataFrame
        Исходный DataFrame, в котором требуется конвертация валюты.

    columns_to_convert : list of str, optional
        Список названий столбцов, которые нужно конвертировать.
        Если не указан или пуст — конвертируются все столбцы, кроме игнорируемых.

    custom_columns_to_ignore : list of str, optional
        Список столбцов, которые НЕ нужно конвертировать.
        Если не указан — используется дефолтный список.

    k_exchange_rate : float, default=1.0
        Курс валюты для конвертации. Один курс применяется для всех выбранных столбцов.
        reverse=False: перевод из валюты в рубли (умножение).
        reverse=True: перевод из рублей в валюту (деление).

    reverse : bool, default=False
        reverse=False: перевод из валюты в рубли (умножение).
        reverse=True: перевод из рублей в валюту (деление).

    Возвращает:
    -----------
    pd.DataFrame
        Новый DataFrame с преобразованными значениями в заданных столбцах.
    """

    # Дефолтный игнорируемый список
    default_ignore = [
        'id', 'Артикул WB', 'Артикул продавца', 'Наименование', 'Размер',
        'Дата', 'Тип документа', 'Обоснование для оплаты',
        'Обоснование штрафов и доплат', 'Страна производства',
        'Продажи шт',
        'Возвраты шт',
        'Продажи минус возвраты шт',

    ]

    if custom_columns_to_ignore is None:
        custom_columns_to_ignore = default_ignore

    # Если columns_to_convert не передан или пуст — берем все числовые, не в ignore
    if not columns_to_convert:
        all_num_cols = [col for col in df.select_dtypes(include='number').columns
                        if col not in custom_columns_to_ignore]
        columns = all_num_cols
    else:
        # Если columns_to_convert есть — только из него, пропуская ignore
        columns = [col for col in columns_to_convert if col not in custom_columns_to_ignore]

    df_result = df.copy()
    for col in columns:
        if reverse:
            df_result[col] = df_result[col] / k_exchange_rate
        else:
            df_result[col] = df_result[col] * k_exchange_rate

    return df_result


# Функция выгрузки воронки продаж
def get_sales_funnel(
    headers: dict,
    nm_ids: list,
    date_from: str,
    date_to: str,
    limit: int = 50
) -> pd.DataFrame:
    """
    Выгружает все данные из метода Воронка продаж WB c поддержкой пагинации и лимитами WB.
    """
    url = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"
    logger.info("Старт выгрузки данных по воронке продаж WB с пагинацией и лимитами")
    all_products = []
    offset = 0
    burst_limit = 3
    burst_period = 60  # секунд
    request_timestamps = []

    while True:
        payload = {
            "selectedPeriod": {
                "start": date_from,
                "end": date_to
            },
            "limit": limit,
            "offset": offset
        }
        if nm_ids:
            payload["nmIds"] = nm_ids

        # Контроль лимитов
        now = time.time()
        request_timestamps = [ts for ts in request_timestamps if now - ts < burst_period]
        if len(request_timestamps) >= burst_limit:
            wait_time = burst_period - (now - request_timestamps[0]) + 0.1
            logger.info(f"Лимит запросов исчерпан, ожидаем {wait_time:.1f} секунд...")
            time.sleep(wait_time)
            continue

        # logger.debug(f"POST {url}, payload: {payload}")
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            if response.status_code == 429:
                logger.info("Превышен лимит запросов! Ждем 21 секунду и пробуем снова.")
                time.sleep(21)
                continue
            response.raise_for_status()
            resp_json = response.json()
            products = resp_json.get("data", {}).get("products", [])
            request_timestamps.append(time.time())
        except Exception as e:
            logger.exception(f"Ошибка при выполнении запроса: {e}")
            raise

        logger.info(f"Получено {len(products)} товаров, offset={offset}")
        all_products.extend(products)

        # Если получено меньше, чем limit — достигли конца списка
        if len(products) < limit:
            break
        offset += limit

    # Сборка данных по нужным полям
    rows = []
    for prod in all_products:
        product = prod.get("product", {})
        stats = prod.get("statistic", {}).get("selected", {})
        rows.append({
            "Артикул WB": product.get("nmId"),
            "Артикул продавца": product.get("vendorCode"),
            "Название карточки товара": product.get("title"),
            "Количество переходов в карточку товара": stats.get("openCount"),
            "Положили в корзину, шт.": stats.get("cartCount"),
            "Заказали товаров, шт.": stats.get("orderCount"),
            "Заказали на сумму": stats.get("orderSum"),
            "Выкупили товаров, шт.": stats.get("buyoutCount"),
            "Выкупили на сумму": stats.get("buyoutSum"),
            "Отменили товаров, шт.": stats.get("cancelCount"),
            "Отменили на сумму": stats.get("cancelSum"),
            "Средняя цена": stats.get("avgPrice"),
        })
    df = pd.DataFrame(rows)

    return df

# Функция получения остатков с опциональной группировкой
def get_reminders_by_product(
    df_products: pd.DataFrame,
    headers: dict,
    agg_col: list[str] | None = ['Артикул продавца', 'Артикул WB', 'Размер'],
    split_stocks: bool = False,
    date_start: str = "2000-01-01",  # для совместимости с сигнатурой getStockRemindersWB
    date_end: str = "2000-01-01",    # для совместимости
) -> pd.DataFrame:
    """
    Выгружает и обрабатывает остатки товаров Wildberries (FBO и FBS),
    маппя их на товары строго по chrtId.

    Ожидается, что df_products содержит как минимум колонки:
    - 'Артикул продавца'
    - 'Артикул WB'
    - 'Размер'
    - 'chrtId' или 'chrtID'

    Логика:
    1. Получаем FBO-остатки через getStockRemindersWB, берём поле quantity.
    2. Получаем FBS-остатки через getStockReminders_fbs, берём поле amount.
    3. Оба набора остатков агрегируем по chrtId.
    4. Мерджим остатки к df_products ТОЛЬКО по chrtId.
       Вся информация о товаре (Артикул продавца, Артикул WB, Размер)
       берётся исключительно из df_products.
    5. Если agg_col не None и не пустой — агрегируем по этим колонкам.

    Итоговый DataFrame (после возможной агрегации) содержит колонки:
    - 'Артикул продавца'
    - 'Артикул WB'
    - 'Размер'
    - 'Остатки FBO'
    - 'Остатки FBS'
    (+ 'Остаток', если split_stocks == False)
    """
    # Импорт WB-функций
    from wb.scripts.uploadDataFromWB import getStockRemindersWB, getStockReminders_fbs

    df_products = df_products.copy()

    # ---------- 0. Подготовка df_products ----------
    # Приводим chrtId/chrtID к единому имени
    if "chrtId" not in df_products.columns and "chrtID" in df_products.columns:
        df_products = df_products.rename(columns={"chrtID": "chrtId"})

    required_cols = ["Артикул продавца", "Артикул WB", "Размер", "chrtId"]
    missing = [c for c in required_cols if c not in df_products.columns]
    if missing:
        raise ValueError(
            f"В df_products отсутствуют необходимые колонки: {missing}. "
            "Ожидаются: 'Артикул продавца', 'Артикул WB', 'Размер', 'chrtId' (или 'chrtID')."
        )

    df_products["chrtId"] = df_products["chrtId"].astype(str)

    # ---------- 1. Остатки FBO ----------
    df_fbo = getStockRemindersWB(
        headers=headers,
        date_start=date_start,
        date_end=date_end,
        to_save=False,
    )

    if df_fbo.empty:
        df_fbo = pd.DataFrame(columns=["chrtId", "Остатки FBO"])
    else:
        # Нормализуем chrtId
        if "chrtID" in df_fbo.columns and "chrtId" not in df_fbo.columns:
            df_fbo = df_fbo.rename(columns={"chrtID": "chrtId"})
        if "chrtId" in df_fbo.columns:
            df_fbo["chrtId"] = df_fbo["chrtId"].astype(str)
        # quantity → Остатки FBO
        if "quantity" in df_fbo.columns:
            df_fbo = df_fbo.rename(columns={"quantity": "Остатки FBO"})
        else:
            df_fbo["Остатки FBO"] = 0
        # Оставляем только chrtId + остаток и агрегируем по chrtId
        df_fbo = (
            df_fbo.loc[:, ["chrtId", "Остатки FBO"]]
            .groupby("chrtId", as_index=False)["Остатки FBO"]
            .sum()
        )

    # ---------- 2. Остатки FBS ----------
    df_fbs = getStockReminders_fbs(headers=headers, df_products=df_products, to_save=False)

    if df_fbs.empty:
        df_fbs = pd.DataFrame(columns=["chrtId", "Остатки FBS"])
    else:
        # Нормализуем chrtId
        if "chrtID" in df_fbs.columns and "chrtId" not in df_fbs.columns:
            df_fbs = df_fbs.rename(columns={"chrtID": "chrtId"})
        if "chrtId" in df_fbs.columns:
            df_fbs["chrtId"] = df_fbs["chrtId"].astype(str)
        # amount → Остатки FBS
        if "amount" in df_fbs.columns:
            df_fbs = df_fbs.rename(columns={"amount": "Остатки FBS"})
        else:
            df_fbs["Остатки FBS"] = 0
        df_fbs = (
            df_fbs.loc[:, ["chrtId", "Остатки FBS"]]
            .groupby("chrtId", as_index=False)["Остатки FBS"]
            .sum()
        )

    # ---------- 3. Мердж остатков с df_products по chrtId ----------
    df_with_stocks = (
        df_products
        .merge(df_fbo, on="chrtId", how="left")
        .merge(df_fbs, on="chrtId", how="left")
    )

    # Оставляем только нужные поля
    result_cols = ["Артикул продавца", "Артикул WB", "Размер", "Остатки FBO", "Остатки FBS"]
    for col in ["Остатки FBO", "Остатки FBS"]:
        if col not in df_with_stocks.columns:
            df_with_stocks[col] = 0

    df_result = df_with_stocks[result_cols].copy()

    # Приводим остатки к числам
    for col in ["Остатки FBO", "Остатки FBS"]:
        df_result[col] = pd.to_numeric(df_result[col], errors="coerce").fillna(0)

    # ---------- 4. Агрегация (если указан agg_col) ----------
    # Если agg_col не передан (None или пустой список) — НЕ агрегируем
    if agg_col:
        # Разрешаем агрегировать только по этим трём полям
        allowed_group_cols = ["Артикул продавца", "Артикул WB", "Размер"]
        group_cols = [c for c in agg_col if c in allowed_group_cols]
        if not group_cols:
            group_cols = allowed_group_cols  # запасной вариант, если передали что-то не то

        df_agg = (
            df_result
            .groupby(group_cols, as_index=False)[["Остатки FBO", "Остатки FBS"]]
            .sum()
        )

        if not split_stocks:
            df_agg = df_agg.assign(
                Остаток=lambda d: d[["Остатки FBO", "Остатки FBS"]].sum(axis=1)
            )
        return df_agg

    # Без агрегации — по строкам df_products
    if not split_stocks:
        df_result = df_result.assign(
            Остаток=lambda d: d[["Остатки FBO", "Остатки FBS"]].sum(axis=1)
        )

    return df_result

class WbRateLimiterMoreLimit:
    """
    Универсальный лимитер для методов WB API с учётом:
      - периода и лимита (для расчёта интервала),
      - всплеска (burst),
      - заголовков X-Ratelimit-Remaining / X-Ratelimit-Retry / X-Ratelimit-Reset.

    Параметры:
      period_seconds  – длительность окна (например, 60 секунд),
      period_limit    – макс. количество запросов за окно,
      max_burst       – теоретический burst по тарифу (из таблицы WB).
                        Нужен лишь как «верхняя планка», реальный burst
                        контролируется X-Ratelimit-Remaining.
    """

    def __init__(self, period_seconds, period_limit, max_burst):
        self.period_seconds = period_seconds
        self.period_limit = period_limit
        self.max_burst = max_burst

        # Интервал, который WB рекомендует для ровного распределения
        self.interval = float(period_seconds) / float(period_limit)

        # Простое окно по времени (для подстраховки)
        self._window_start_ts = None
        self._requests_in_window = 0

    def _wait_for_window_if_needed(self):
        """Грубый контроль по количеству запросов за период (подстраховка)."""
        now = time.time()

        if self._window_start_ts is None:
            self._window_start_ts = now
            self._requests_in_window = 0
            return

        if now - self._window_start_ts >= self.period_seconds:
            # Новое окно
            logger.debug("Новое окно лимита WB. Сбрасываю счётчик запросов.")
            self._window_start_ts = now
            self._requests_in_window = 0
            return

        if self._requests_in_window >= self.period_limit:
            sleep_time = self.period_seconds - (now - self._window_start_ts)
            if sleep_time > 0:
                logger.warning(
                    f"Достигнут лимит {self.period_limit} запросов за {self.period_seconds} секунд. "
                    f"Пауза {sleep_time:.2f} сек (по окну)."
                )
                time.sleep(sleep_time)
            # новое окно
            self._window_start_ts = time.time()
            self._requests_in_window = 0

    def call(self, func, max_retries=3, *args, **kwargs):
        """
        Вызов произвольной функции func(*args, **kwargs) с учётом лимита WB.
        Ожидается, что func возвращает объект requests.Response.

        Возвращает (response, success_flag).
        """
        for attempt in range(1, max_retries + 1):
            # Подстраховка по количеству запросов за период
            self._wait_for_window_if_needed()

            try:
                resp = func(*args, **kwargs)
            except Exception as e:
                func_name = getattr(func, "__name__", "request")
                logger.error(
                    f"Ошибка соединения при вызове {func_name} "
                    f"(попытка {attempt}/{max_retries}): {e}"
                )
                if attempt == max_retries:
                    return None, False
                time.sleep(1)
                continue

            # Обновляем счётчик окна
            now = time.time()
            if self._window_start_ts is None:
                self._window_start_ts = now
                self._requests_in_window = 0
            self._requests_in_window += 1

            status = resp.status_code
            headers = resp.headers or {}

            # Читаем заголовки WB
            rl_remaining = headers.get("X-Ratelimit-Remaining")
            rl_retry = headers.get("X-Ratelimit-Retry")
            rl_reset = headers.get("X-Ratelimit-Reset")
            rl_limit_burst = headers.get("X-Ratelimit-Limit")

            # 429 — слишком часто, WB сам говорит, когда можно повторить
            if status == 429:
                retry_after = None
                if rl_retry is not None:
                    try:
                        retry_after = float(rl_retry)
                    except Exception:
                        retry_after = None

                if retry_after is None:
                    # Нет явного X-Ratelimit-Retry — ждём консервативно 2 сек
                    retry_after = 2.0

                logger.warning(
                    f"Получен 429 Too Many Requests. "
                    f"X-Ratelimit-Retry={rl_retry}, X-Ratelimit-Reset={rl_reset}, "
                    f"X-Ratelimit-Limit={rl_limit_burst}. "
                    f"Ждём {retry_after:.2f} сек перед повтором."
                )
                time.sleep(retry_after)

                if attempt == max_retries:
                    return resp, False
                continue

            # Не 429 — значит, запрос принят. Дальше решаем, делать ли паузу
            # по интервалу (если rl_remaining уже 0).
            # rl_remaining показывает, сколько запросов ещё можно сделать БЕЗ пауз.
            need_interval_sleep = False

            if rl_remaining is not None:
                try:
                    rl_remaining_val = int(rl_remaining)
                except Exception:
                    rl_remaining_val = None
            else:
                rl_remaining_val = None

            if rl_remaining_val is not None:
                # Если всплеск разрешён (remaining > 0) — можно следующий запрос без паузы.
                # Если 0 — дальше запрос без паузы приведёт к 429,
                # поэтому подождём примерно интервал.
                if rl_remaining_val <= 0:
                    need_interval_sleep = True
            else:
                # WB не прислал remaining — ориентируемся только на интервал
                need_interval_sleep = True

            if need_interval_sleep:
                # Интервал — рекомендация для ровного распределения.
                # Чтобы не спамить, выдерживаем не меньше interval.
                logger.debug(
                    f"X-Ratelimit-Remaining={rl_remaining}; выдерживаем интервал {self.interval:.3f} сек."
                )
                time.sleep(self.interval)

            return resp, True

        return None, False

class WbRateLimiter:
    """
    Универсальный лимитер для запросов к API Wildberries.

    Логика работы основана на официальной модели лимитов WB:
    https://openapi.wildberries.ru/

    Основные понятия:
    ------------------
    - Период (period_seconds):
        Временное окно, в течение которого можно выполнить не более period_limit запросов.
        Для WB обычно 60 секунд.

    - Лимит (period_limit):
        Максимальное количество запросов за период. Например:
          * персональный / сервисный токен: 300 запросов / 60 сек
          * базовый / тестовый токен:      150 запросов / 60 сек

        Исходя из этого считается рекомендуемый интервал:
            interval = period_seconds / period_limit

    - Интервал (interval):
        Рекомендуемая пауза между запросами при равномерной нагрузке.
        Например:
          60 сек / 300 запросов = 0.2 сек.

        Лимитер использует этот интервал как fallback:
          * когда WB говорит, что разрешённый всплеск исчерпан (X-Ratelimit-Remaining=0);
          * или когда WB не прислал заголовок X-Ratelimit-Remaining.

    - Всплеск (burst, max_burst и X-Ratelimit-Remaining):
        Позволяет отправить несколько запросов подряд БЕЗ интервальной паузы.
        В документации WB теоретический burst указан в тарифе (20 или 10),
        а фактический остаток всплеска отдаётся в заголовке:

          X-Ratelimit-Remaining — сколько запросов ещё можно сделать без пауз.

        Лимитер не считает burst сам, а доверяет заголовку:
          * если X-Ratelimit-Remaining > 0 — следующий запрос можно слать без ожидания;
          * если X-Ratelimit-Remaining == 0 — дальнейший запрос без паузы с большой
            вероятностью приведёт к 429, поэтому мы выдерживаем интервал.

    - Ошибка 429 Too Many Requests:
        Если вы превысили мгновенную скорость запросов, WB вернёт 429 и укажет,
        когда можно повторить запрос:

          * X-Ratelimit-Retry — через сколько секунд можно повторить запрос;
          * X-Ratelimit-Reset — через сколько секунд burst восстановится;
          * X-Ratelimit-Limit — максимальный burst.

        Лимитер при 429:
          * читает X-Ratelimit-Retry (если есть);
          * ждёт указанное количество секунд (или interval, если заголовка нет);
          * повторяет запрос до max_retries раз.

    Параметры конструктора:
    -----------------------
    :param period_seconds: int
        Длительность окна лимита в секундах (обычно 60).

    :param period_limit: int
        Максимальное количество запросов за period_seconds.
        Примеры:
          * персональный токен: 300
          * базовый токен: 150
          * метод с жёстким ограничением: 3 и т.п.

        Используется только для расчёта interval = period_seconds / period_limit.

    :param max_burst: int
        Теоретический burst по тарифу (20 или 10).
        Нужен как верхняя планка "разумного" всплеска, но реальный остаток всплеска
        контролируется через X-Ratelimit-Remaining.

    Методы:
    -------
    call(func, max_retries=3, skip_sleep=False, *args, **kwargs)

        Обёртка над вызовом HTTP‑клиента (например, requests.Session.get / post)
        с учётом лимитов WB.

        Ожидается, что func(*args, **kwargs) возвращает объект requests.Response.

        Параметры:
          :param func: callable
              Функция/метод, выполняющий запрос, например session.get или session.post.

          :param max_retries: int
              Сколько раз повторять запрос при временных ошибках (в том числе 429).

          :param skip_sleep: bool
              Если True, лимитер НЕ будет делать финальную паузу по интервалу после
              успешного запроса, даже если X-Ratelimit-Remaining == 0.
              Полезно для последнего запроса серии, чтобы не ждать зря.

          :param *args, **kwargs:
              Аргументы, которые будут переданы в func. Например:
                  limiter.call(session.get, url=..., params=...)
                  limiter.call(session.post, url=..., json=...)

        Возвращает:
          (response, success_flag), где:
            * response — объект requests.Response или None при фатальной ошибке;
            * success_flag — bool, True если запрос успешно выполнен (200, 201 и т.п.,
              но не 429 после исчерпания max_retries).

    Примеры использования:
    ----------------------

    1) Лимитер для персонального токена (300 запросов в минуту, burst=20):

        import requests

        headers = {"Authorization": "Bearer <TOKEN>"}
        session = requests.Session()
        session.headers.update(headers)

        limiter = WbRateLimiter(
            period_seconds=60,
            period_limit=300,  # 300 запросов в минуту
            max_burst=20       # теоретический burst
        )

        # GET запрос с учётом лимита:
        resp, ok = limiter.call(
            session.get,
            url="https://advert-api.wildberries.ru/adv/v3/campaigns"
        )
        if ok and resp.status_code == 200:
            data = resp.json()

    2) Лимитер для метода с жёстким лимитом 3 запроса/минуту без bursts:

        limiter = WbRateLimiter(
            period_seconds=60,
            period_limit=3,  # не больше 3 запросов за минуту
            max_burst=1      # всплеск фактически отсутствует
        )

        # В цикле:
        for params in list_of_params:
            resp, ok = limiter.call(
                session.get,
                url="https://advert-api.wildberries.ru/adv/v3/fullstats",
                params=params
            )
            if not ok:
                # Обработка ошибки / логирование
                continue
            # Обработка ответа resp

    3) Пример использования skip_sleep для последнего запроса:

        # Пусть мы знаем, что это последний запрос в серии:
        is_last_request = True

        resp, ok = limiter.call(
            session.post,
            url="https://advert-api.wildberries.ru/adv/v0/normquery/stats",
            json=payload,
            skip_sleep=is_last_request,  # не ждать интервал после последнего запроса
        )

    Детали поведения:
    -----------------
    - При успешном запросе (не 429):
        * Лимитер читает X-Ratelimit-Remaining:
            - если значение > 0 и skip_sleep=False — следующий запрос можно отправить сразу;
            - если значение == 0 и skip_sleep=False — лимитер выдержит паузу interval
              перед тем, как разрешить следующий запрос (иначе возможен 429);
            - если skip_sleep=True — финальная пауза не делается, даже при remaining=0.

        * Если заголовок X-Ratelimit-Remaining отсутствует, лимитер использует interval
          как «мягкую» рекомендацию и выдерживает паузу, если skip_sleep=False.

    - При 429 Too Many Requests:
        * Лимитер считывает X-Ratelimit-Retry и ждёт указанное количество секунд.
          Если заголовка нет — ждёт interval.
        * Пытается повторить запрос до max_retries раз, после чего вернёт (resp, False).

    - Остальные коды ответа (4xx/5xx, кроме 429) передаются наружу как есть.
    """

    def __init__(self, period_seconds, period_limit, max_burst):
        self.period_seconds = period_seconds
        self.period_limit = period_limit
        self.max_burst = max_burst
        # Интервал для равномерного распределения запросов
        self.interval = float(period_seconds) / float(period_limit)

    def call(self, func, max_retries=3, skip_sleep=False, *args, **kwargs):
        """
        Вызов произвольной функции func(*args, **kwargs) с учётом лимита WB.

        Ожидается, что func возвращает объект requests.Response.

        :param func: callable
            HTTP‑функция (например, session.get или session.post).

        :param max_retries: int
            Максимальное количество попыток при временных ошибках (включая 429).

        :param skip_sleep: bool
            Если True, не делать финальную паузу по интервалу после успешного
            запроса (удобно для последнего запроса в серии).

        :param *args, **kwargs:
            Аргументы, которые будут переданы в func.

        :return: (response, success_flag)
        """
        for attempt in range(1, max_retries + 1):
            try:
                resp = func(*args, **kwargs)
            except Exception as e:
                func_name = getattr(func, "__name__", "request")
                logger.error(
                    f"Ошибка соединения при вызове {func_name} "
                    f"(попытка {attempt}/{max_retries}): {e}"
                )
                if attempt == max_retries:
                    return None, False
                time.sleep(1)
                continue

            status = resp.status_code
            headers = resp.headers or {}

            rl_remaining = headers.get("X-Ratelimit-Remaining")
            rl_retry = headers.get("X-Ratelimit-Retry")
            rl_reset = headers.get("X-Ratelimit-Reset")
            rl_limit_burst = headers.get("X-Ratelimit-Limit")

            if status == 429:
                # WB сам говорит, когда можно повторить
                retry_after = None
                if rl_retry is not None:
                    try:
                        retry_after = float(rl_retry)
                    except Exception:
                        retry_after = None

                if retry_after is None:
                    retry_after = self.interval  # консервативно ждём хотя бы интервал

                logger.warning(
                    "Получен 429 Too Many Requests. "
                    f"X-Ratelimit-Retry={rl_retry}, X-Ratelimit-Reset={rl_reset}, "
                    f"X-Ratelimit-Limit={rl_limit_burst}. "
                    f"Ждём {retry_after:.2f} сек перед повтором."
                )
                time.sleep(retry_after)

                if attempt == max_retries:
                    return resp, False
                continue

            # Не 429 — запрос принят. Смотрим на remaining
            need_interval_sleep = False
            rl_remaining_val = None

            if rl_remaining is not None:
                try:
                    rl_remaining_val = int(rl_remaining)
                except Exception:
                    rl_remaining_val = None

            if rl_remaining_val is not None:
                # Если remaining == 0, дальше без паузы будет 429.
                if rl_remaining_val <= 0:
                    need_interval_sleep = True
            else:
                # WB не прислал remaining — ориентируемся на интервал как на рекомендацию
                need_interval_sleep = True

            if need_interval_sleep and not skip_sleep:
                logger.debug(
                    f"X-Ratelimit-Remaining={rl_remaining}; выдерживаем интервал {self.interval:.3f} сек."
                )
                time.sleep(self.interval)

            return resp, True

        return None, False
