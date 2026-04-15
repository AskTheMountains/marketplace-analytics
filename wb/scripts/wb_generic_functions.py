# %% Определение функций
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

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent


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
    last_call_time=None, max_retries=5
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
    Возвращает:
      tuple: (response_data, время успешного запроса).
    """

    url = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod'
    params = {
        'dateFrom': date_from,
        'dateTo': date_to,
        'rrdid': rrdid,
        'limit': limit,
        'period': 'weekly',
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
            response = requests.get(url, headers=headers, params=params)
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

## Основная логика (обработка недель, rrd_id и бинарный сплит)
def upload_realization_report(headers, date_start, date_end, verbose=False):
    """
    Выгружает финансовый отчет Wildberries по частям (по недельным окнам) с контролем лимита
    1 запрос в минуту.

    Аргументы:
      headers (dict): Заголовки запроса с токеном.
      date_start (str): Дата начала ('YYYY-MM-DD').
      date_end (str): Дата конца ('YYYY-MM-DD').
      verbose (bool): Выводить ли подробный лог.

    Возвращает:
      pd.DataFrame со всеми уникальными строками за указанный период.
    """

    all_rows = []             # Список для сбора всех строк отчета
    dedup_key_set = set()     # Для дедупликации по (realizationreport_id, rrd_id)

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
                    last_call_time=last_call_time)

                rows = json_resp
                if not rows:
                    # Нет данных — выходим из пагинации для этого окна
                    break

                # Дедупликация по ключу (realizationreport_id, rrd_id)
                for row in rows:
                    key = (row.get('realizationreport_id'), row.get('rrd_id'))
                    if key not in dedup_key_set:
                        dedup_key_set.add(key)
                        all_rows.append(row)
                        rows_in_window.append(row)

                # Обновляем rrdid для следующего запроса по пагинации
                max_rrd = max(row['rrd_id'] for row in rows)

                # Запоминаем максимальную дату из выгруженных данных
                last_max_date_to = max(pd.to_datetime(row['date_to']) for row in rows)

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
# %%
