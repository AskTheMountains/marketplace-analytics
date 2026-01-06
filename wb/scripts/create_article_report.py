# %% Определение всех функций
import argparse
import re
from turtle import st
from unittest import result
import warnings
from collections import defaultdict, deque
from datetime import date, datetime, timedelta
from pathlib import Path
import requests
import time

import pandas as pd
import numpy as np
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

warnings.simplefilter(action='ignore', category=FutureWarning)
pd.set_option('future.no_silent_downcasting', True)


# Словарь русских месяцев
months_ru = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь',
    7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}



# Некоторые вспомогательные функции
from generic_functions import move_columns
# Парсинг аргументов командной строки
from wb.scripts.command_args_parser import cli_article_and_companies
# Файл с настройками и номером клиента
from wb.scripts.constants import (
    client_name,
    headers,
    ARTICLE_REPORT_LIST,
    marketplace_dir_path
)
# Функции выгрузки данных по API WB
from wb.scripts.uploadDataFromWB import getOrdersWB as get_orders
from wb.scripts.uploadDataFromWB import getWBProduct

# С какой строки идет запись таблицы
ROW_START_EXCEL = 15
# COL_START_EXCEL = 1


# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Create promotion report (WB)",
    usage=(
        "python.exe -m wb.scripts.create_campaigns_report [start_date] [end_date] "
        "[--start_date YYYY-MM-DD --end_date YYYY-MM-DD --reference_date YYYY-MM-DD] "
        "[-c N | --client N] "
        "[--promotion_companies \"ID1, ID2, ...\"] "
        "[--wb_skus \"123456, 234567, ...\"]"
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
        "  python.exe -m wb.scripts.create_campaigns_report --wb_skus \"123456, 234567, 345678\"\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --wb_skus \"123456; 234567 345678\"\n\n"
        "  python.exe -m wb.scripts.create_campaigns_report --wb_skus \"123456\\n234567\\n345678\"\n\n"
        "Можно комбинировать любые разделители для обоих списков: запятая, пробел, точка с запятой, перевод строки.\n"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


# Функция создания директории, куда будет помещен итоговый отчет
def create_article_svod_dir(client_name, date_report=str(date.today())):
    articles_svod_path = (
        marketplace_dir_path
        / "Clients"
        / client_name
        / "ArticleReport"
    )
    if not articles_svod_path.exists():
        articles_svod_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"Creating Articles Report directory:{articles_svod_path}")

    return articles_svod_path


# Функция парсинга аргументов командной строки
def parse_cli_companies(cli_params, default, client_name):
    """
    Аргументы:
        cli_params — dict, ключи: 'promotion_companies', 'wb_skus', 'start_dates', 'end_dates', 'start_date', 'end_date'
        default — словарь вида {client_name: {кампания: {...}}}
        client_name — строка

    Возвращает:
        - Словарь {компания: {'skus': ..., 'start_date': ..., 'end_date': ...}, ...}
        - Или {'start_date': ..., 'end_date': ...}
        - Или пустой словарь.
    """
    companies_str  = cli_params.get('promotion_companies')
    skus_str       = cli_params.get('wb_skus')
    multi_starts   = cli_params.get('start_dates')    # блок с РК
    multi_ends     = cli_params.get('end_dates')      # блок с РК
    single_start   = cli_params.get('start_date')     # общий блок дат
    single_end     = cli_params.get('end_date')

    logger.info(f'\nПараметры для клиента {client_name}:\n{cli_params}')

    # --- Обработка по бизнес-логике ---

    # 2. Если есть блок кампаний и заданы общие даты — игнорируем multi_dates, применяем single_start/single_end к каждой кампании
    if companies_str and skus_str and single_start and single_end:
        companies = [int(x) for x in re.split(r'[\s,]+', companies_str.strip()) if x]
        matches = re.findall(r'\[([^\]]+)\]', skus_str or '')
        skus = [[int(sku.strip()) for sku in match.split(',') if sku.strip()] for match in matches]
        if len(companies) != len(skus):
            raise ValueError(
                f'Длины кампаний ({len(companies)}) и списков SKU ({len(skus)}) должны совпадать!'
            )
        return {
            comp: {
                'skus': sku_list,
                'start_date': single_start,
                'end_date': single_end
            }
            for comp, sku_list in zip(companies, skus)
        }

    # 1. Есть кампании + индивидуальные даты (но нет общих дат): проверка, что у каждой кампании даты заданы
    if companies_str and skus_str and multi_starts and multi_ends:
        companies = [int(x) for x in re.split(r'[\s,]+', companies_str.strip()) if x]
        matches = re.findall(r'\[([^\]]+)\]', skus_str or '')
        skus = [[int(sku.strip()) for sku in match.split(',') if sku.strip()] for match in matches]
        starts = [s.strip() for s in re.split(r'[\s,]+', multi_starts.strip()) if s.strip()]
        ends = [e.strip() for e in re.split(r'[\s,]+', multi_ends.strip()) if e.strip()]
        if not (len(companies) == len(skus) == len(starts) == len(ends)):
            raise ValueError(
                f'Длины кампаний ({len(companies)}), SKU ({len(skus)}), '
                f'start_dates ({len(starts)}), end_dates ({len(ends)}) должны совпадать!'
            )
        return {
            comp: {
                'skus': sku_list,
                'start_date': start,
                'end_date': end
            }
            for comp, sku_list, start, end in zip(companies, skus, starts, ends)
        }
    # Если заданы кампании+артикулы, но хотя бы у одной кампании не переданы обе даты - ошибка
    if (companies_str and skus_str) and (multi_starts or multi_ends):
        message = 'Даты должны быть заданы для всех кампаний, если указан хотя бы один набор дат.'
        logger.error(message)
        raise ValueError(message)

    # 3. Есть только общий блок дат (без РК) — возвращаем просто даты
    if single_start and single_end and not (companies_str and skus_str):
        return {'start_date': single_start, 'end_date': single_end}

    # 4. Если не передано ничего или неполные данные — пустой словарь
    return {}


# Функция чтения данных последнего отчета
def read_wb_report_articles_info(reports_dir, client_name):
    """
    Читает самый свежий отчет WB по клиенту и возвращает общий словарь:
        {
            <артикул (str)>: {
                "max_date": str ('YYYY-MM-DD') или None,
                "dates": set(date) — все уникальные даты строк по этому артикулу,
                "campaign_history": {date: ID_кампании, ...},
                "df": исходный DataFrame листа артикула
            },
            ...
        }
    Если отчетов нет — вернет пустой словарь.

    :param reports_dir: каталог с Excel-отчетами Wildberries
    :param client_name: имя клиента (используется в паттерне поиска файла)
    :return: словарь с данными по каждому артикулу
    """
    reports_dir = Path(reports_dir)
    pattern = f"*Свод_Артикулы_РК_{client_name}_WB.xlsx"
    files = list(reports_dir.glob(pattern))
    if not files:
        logger.warning(f"Не найдено ни одного отчета для клиента '{client_name}' в папке '{reports_dir}'.")
        return {}

    # Сортировка файлов по дате изменения (последний — самый свежий)
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    latest_path = files[0]
    logger.info(f"Используется последний отчет '{latest_path.name}' (дата изменения: {latest_path.stat().st_mtime}).")

    try:
        xl = pd.ExcelFile(latest_path, engine='openpyxl')
    except Exception as e:
        logger.error(f"Ошибка при открытии Excel-файла '{latest_path}': {e}")
        return {}

    articles_info = {}

    # Обрабатываем каждый лист (один артикул = один лист)
    for sheet in xl.sheet_names:
        try:
            # Читаем лист, пропуская шапку до табличных данных
            df = pd.read_excel(
                latest_path,
                sheet_name=sheet,
                skiprows=ROW_START_EXCEL - 1,
                engine='openpyxl'
            )
        except Exception as e:
            logger.error(f"Ошибка при чтении листа '{sheet}' из файла '{latest_path}': {e}")
            continue

        # Если в листе есть столбец "Дата"
        if 'Дата' in df.columns:
            # Преобразуем 'Дата' к типу datetime (нормализуем формат)
            df['date_parsed'] = pd.to_datetime(df['Дата'], format="%d.%m.%Y", dayfirst=True, errors='coerce')
            max_dt = df['date_parsed'].max()
            max_date_str = max_dt.strftime('%Y-%m-%d') if pd.notnull(max_dt) else None
            dates_set = set(df['date_parsed'].dropna().dt.date)
            logger.info(f"Лист '{sheet}': найдено {len(dates_set)} уникальных дат, максимальная дата: {max_date_str}.")
        else:
            df['date_parsed'] = pd.NaT
            max_date_str = None
            dates_set = set()
            logger.warning(f"Лист '{sheet}': отсутствует столбец 'Дата', все даты будут пустыми.")

        # "История кампаний": только строки, где и дата, и ID кампании заполнены
        if 'ID Кампании' in df.columns:
            campaign_series = df.loc[
                df['date_parsed'].notnull() & df['ID Кампании'].notnull()
            ].set_index('date_parsed')['ID Кампании']
            campaign_map = {dt.date(): cid for dt, cid in campaign_series.items()}
            logger.info(f"Лист '{sheet}': найдено {len(campaign_map)} записей истории кампаний.")
        else:
            campaign_map = {}
            logger.warning(f"Лист '{sheet}': отсутствует столбец 'ID Кампании', история кампаний будет пустой.")

        articles_info[sheet] = {
            "max_date": max_date_str,                 # максимальная дата по этому артикулу
            "dates": dates_set,                       # множество всех уникальных дат
            "campaign_history": campaign_map,         # дата -> ID кампании
            "df": df                                  # исходный датафрейм листа артикула
        }


# Функция определения номеров РК для отправки в запрос статистики РК
def get_campaigns_and_period_for_update(articles_info, cli_result, today=None):
    """
    Определяет, какие номера РК и даты нужны для обновления.

    articles_info — dict, полученный из read_wb_report_articles_info
    cli_result — результат парсера CLI (dict: кампании<->артикулы<->даты/или просто даты)
    today — дата "сегодня" (date). Если None, будет использоваться date.today().

    Возвращает:
    {
        'campaign_ids': set(номеров кампаний для запроса),
        'start_date': дата начала,
        'end_date': дата окончания,
        'article_to_campaign_map': {артикул: {дата: номер_РК, ...}, ...}
    }
    """
    if today is None:
        today = date.today()

    logger.info("Старт определения кампаний и дат для обновления отчета")

    # Итоговые кампании для WB API
    campaigns_to_query = set()
    # Итоговые даты
    start_date_api = None
    end_date_api = None
    # Карта для лога: артикул — даты — какая кампания
    article_to_campaign_map = {}

    # Если нет CLI и данных предыдущего отчета, будет ошибка
    if not cli_result and not articles_info:
        logger.error("Нет данных для обновления отчета: отсутствуют параметры и информация из предыдущих отчетов.")
        raise ValueError("Нет данных для обновления отчета.")

    # СЦЕНАРИЙ 1. CLI пуст — автообновление по всем артикулам
    if not cli_result:
        logger.info("Автоматический режим: обновление по всем артикулам и датам, которые еще не выгружены")
        yesterday = today - timedelta(days=1)
        for article, art_data in articles_info.items():
            max_date_str = art_data.get("max_date")
            known_dates = art_data.get("dates", set())
            if max_date_str:
                d = date.fromisoformat(max_date_str) + timedelta(days=1)
            else:
                d = yesterday - timedelta(days=6)  # если нет отчета — неделя назад
                logger.warning(f"Артикул {article}: нет max_date, обновляются последние 7 дней ({d} - {yesterday})")
            dates_needed = []
            while d <= yesterday:
                if d not in known_dates:
                    dates_needed.append(d)
                d += timedelta(days=1)
            camp_history = art_data["campaign_history"]
            article_to_campaign_map[article] = {}
            camp_dates = []
            for dt in dates_needed:
                prev_camps = [c for d_, c in camp_history.items() if d_ <= dt]
                if prev_camps:
                    camp_id = prev_camps[-1]
                    article_to_campaign_map[article][dt] = camp_id
                    campaigns_to_query.add(camp_id)
                    camp_dates.append(dt)
                else:
                    logger.warning(f"Артикул {article}: для {dt} неизвестна кампания по истории, обновление не будет выполнено по этой дате.")
            if camp_dates:
                if (start_date_api is None) or (min(camp_dates) < start_date_api):
                    start_date_api = min(camp_dates)
                if (end_date_api is None) or (max(camp_dates) > end_date_api):
                    end_date_api = max(camp_dates)
            else:
                logger.info(f"Артикул {article}: все даты уже актуальны или нет подходящих дат к обновлению.")

    # СЦЕНАРИЙ 2. Только диапазон дат от пользователя
    elif 'start_date' in cli_result and 'end_date' in cli_result:
        s = date.fromisoformat(cli_result['start_date'])
        e = date.fromisoformat(cli_result['end_date'])
        logger.info(f"Пользователь задал диапазон дат {s} — {e} (для всех известных артикулов)")
        for article, art_data in articles_info.items():
            known_dates = art_data.get("dates", set())
            camp_history = art_data["campaign_history"]
            article_to_campaign_map[article] = {}
            camp_dates = []
            d = s
            while d <= e:
                if d not in known_dates:
                    prev_camps = [c for d_, c in camp_history.items() if d_ <= d]
                    if prev_camps:
                        camp_id = prev_camps[-1]
                        article_to_campaign_map[article][d] = camp_id
                        campaigns_to_query.add(camp_id)
                        camp_dates.append(d)
                    else:
                        logger.warning(f"Артикул {article}: для {d} неизвестна кампания по истории, дата пропущена.")
                d += timedelta(days=1)
            if camp_dates:
                if (start_date_api is None) or (min(camp_dates) < start_date_api):
                    start_date_api = min(camp_dates)
                if (end_date_api is None) or (max(camp_dates) > end_date_api):
                    end_date_api = max(camp_dates)
            else:
                logger.info(f"Артикул {article}: все даты в диапазоне уже актуальны или отсутствуют кампании для обновления.")

    # СЦЕНАРИЙ 3. Пользователь явно задал кампании и даты
    else:
        logger.info("Пользователь задал кампании, даты и/или артикула явно")
        for campaign_id, params in cli_result.items():
            skus = params['skus']
            s = date.fromisoformat(params['start_date'])
            e = date.fromisoformat(params['end_date'])
            campaigns_to_query.add(campaign_id)
            logger.info(f"Кампания {campaign_id}: обновление по артикулам {skus} за период {s} — {e}")
            for sku in skus:
                article = str(sku)
                art_data = articles_info.get(article, {})
                known_dates = art_data.get("dates", set())
                article_to_campaign_map.setdefault(article, {})
                camp_dates = []
                d = s
                while d <= e:
                    if d not in known_dates:
                        article_to_campaign_map[article][d] = campaign_id
                        camp_dates.append(d)
                    d += timedelta(days=1)
                if camp_dates:
                    if (start_date_api is None) or (min(camp_dates) < start_date_api):
                        start_date_api = min(camp_dates)
                    if (end_date_api is None) or (max(camp_dates) > end_date_api):
                        end_date_api = max(camp_dates)
                else:
                    logger.info(f"Артикул {article}, кампания {campaign_id}: все даты в диапазоне уже актуальны.")

    logger.info(f"Итоговый набор кампаний для запроса: {campaigns_to_query}")
    logger.info(f"Дата начала: {start_date_api}, дата окончания: {end_date_api}")
    logger.info("Формирование карты артикул — дата — кампания завершено.")

    result = {
        'campaign_ids': campaigns_to_query,
        'start_date': start_date_api,
        'end_date': end_date_api,
        'article_to_campaign_map': article_to_campaign_map
    }
    return result

# Вспомогательная функция для получения статистики по РК
def get_wb_adv_stats(
    campaign_ids,
    begin_date,
    end_date,
    headers,
    url = "https://advert-api.wildberries.ru/adv/v3/fullstats",
    period=60,
    max_requests=300,
    default_burst=10,
    burst_pause=0.2,
    max_retries=5
):
    """
    Запрашивает статистику РК WB с поддержкой лимитов, burst и разбивкой датовых интервалов на <=31 день.
    Автоматически бьёт кампании на чанки по 50 и даты на куски по 31 день.
    """
    companies_stats_api = []

    # Преобразуем к списку и бьём на чанки по 50 кампаний
    campaign_ids = list(map(str, campaign_ids))
    chunk_size = 50
    chunks = [campaign_ids[i:i + chunk_size] for i in range(0, len(campaign_ids), chunk_size)]

    # Делаем список всех [start, end] интервалов <=31 дня
    date_chunks = []
    dt_format = '%Y-%m-%d'
    start = datetime.strptime(begin_date, dt_format).date()
    end = datetime.strptime(end_date, dt_format).date()
    cur_start = start
    while cur_start <= end:
        cur_end = min(cur_start + timedelta(days=30), end)
        date_chunks.append((cur_start.isoformat(), cur_end.isoformat()))
        cur_start = cur_end + timedelta(days=1)

    logger.info(f"Всего кампаний: {len(campaign_ids)} (чанки по {chunk_size}), периодов: {len(date_chunks)} (по 31 день макс.)")

    session = requests.Session()
    req_count = 0
    for chunk_num, chunk in enumerate(chunks, 1):
        for period_num, (period_start, period_end) in enumerate(date_chunks, 1):
            params = {
                "ids": ",".join(chunk),
                "beginDate": period_start,
                "endDate": period_end
            }
            retry = 0
            while retry <= max_retries:
                try:
                    logger.info(
                        f"Чанк {chunk_num}/{len(chunks)}, период {period_num}/{len(date_chunks)}: "
                        f"{period_start} — {period_end}, кампании: {params['ids']} (попытка #{retry+1})"
                    )
                    response = session.get(url, params=params, headers=headers, timeout=60)
                    if response.status_code == 200:
                        remaining = int(response.headers.get('X-Ratelimit-Remaining', default_burst))
                        logger.debug(f"Ответ 200. Осталось {remaining} запросов до burst-лимита.")
                        companies_stats_api.extend(response.json())
                        req_count += 1
                        if remaining <= 1:
                            logger.info(f"Остался 1 запрос до burst-лимита. Пауза {burst_pause} сек.")
                            time.sleep(burst_pause)
                        else:
                            time.sleep(burst_pause)
                        break
                    elif response.status_code == 429:
                        retry_after = int(response.headers.get('X-Ratelimit-Retry', 2))
                        burst_reset = int(response.headers.get('X-Ratelimit-Reset', 1))
                        limit = response.headers.get('X-Ratelimit-Limit')
                        logger.warning(
                            f"Ответ 429 Too Many Requests. Ждём {retry_after} сек (X-Ratelimit-Retry), "
                            f"burst будет полностью восстановлен через {burst_reset} сек (лимит {limit})."
                        )
                        time.sleep(max(retry_after, 1))
                        retry += 1
                    else:
                        logger.error(
                            f"Ошибка {response.status_code} при запросе WB API: {response.text}"
                        )
                        break
                except Exception as e:
                    logger.exception(f"Ошибка при запросе WB API: {e}")
                    time.sleep(2)
                    retry += 1
            else:
                logger.error(
                    f"Не удалось получить статистику по кампаниям {params['ids']} и датам {period_start}-{period_end} после {max_retries} попыток."
                )

    logger.success(f"Статистика получена по {req_count} запросам, всего кампаний: {len(campaign_ids)}.")

    return companies_stats_api

# Основная функция получения статистики РК
def get_companies_stats_all(companies_to_update):

    campaign_ids = companies_to_update['campaign_ids']
    start_date = companies_to_update['start_date']
    end_date = companies_to_update['end_date']

    if campaign_ids and start_date and end_date:
        # 2. Получаем статистику только для нужных кампаний и дат:
        start_date_promotion = start_date.isoformat() if hasattr(start_date, "isoformat") else str(start_date)
        end_date_promotion = end_date.isoformat() if hasattr(end_date, "isoformat") else str(end_date)
        companies_stats_api = get_wb_adv_stats(
            campaign_ids=campaign_ids,
            begin_date=start_date_promotion,
            end_date=end_date_promotion,
            headers=headers
        )
    else:
        logger.warning("Нет кампаний и/или даты для обновления, запрос к WB API не будет выполнен.")
        companies_stats_api = []

    # 3. Возвращай что нужно для дальнейшей обработки — распределения по артикулам и датам,
    #    вставки в DataFrame/отчёт и т.д.
    return companies_stats_api

# Функция обработки статистики РК
def process_companies_stats_api(companies_stats_api):
    """
    Преобразует список WB /adv/v3/fullstats в плоский DataFrame через json_normalize и explode+concat.
    days→apps→nms раскрываются для каждого дня, app и артикула.
    """

    # 1. Преобразуем в DataFrame с раскрытием дней (campaign-level)
    df_companies_stats_processed = pd.json_normalize(
        companies_stats_api,
        record_path=['days'],
        meta=['advertId'],
        errors='ignore',
        record_prefix='day_'
    )
    if df_companies_stats_processed.empty:
        return df_companies_stats_processed

    # 2. Раскрываем список apps через explode и concat, добавляя префикс app_
    if 'day_apps' in df_companies_stats_processed.columns:
        df_companies_stats_processed = df_companies_stats_processed.explode('day_apps', ignore_index=True)
        app_df = pd.json_normalize(df_companies_stats_processed['day_apps'], sep='_')
        app_df = app_df.add_prefix('app_').reset_index(drop=True)
        df_companies_stats_processed = df_companies_stats_processed.reset_index(drop=True)
        df_companies_stats_processed = pd.concat([df_companies_stats_processed, app_df], axis=1)
        df_companies_stats_processed = df_companies_stats_processed.drop(columns=['day_apps'])

    # 3. Раскрываем список nms через explode и concat, добавляя префикс nms_
    if 'app_nms' in df_companies_stats_processed.columns:
        df_companies_stats_processed = df_companies_stats_processed.explode('app_nms', ignore_index=True)
        nms_df = pd.json_normalize(df_companies_stats_processed['app_nms'], sep='_')
        nms_df = nms_df.add_prefix('nms_').reset_index(drop=True)
        df_companies_stats_processed = df_companies_stats_processed.reset_index(drop=True)
        df_companies_stats_processed = pd.concat([df_companies_stats_processed, nms_df], axis=1)
        df_companies_stats_processed = df_companies_stats_processed.drop(columns=['app_nms'])

    # Приводим дату к нужному формату
    df_companies_stats_processed['Дата'] = (
        pd.to_datetime(df_companies_stats_processed['day_date'])
        .dt.tz_localize(None)
        .dt.date
    )

    # 4. Выбираем нужные колонки и переименовываем
    columns_mapping = {
        'advertId': 'campaign_id',
        'Дата': 'Дата',
        'nms_nmId': 'Артикул WB',
        # 'nms_name': 'Наименование товара',
        'nms_views': 'Показы',
        'nms_clicks': 'Клики',
        'nms_atbs': 'В корзину',
        'nms_canceled': 'Отмены',
        'nms_cpc': 'CPC',
        'nms_cr': 'CR',
        'nms_ctr': 'CTR',
        'nms_orders': 'Заказы РК шт',
        'nms_sum_price': 'Заказы РК руб',
        # 'nms_shks': 'Количество заказанных товаров',
        'nms_sum': 'Расход',
    }
    df_companies_stats_result = df_companies_stats_processed.rename(columns=columns_mapping)
    df_companies_stats_result = df_companies_stats_result[list(columns_mapping.values())]


    # Приведи типы, если нужно:
    # df_companies_stats_processed['date'] = pd.to_datetime(df_companies_stats_processed['date']).dt.date

    return df_companies_stats_processed


# Функция получения списка кампаний
def get_company_list(
     headers,
     cli_companies_parsed
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
    # Переименовываем колонку с номером кампании
    df_company_list = df_company_list.rename(columns={
        'advertId': 'Номер кампании'
    })
    # Если задан список кампаний CLI, то используем только кампании из CLI
    # if cli_companies_parsed:
    #     df_company_list = df_company_list.loc[
    #         df_company_list['Номер кампании'].isin(cli_companies_parsed),
    #         :
    #     ]
    # # Выводим предупреждение, если какой-то кампании не было найдено в списке кампаний API
    # if not df_company_list['Номер кампании'].isin(cli_companies_parsed).any():
    #     logger.warning("Some campaigns from CLI were not found in API response")

    return df_company_list

# Функция получения информации об артикулах кампании в зависимости от типа
def extract_nms_from_campaign(camp):
    nms = []
    if 'autoParams' in camp and 'nms' in camp['autoParams']:
        if isinstance(camp['autoParams']['nms'], list):
            nms.extend(camp['autoParams']['nms'])
    if 'params' in camp and isinstance(camp['params'], list):
        for param in camp['params']:
            if 'nms' in param and isinstance(param['nms'], list):
                for nm_entry in param['nms']:
                    if isinstance(nm_entry, dict) and 'nm' in nm_entry:
                        nms.append(nm_entry['nm'])
                    elif isinstance(nm_entry, int):
                        nms.append(nm_entry)
    if 'autoParams' in camp and 'nmCPM' in camp['autoParams']:
        for item in camp['autoParams']['nmCPM']:
            if 'nm' in item:
                nms.append(item['nm'])
    return list(set(nms))


def get_companies_info_not_9(headers, df_company_list):
    """
    Получает по API WB инфо по списку кампаний (пачками по 50), парсит результат с учетом разной структуры кампаний.
    Возвращает итоговый DataFrame, где каждая строка — отдельный артикул кампании:
    ID кампании, Тип кампании, Название кампании,
    Дата создания, Дата запуска, Дата последнего изменения,
    Артикул WB (nms)
    """
    # Разбиваем на чанки по 50
    step = 50
    df_company_list_ = df_company_list.copy()
    df_company_list_['id'] = np.arange(len(df_company_list_))
    df_company_list_['chunks'] = (df_company_list_['id'] // step) + 1
    # Фильтруем тип кампании, убирая кампании с типом 9
    df_company_list_ = df_company_list_[df_company_list_['type'] != 9]
    logger.info("Uploading companies info")

    all_rows = []

    # Лимиты WB API
    REQUEST_LIMIT = 5
    REQUEST_INTERVAL = 0.2  # 200 мс
    TIME_WINDOW = 1.0
    request_times = deque()

    if not df_company_list_.empty:
        for chunk in df_company_list_['chunks'].unique():
            now = time.time()
            while request_times and now - request_times[0] > TIME_WINDOW:
                request_times.popleft()
            if len(request_times) >= REQUEST_LIMIT:
                time_to_wait = TIME_WINDOW - (now - request_times[0])
                logger.info(f"Rate limit reached, sleeping for {round(time_to_wait,3)} seconds")
                time.sleep(max(time_to_wait, REQUEST_INTERVAL))
            else:
                time.sleep(REQUEST_INTERVAL)
            request_times.append(time.time())

            df_chunk = df_company_list_[df_company_list_['chunks'] == chunk]
            company_list = df_chunk['Номер кампании'].tolist()
            try:
                resp = requests.post(
                    "https://advert-api.wildberries.ru/adv/v1/promotion/adverts",
                    headers=headers,
                    json=company_list, timeout=30
                )
                resp.raise_for_status()
                campaigns = resp.json()
            except Exception as e:
                logger.warning(f"Request failed: {e}")
                continue

            # Обработка каждой кампании
            for camp in campaigns:
                advert_id = camp.get('advertId', None)
                camp_type = camp.get('type', None)
                name = camp.get('name', None)
                start_time = camp.get('startTime', None)
                create_time = camp.get('createTime', None)
                change_time = camp.get('changeTime', None)
                nms = extract_nms_from_campaign(camp)
                all_rows.append({
                    "ID кампании": advert_id,
                    "Тип кампании": camp_type,
                    "Название кампании": name,
                    "Дата запуска": start_time,
                    "Дата создания": create_time,
                    "Дата последнего изменения": change_time,
                    "Артикул WB": nms  # теперь поле для explode
                })
        df_companies_info_not_9 = pd.DataFrame(all_rows)

        # Преобразуем поле "Артикул WB" из списка в строки (explode)
        df_companies_info_not_9 = df_companies_info_not_9.explode("Артикул WB").reset_index(drop=True)
        # Можно опционально убрать пустые значения артикулов:
        df_companies_info_not_9 = df_companies_info_not_9[df_companies_info_not_9["Артикул WB"].notna()]

        # Преобразуем даты к формату pandas.Timestamp (без таймзоны)
        for col in ["Дата запуска", "Дата создания", "Дата последнего изменения"]:
            df_companies_info_not_9[col] = pd.to_datetime(
                df_companies_info_not_9[col]
            ).dt.tz_localize(None)
    else:
        df_companies_info_not_9 = pd.DataFrame(columns=[
            "ID кампании", "Тип кампании", "Название кампании",
            "Дата создания", "Дата запуска", "Дата последнего изменения",
            "Артикул WB"
        ])

    return df_companies_info_not_9


def get_companies_info_type_9(headers, df_company_list, statuses='4,7,8,9,11', payment_type=None):
    """
    Получает по API WB инфо по кампаниям с типом 9.
    Использует pandas.json_normalize + explode для структурирования данных по каждой паре (ID кампании, nm_id).
    """
    step = 50
    df_company_list_ = df_company_list.copy()
    df_company_list_['id'] = np.arange(len(df_company_list_))
    df_company_list_['chunks'] = (df_company_list_['id'] // step) + 1
    df_company_list_ = df_company_list_[df_company_list_['type'] == 9]
    logger.info("Uploading companies info (type 9)")

    # Лимиты WB API
    REQUEST_LIMIT = 5
    REQUEST_INTERVAL = 0.2
    TIME_WINDOW = 1.0
    request_times = deque()

    # ФИНАЛЬНЫЙ СПИСОК всех полученных рекламных кампаний
    all_adverts = []

    if not df_company_list_.empty:
        for chunk in df_company_list_['chunks'].unique():
            # Лимит запросов
            now = time.time()
            while request_times and now - request_times[0] > TIME_WINDOW:
                request_times.popleft()
            if len(request_times) >= REQUEST_LIMIT:
                time_to_wait = TIME_WINDOW - (now - request_times[0])
                logger.info(f"Rate limit reached, sleeping for {round(time_to_wait,3)} seconds")
                time.sleep(max(time_to_wait, REQUEST_INTERVAL))
            else:
                time.sleep(REQUEST_INTERVAL)
            request_times.append(time.time())

            df_chunk = df_company_list_[df_company_list_['chunks'] == chunk]
            company_ids = df_chunk['Номер кампании'].tolist()
            ids_str = ",".join(str(x) for x in company_ids)
            params = {"ids": ids_str, "statuses": statuses}
            if payment_type: params["payment_type"] = payment_type

            try:
                resp = requests.get(
                    "https://advert-api.wildberries.ru/adv/v0/auction/adverts",
                    headers=headers,
                    params=params,
                    timeout=30
                )
                resp.raise_for_status()
                result = resp.json()
                adverts = result.get("adverts", [])
                all_adverts.extend(adverts)
            except Exception as e:
                logger.warning(f"Request failed: {e}")
                continue

        # json_normalize подготовит "плоский" df
        # Возможно, какая-то часть записей не содержит nm_settings — после explode будут строки с NaN
        df_adverts_type_9 = pd.json_normalize(
            all_adverts,
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

        # Переименуем столбцы для удобства
        df_adverts_type_9 = df_adverts_type_9.rename(columns={
            'id': 'ID кампании',
            'nm_id': 'Артикул WB',
            'bids.search': 'Ставка в поиске',
            'bids.recommendations': 'Ставка в рекомендациях',
            'status': 'Статус кампании',
            'settings.name': 'Название кампании',
            'settings.payment_type': 'Тип оплаты',
            'timestamps.created': 'Дата создания',
            'timestamps.started': 'Дата запуска',
            'timestamps.updated': 'Дата последнего изменения',
            'timestamps.deleted': 'Дата удаления',
            'subject.id': 'ID категории',
            'subject.name': 'Категория'
        })
        # Приводим даты к pandas.Timestamp
        for col in [
            'Дата создания', 'Дата запуска', 'Дата последнего изменения', 'Дата удаления'
        ]:
            if col in df_adverts_type_9:
                df_adverts_type_9[col] = (
                    pd.
                    to_datetime(df_adverts_type_9[col], format='mixed')
                    .dt.tz_localize(None)
                )

        # Убираем дубли, если вдруг nm_settings был пуст — в результате будет строка только с campaign_id, их можно убрать:
        df_adverts_type_9 = df_adverts_type_9.dropna(subset=["Артикул WB"])

        # Опционально типы:
        df_adverts_type_9['ID кампании'] = df_adverts_type_9['ID кампании'].astype(int)
        df_adverts_type_9['Артикул WB'] = df_adverts_type_9['Артикул WB'].astype(int)
        # Сбрасываем index
        df_adverts_type_9 = df_adverts_type_9.reset_index(drop=True)

    return df_adverts_type_9


# Функция получения "сырых" данных по доп. информации по кампаниям из WB API
def fetch_companies_info_raw(headers, df_company_list_filtered, statuses='-1,4,7,8,9,11', payment_type=None):
    """
    Выполняет запросы к API WB /api/advert/v2/adverts для получения данных по кампаниям.

    :param headers: dict с обязательным Authorization
    :param df_company_list_filtered: pd.DataFrame с колонкой 'Номер кампании', содержащий кампании для запроса
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
        ids_str = ",".join(str(x) for x in df_chunk['Номер кампании'])

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
        'settings.payment_type': 'Тип оплаты',
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
    df_adverts_info =  df_adverts_info.assign(
        **{
            'Ставка в рекомендациях (руб.)': df_adverts_info['Ставка в рекомендациях (коп.)'] / 100,
            'Ставка в поиске (руб.)': df_adverts_info['Ставка в поиске (коп.)'] / 100,
        }
    )
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
    # Выбираем нужные колонки
    df_adverts_info = df_adverts_info[[
        'ID кампании',
        'Артикул WB',
        'Ставка в рекомендациях (руб.)',
        'Ставка в поиске (руб.)'
    ]]

    df_adverts_info = df_adverts_info.reset_index(drop=True)

    return df_adverts_info

# Функция получения дополнительной информации по РК (общая)
def get_companies_info_all(headers, companies_to_update):
    """
    Получает полную информацию по кампаниям из WB через новый API v2 (https://advert-api.wildberries.ru/api/advert/v2/adverts).

    1) Получает список кампаний
    2) Фильтрует по кампаниям из companies_to_update
    3) Запрашивает подробную информацию по этим кампаниям
    4) Возвращает DataFrame с детализацией

    :param headers: dict с заголовками, включая авторизацию
    :param companies_to_update: dict, результат get_campaigns_and_period_for_update (ключ 'campaign_ids')
    :return: pd.DataFrame с информацией по кампаниям и артикулам
    """
    campaign_ids = companies_to_update.get('campaign_ids', set())
    if not campaign_ids:
        logger.warning("Нет кампаний для получения информации по кампаниям.")
        return pd.DataFrame()

    # Получаем полный список кампаний для клиента
    df_company_list = get_company_list(headers, cli_companies_parsed=None)

    # Фильтруем список кампаний для обновления по campaign_ids
    df_company_list_filtered = df_company_list[df_company_list['Номер кампании'].isin(campaign_ids)].copy()
    logger.info(f"Фильтруем список кампаний: выбранных {len(df_company_list_filtered)} из {len(df_company_list)}")

    if df_company_list_filtered.empty:
        logger.warning("После фильтрации не осталось кампаний для запроса детальной информации.")
        return pd.DataFrame()

    # Запрашиваем подробную информацию по отфильтрованным кампаниям через новый API
    all_adverts_info = fetch_companies_info_raw(
        headers,
        df_company_list_filtered,
    )
    # Обрабатываем "сырые" данные в DataFrame
    df_companies_info_all = process_companies_info_to_df(all_adverts_info)

    logger.info(f"Получена доп. информация по {len(df_companies_info_all)} кампаниям.")

    return df_companies_info_all


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
        companies = ARTICLE_REPORT_LIST.get(client_name, [])
        if not companies:
            raise ValueError(f"Список кампаний пуст для клиента {client_name}!")
        input_companies = [[int(company)] for company in companies]

    # Получаем даты для отчёта РК
    date_start_performance, date_end_performance = get_start_end_date(df_date_range, type_dates='companies')
    # Добавляем даты к каждой кампании
    for company in input_companies:
        company.extend([date_start_performance, date_end_performance])

    return input_companies


# Функция создания началной и конечной даты
# GPT START ----
def generate_dates(
        start_date: str,
        end_date: str
    ) -> pd.DataFrame:
    """
    Генерирует диапазон дат между start_date и end_date включительно (по времени: 00:00:00 и 23:59:59).
    Даты передаются в формате 'YYYY-MM-DD'.
    Возвращает DataFrame с полями: date_start, date_end, datetime_start, datetime_end.
    """
    def parse_date(date_str):
        return datetime.strptime(date_str, "%Y-%m-%d")

    start_date_dt = parse_date(start_date)
    end_date_dt = parse_date(end_date)

    # if start_date_dt >= end_date_dt:
    #     raise ValueError("Нет новых данных для генерации отчета, проверьте даты в текущем отчете")

    df_dates = pd.DataFrame([{
        'date_start': start_date_dt.strftime('%Y-%m-%d'),
        'date_end': end_date_dt.strftime('%Y-%m-%d'),
        'datetime_start': start_date_dt.strftime('%Y-%m-%dT00:00:00Z'),
        'datetime_end': end_date_dt.strftime('%Y-%m-%dT23:59:59Z')
    }])
    return df_dates
# GPT END ----


# Функция получения начальной и конечной даты в формате строки
def get_start_end_date(df_dates, type_dates='companies'):
    # Если задан тип дат - даты для заказов, то берем даты со временем
    if type_dates == 'orders':
        date_start, date_end = df_dates.loc[0, ['datetime_start', 'datetime_end']]
    # В остальных случаях берем даты без времени
    else:
        date_start, date_end = df_dates.loc[0, ['date_start', 'date_end']]

    return date_start, date_end


# Функция создания диапазона дат с частотой одного дня
def generate_date_range(df_dates):
    # Получаем начальную и конечную дату
    date_start_range, date_end_range = get_start_end_date(df_dates)
    # Генерируем диапазон дат от начальной до конечной даты с интервалом в один день
    df_date_range = pd.DataFrame({'datetime': pd.date_range(date_start_range, date_end_range, freq='d')})
    # Переводим в формат даты, чтобы убрать время
    df_date_range['date'] = df_date_range['datetime'].dt.date

    return df_date_range





def split_into_intervals(start_date, end_date, max_days=31):
    """Разбивает интервал на интервалы не длиннее max_days"""
    intervals = []
    current_start = start_date
    while current_start <= end_date:
        current_end = min(current_start + timedelta(days=max_days-1), end_date)
        intervals.append({'begin': current_start.strftime('%Y-%m-%d'), 'end': current_end.strftime('%Y-%m-%d')})
        current_start = current_end + timedelta(days=1)
    return intervals



# %% Вызов всех функций
if __name__ == '__main__':
    # # Парсим аргументы командной строки
    cli_params = cli_article_and_companies(parser)
    # Создаем директорию, куда будет помещен итоговый отчет
    article_report_dir = create_article_svod_dir(client_name)
    # # Даты
    # date_start_cli = cli_params.get('start_date')
    # date_end_cli = cli_params.get('end_date')
    # # Список РК
    cli_companies = cli_params.get('promotion_companies')
    # # Список артикулов
    cli_sku = cli_params.get('wb_skus')
    # # Парсим список РК
    # cli_companies_parsed = parse_cli_companies(
    #     cli_params,
    #     ARTICLE_REPORT_LIST,
    #     client_name
    # )
    cli_companies_parsed = ARTICLE_REPORT_LIST.get(client_name)

    # Считываем данные последнего отчета
    articles_info = read_wb_report_articles_info(article_report_dir, client_name)
    # Определяем номера РК для получения статистики
    companies_to_update = get_campaigns_and_period_for_update(
        articles_info=articles_info,
        cli_result=cli_companies_parsed
    )
    # Получаем статистику РК
    companies_stats_api = get_companies_stats_all(companies_to_update)
    # Обрабатываем статистику РК
    df_companies_stats_processed = process_companies_stats_api(companies_stats_api)

    logger.info(f"Creating article svod for client {client_name}")
    # Дата, за которую формируется (или формировался) отчет
    date_report = str(date.today())
