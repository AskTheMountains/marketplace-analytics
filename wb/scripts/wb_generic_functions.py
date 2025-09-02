# %% Определение функций
import requests
import json
import time
from datetime import date,datetime,timedelta
import pandas as pd
import shutil
import os
from pathlib import Path
import csv
import numpy as np
import openpyxl
from openpyxl.styles import Border, Side, Alignment, numbers
from openpyxl.utils import get_column_letter
import re
from loguru import logger
import getopt
import sys
pd.options.mode.chained_assignment = None  # default='warn'

# Файл с настройками и номером клиента
from wb.scripts.constants import (
    headers,
    client_name,
    promotion_companies,
    marketplace_dir_name
)
# Функции выгрузки данных
from wb.scripts.uploadDataFromWB import(
    getWBProduct,
    getOrdersWB
)
# Некоторые вспомогательные функции
from generic_functions import move_columns
# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent

# Функция выгрузки статистики РК
# Функция получения статистики кампаний
def get_company_stats(
        headers,
        df_input_companies,
        request_type='interval',
        groupby_cols = ['nmId', 'name']
):
    # Создаем копию для избежания изменений в оригинальном df
    df_companies_list = df_input_companies.copy()
    # df, в который будем помещать результаты выгрузки по апи
    df_companies_stats = pd.DataFrame()

    # Для запроса с интервалом
    # Для каждой кампании формируем интервал запроса данных статистики
    df_companies_list['interval'] = df_companies_list.apply(
        lambda row:  {'begin': row['date_start'], 'end': row['date_end']},
        axis=1
    )

    # Для запроса по дням
    # Переводим в timestamp
    # df_companies_list['datetime_start'] = pd.to_datetime(df_companies_list['date_start'])
    # df_companies_list['datetime_end'] = pd.to_datetime(df_companies_list['date_end'])
    # Создаем список дат с частотой одного дня
    df_companies_list['dates'] = df_companies_list.apply(
        lambda row: pd.date_range(row['date_start'], row['date_end'], freq='D').strftime('%Y-%m-%d').tolist(),
        axis=1
    )
    # Удаляем вспомогательные колонки
    df_companies_list = df_companies_list.loc[:, ~df_companies_list.columns.isin([
        'date_start', 'date_end', 'datetime_start', 'datetime_end'
    ])]

    # Разбиваем df на диапазоны по 100 измерений
    step = 100
    df_companies_list['id'] = np.arange(0, len(df_companies_list))
    df_companies_list['chunks'] = df_companies_list['id'].apply(lambda x: int(x/step) + 1)
    max_chunks = max(df_companies_list['chunks'])

    logger.info("Uploading companies stats")
    # Цикл по каждому диапазону
    for chunk in df_companies_list['chunks'].unique():
        # Выбираем нужный диапазон и колонки
        tmp_df_companies_list = df_companies_list.loc[
            df_companies_list['chunks'] == chunk,
            ['Номер кампании'] + [request_type]
        ]
        # Переименовываем колонку для соответствия параметрам запроса
        tmp_df_companies_list = tmp_df_companies_list.rename(columns={'Номер кампании': 'id'})
        # Переводим в словарь (это и будут параметры запроса)
        params_companies_stats = tmp_df_companies_list.to_dict(orient='records')
        logger.info(f"Uploading {chunk} of {max_chunks} chunks")
        # Делаем запрос
        resp_data_companies_stats = requests.post("https://advert-api.wildberries.ru/adv/v2/fullstats", headers=headers, json=params_companies_stats)
        # Если пришли данные по кампаниям
        if resp_data_companies_stats.status_code == 200:
            # Переводим в словарь
            resp_data_companies_stats = resp_data_companies_stats.json()
            # Переводим в df
            tmp_df_companies_stats = pd.DataFrame(resp_data_companies_stats)
            # Объединяем с предыдущим проходом цикла
            df_companies_stats = pd.concat([df_companies_stats, tmp_df_companies_stats])
            # Ждем 1 минуту до выгрузки следующего диапазона кампаний (если это не последний диапазон)
            logger.info("Waiting 1 minute before uploading next chunk")
            if chunk != max_chunks:
                time.sleep(60)
        # Если не пришли данные по кампаниям, выводим ошибку
        else:
            print(f"{resp_data_companies_stats.json()}")
            if chunk != max_chunks:
                time.sleep(60)
                # Ждем 1 минуту до выгрузки следующего диапазона кампаний (если это не последний диапазон)
                logger.info("Waiting 1 minute before uploading next chunk")

    # Убираем лишние колонки
    # df_companies_stats.loc[:, ~df_companies_stats.columns.isin(['dates', 'days', 'boosterStats'])]

    # Получаем df с номенклатурами кампаний
    # df_companies_nms = result_companies_info['df_companies_nms_all']

    # Добавляем в df статистики кампаний исходные интервалы запроса
    # поскольку интервал, возвращаемый в ответе, не соответствует интервалу запроса
    df_request_interval = (
        pd.concat([
            df_companies_list[['Номер кампании']],
            pd.json_normalize(df_companies_list['interval'])
        ], axis=1)
        .rename(columns={
            'Номер кампании': 'advertId',
            'begin': 'Начальная дата запроса статистики РК',
            'end': 'Конечная дата запроса статистики РК'
        })
    )
    df_companies_stats_with_request_interval = (
        df_companies_stats
        .merge(
            df_request_interval,
            on='advertId',
            how='left'
        )
    )

    # Достаем interval из словаря
    df_companies_stats_with_response_interval = (
        pd.concat([
            df_companies_stats_with_request_interval,
            pd.json_normalize(df_companies_stats_with_request_interval['interval'])
        ], axis=1)
        # .rename(columns={
        #     'begin': 'date_start',
        #     'end': 'date_end',
        # })
    )
    # Переводим даты в timestamp
    # Создаем колонки с датой
    df_companies_stats_with_response_interval['Начальная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['begin']).dt.date
    df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['end']).dt.date
    # У конечной даты добавляем 23:59:59
    # time_to_add = timedelta(hours=23, minutes=29, seconds=59)
    # df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] = (
    #      df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] + time_to_add
    # )

    # Перемещаем колонки с ID кампании и датами в начало df
    companies_columns = [
        'advertId',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'Начальная дата ответа статистики РК', 'Конечная дата ответа статистики РК'
    ]
    df_companies_stats_with_response_interval = move_columns(
        df_companies_stats_with_response_interval,
        companies_columns,
        position=0
    )

    # Распаковываем статистику по дням
    df_companies_stats_unpacked = df_companies_stats_with_response_interval.explode(column='days', ignore_index=True)
    df_companies_stats_by_days = pd.concat([
        df_companies_stats_unpacked[companies_columns],
        pd.json_normalize(df_companies_stats_unpacked['days']),
        ], axis=1)

    # Распаковываем статистику по платформам
    df_by_days_unpacked = df_companies_stats_by_days.explode(column='apps', ignore_index=True)
    df_by_apps = pd.concat([
        df_by_days_unpacked[companies_columns],
        pd.json_normalize(df_by_days_unpacked['apps']),
        ], axis=1)

    # Распаковываем статистику по артикулам
    df_by_apps_unpacked = df_by_apps.explode(column='nm', ignore_index=True)
    df_by_nms = pd.concat([
        df_by_apps_unpacked[companies_columns],
        pd.json_normalize(df_by_apps_unpacked['nm']),
        ], axis=1)

    # Колонки, по которым производится группировка
    group_cols = companies_columns + groupby_cols
    # Считаем суммы по кампании и артикулу
    df_companies_stats_nms = (
        df_by_nms
        .loc[:,
             group_cols +
             ['views', 'clicks', 'sum', 'atbs', 'orders', 'shks', 'sum_price']]
        .groupby(group_cols)
        .sum()
        # .agg(**{
        #     'views': ('views', 'sum'),
        #     'clicks': ('clicks', 'sum')
        # })
        .reset_index()
    )
    # Считаем оставшиеся метрики
    df_companies_stats_nms = df_companies_stats_nms.assign(
        ctr=lambda df: df['clicks'] / df['views'] * 100,
        cr=lambda df: df['shks'] / df['clicks'] * 100,
        cpc=lambda df: df['sum'] / df['clicks'],
        cpm=lambda df: df['sum'] / df['views'] * 1000,
    )
    # Переименовываем колонки
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

    # Создаем столбец с уникальным ID кампании по дате
    df_companies_stats_nms['unique_company_id'] = (
        df_companies_stats_nms
        .groupby(['Номер кампании', 'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК'], as_index=False)
        .ngroup() + 1
    )
    # Порядок колонок
    df_companies_stats_nms = df_companies_stats_nms.loc[:,
        [
            'unique_company_id',
            'Номер кампании',
            'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
            'Начальная дата ответа статистики РК', 'Конечная дата ответа статистики РК',
            'Артикул WB', 'Наименование товара',
            'Показы', 'Клики',
            'CTR, %', 'CR, %', 'CPC','CPM',
            'Затраты руб',
            'Добавлений в корзину',
            'Заказы по рекламе шт (из отчета РК)',
            'Заказанные товары шт (из отчета РК)',
            'Заказы по рекламе руб (из отчета РК)',
       ]
    ]
    # Перемещаем колонку в начало df
    # df_companies_stats_nms = move_columns(
    #     df_companies_stats_nms,
    #     columns_to_move=['unique_company_id'],
    #     position=0
    # )

    return df_companies_stats_nms


def split_into_intervals(start_date, end_date, max_days=31):
    """Разбивает интервал на интервалы не длиннее max_days"""
    intervals = []
    current_start = start_date
    while current_start <= end_date:
        current_end = min(current_start + timedelta(days=max_days-1), end_date)
        intervals.append({'begin': current_start.strftime('%Y-%m-%d'), 'end': current_end.strftime('%Y-%m-%d')})
        current_start = current_end + timedelta(days=1)
    return intervals


# Функция получения статистики кампаний (v2)
def get_company_stats_v2(
        headers,
        df_input_companies,
        request_type='interval',
        groupby_cols = ['advertId','nmId', 'name']
):
    # Создаем копию для избежания изменений в оригинальном df
    df_companies_list = df_input_companies.copy()

    # Для запроса с интервалом
    # Для каждой кампании формируем интервал запроса данных статистики
    df_companies_list['interval'] = df_companies_list.apply(
        lambda row:  {'begin': row['date_start'], 'end': row['date_end']},
        axis=1
    )
    # Формируем интервалы не более 31 дня
    df_companies_list['intervals'] = df_companies_list.apply(
        lambda row: split_into_intervals(
            pd.to_datetime(row['date_start']),
            pd.to_datetime(row['date_end'])
        ), axis=1
    )
    # Создаем распакованный df со списком кампаний по интервалам не более 31 дня
    df_companies_list_full = (
        df_companies_list
        .explode('intervals')
        .reset_index(drop=True)
    )
    # Создаем даты для запроса
    df_companies_list_full['interval_begin'] = df_companies_list_full['intervals'].apply(lambda x: x['begin'])
    df_companies_list_full['interval_end'] = df_companies_list_full['intervals'].apply(lambda x: x['end'])
    # Для запроса по дням
    # Переводим в timestamp
    # df_companies_list['datetime_start'] = pd.to_datetime(df_companies_list['date_start'])
    # df_companies_list['datetime_end'] = pd.to_datetime(df_companies_list['date_end'])
    # Создаем список дат с частотой одного дня
    # df_companies_list['dates'] = df_companies_list.apply(
    #     lambda row: pd.date_range(row['date_start'], row['date_end'], freq='D').strftime('%Y-%m-%d').tolist(),
    #     axis=1
    # )

    # Удаляем вспомогательные колонки
    # df_companies_list_full = df_companies_list_full.loc[:, ~df_companies_list_full.columns.isin([
    #     'date_start', 'date_end', 'datetime_start', 'datetime_end'
    # ])]

    # # Разбиваем df на диапазоны по 100 измерений
    # step = 100
    # df_companies_list_full['id'] = np.arange(0, len(df_companies_list_full))
    # df_companies_list_full['chunks'] = df_companies_list_full['id'].apply(lambda x: int(x/step) + 1)
    # max_chunks = max(df_companies_list_full['chunks'])

    step = 100  # максимум кампаний в одном запросе
    df_companies_stats = pd.DataFrame()  # df, куда будем складывать результаты выгрузки АПИ
    # получаем уникальные интервалы
    unique_intervals = df_companies_list_full[['interval_begin', 'interval_end']].drop_duplicates()
    # Считаем количество отправленных запросов подряд без паузы
    request_counter = 0

    logger.info("Uploading companies stats")

    # Цикл по каждому интервалу
    for interval_idx, interval_row in unique_intervals.iterrows():
        begin = interval_row['interval_begin']
        end = interval_row['interval_end']
        # список уникальных кампаний для этого интервала
        company_ids = df_companies_list_full[
            (df_companies_list_full['interval_begin'] == begin) &
            (df_companies_list_full['interval_end'] == end)
        ]['Номер кампании'].unique()

        # разбиваем на интервалы по 100 кампаний
        num_chunks = (len(company_ids) + step - 1) // step

        for chunk_idx in range(num_chunks):
            i = chunk_idx * step
            # Формируем параметры запроса
            chunk = company_ids[i:i+step]
            params_companies_stats = [
                {'id': int(company_id), 'interval': {'begin': begin, 'end': end}}
                for company_id in chunk
            ]
            logger.info(
                f"\nUploading companies {list(map(int, chunk))}\n"
                f"Interval: {begin} - {end} "
                # f"(кампаний в чанке: {len(chunk)})"
            )
            # Отправляем запрос
            resp_data_companies_stats = requests.post(
                "https://advert-api.wildberries.ru/adv/v2/fullstats",
                headers=headers,
                json=params_companies_stats
            )
            # Если запрос прошел успешно, добавляем данные в общий df
            if resp_data_companies_stats.status_code == 200:
                # Переводим в словарь
                resp_data_companies_stats = resp_data_companies_stats.json()
                # Переводим в df
                tmp_df_companies_stats = pd.DataFrame(resp_data_companies_stats)
                # Добавляем даты текущего интервала
                tmp_df_companies_stats['request_start_date_chunk'] = begin
                tmp_df_companies_stats['request_end_date_chunk'] = end
                # Объединяем в общий df
                df_companies_stats = pd.concat([df_companies_stats, tmp_df_companies_stats])
            # Если запрос вернулся с ошибкой, выводим ошибку
            else:
                logger.error(f"{resp_data_companies_stats.json()}")

            # После КАЖДОГО запроса увеличиваем счетчик
            request_counter += 1

            # Проверяем, последний ли это чанк в последнем интервале
            is_last_interval = (interval_idx == len(unique_intervals) - 1)
            is_last_chunk = (chunk_idx == num_chunks - 1)
            is_last_request = is_last_interval and is_last_chunk

            # После каждого 5го запроса, если НЕ последний, ставим паузу и обнуляем счетчик
            # if request_counter == 5 and not is_last_request:
            if not is_last_request:
                logger.info("Reached 5 requests limit, waiting 1 minute")
                time.sleep(60)
                request_counter = 0  # Обнуляем счетчик

    # Делаем сортировку по номеру кампании
    df_companies_stats = df_companies_stats.sort_values(by='advertId', ignore_index=True)
    # Добавляем в df статистики кампаний исходные интервалы запроса
    # поскольку интервал, возвращаемый в ответе, не соответствует интервалу запроса
    df_request_interval = (
        df_input_companies
        .rename(columns={
            'Номер кампании': 'advertId',
            'date_start': 'Начальная дата запроса статистики РК',
            'date_end': 'Конечная дата запроса статистики РК'
        })
        .loc[:,
             [
            'advertId',
            'Начальная дата запроса статистики РК',
            'Конечная дата запроса статистики РК'
            ]
        ]
    )
    df_companies_stats_with_request_interval = (
        df_companies_stats
        .merge(
            df_request_interval,
            on='advertId',
            how='left'
        )
    )

    # Достаем interval ответа статистики РК из словаря
    df_companies_stats_with_response_interval = (
        pd.concat([
            df_companies_stats_with_request_interval,
            pd.json_normalize(df_companies_stats_with_request_interval['interval'])
        ], axis=1)
        .rename(columns={
            'begin': 'response_start_date',
            'end': 'response_end_date',
        })
    )
    # Переводим даты в timestamp
    # Создаем колонки с датой
    df_companies_stats_with_response_interval['Начальная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['response_start_date']).dt.date
    df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['response_end_date']).dt.date
    # У конечной даты добавляем 23:59:59
    # time_to_add = timedelta(hours=23, minutes=29, seconds=59)
    # df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] = (
    #      df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] + time_to_add
    # )

    # Перемещаем колонки с ID кампании и датами в начало df
    companies_columns = [
        'advertId',
        'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
        'request_start_date_chunk', 'request_end_date_chunk',
        'Начальная дата ответа статистики РК', 'Конечная дата ответа статистики РК'
    ]
    df_companies_stats_with_response_interval = move_columns(
        df_companies_stats_with_response_interval,
        companies_columns,
        position=0
    )

    # Распаковываем статистику по дням
    df_companies_stats_unpacked = df_companies_stats_with_response_interval.explode(column='days', ignore_index=True)
    df_companies_stats_by_days = pd.concat([
        df_companies_stats_unpacked[companies_columns],
        pd.json_normalize(df_companies_stats_unpacked['days']),
        ], axis=1)

    # Распаковываем статистику по платформам
    df_by_days_unpacked = df_companies_stats_by_days.explode(column='apps', ignore_index=True)
    df_by_apps = pd.concat([
        df_by_days_unpacked[companies_columns],
        pd.json_normalize(df_by_days_unpacked['apps']),
        ], axis=1)

    # Распаковываем статистику по артикулам
    df_by_apps_unpacked = df_by_apps.explode(column='nm', ignore_index=True)
    df_by_nms = pd.concat([
        df_by_apps_unpacked[companies_columns],
        pd.json_normalize(df_by_apps_unpacked['nm']),
        ], axis=1)

    # Колонки, по которым производится группировка
    group_cols = companies_columns + groupby_cols

    # Считаем суммы по кампании и артикулу
    df_companies_stats_nms = (
        df_by_nms
        .loc[:,
             groupby_cols +
             ['views', 'clicks', 'sum', 'atbs', 'orders', 'shks', 'sum_price']]
        .groupby(groupby_cols)
        .sum()
        # .agg(**{
        #     'views': ('views', 'sum'),
        #     'clicks': ('clicks', 'sum')
        # })
        .reset_index()
    )
    # Считаем оставшиеся метрики
    df_companies_stats_nms = df_companies_stats_nms.assign(
        ctr=lambda df: df['clicks'] / df['views'] * 100,
        cr=lambda df: df['shks'] / df['clicks'] * 100,
        cpc=lambda df: df['sum'] / df['clicks'],
        cpm=lambda df: df['sum'] / df['views'] * 1000,
    )
    # Переименовываем колонки
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

    # Создаем столбец с уникальным ID кампании по дате
    # df_companies_stats_nms['unique_company_id'] = (
    #     df_companies_stats_nms
    #     .groupby(['Номер кампании', 'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК'], as_index=False)
    #     .ngroup() + 1
    # )
    # # Порядок колонок
    # df_companies_stats_nms = df_companies_stats_nms.loc[:,
    #     [
    #         'unique_company_id',
    #         'Номер кампании',
    #         'Начальная дата запроса статистики РК', 'Конечная дата запроса статистики РК',
    #         # 'Начальная дата ответа статистики РК', 'Конечная дата ответа статистики РК',
    #         'Артикул WB', 'Наименование товара',
    #         'Показы', 'Клики',
    #         'CTR, %', 'CR, %', 'CPC','CPM',
    #         'Затраты руб',
    #         'Добавлений в корзину',
    #         'Заказы по рекламе шт (из отчета РК)',
    #         'Заказанные товары шт (из отчета РК)',
    #         'Заказы по рекламе руб (из отчета РК)',
    #    ]
    # ]
    # Перемещаем колонку в начало df
    # df_companies_stats_nms = move_columns(
    #     df_companies_stats_nms,
    #     columns_to_move=['unique_company_id'],
    #     position=0
    # )

    return df_companies_stats_nms
