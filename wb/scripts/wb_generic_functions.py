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

# Функция получения статистики рекламных кампаний
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


# Функция получения статистики рекламных кампаний (v2)
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
            params_companies_stats = {
                'ids': ','.join(str(i) for i in company_ids),
                'beginDate': begin,
                'endDate': end
            }

            logger.info(
                f"\nUploading companies {list(map(int, chunk))}\n"
                f"Interval: {begin} - {end} "
                # f"(кампаний в чанке: {len(chunk)})"
            )
            # Отправляем запрос
            resp_data_companies_stats = requests.get(
                "https://advert-api.wildberries.ru/adv/v3/fullstats",
                headers=headers,
                params=params_companies_stats
            )            # Если запрос прошел успешно, добавляем данные в общий df
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

    # Если статистика по РК не пустая, начинаем обработку статистики
    if not df_companies_stats.empty:
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
        # df_companies_stats_with_response_interval = (
        #     pd.concat([
        #         df_companies_stats_with_request_interval,
        #         pd.json_normalize(df_companies_stats_with_request_interval['interval'])
        #     ], axis=1)
        #     .rename(columns={
        #         'begin': 'response_start_date',
        #         'end': 'response_end_date',
        #     })
        # )
        # Переводим даты в timestamp
        # Создаем колонки с датой
        # df_companies_stats_with_response_interval['Начальная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['response_start_date']).dt.date
        # df_companies_stats_with_response_interval['Конечная дата ответа статистики РК'] = pd.to_datetime(df_companies_stats_with_response_interval['response_end_date']).dt.date
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
            # 'Начальная дата ответа статистики РК', 'Конечная дата ответа статистики РК'
        ]
        df_companies_stats_with_request_interval = move_columns(
            df_companies_stats_with_request_interval,
            companies_columns,
            position=0
        )


        # Распаковываем статистику по дням
        df_companies_stats_unpacked = df_companies_stats_with_request_interval.explode(column='days', ignore_index=True)
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
        df_by_apps_unpacked = df_by_apps.explode(column='nms', ignore_index=True)
        df_by_nms = pd.concat([
            df_by_apps_unpacked[companies_columns],
            pd.json_normalize(df_by_apps_unpacked['nms']),
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
        # Фильтруем статистику только по тем артикулам, по которым были расходы
        df_companies_stats_nms = (
            df_companies_stats_nms
            .loc[df_companies_stats_nms['Затраты руб'] > 0, :]
            .reset_index(drop=True)
        )
    # Если статистика пришла пустая, то возвращаем пустой df
    else:
        logger.warning("Empty promotion statistics")
        df_companies_stats_nms = pd.DataFrame()

    return df_companies_stats_nms


# --- Выгрузка отчета реализации --- #
## Вспомогательные функции
def wildberries_get_report_detail(
    headers, date_from, date_to, rrdid=0, limit=100000,
    last_call_time=None, max_retries=5
):
    """
    Запрос за один кусок отчёта c лимитом – 1 запрос в минуту и обработкой 429-ошибки.
    """
    url = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod'
    params = {
        'dateFrom': date_from,
        'dateTo': date_to,
        'rrdid': rrdid,
        'limit': limit
    }

    attempt = 0
    while True:
        # Лимит "1 в минуту" между запросами (если не первый)
        if last_call_time is not None:
            elapsed = time.time() - last_call_time
            wait = 60 - elapsed
            if wait > 0:
                if logger:
                    logger.info(f'Пауза {int(wait)} секунд между запросами (лимит WB)')
                time.sleep(wait)
        try:
            response = requests.get(url, headers=headers, params=params)
        except Exception as e:
            if logger:
                logger.error(f"Ошибка соединения: {str(e)}")
            raise
        # Проверим код возврата
        if response.status_code == 429:
            logger.warning(f'Получен 429 Too Many Requests. Ждём 60 сек и повторяем запрос.')
            time.sleep(60)

            attempt += 1
            if attempt > max_retries:
                logger.error('Превышено максимальное число повторов после 429')
                response.raise_for_status()
        else:
            response.raise_for_status()
            break  # Ошибки нет – выходим из цикла

    return response.json(), time.time()

## Основная логика (обработка недель, rrd_id и бинарный сплит)
def upload_realization_report(headers, date_start, date_end, verbose=False):
    """
    Выгружает финансовый отчет Wildberries по частям (по недельным окнам, с разбивкой, пагинацией
    по rrd_id и автодроблением, если данных слишком много).

    Аргументы:
      token (str): Ваш API-токен WB.
      date_start (str): Дата начала ('YYYY-MM-DD').
      date_end (str): Дата конца ('YYYY-MM-DD').
      verbose (bool): Выводить ли подробный лог.

    Возвращает:
      pd.DataFrame со всеми уникальными строками за указанный период.
    """
    all_rows = []             # Список для сбора всех строк отчета
    dedup_key_set = set()     # Для дедупликации: уникальные пары (realizationreport_id, rrd_id)

    cursor_start = datetime.strptime(date_start, "%Y-%m-%d")  # Текущая начальная дата окна
    end_date = datetime.strptime(date_end, "%Y-%m-%d")         # Финальная дата всего периода

    while cursor_start <= end_date:
        # Определяем дату конца текущего окна: максимум неделя или до конца всего периода
        window_end = min(cursor_start + timedelta(days=6), end_date)
        if verbose:
            logger.info(f"\nОкно: {cursor_start} - {window_end}")

        temp_from = cursor_start    # Начальная дата дайжеста (окна)
        temp_to = window_end        # Конечная дата дайжеста (окна)

        while True:  # Цикл автодробления окна, если оно слишком объемное
            rrdid = 0                # Начинаем с нуля (с начала страницы)
            last_max_date_to = None  # Для отслеживания до какой даты реально дошли
            rows_in_window = []      # Сюда собираем строки ТОЛЬКО за это окно
            last_call_time = None  # Пауза между запросами

            while True:  # Цикл пагинации по rrd_id
                # Делаем запрос в API и ждем (sleep внутри функции)
                json_resp, last_call_time = wildberries_get_report_detail(
                    headers,
                    temp_from.strftime("%Y-%m-%d"),
                    temp_to.strftime("%Y-%m-%d"),
                    rrdid=rrdid,
                    last_call_time=last_call_time)

                rows = json_resp
                if not rows:     # Если нет данных — выходим из пагинации
                    break

                # Дедупликация по (realizationreport_id, rrd_id)
                for row in rows:
                    key = (row.get('realizationreport_id'), row.get('rrd_id'))
                    if key not in dedup_key_set:
                        dedup_key_set.add(key)
                        all_rows.append(row)
                        rows_in_window.append(row)

                # Обновляем rrdid для следующего запроса по пачке
                max_rrd = max(row['rrd_id'] for row in rows)
                # Смотрим, по какие даты реально дошли — это нужно для контроля дробления окна
                last_max_date_to = max(pd.to_datetime(row['date_to']) for row in rows)

                if verbose:
                   logger.info(f"  rrdid={rrdid}: получено {len(rows)} строк; max_rrd={max_rrd}")

                # Если получили меньше лимита — все данные окна получены, выходим из rrdid-цикла
                if len(rows) < 100000:
                    break

                rrdid = max_rrd  # Следующий запрос с обновлённым rrdid

            # Если за это окно подняли меньше лимита — окно пройдено, дробление не требуется
            if len(rows_in_window) < 100000:
                break

            # Если лимит выбран, но окно не "сдвигается" по дате — значит слишком плотно, дробим окно меньше
            date_progress = (last_max_date_to - temp_from).days
            window_days = (temp_to - temp_from).days
            if verbose:
                logger.info(
                    f"\n!!! Нужно дробление: внутри окна {len(rows_in_window)} строк,\n"
                    f"last_max_date_to={last_max_date_to}, date_progress={date_progress}, window_days={window_days}"
                )

            # Бинарное деление (уменьшаем правую границу окна вдвое)
            temp_to = temp_from + timedelta(days=window_days//2)
            # Если уже нельзя уменьшать — выходим (что-то с данными, обычно не бывает)
            if temp_to <= temp_from:
                break

        # Переходим к следующему недельному окну
        cursor_start = window_end + timedelta(days=1)

    # Формируем итоговый датафрейм (уже без дублей)
    df = pd.DataFrame(all_rows)
    # Выводим число строк в отчете
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
