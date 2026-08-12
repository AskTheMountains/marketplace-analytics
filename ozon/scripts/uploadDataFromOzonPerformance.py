
# %% Определение всех функций
import shutil
import math
import time
from datetime import date, datetime, timedelta
from zipfile import ZipFile

import pandas as pd
import requests
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger

from ozon.scripts.constants import (client_id_performance, client_name,
                                    client_secret_performance,
                                    marketplace_dir_path,
                                    ozon_performance_api_url)
from ozon.scripts.ozon_generic_functions import msk_boundary_to_utc

# Файл с настройками и номером клиента
# from options import settings, client_number


# Функция создания нужных директорий
def create_dirs_performance():
    print('jopa')


# Функция генерации дат
def generateDates(days_ago = 30):
    date_end = str(date.today()- timedelta(days=1)) + 'T23:59:59.000Z'
    date_start = str(date.today()- timedelta(days=1) - timedelta(days_ago)) + 'T00:00:00.000Z'
    logger.info(f"Uploading orders from "+ date_start + " to " + date_end)
    return date_end,date_start #,date_end_file,date_start_file

# Функция сохранения дат в csv
def save_dates_to_csv(date_start, date_end):
    # Формат дат для названий файла
    date_end_file = datetime.fromisoformat(date_end).strftime('%Y-%m-%d')
    date_start_file = datetime.fromisoformat(date_start).strftime('%Y-%m-%d')
    # Датафрейм с датами начала и окончания выгрузки данных
    df_dates = pd.DataFrame({'date_end': date_end,
                             'date_start': date_start,
                             'date_end_file': date_end_file,
                             'date_start_file': date_start_file},
                             index=[0])
    df_dates.to_csv(f"{uploaddir_today}/{str(date.today())}_dates_from_to.csv", sep=';', index=False)


# Функция получения токена авторизации
def getAuthorizationToken(client_id_performance, client_secret_performance):
    # URL запроса
    url = ozon_performance_api_url
    endpoint = "/api/client/token"
    # Заголовки запроса
    headers = {
    "Content-Type": "application/json",
    "Accept": "application/json"
    }
    # Параметры запроса
    params = {
    "client_id": client_id_performance,
    "client_secret": client_secret_performance,
    "grant_type":"client_credentials"
    }
    token_dict = requests.post(url + endpoint, headers=headers, json=params).json()
    # Добавляем поля даты и времени создания токена и истечения срока его действия
    token_creation_date = datetime.now()
    token_expiration_date = token_creation_date + timedelta(seconds=token_dict['expires_in']) - timedelta(seconds=60)
    token_dict['date_created'] = token_creation_date
    token_dict['date_expires'] = token_expiration_date
    if 'access_token' not in token_dict.keys():
        print('Error getting access token')
    else:
        return token_dict


# Функция проверки действительности токена
def check_token_exp(auth_token, client_id_performance, client_secret_performance):
    """
    Проверяет срок действия токена. Если срок истёк — запрашивает новый токен.

    Параметры:
        auth_token             - текущий токен авторизации (dict)
        client_id_performance  - client_id
        client_secret_performance - client_secret

    Возвращает:
        dict с (возможно обновлённым) токеном, содержащий ключи:
            - access_token
            - token_type
            - date_expires (datetime)
    """
    # Делаем копию для избежания изменений в оригинале
    new_auth_token = auth_token.copy()

    # Текущая дата и время
    current_datetime = datetime.now()

    # Если срок действия истёк — получаем новый токен
    if current_datetime >= auth_token['date_expires']:
        logger.info("Срок действия токена истёк, запрашиваем новый токен авторизации")
        new_auth_token = getAuthorizationToken(client_id_performance, client_secret_performance)

    return new_auth_token

# Функция выгрузки списка отчетов с их статусами, сгенерированных по API
def get_performance_reports_list(
        auth_token,
        page_size=50,
        max_pages=None,
):
    """
    Выгружает список отчётов, сгенерированных по API Performance, и возвращает их в виде DataFrame.

    Параметры:
        auth_token                 - токен авторизации
        page_size                  - количество отчётов на одной странице (pageSize)
        max_pages                  - максимальное количество страниц для выгрузки
                                     (если None — идём до конца, пока не выгрузим total)

    Возвращает:
        pd.DataFrame с колонками:
            - UUID
            - state
            - createdAt
            - updatedAt
            - error
            - link
            - kind
            - name
            - campaigns_ids   (список ID кампаний из блока items[].campaigns)
            - campaigns_titles(список названий кампаний)
            - request_from
            - request_to
            - request_dateFrom
            - request_dateTo
            - request_groupBy
    """

    # Получаем токен авторизации
    auth_token = check_token_exp(auth_token, client_id_performance, client_secret_performance)
    headers = {
        'Authorization': f"{auth_token['token_type']} {auth_token['access_token']}"
    }

    base_url = f"{ozon_performance_api_url}:443/api/client/statistics/externallist"

    reports = []
    page = 1
    total = None  # всего отчётов (uint64 в ответе)

    logger.info(
        f"Начинаем выгрузку списка отчётов Performance через externallist "
        f"(pageSize={page_size})"
    )

    while True:
        params = {
            "page": page,
            "pageSize": page_size,
        }

        resp = requests.get(base_url, headers=headers, params=params)
        try:
            data = resp.json()
        except ValueError:
            logger.error(
                f"Не удалось распарсить ответ externallist как JSON. "
                f"HTTP статус: {resp.status_code}, текст: {resp.text[:500]}"
            )
            break

        if resp.status_code != 200:
            logger.error(
                f"Ошибка при запросе списка отчётов externallist. "
                f"HTTP статус: {resp.status_code}, ответ: {data}"
            )
            break

        # items — список отчётов, total — общее количество
        items = data.get("items", [])
        total_str = data.get("total", "0")
        try:
            total = int(total_str)
        except (TypeError, ValueError):
            total = None

        logger.info(
            f"Загружена страница {page}. "
            f"Количество отчётов на странице: {len(items)}. "
            f"Всего отчётов (total): {total_str}."
        )

        # Собираем данные по каждому отчёту
        for item in items:
            meta = item.get("meta", {}) or {}
            campaigns_info = item.get("campaigns", []) or []
            request_info = meta.get("request", {}) or {}

            campaigns_ids = [c.get("id") for c in campaigns_info if isinstance(c, dict)]
            campaigns_titles = [c.get("title") for c in campaigns_info if isinstance(c, dict)]

            reports.append({
                "UUID": meta.get("UUID"),
                "state": meta.get("state"),
                "createdAt": meta.get("createdAt"),
                "updatedAt": meta.get("updatedAt"),
                "error": meta.get("error"),
                "link": meta.get("link"),
                "kind": meta.get("kind"),
                "name": item.get("name"),
                "campaigns_ids": campaigns_ids,
                "campaigns_titles": campaigns_titles,
                "request_from": request_info.get("from"),
                "request_to": request_info.get("to"),
                "request_dateFrom": request_info.get("dateFrom"),
                "request_dateTo": request_info.get("dateTo"),
                "request_groupBy": request_info.get("groupBy"),
            })

        # Если на странице ничего не вернулось — выходим
        if not items:
            logger.info("Получен пустой список items — дальнейших страниц нет.")
            break

        # Если total известен — считаем, сколько максимум страниц
        if total is not None:
            max_pages_by_total = math.ceil(total / page_size)
        else:
            max_pages_by_total = None

        # Условие остановки по количеству страниц
        if max_pages is not None and page >= max_pages:
            logger.info(
                f"Достигнут лимит по количеству страниц: {page} из {max_pages}. "
                f"Останавливаем выгрузку."
            )
            break

        # Условие остановки по total
        if max_pages_by_total is not None and page >= max_pages_by_total:
            logger.info(
                f"Достигнута последняя страница по расчёту из total: "
                f"{page} из {max_pages_by_total}. Останавливаем выгрузку."
            )
            break

        # Переходим к следующей странице
        page += 1

    if not reports:
        logger.info("Список отчётов externallist пуст — возвращаем пустой DataFrame.")
        return pd.DataFrame()

    df_reports = pd.DataFrame(reports)
    logger.info(
        f"Выгрузка списка отчётов завершена. "
        f"Получено записей: {df_reports.shape[0]}."
    )
    return df_reports

# Список кампаний
def getCompanyList(auth_token, companies_list=[]):
    # Авторизация
    headers = {
        'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
    }
    # Параметры запроса
    params_companies = {
        "campaignIds": companies_list,
        "advObjectType": [],
        "state": [],
    }
    resp_data_companies = (
        requests.get(
            f"{ozon_performance_api_url}:443/api/client/campaign",
            headers=headers, params=params_companies
        )
        .json()
    )
    df_companies = pd.DataFrame(resp_data_companies['list'])
    df_companies = df_companies.rename(columns={
        'id': 'Номер кампании'
    })
    return df_companies


# Статистика по кампаниям
def getCompanyStatistics(
        date_start,
        date_end,
        auth_token,
        df_companies,
        companies_stats_dir,
        date_from='',
        date_to='',
        task_id=None
):
    """
    Выгружает отчёты по статистике рекламных кампаний из Ozon Performance.

    Логика работы:
      1. Если передан только task_id (df_companies = None или пустой):
         - игнорируются параметры дат и список кампаний;
         - функция дожидается готовности отчёта по этому task_id и
           скачивает его (или завершает работу при статусе ERROR).

      2. Если передан task_id и одновременно передан непустой df_companies:
         - сначала обрабатывается отчёт по этому task_id (как в п.1);
         - затем выполняется обычная выгрузка по списку кампаний и датам
           с разбиением на чанки;
         - если по task_id статус ERROR — это логируется, но выгрузка по
           списку кампаний всё равно выполняется.

      3. Если task_id не передан:
         - выполняется только обычная логика по списку кампаний с разбиением
           на чанки и запросом новых заданий на формирование отчётов.

    Параметры:
        date_start : str
            Дата/дата-время начала периода в формате, поддерживаемом API.
            Используется в параметре "from" при создании задания отчёта.
            Если пустая строка, поле "from" в запросе не отправляется.

        date_end : str
            Дата/дата-время окончания периода в формате, поддерживаемом API.
            Используется в параметре "to" при создании задания отчёта.
            Если пустая строка, поле "to" в запросе не отправляется.

        auth_token : dict
            Текущий токен авторизации Performance API со структурой:
                {
                    'access_token': str,
                    'token_type': str,
                    'date_expires': datetime
                }
            Перед каждым запросом токен обновляется через check_token_exp.

        df_companies : pandas.DataFrame или None
            DataFrame со списком кампаний для выгрузки статистики.
            Обязательно должен содержать колонку "Номер кампании".
            Если df_companies = None или пустой и task_id не задан,
            функция завершится без выгрузки.
            Если df_companies = None или пустой и task_id задан — будет
            обработан только отчёт по task_id.

        companies_stats_dir : pathlib.Path
            Путь к директории, в которую будут сохраняться файлы отчётов
            по кампаниям (csv или распакованные из zip).

        date_from : str, по умолчанию ''
            Дата начала периода без времени (например, '2026-02-01').
            Используется в параметре "dateFrom" при создании задания отчёта.
            Если пустая строка, поле "dateFrom" в запросе не отправляется.

        date_to : str, по умолчанию ''
            Дата окончания периода без времени (например, '2026-02-28').
            Используется в параметре "dateTo" при создании задания отчёта.
            Если пустая строка, поле "dateTo" в запросе не отправляется.

        task_id : str или None, по умолчанию None
            Идентификатор уже созданного задания на формирование отчёта
            по статистике кампаний в Ozon Performance.
            Используется следующим образом:
              - если задан и df_companies пустой/None — игнорируются даты и кампании,
                функция ждёт завершения этого задания и скачивает отчёт;
              - если задан и df_companies непустой — сначала обрабатывается этот
                отчёт, затем создаются и обрабатываются новые задания по списку
                кампаний (обычная логика по чанкам);
              - если не задан — для каждого чанка создаётся своё задание
                на формирование отчёта.
    """

    # ===== 1. Если передан task_id — сначала работаем только с ним =====
    if task_id is not None:
        logger.info(
            "Передан task_id для отчёта по кампаниям. "
            "Сначала обработаем этот отчёт."
        )

        # Проверяем/обновляем токен
        auth_token = check_token_exp(auth_token, client_id_performance, client_secret_performance)
        headers = {
            'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
        }

        report_status = ''
        logger.info(f"Проверяем статус ранее созданного отчёта по кампаниям. ID задания: {task_id}")

        while report_status not in ['OK', 'ERROR']:
            auth_token = check_token_exp(auth_token, client_id_performance, client_secret_performance)
            headers = {
                'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
            }
            resp_data_report_status = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/{task_id}",
                headers=headers
            ).json()

            report_status = resp_data_report_status.get('state', '')

            logger.info(
                "Статус ранее созданного отчёта по статистике кампаний:\n"
                f"  Текущий статус: {report_status}\n"
                f"  ID задания: {task_id}"
            )
            time.sleep(5)

        if report_status == 'OK':
            logger.info("Ранее созданный отчёт по кампаниям сформирован, скачиваем файл")

            params_get_report = {
                "UUID": resp_data_report_status['UUID']
            }
            resp_data_report = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/report",
                params=params_get_report,
                headers=headers
            )

            report_type = resp_data_report.headers['Content-Type'].split(';')[0]
            file_name = resp_data_report.headers['Content-Disposition'].split(';')[1]

            replace_dict = {
                '"': '',
                ' ': '',
                'filename=': '',
            }
            for old, new in replace_dict.items():
                file_name = file_name.replace(old, new)

            if 'csv' in report_type:
                file_path = companies_stats_dir / file_name
                with open(file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                logger.info(f"Файл ранее созданного отчёта по кампаниям сохранён: {file_path}")
            else:
                zip_file_path = companies_stats_dir / f"{str(date.today())}_Кампании_task.{report_type.split('/')[-1]}"
                with open(zip_file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                with ZipFile(zip_file_path, 'r') as zip_ref:
                    zip_ref.extractall(companies_stats_dir)
                logger.info(
                    f"Архив ранее созданного отчёта по кампаниям сохранён: {zip_file_path} "
                    f"(содержимое распаковано в {companies_stats_dir})"
                )
        else:
            logger.error(
                "Ранее созданный отчёт по статистике кампаний не был сформирован.\n"
                f"Итоговый статус: {report_status}\n"
                f"ID задания: {task_id}"
            )

        # Если список кампаний не передан или пуст — на этом всё
        if df_companies is None or getattr(df_companies, "shape", (0,))[0] == 0:
            logger.info(
                "Список кампаний не передан (или пуст). "
                "После обработки task_id дополнительная выгрузка по кампаниям не выполняется."
            )
            return

        # Если кампании есть — будем дальше работать по чанкам, но с НОВЫМИ заданиями
        logger.info(
            "После обработки ранее созданного отчёта продолжаем выгрузку "
            "по переданным кампаниям и датам."
        )
        task_id = None  # сбрасываем, дальше работаем как обычно

    # ===== 2. Дальше — обычная логика по чанкам, ТОЛЬКО если есть df_companies =====

    if df_companies is None:
        logger.info("df_companies не передан, пропускаем выгрузку статистики по кампаниям")
        return

    # Копия df, чтобы не портить исходный
    df_companies_ = df_companies.copy().reset_index(drop=True)

    # Если кампаний нет — сразу выходим
    total_companies = df_companies_.shape[0]
    if total_companies == 0:
        logger.info("Нет кампаний для выгрузки статистики")
        return

    # Размер чанка (сколько кампаний запрашиваем за один раз)
    step = 10
    # Нумеруем чанки: 0..N → 1..M
    df_companies_['chunk'] = df_companies_.index // step + 1
    # Всего чанков
    total_chunks = int(df_companies_['chunk'].max())
    # Сколько кампаний уже выгружено (суммарно по всем завершённым чанкам)
    uploaded_companies_amount = 0
    # Даты для логов
    log_date_start = date_from if date_from else date_start
    log_date_end = date_to if date_to else date_end

    # Цикл по чанкам (по порядку)
    for chunk in sorted(df_companies_['chunk'].unique()):
        # Текущий поднабор кампаний
        chunk_df = df_companies_.loc[df_companies_['chunk'] == chunk]
        # Список кампаний (строками)
        companies_list_str = chunk_df['Номер кампании'].astype(str).tolist()
        # Фактический размер текущего чанка
        current_companies_amount = len(companies_list_str)

        logger.info(
            f"Старт выгрузки статистики по кампаниям. "
            f"Текущий чанк: {chunk} из {total_chunks} "
            f"(кампаний в чанке: {current_companies_amount}, "
            f"максимальный размер чанка: {step}). "
            f"Всего кампаний: {total_companies}."
        )

        # Проверяем/обновляем токен
        auth_token = check_token_exp(auth_token, client_id_performance, client_secret_performance)
        headers = {
            'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
        }

        # Параметры запроса
        params_companies_stats = {
            "campaigns": companies_list_str,
            "from": date_start,
            "to": date_end,
            "dateFrom": date_from,
            "dateTo": date_to,
            "groupBy": "NO_GROUP_BY"
        }
        # Удаляем from/to, если они не заданы
        if date_start == '' or date_end == '':
            params_companies_stats.pop('from')
            params_companies_stats.pop('to')

        # Если task_id ещё не задан, создаём новое задание на формирование отчёта
        if task_id is None:
            report_task_ok = False
            while not report_task_ok:
                resp_data_company_stats = (
                    requests
                    .post(
                        f"{ozon_performance_api_url}:443/api/client/statistics",
                        headers=headers,
                        json=params_companies_stats
                    )
                    .json()
                )
                # Если при запросе выходит ошибка, выводим текст ошибки и пробуем ещё раз
                if 'error' in resp_data_company_stats:
                    error_message = resp_data_company_stats['error']
                    logger.warning(
                        f"Ошибка при создании задания на выгрузку статистики кампаний: "
                        f"{error_message}. Повтор через 5 секунд."
                    )
                    time.sleep(5)
                else:
                    report_task_ok = True
            # Получаем ID задания на генерацию отчёта
            task_id = resp_data_company_stats['UUID']

        # Ожидаем готовности отчёта
        report_status = ''  # начальное значение статуса
        logger.info("Проверяем статус формирования отчёта по кампаниям")
        while report_status not in ['OK', 'ERROR']:
            # Обновляем токен при необходимости
            auth_token = check_token_exp(auth_token, client_id_performance, client_secret_performance)
            headers = {
                'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
            }
            # Запрашиваем статус отчёта
            resp_data_report_status = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/{task_id}",
                headers=headers
            ).json()
            report_status = resp_data_report_status.get('state', '')
            logger.info(
                "\nСтатус формирования отчёта по статистике кампаний:\n"
                f"  Даты отчёта: {log_date_start} - {log_date_end}\n"
                f"  Текущий чанк: {chunk} из {total_chunks} "
                f"(размер чанка: {current_companies_amount})\n"
                f"  Всего выгружено кампаний: {uploaded_companies_amount} из {total_companies}\n"
                f"  Текущий статус отчёта: {report_status}\n"
                f"  ID задания: {task_id}"
            )
            # Ограничение API — не чаще одного запроса раз в 5 секунд
            time.sleep(5)

        # Если отчёт сформирован успешно — скачиваем его
        if resp_data_report_status.get('state') == 'OK':
            logger.info("Отчёт по статистике кампаний сформирован, скачиваем файл")
            params_get_report = {
                "UUID": resp_data_report_status['UUID']
            }
            resp_data_report = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/report",
                params=params_get_report,
                headers=headers
            )
            # Тип содержимого (csv или zip)
            report_type = resp_data_report.headers['Content-Type'].split(';')[0]
            # Имя файла из заголовка
            file_name = resp_data_report.headers['Content-Disposition'].split(';')[1]
            # Чистим имя файла
            replace_dict = {
                '"': '',
                ' ': '',
                'filename=': '',
            }
            for old, new in replace_dict.items():
                file_name = file_name.replace(old, new)
            # Если csv, просто сохраняем
            if 'csv' in report_type:
                file_path = companies_stats_dir / file_name
                with open(file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                logger.info(f"Файл отчёта по кампаниям сохранён: {file_path}")
            else:
                # Иначе считаем, что пришёл zip-архив
                zip_file_path = companies_stats_dir / f"{str(date.today())}_Кампании_{chunk}.zip"
                with open(zip_file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                with ZipFile(zip_file_path, 'r') as zip_ref:
                    zip_ref.extractall(companies_stats_dir)
                logger.info(
                    f"Архив отчёта по кампаниям сохранён: {zip_file_path} "
                    f"(содержимое распаковано в {companies_stats_dir})"
                )

            # Увеличиваем счётчик успешно выгруженных кампаний
            uploaded_companies_amount += current_companies_amount
            logger.info(
                f"Чанк {chunk} завершён. "
                f"Всего выгружено кампаний: {uploaded_companies_amount} из {total_companies}."
            )
            # Сбрасываем task_id для следующего чанка
            task_id = None
        else:
            logger.error(
                f"Отчёт по статистике кампаний не был сформирован (статус: {report_status}). "
                f"Чанк: {chunk} из {total_chunks}, ID задания: {task_id}"
            )
            # В случае ERROR просто переходим к следующему чанку
            task_id = None

# Отчёт по товарам в продвижении в поиске
def getProductsSearchPromote(auth_token, date_start, date_end):

    # Преобразуем строки в datetime с часовым поясом Moscow
    dt_start = pd.to_datetime(date_start).tz_localize('Europe/Moscow').replace(hour=0, minute=0, second=0)
    dt_end = pd.to_datetime(date_end).tz_localize('Europe/Moscow').replace(hour=23, minute=59, second=59)

    # Переводим в UTC и форматируем как строку ISO с Z
    date_start_msk = dt_start.tz_convert('UTC').strftime('%Y-%m-%dT%H:%M:%SZ')
    date_end_msk = dt_end.tz_convert('UTC').strftime('%Y-%m-%dT%H:%M:%SZ')
    # Даты формирования отчета
    params_report_id = {
        "from": date_start_msk,
        "to": date_end_msk,
    }
    # Авторизация
    headers = {
        'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
    }
    # Получаем идентификатор отчета
    resp_data_report_id = requests.post(f"{ozon_performance_api_url}:443/api/client/statistic/products/generate",
                              headers=headers,
                              json=params_report_id).json()
    # Получаем статус отчета
    report_status = ''
    logger.info(
        "\nGetting products promote report status\n",
        f"Dates: {date_start_msk} - {date_end_msk}"
    )
    while report_status not in ['OK', 'ERROR']:
        resp_data_report_status = requests.get(f"{ozon_performance_api_url}:443/api/client/statistics/{resp_data_report_id['UUID']}",
                                            headers=headers).json()
        report_status = resp_data_report_status['state']
        time.sleep(5)

    # Если отчет сформирован успешно, то скачиваем его
    if resp_data_report_status['state'] == 'OK':
        logger.info("Products promote report created, downloading")
        params_get_report = {
            "UUID": resp_data_report_status['UUID']
        }
        resp_data_report = requests.get(
            f"{ozon_performance_api_url}:443/api/client/statistics/report",
            params=params_get_report,
            headers=headers
        )
        file_path = uploaddir_today / f"{str(date.today())}_Отчет_Товары_Продвижение.csv"
        with open(file_path, 'wb') as w:
            w.write(resp_data_report.content)
        logger.info("Done uploading products promote report")
    else:
        logger.info("Error creating products promote report")


def get_search_promo_orders(
        auth_token,
        date_start,
        date_end,
        companies_stats_dir,
        task_id=None,
        report_type: str = "both",
):
    """
    Выгружает отчёт(ы) по продвижению с оплатой за заказ.

    Варианты отчётов (параметр report_type, используется ТОЛЬКО если task_id не задан):
        - "all_sku"
            Отчёт по всем товарам.
            Задача на формирование отчёта создаётся через GET:
                /api/client/statistics/all_sku_promo/orders/generate
            Параметры:
                timeBounds.from = date_start
                timeBounds.to   = date_end

        - "selected_sku"
            Отчёт по выбранным товарам.
            Задача на формирование отчёта создаётся через POST:
                /api/client/statistics/orders/generate
            Параметры JSON:
                {
                    "from": date_start,
                    "to":   date_end
                }

        - "both"
            Последовательно выгружает ДВА отчёта:
                1) отчёт по всем товарам ("all_sku"),
                2) отчёт по выбранным товарам ("selected_sku").

    Особое поведение при task_id:
        - если task_id ПЕРЕДАН:
            Параметр report_type игнорируется полностью.
            Функция НЕ создаёт новые задания, а только:
                * ждёт завершения отчёта по этому task_id;
                * при статусе OK — скачивает CSV;
                * при статусе ERROR — логирует ошибку и завершает работу.

        - если task_id НЕ передан:
            Используется параметр report_type и создаются новые задания
            на формирование отчётов в соответствии с выбранным типом.

    Параметры:
        auth_token : dict
            Текущий токен авторизации Performance API со структурой:
                {
                    'access_token': str,
                    'token_type': str,
                    'date_expires': datetime
                }
            Внутри функции перед каждым сетевым запросом токен обновляется
            через check_token_exp и возвращается в актуальном виде.

        date_start : str
            Дата/дата-время начала периода в формате, поддерживаемом API.
            Для:
                - "selected_sku" → используется в теле POST-запроса как "from".
                - "all_sku"      → используется в GET-параметре "timeBounds.from".
            При использовании task_id напрямую не участвует в запросе,
            но может использоваться только в логировании.

        date_end : str
            Дата/дата-время окончания периода в формате, поддерживаемом API.
            Для:
                - "selected_sku" → используется в теле POST-запроса как "to".
                - "all_sku"      → используется в GET-параметре "timeBounds.to".
            При использовании task_id напрямую не участвует в запросе,
            но может использоваться только в логировании.

        companies_stats_dir : pathlib.Path
            Путь к директории, в которую будут сохраняться CSV-файлы
            отчётов по оплате за заказ. Имена файлов берутся из заголовка
            Content-Disposition ответа API или формируются вручную при
            обработке специфической ошибки "Кампания не найдена".

        task_id : str или None, по умолчанию None
            Идентификатор уже созданного задания на формирование отчёта.
            Если задан:
                - параметр report_type игнорируется;
                - новые задания НЕ создаются;
                - функция только ждёт завершения отчёта по этому task_id
                  и скачивает его (или логирует ERROR).

        report_type : str, по умолчанию "both"
            Тип(ы) отчёта, которые нужно выгрузить, когда task_id НЕ задан.
            Допустимые значения:
                - "all_sku"      — только отчёт по всем товарам;
                - "selected_sku" — только отчёт по выбранным товарам;
                - "both"         — оба отчёта, сначала "all_sku", затем "selected_sku".
            Если task_id передан, этот параметр игнорируется.

    Возвращает:
        dict — актуальный auth_token (возможно обновлённый в процессе работы).
    """

    # ==============================
    # Если передан task_id — игнорируем report_type и только ждём/скачиваем этот отчёт
    # ==============================
    if task_id is not None:
        logger.info(
            "Передан task_id для отчёта по оплате за заказ. "
            "Параметр report_type игнорируется — создаём/используем только этот отчёт.\n"
            f"Диапазон дат (для логов): {date_start} — {date_end}\n"
            f"ID задания: {task_id}"
        )

        report_status = ''

        while report_status not in ['OK', 'ERROR']:
            # Обновляем токен
            auth_token = check_token_exp(
                auth_token,
                client_id_performance,
                client_secret_performance
            )
            headers = {
                'Authorization': auth_token['token_type'] + ' ' + auth_token['access_token']
            }

            resp_data_report_status = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/{task_id}",
                headers=headers
            ).json()

            report_status = resp_data_report_status.get('state', '')

            logger.info(
                "Статус формирования отчёта по продвижению с оплатой за заказ:\n"
                f"  Диапазон дат (лог): {date_start} — {date_end}\n"
                f"  Текущий статус: {report_status}\n"
                f"  ID задания: {task_id}"
            )

            time.sleep(5)

        if report_status == 'OK':
            logger.info("Отчёт по оплате за заказ сформирован, скачиваем файл")

            params_get_report = {
                "UUID": resp_data_report_status['UUID']
            }
            resp_data_report = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/report",
                params=params_get_report,
                headers=headers
            )

            report_type_header = resp_data_report.headers['Content-Type'].split(';')[0]
            file_name = resp_data_report.headers['Content-Disposition'].split(';')[1]

            replace_dict = {
                '"': '',
                ' ': '',
                'filename=': '',
            }
            for old, new in replace_dict.items():
                file_name = file_name.replace(old, new)

            if 'csv' in report_type_header:
                file_path = companies_stats_dir / file_name
                with open(file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                logger.info(f"Отчёт по оплате за заказ сохранён: {file_path}")
            else:
                raise ValueError(
                    f"Неизвестный формат отчёта по оплате за заказ: {report_type_header}"
                )
        else:
            logger.error(
                "Отчёт по продвижению с оплатой за заказ не был сформирован.\n"
                f"Итоговый статус: {report_status}\n"
                f"ID задания: {task_id}"
            )

        return auth_token

    # ==============================
    # Ниже — прежняя логика, но ТОЛЬКО когда task_id не передан
    # ==============================

    report_type = report_type.lower()
    if report_type not in ("all_sku", "selected_sku", "both"):
        raise ValueError("Параметр report_type должен быть 'all_sku', 'selected_sku' или 'both'")

    def _run_single_report(single_type: str, single_task_id: str | None, auth_token_local):
        """
        Обработать один отчёт указанного типа (без использования внешнего task_id, здесь он всегда None):
          - создать задачу на отчёт;
          - дождаться готовности;
          - скачать CSV или сохранить "пустой" отчёт в случае 'Кампания не найдена'.

        single_type: 'all_sku' или 'selected_sku'

        Возвращает:
            обновлённый auth_token_local
        """
        logger.info(
            "Начинаем обработку отчёта по продвижению с оплатой за заказ\n"
            f"Диапазон дат (с временем): {date_start} — {date_end}\n"
            f"Тип отчёта: {single_type}"
        )

        # Формируем параметры для создания задачи
        if single_type == "selected_sku":
            params_products_pay_for_orders = {
                'from': date_start,
                'to': date_end,
            }
        else:
            params_products_pay_for_orders = {
                'timeBounds.from': date_start,
                'timeBounds.to': date_end,
            }

        # Создаём задачу
        report_task_ok = False
        while not report_task_ok:
            auth_token_local = check_token_exp(
                auth_token_local,
                client_id_performance,
                client_secret_performance
            )
            headers = {
                'Authorization': auth_token_local['token_type'] + ' ' + auth_token_local['access_token']
            }

            if single_type == "selected_sku":
                resp_data_products_pay_for_orders = (
                    requests
                    .post(
                        f"{ozon_performance_api_url}:443/api/client/statistics/orders/generate",
                        headers=headers,
                        json=params_products_pay_for_orders,
                    )
                    .json()
                )
            else:
                resp_data_products_pay_for_orders = (
                    requests
                    .get(
                        f"{ozon_performance_api_url}:443/api/client/statistics/all_sku_promo/orders/generate",
                        headers=headers,
                        params=params_products_pay_for_orders,
                    )
                    .json()
                )

            if 'error' in resp_data_products_pay_for_orders:
                error_message = resp_data_products_pay_for_orders['error']

                if error_message == 'Кампания не найдена':
                    logger.warning(
                        "Ozon Performance вернул ошибку 'Кампания не найдена' "
                        f"для отчёта по оплате за заказ (тип: {single_type}). "
                        "Сохраняем пустой отчёт и пропускаем выгрузку."
                    )

                    # Формируем период dd.mm.yyyy-dd.mm.yyyy
                    if single_type == "selected_sku":
                        from_date = params_products_pay_for_orders['from'][:10]
                        to_date = params_products_pay_for_orders['to'][:10]
                    else:
                        from_date = params_products_pay_for_orders['timeBounds.from'][:10]
                        to_date = params_products_pay_for_orders['timeBounds.to'][:10]

                    from_str = datetime.strptime(from_date, "%Y-%m-%d").strftime("%d.%m.%Y")
                    to_str = datetime.strptime(to_date, "%Y-%m-%d").strftime("%d.%m.%Y")
                    period_str = f"{from_str}-{to_str}"

                    if single_type == "selected_sku":
                        file_name = f"search_promo_organisation_orders_report_{period_str}.csv"
                        header_title = f"Отчёт по заказам, период {period_str}"
                        columns = [
                            'Дата',
                            'ID заказа',
                            'Номер заказа',
                            'SKU',
                            'SKU продвигаемого товара',
                            'Артикул',
                            'Источник заказов',
                            'Название товара',
                            'Количество',
                            'Стоимость продажи, ₽',
                            'Стоимость, ₽',
                            'Ставка, %',
                            'Ставка, ₽',
                            'Расход, ₽',
                        ]
                    else:
                        file_name = f"search_promo_all_skus_orders_report_{period_str}.csv"
                        header_title = (
                            f"Оплата за заказ (все товары). Отчёт по заказам, Период {period_str}"
                        )
                        columns = [
                            'Дата',
                            'ID заказа',
                            'Номер заказа',
                            'SKU',
                            'SKU продвигаемого товара',
                            'Артикул',
                            'Название товара',
                            'Количество',
                            'Стоимость продажи, ₽',
                            'Стоимость, ₽',
                            'Ставка, %',
                            'Ставка, ₽',
                            'Расход, ₽',
                        ]

                    file_path = companies_stats_dir / file_name
                    with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                        f.write(header_title)
                        if columns:
                            f.write(';' * (len(columns) - 1))
                        f.write('\n')
                        f.write(';'.join(columns))
                        f.write('\n')

                    logger.info(
                        f"Пустой отчёт по оплате за заказ сохранён (тип: {single_type}) "
                        f"из-за ошибки 'Кампания не найдена': {file_path}"
                    )
                    return auth_token_local

                logger.warning(
                    "Ошибка при создании задания на отчёт по оплате за заказ "
                    f"({single_type}): {error_message}. Повтор через 5 секунд."
                )
                time.sleep(5)
            else:
                report_task_ok = True

        single_task_id = resp_data_products_pay_for_orders['UUID']
        logger.info(
            "Задание на формирование отчёта по оплате за заказ создано. "
            f"Тип: {single_type}, ID задания: {single_task_id}"
        )

        # ждём готовности
        report_status = ''
        logger.info(
            "Проверяем статус формирования отчёта по оплате за заказ\n"
            f"Тип: {single_type}"
        )

        while report_status not in ['OK', 'ERROR']:
            auth_token_local = check_token_exp(
                auth_token_local,
                client_id_performance,
                client_secret_performance
            )
            headers = {
                'Authorization': auth_token_local['token_type'] + ' ' + auth_token_local['access_token']
            }
            resp_data_report_status = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/{single_task_id}",
                headers=headers
            ).json()
            report_status = resp_data_report_status.get('state', '')

            logger.info(
                "\nСтатус формирования отчёта по продвижению с оплатой за заказ:\n"
                f"  Диапазон дат: {date_start} — {date_end}\n"
                f"  Текущий статус: {report_status}\n"
                f"  Тип: {single_type}\n"
                f"  ID задания: {single_task_id}"
            )
            time.sleep(5)

        if report_status == 'OK':
            logger.info(
                f"Отчёт по оплате за заказ сформирован (тип: {single_type}), скачиваем файл"
            )
            params_get_report = {
                "UUID": resp_data_report_status['UUID']
            }
            resp_data_report = requests.get(
                f"{ozon_performance_api_url}:443/api/client/statistics/report",
                params=params_get_report,
                headers=headers
            )
            report_type_header = resp_data_report.headers['Content-Type'].split(';')[0]
            file_name = resp_data_report.headers['Content-Disposition'].split(';')[1]

            replace_dict = {
                '"': '',
                ' ': '',
                'filename=': '',
            }
            for old, new in replace_dict.items():
                file_name = file_name.replace(old, new)

            if 'csv' in report_type_header:
                file_path = companies_stats_dir / file_name
                with open(file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                logger.info(
                    f"Отчёт по продвижению с оплатой за заказ сохранён "
                    f"(тип: {single_type}): {file_path}"
                )
            else:
                raise ValueError(
                    f"Неизвестный формат отчёта по оплате за заказ "
                    f"(тип: {single_type}): {report_type_header}"
                )
        else:
            logger.error(
                "Отчёт по продвижению с оплатой за заказ не был сформирован.\n"
                f"Итоговый статус: {report_status}\n"
                f"Тип: {single_type}\n"
                f"ID задания: {single_task_id}"
            )

        return auth_token_local

    # ===== Основная логика, когда task_id не передан =====
    if report_type in ("all_sku", "selected_sku"):
        auth_token = _run_single_report(report_type, None, auth_token)
    else:
        logger.info("Режим 'both': начинаем выгрузку двух отчётов (all_sku и selected_sku)")
        auth_token = _run_single_report("all_sku", None, auth_token)
        auth_token = _run_single_report("selected_sku", None, auth_token)

    return auth_token

# Вызов всех функций
def upload_data_performane(
        date_start,
        date_end,
        client_id_performance,
        client_secret_performance,
        companies_stats_dir,
        date_from='',
        date_to='',
        df_companies=None,
        task_id=None,
        mode='all',  # 'all' | 'campaigns' | 'orders'
):
    """
    Выгружает данные из Ozon Performance API:
      - статистику по рекламным кампаниям,
      - отчёты по продвижению с оплатой за заказ.

    Параметры
    ---------
    date_start : str
        Дата/время начала периода в формате, ожидаемом Performance API
        (например, '2025-12-01T00:00:00Z'). Может быть пустой строкой, если
        используется пара date_from/date_to.
    date_end : str
        Дата/время окончания периода в формате, ожидаемом Performance API
        (например, '2025-12-31T23:59:59Z'). Может быть пустой строкой.
    client_id_performance : str
        Идентификатор клиента для Ozon Performance API.
    client_secret_performance : str
        Секретный ключ клиента для Ozon Performance API.
    companies_stats_dir : pathlib.Path
        Директория, куда будут сохранены выгруженные файлы статистики.
    date_from : str, optional
        Дата начала периода без времени (обычно 'YYYY-MM-DD').
        Используется для статистики кампаний и, при отсутствии date_start/date_end,
        для формирования периода отчёта по оплате за заказ.
    date_to : str, optional
        Дата окончания периода без времени (обычно 'YYYY-MM-DD').
    df_companies : pandas.DataFrame | None, optional
        DataFrame со списком кампаний (как минимум с колонкой "Номер кампании").
        Обязателен, если mode включает выгрузку кампаний ('all' или 'campaigns').
    task_id : str | None, optional
        ID уже созданного задания на отчёт по кампаниям или оплате за заказ.
        Если передан, функция попытается продолжить с этого задания;
        если None, будет создано новое задание.
    mode : {'all', 'campaigns', 'orders'}, optional
        Режим работы:
          - 'all'       — выгружать и статистику по кампаниям, и отчёты по оплате за заказ;
          - 'campaigns' — выгружать только статистику по кампаниям;
          - 'orders'    — выгружать только отчёты по оплате за заказ.

    Использование дат
    ------------------
    Статистика по кампаниям:
      - можно передать либо date_start/date_end (с временем),
      - либо оставить их пустыми и передать date_from/date_to (только даты).

    Отчёты по оплате за заказ:
      - всегда требуют дат с временем;
      - если date_start/date_end не заданы, но указаны date_from/date_to
        (в формате 'YYYY-MM-DD'), период для отчёта по оплате за заказ
        формируется как:
            orders_date_start = date_from + 'T00:00:00Z'
            orders_date_end   = date_to   + 'T23:59:59Z'
    """

    if mode not in {'all', 'campaigns', 'orders'}:
        raise ValueError(
            f"Не поддерживаемый режим выгрузки: {mode}. "
            f"Используйте: 'all', 'campaigns' или 'orders'."
        )

    # Получаем токен авторизации
    auth_token = getAuthorizationToken(client_id_performance, client_secret_performance)

    # -------------------------
    # 1. Выгрузка кампаний (если требуется)
    # -------------------------
    if mode in {'all', 'campaigns'}:
        # Если df со списком кампаний не задан, то список кампаний должны были передать явно
        if df_companies is None:
            logger.error("Не передан DataFrame со списком кампаний (df_companies).")
            raise ValueError(
                "Для выгрузки статистики кампаний необходимо передать "
                "DataFrame со списком кампаний в параметр df_companies."
            )

        df_companies_for_statistic = df_companies.copy()

        logger.info(
            f"Начинаем выгрузку статистики по кампаниям "
            f"(количество кампаний: {len(df_companies_for_statistic)})"
        )

        # Здесь сохраняем прежнее поведение:
        # date_start/date_end могут быть как с временем, так и пустыми,
        # а date_from/date_to — просто даты.
        getCompanyStatistics(
            date_start,
            date_end,
            auth_token,
            df_companies_for_statistic,
            companies_stats_dir,
            date_from,
            date_to,
            task_id
        )

    # -------------------------
    # 2. Выгрузка отчёта по оплате за заказ (если требуется)
    # -------------------------
    if mode in {'all', 'orders'}:
        # Для оплаты за заказ нужны даты с временем.
        # Пытаемся использовать date_start/date_end, если они заданы.
        if date_start and date_end:
            orders_date_start = date_start
            orders_date_end = date_end
        # Если они не заданы, но есть date_from/date_to (только даты),
        # формируем строки с временем.
        elif date_from and date_to:
            # здесь считаем, что date_from / date_to в формате 'YYYY-MM-DD'
            # orders_date_start = f"{date_from}T00:00:00Z"
            # orders_date_end = f"{date_to}T23:59:59Z"
            orders_date_start = msk_boundary_to_utc(date_from)
            orders_date_end = msk_boundary_to_utc(date_to, end_of_day=True)
        else:
            raise ValueError(
                "Для выгрузки отчёта по оплате за заказ необходимо указать либо "
                "date_start/date_end (с временем), либо date_from/date_to (без времени)."
            )

        logger.info(
            "Начинаем выгрузку отчёта по продвижению с оплатой за заказ: "
            f"{orders_date_start} — {orders_date_end}"
        )

        get_search_promo_orders(
            auth_token,
            orders_date_start,
            orders_date_end,
            companies_stats_dir,
            task_id
        )

    print(f"\033[44m\033[37mDONE UPLOADING OZON PERFORMANCE FOR CLIENT {client_name}")


# %% Вызов всех функций
if __name__ == "__main__":
    # Создаем директорию для файлов выгрузки по API Performance
    uploaddir = marketplace_dir_path / "Clients" / client_name / "UploadFilesPerformance"
    uploaddir.mkdir(parents=True, exist_ok=True)

    # Создаем директорию для выгрузки за сегодняшний день
    uploaddir_today = uploaddir / f"UploadFiles_{str(date.today())}"
    if uploaddir_today.exists():
        # Если директория уже существует, удаляем её и создаём заново
        shutil.rmtree(uploaddir_today)
    uploaddir_today.mkdir()

    # Создаем директорию для статистики кампаний
    companies_stats_dir = uploaddir_today / f"{str(date.today())}_Кампании"
    companies_stats_dir.mkdir()

    date_start = '2025-08-17T00:00:00Z'
    date_end = '2025-08-27T23:59:59Z'
    auth_token = getAuthorizationToken(client_id_performance, client_secret_performance,)

    df_companies = getCompanyList(auth_token)
    getCompanyStatistics(date_start, date_end, auth_token, df_companies, companies_stats_dir)
    print(f"\033[44m\033[37mDONE UPLOADING OZON PERFORMANCE FOR CLIENT {client_name}")
    # getProductsSearchPromote(auth_token, date_start, date_end)

# %%
