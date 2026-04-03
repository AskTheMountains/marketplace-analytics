
# %% Определение всех функций
import shutil
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
    # Делаем копию для избежания изменений в оригинале
    new_auth_token = auth_token.copy()
    # Сохраняем текущую дату и время в переменную
    current_datetime = datetime.now()
    # Если текущая дата больше даты истечения срока действия токена, запрашиваем новый
    if current_datetime >= auth_token['date_expires']:
        logger.info("Auth token expired, requesting new token")
        new_auth_token = getAuthorizationToken(client_id_performance, client_secret_performance)

    return new_auth_token


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
                file_type = 'csv'
                file_path = companies_stats_dir / file_name
                with open(file_path, 'wb') as w:
                    w.write(resp_data_report.content)
                logger.info(f"Файл отчёта по кампаниям сохранён: {file_path}")
            else:
                # Иначе считаем, что пришёл zip-архив
                file_type = 'zip'
                zip_file_path = companies_stats_dir / f"{str(date.today())}_Кампании_{chunk}.{file_type}"
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

    Варианты отчёта:
      - "all_sku"      — отчёт по всем товарам (создание задачи через GET)
      - "selected_sku" — отчёт по выбранным товарам (создание задачи через POST)
      - "both"         — поочерёдно выгружает оба отчёта: all_sku и selected_sku
    """
    report_type = report_type.lower()
    if report_type not in ("all_sku", "selected_sku", "both"):
        raise ValueError("Параметр report_type должен быть 'all_sku', 'selected_sku' или 'both'")

    def _run_single_report(single_type: str, single_task_id: str | None, auth_token_local):
        """
        Создать задачу на отчёт указанного типа, дождаться готовности и скачать CSV.
        single_type: 'all_sku' или 'selected_sku'
        Возвращает обновлённый auth_token.
        """
        logger.info(
            "Начинаем выгрузку отчёта по продвижению с оплатой за заказ\n"
            f"Диапазон дат (с временем): {date_start} — {date_end}\n"
            f"Тип отчёта: {single_type}"
        )

        # Формируем параметры запроса для создания задачи
        if single_type == "selected_sku":
            # вариант с POST, для выбранных товаров
            params_products_pay_for_orders = {
                'from': date_start,
                'to': date_end,
            }
        else:
            # вариант с GET, для всех товаров
            params_products_pay_for_orders = {
                'timeBounds.from': date_start,
                'timeBounds.to': date_end,
            }

        # Если не задан ID задания на генерацию отчёта, создаём новое задание
        if single_task_id is None:
            report_task_ok = False
            while not report_task_ok:
                # Проверяем/обновляем токен (как в getCompanyStatistics)
                auth_token_local = check_token_exp(
                    auth_token_local,
                    client_id_performance,
                    client_secret_performance
                )
                headers = {
                    'Authorization': auth_token_local['token_type'] + ' ' + auth_token_local['access_token']
                }
                # Создаём задание на генерацию отчёта в зависимости от типа
                if single_type == "selected_sku":
                    # POST — отчёт по выбранным товарам
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
                    # GET — отчёт по всем товарам
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

        # ----------------------------
        # Ожидаем готовности отчёта
        # ----------------------------
        report_status = ''  # начальное значение статуса
        logger.info(
            "Проверяем статус формирования отчёта по оплате за заказ\n"
            f"Тип: {single_type}"
        )
        while report_status not in ['OK', 'ERROR']:
            # Обновляем токен при необходимости (как в getCompanyStatistics)
            auth_token_local = check_token_exp(
                auth_token_local,
                client_id_performance,
                client_secret_performance
            )
            headers = {
                'Authorization': auth_token_local['token_type'] + ' ' + auth_token_local['access_token']
            }
            # Проверяем статус задания
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

        # ----------------------------
        # Если отчёт сформирован — скачиваем
        # ----------------------------
        if resp_data_report_status.get('state') == 'OK':
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

    # ----------------------------
    # Основная логика по report_type
    # ----------------------------
    if report_type in ("all_sku", "selected_sku"):
        # Один отчёт, можно передать task_id
        auth_token = _run_single_report(report_type, task_id, auth_token)
    else:
        # Новый режим: выгружаем оба отчёта независимо
        logger.info("Режим 'both': начинаем выгрузку двух отчётов (all_sku и selected_sku)")
        auth_token = _run_single_report("all_sku", None, auth_token)
        auth_token = _run_single_report("selected_sku", None, auth_token)

    # при желании можно возвращать обновлённый токен
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
    mode:
        'all'       - выгружать и кампании, и оплату за заказ
        'campaigns' - только кампании
        'orders'    - только оплату за заказ

    Пары дат:
        - для кампаний:
            можно передавать либо date_start/date_end (с временем),
            либо оставить их пустыми и передать date_from/date_to (только дата).

        - для оплаты за заказ:
            всегда требуются даты с временем;
            если date_start/date_end пустые, но заданы date_from/date_to,
            они будут автоматически преобразованы в
            date_from + 'T00:00:00Z' и date_to + 'T23:59:59Z'.
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
            orders_date_start = f"{date_from}T00:00:00Z"
            orders_date_end = f"{date_to}T23:59:59Z"
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
