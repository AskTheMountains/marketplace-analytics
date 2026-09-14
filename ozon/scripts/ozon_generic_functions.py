
# %% Определение функций
import json
from pathlib import Path

import pandas as pd
import numpy as np
import requests
import re
import time

# Логгер клиента озон
from ozon.scripts.ozon_logger import ozon_client_logger as logger

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent

# Некоторые константы
from ozon.scripts.constants import (
    client_name,
    article_replacements
)

# Функция очистки артикула от всякого тупого вонючего говна, гавна и дерьма
def clean_ozon_article(val):
    # Пропуски оставляем пропусками
    if pd.isna(val):
        return np.nan

    # Если это число — обрабатываем как «артикул, случайно прочитанный числом»
    if isinstance(val, (int, float, np.integer, np.floating)):
        # Красиво превращаем в строку без лишних .0, .00 и т.п.
        s = ('%g' % val)  # 123.0 -> "123", 123.50 -> "123.5"
        return s

    # Всё остальное — строки / объекты — обрабатываем как текстовый артикул
    s = str(val)

    # Применяем regex-словарь ТОЛЬКО к строке
    # Оборачиваем в Series, чтобы использовать .replace(regex=True)
    s_clean = (
        pd.Series([s])
        .replace(article_replacements, regex=True)
        .iloc[0]
    )

    return s_clean


# Функция выгрузки отчета о заказах
def get_orders(headers, date_start, date_end):
    """
    Выгружает отчет о заказах за указанный период (FBO + FBS)

    Параметры:
    - headers: заголовки для запросов к API Ozon
    - date_start: дата начала периода (строка в формате ISO 8601)
    - date_end: дата окончания периода (строка в формате ISO 8601)
    Возвращает объединенный датафрейм с заказами
    """
    from ozon.scripts.uploadDataFromOzon import getOrders
    # Выгружаем отчет об отправлениях fbo и fbs
    df_orders_fbo = getOrders(headers, date_start, date_end, delivery_schema='fbo', to_save=False)
    df_orders_fbs = getOrders(headers, date_start, date_end, delivery_schema='fbs', to_save=False)
    # Убираем лишние колонки из отчета fbo
    df_orders_fbo = df_orders_fbo.loc[:, ~df_orders_fbo.columns.isin(['Объемный вес товаров, кг'])]

    # Объединяем два отчета в один
    common_columns = [
        'Артикул',
        'Название товара',
        'OZON id', 'Статус',
        'Принят в обработку',
        'Количество',
        'Сумма отправления',
        'Номер заказа',
        'Схема доставки',
    ]

    # df_orders_fbo, df_orders_fbs = getOrders_v2(headers, date_start, date_end, to_save=False)
    # # Объединяем два отчета в один
    # common_columns = [
    #     'Схема доставки',
    #     'Артикул',
    #     'Наименование товара',
    #     'SKU',
    #     'Статус',
    #     'Дата и время начала обработки отправления',
    #     'Сумма отправления'
    # ]

    # GPT START----
    dfs = []
    for df in [df_orders_fbo, df_orders_fbs]:
        if df is not None and not df.empty:
            # Берем только те common_columns, которые реально есть
            cols = [col for col in common_columns if col in df.columns]
            dfs.append(df.loc[:, cols])

    if dfs:
        df_orders_all = pd.concat(dfs, ignore_index=True)
    else:
        # Если оба датафрейма пустые — создаём пустой датафрейм с нужными колонками
        df_orders_all = pd.DataFrame(columns=common_columns)
    # GPT END----

    # Переводим колонку с датой заказа в datetime
    df_orders_all['datetime_orders'] = (
        pd.to_datetime(df_orders_all['Принят в обработку'], format='mixed')
        .dt.tz_localize(None)
    )
    # Указываем, какие колонки считаем за заказы в штуках и рублях
    df_orders_all['Заказы шт'] = df_orders_all['Количество'] if 'Количество' in df_orders_all.columns else 1
    df_orders_all['Заказы руб'] = df_orders_all['Сумма отправления']
    # Переименовываем колонку с Ozon ID
    df_orders_all = df_orders_all.rename(columns={
        'OZON id': 'Ozon Product ID',
        'Название товара': 'Наименование товара',
    })
    # Переводим столбец с артикулом в строку
    df_orders_all['Артикул'] = df_orders_all['Артикул'].astype(str)

    return df_orders_all


# Функция выгрузки остатков потоварно
def get_reminders_by_product(
    df_products,
    headers,
    to_agg=True,
    agg_col=['Артикул'],
    split_stocks=False,  # новый флаг
):
    """
    Выгружает и обрабатывает остатки товаров по данным из Ozon (FBO и FBS).
    """
    from ozon.scripts.uploadDataFromOzon import (
        getStockReminders_fbo_v2,
        getStockReminders_fbs,
    )
    import pandas as pd

    # Выгружаем остатки FBO и FBS
    df_reminders_fbo = getStockReminders_fbo_v2(headers, df_products, to_save=False)
    df_reminders_fbs = getStockReminders_fbs(headers, df_products, to_save=False)

    # Переименовываем колонки для удобства
    df_reminders_fbo = df_reminders_fbo.rename(
        columns={
            'Доступный к продаже товар': 'Остатки FBO',
        }
    )
    df_reminders_fbs = df_reminders_fbs.rename(
        columns={
            'Всего на складе, шт': 'Остатки FBS',
        }
    )

    # Приводим артикулы к строке, чтобы ключ был единообразным
    if 'Артикул' in df_reminders_fbo.columns:
        df_reminders_fbo['Артикул'] = df_reminders_fbo['Артикул'].astype(str)
    if 'Артикул' in df_reminders_fbs.columns:
        df_reminders_fbs['Артикул'] = df_reminders_fbs['Артикул'].astype(str)

    # 1. Предварительная агрегация по Артикулу внутри каждого источника,
    #    чтобы не было дублей по одному артикулу (разные склады и т.п.)
    if not df_reminders_fbo.empty:
        df_reminders_fbo = (
            df_reminders_fbo
            .groupby('Артикул', as_index=False)[['Остатки FBO']]
            .sum()
        )
    else:
        df_reminders_fbo = pd.DataFrame(columns=['Артикул', 'Остатки FBO'])

    if not df_reminders_fbs.empty:
        df_reminders_fbs = (
            df_reminders_fbs
            .groupby('Артикул', as_index=False)[['Остатки FBS']]
            .sum()
        )
    else:
        df_reminders_fbs = pd.DataFrame(columns=['Артикул', 'Остатки FBS'])

    # 2. Объединяем агрегированные FBO и FBS по артикулу (outer, чтобы не потерять данные)
    df_reminders_all = pd.merge(
        df_reminders_fbo[['Артикул', 'Остатки FBO']],
        df_reminders_fbs[['Артикул', 'Остатки FBS']],
        on='Артикул',
        how='outer',
    )

    # Страхуемся от NaN и приводим к числам
    for col in ['Остатки FBO', 'Остатки FBS']:
        if col in df_reminders_all.columns:
            df_reminders_all[col] = (
                pd.to_numeric(df_reminders_all[col], errors='coerce')
                .fillna(0)
            )

    # Если стоит флаг агрегации по другим колонкам
    if to_agg:
        df_agg = (
            df_reminders_all
            .groupby(agg_col, as_index=False)[['Остатки FBO', 'Остатки FBS']]
            .sum()
        )

        if not split_stocks:
            df_agg = df_agg.assign(
                Остаток=lambda df: df[['Остатки FBO', 'Остатки FBS']].sum(axis=1)
            )

        return df_agg

    # Если флаг агрегации не стоит – возвращаем неагрегированные остатки
    return df_reminders_all

# Функция выгрузки отчета реализации товаров
def get_products_realization_report(headers, year, month):
    params_finance_report = json.dumps({
    "year": year,
    "month": month
    })
    resp_data_finance_report = requests.post("https://api-seller.ozon.ru/v2/finance/realization", headers=headers, data=params_finance_report).json()
    # result_finance_report_title = pd.DataFrame(resp_data_finance_report['result']['header'], index=[0])
    result_finance_report = pd.DataFrame(resp_data_finance_report['result']['rows'])
    # Получаем данные по товару
    df_products_realization = (
        pd
        .json_normalize(result_finance_report['item'])
        .rename(columns={
            'name': 'Наименование товара',
            'offer_id': 'Артикул',
            'barcode': 'Штрихкод',
            'sku': 'SKU'
        })
    )
    # Словарики для переименовывания
    sales_rename_columns = {
        'amount': 'Сумма',
        'bonus': 'Баллы за скидки',
        'commission': 'Комиссия',
        'compensation': 'Доплата за счёт Ozon',
        'price_per_instance': 'Цена за экземпляр',
        'quantity': 'Количество товара',
        'standard_fee': 'Базовое вознаграждение Ozon',
        'bank_coinvestment': 'Механики лояльности партнеров: зелёные цены',
        'stars': 'Механики лояльности партнеров: звезды',
        'pick_up_point_coinvestment': 'Механики лояльности партнеров: АПВЗ',
        'total': 'Итого к начислению',
    }
    return_rename_columns = sales_rename_columns.copy()
    # Получаем данные по продажам
    df_sales_realization = (
        pd
        .json_normalize(result_finance_report['delivery_commission'])
        .rename(columns=sales_rename_columns)
    )
    # Добавляем префикс к колонкам продаж
    df_sales_realization.columns = [f"{col} (Продажи)" for col in df_sales_realization.columns]

    # Получаем данные по возвратам
    df_returns_realization = (
        pd
        .json_normalize(result_finance_report['return_commission'])
        .rename(columns=return_rename_columns)
    )
    # Добавляем префикс к колонкам возвратов
    df_returns_realization.columns = [f"{col} (Возвраты)" for col in df_returns_realization.columns]

    # Объединяем отчет в один df
    df_realization_report = pd.concat(
        [df_products_realization, df_sales_realization, df_returns_realization],
        axis=1
    )
    # Добавляем год и месяц
    df_realization_report['Год'] = year
    df_realization_report['Месяц'] = month
    # Получаем нужные данные из отчета апи
    # result_finance_report = result_finance_report.assign(
    #     # Номенклатура
    #     Наименование_товара=[d.get('name') for d in result_finance_report.item],
    #     Артикул_продавца=[d.get('offer_id') for d in result_finance_report.item],
    #     Штрихкод=[d.get('barcode') for d in result_finance_report.item],
    #     SKU=[d.get('sku') for d in result_finance_report.item],
    #     Количество=[d.get('quantity') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     # Продажи
    #     Количество_продажи=[d.get('quantity') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Цена_продажи=[d.get('price_per_instance') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Сумма_продажи=[d.get('amount') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Баллы_за_скидки_продажи=[d.get('bonus') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Комиссия_продажи=[d.get('comission') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Доплата_за_счет_озон_продажи=[d.get('compensation') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Базовое_вознаграждение_озон_продажи=[d.get('standard_fee') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_звёзды_продажи=[d.get('stars') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_продажи=[d.get('bank_coinvestment') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_апвз_продажи=[d.get('pick_up_point_coinvestment') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     Итого_к_начислению_продажи=[d.get('total') if d != None else 0 for d in result_finance_report.delivery_commission],
    #     # Возвраты
    #     Количество_возвраты=[d.get('quantity') if d != None else 0 for d in result_finance_report.return_commission ],
    #     Цена_возвраты=[d.get('price_per_instance') if d != None else 0 for d in result_finance_report.return_commission],
    #     Сумма_возвраты=[d.get('amount') if d != None else 0 for d in result_finance_report.return_commission],
    #     Баллы_за_скидки_возвраты=[d.get('bonus') if d != None else 0 for d in result_finance_report.return_commission],
    #     Комиссия_возвраты=[d.get('comission') if d != None else 0 for d in result_finance_report.return_commission],
    #     Доплата_за_счет_озон_возвраты=[d.get('compensation') if d != None else 0 for d in result_finance_report.return_commission],
    #     Базовое_вознаграждение_озон_возвраты=[d.get('standard_fee') if d != None else 0 for d in result_finance_report.return_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_звёзды_возвраты=[d.get('stars') if d != None else 0 for d in result_finance_report.return_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_возвраты=[d.get('bank_coinvestment') if d != None else 0 for d in result_finance_report.return_commission],
    #     Выплаты_по_механикам_лояльности_партнёров_апвз_возвраты=[d.get('pick_up_point_coinvestment') if d != None else 0 for d in result_finance_report.return_commission],
    #     Итого_к_начислению_возвраты=[d.get('total') if d != None else 0 for d in result_finance_report.return_commission],
    #     Год=year,
    #     Месяц=month
    # )
    # for col in ['Количество_продажи', 'Сумма_продажи', 'Количество_возвраты', 'Сумма_возвраты']:
    #     df_finance_report[col] = df_finance_report[col].fillna(0)

    # df_finance_report['Механики партнеров Продажи'] = df_finance_report[['Выплаты_по_механикам_лояльности_партнёров_звёзды_продажи',
    #                                                             'Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_продажи',
    #                                                             'Выплаты_по_механикам_лояльности_партнёров_апвз_продажи']].sum(axis=1)
    # df_finance_report['Механики партнеров Возвраты'] = df_finance_report[['Выплаты_по_механикам_лояльности_партнёров_звёзды_возвраты',
    #                                                             'Выплаты_по_механикам_лояльности_партнёров_зелёные_цены_возвраты',
    #                                                             'Выплаты_по_механикам_лояльности_партнёров_апвз_возвраты']].sum(axis=1)
    # df_finance_report['Продажи минус возвраты шт Озон'] = df_finance_report['Количество_продажи'] - df_finance_report['Количество_возвраты']
    # df_finance_report['Продажи минус возвраты руб Озон'] = df_finance_report['Сумма_продажи'] - df_finance_report['Сумма_возвраты']
    # df_finance_report['Баллы Озон Итог'] = df_finance_report['Баллы_Озон_продажи'] - df_finance_report['Баллы_Озон_возвраты']

    # df_finance_report_by_month = (
    #     df_finance_report
    #     .groupby(['Артикул_продавца', 'SKU', 'Штрихкод', 'Год', 'Месяц'])
    #     .agg(**{
    #     'Продажи шт': ('Количество_продажи', 'sum'),
    #     'Продажи руб': ('Сумма_продажи', 'sum'),
    #     'Баллы Озон Продажи': ('Баллы_Озон_продажи', 'sum'),
    #     'Возвраты шт': ('Количество_возвраты', 'sum'),
    #     'Возвраты руб': ('Сумма_возвраты', 'sum'),
    #     'Баллы Озон Возвраты': ('Баллы_Озон_возвраты', 'sum'),
    #     'Продажи минус возвраты шт Озон': ('Продажи минус возвраты шт Озон', 'sum'),
    #     'Продажи минус возвраты руб Озон': ('Продажи минус возвраты руб Озон', 'sum')
    #     })
    #     .reset_index()
    #     .rename(columns={'Артикул_продавца': 'Артикул продавца'})
    #     # .assign(
    #     #     Размер=df_finance_report_by_month['Артикул продавца'].str.split('_', n=2)
    #     # )
    # )

    # if to_save:
    #     with pd.ExcelWriter(f"{uploaddir_finance_report_today}/finance_report_{client_name}_Ozon.xlsx") as w:
    #         # result_finance_report_title.to_excel(w, sheet_name='Титульный лист отчета')
    #         df_realization_report.to_excel(w, sheet_name='Таблица отчета')

    return df_realization_report


# Функция перевода дат в часовой пояс UTC (для некоторых методов АПИ)
def msk_boundary_to_utc(date_str: str, end_of_day: bool = False) -> str:
    """
    Принимает:
    - date_str: 'YYYY-MM-DD' (трактуется как MSK, берёт 00:00:00 если end_of_day=False, 23:59:59 если True)
    - date_str вида 'YYYY-MM-DDTHH:MM:SS' (трактуется как MSK, время берётся как есть)
    - date_str вида 'YYYY-MM-DDTHH:MM:SSZ' (в UTC, время берётся как есть — НЕ ДЛЯ КРАЙНИХ ГРАНИЦ!)
    Возвращает строку UTC: YYYY-MM-DDTHH:MM:SSZ
    """
    # Только дата, без времени: трактуем как московские сутки
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str):
        time_part = "23:59:59" if end_of_day else "00:00:00"
        dt_full = f"{date_str} {time_part}"
        dt = pd.to_datetime(dt_full)
        dt = dt.tz_localize('Europe/Moscow')
        dt_utc = dt.tz_convert('UTC')
        return dt_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    # Дата и время, без Z
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", date_str):
        dt = pd.to_datetime(date_str)
        dt = dt.tz_localize('Europe/Moscow')
        dt_utc = dt.tz_convert('UTC')
        return dt_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    # Дата и время с Z (UTC)
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", date_str):
        # Уже в UTC — возвращаем как есть, но можем предупредить
        return date_str
    else:
        raise ValueError('Неподдерживаемый формат даты!')



# Функция выделения общего артикула размерного ряда
def get_common_article(df_products: pd.DataFrame) -> pd.Series:
    df_products_processed = df_products.copy()
    df_products_processed = df_products_processed.rename(columns={
        'Артикул продавца': 'Артикул'
    })

    if client_name in ['SENS_IP']:
        # Разделяем с конца через rsplit, берем суффикс
        last = df_products_processed['Артикул'].str.rsplit('_', n=1, expand=True)
        if last.shape[1] < 2:
            last = last.reindex(columns=[0, 1], fill_value=None)
        last[1] = last[1].astype(str).fillna('')
        mask = last[1].str.contains(r'[A-Za-zА-Яа-я]', na=False)
        df_products_processed['Артикул размерного ряда'] = (
            df_products_processed['Артикул']
            .where(mask, last[0])
        )

    elif client_name in ['SENS']:
        pattern = (
            r'^(?:'
            r'([^/\s]+?)(?=[/\s]\d)'             # короткий артикул типа УБ24
            r'|'
            r'(.+?)(?:(?<=\S)(?=[_\-\s]\d))'     # длинный артикул до разделителя
            r')'
        )
        df_products_processed['Артикул размерного ряда'] = (
            df_products_processed['Артикул']
            .str.extract(pattern, expand=False)
            .bfill(axis=1)
            .iloc[:, 0]
            .fillna(df_products_processed['Артикул'])
        )

    elif client_name in ['KMS']:
        size_pattern = re.compile(
            r'^(?P<base>.+?)'          # всё до последнего разделителя
            r'[\s/_-]'                 # разделитель: пробел, /, _, -
            r'(?P<size>'
            r'\d{2}'                   # 35, 40, 45, 46...
            r'|XXS|XS|S|M|L|XL|XXL'    # размеры одежды
            r'|NP'                     # как в "Шоколад_паста/NP"
            r'|[12]'                   # /1, /2
            r')\s*$',
            flags=re.IGNORECASE
        )

        df_products_processed['Артикул размерного ряда'] = (
            df_products_processed['Артикул']
            .str.replace(size_pattern, r'\g<base>', regex=True)
        )

    else:
        df_products_processed['Артикул размерного ряда'] = df_products_processed['Артикул']

    return df_products_processed['Артикул размерного ряда']


class OzonRateLimiter:
    """
    Простой лимитер для запросов к Ozon API.

    Параметры:
      interval          – интервал между УСПЕШНЫМИ запросами (сек),
      error_interval    – интервал ожидания после неуспешного запроса (сек),
      max_retries       – максимальное количество попыток при ошибках.
    """

    def __init__(self, interval: float = 0.2, error_interval: float = 2, max_retries: int = 3):
        self.interval = float(interval)
        self.error_interval = float(error_interval)
        self.max_retries = int(max_retries)
        self._last_request_ts = None

    def _sleep_if_needed(self):
        """Выдержать интервал между успешными запросами."""
        now = time.time()
        if self._last_request_ts is None:
            self._last_request_ts = now
            return

        dt = now - self._last_request_ts
        if dt < self.interval:
            sleep_time = self.interval - dt
            logger.debug(f"Выдерживаем интервал между запросами {self.interval:.2f} сек. "
                         f"Пауза {sleep_time:.2f} сек.")
            time.sleep(sleep_time)
        self._last_request_ts = time.time()

    @staticmethod
    def _log_ozon_error(resp: requests.Response):
        """
        Пытается красиво вывести ошибку Ozon.
        Если есть поля code / message / details — форматируем их.
        Иначе выводим сырой текст ответа.
        """
        try:
            data = resp.json()
        except Exception:
            logger.error(f"Ошибка Ozon (HTTP {resp.status_code}): {resp.text}")
            return

        code = data.get("code")
        message = data.get("message")
        details = data.get("details")

        # Если структура стандартная
        if code is not None or message is not None or details is not None:
            logger.error(f"Ошибка Ozon (HTTP {resp.status_code}): code={code}, message={message}")

            # details может быть списком объектов с typeUrl / value
            if isinstance(details, list) and details:
                for i, d in enumerate(details, start=1):
                    turl = d.get("typeUrl")
                    val = d.get("value")
                    logger.error(f"  detail[{i}]: typeUrl={turl}, value={val}")
        else:
            # Нестандартная структура
            logger.error(f"Ошибка Ozon (HTTP {resp.status_code}), тело ответа: {data}")

    def call(self, func, *args, **kwargs):
        """
        Вызов HTTP‑функции (session.get / session.post и т.п.) с учётом:
          - интервала между успешными запросами,
          - интервала после ошибки,
          - максимального числа повторов.

        Ожидается, что func(*args, **kwargs) возвращает requests.Response.

        Возвращает:
          (response, success_flag)
        """

        for attempt in range(1, self.max_retries + 1):
            # Интервал между ПОСЛЕДНИМ успешным и текущим запросом
            self._sleep_if_needed()

            try:
                resp = func(*args, **kwargs)
            except Exception as e:
                func_name = getattr(func, "__name__", "request")
                logger.error(
                    f"Ошибка соединения при вызове {func_name} "
                    f"(попытка {attempt}/{self.max_retries}): {e}"
                )
                if attempt == self.max_retries:
                    return None, False
                logger.info(f"Повтор запроса через {self.error_interval:.2f} сек.")
                time.sleep(self.error_interval)
                continue

            status = resp.status_code

            # Успех (считаем, что 2xx — успех)
            if 200 <= status < 300:
                # Фиксируем время успешного запроса для следующего интервала
                self._last_request_ts = time.time()
                return resp, True

            # Ошибка от Ozon
            self._log_ozon_error(resp)

            if attempt == self.max_retries:
                return resp, False

            logger.info(
                f"Ошибка Ozon (HTTP {status}), повтор запроса "
                f"{attempt + 1}/{self.max_retries} через {self.error_interval:.2f} сек."
            )
            time.sleep(self.error_interval)

        # Теоретически сюда не дойдём
        return None, False

# %%
