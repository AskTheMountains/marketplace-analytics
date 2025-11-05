
# %% Определение функций
import json
from pathlib import Path

import pandas as pd
import requests
import re

# Папка, где лежит текущий скрипт
BASE_DIR = Path(__file__).parent.parent

# Некоторые константы
from ozon.scripts.constants import (
    client_name
)

# Функции выгрузки данных по API Seller
from ozon.scripts.uploadDataFromOzon import (
    getOrders,
    getStockReminders_fbo_v2,
    getStockReminders_fbs
)


# Функция выгрузки отчета о заказах
def get_orders(headers, date_start, date_end):
    # Выгружаем отчет об отправлениях fbo и fbs
    df_orders_fbo = getOrders(headers, date_start, date_end, delivery_schema='fbo', to_save=False)
    df_orders_fbs = getOrders(headers, date_start, date_end, delivery_schema='fbs', to_save=False)
    # Убираем лишние колонки из отчета fbo
    df_orders_fbo = df_orders_fbo.loc[:, ~df_orders_fbo.columns.isin(['Объемный вес товаров, кг'])]

    # Объединяем два отчета в один
    common_columns = ['Артикул', 'Название товара', 'OZON id', 'Статус', 'Принят в обработку', 'Сумма отправления']

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
    df_orders_all['datetime_orders'] = pd.to_datetime(df_orders_all['Принят в обработку'])
    # Указываем, какие колонки считаем за заказы в штуках и рублях
    df_orders_all['Заказы шт'] = 1
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
def get_reminders_by_product(df_products, headers, to_agg=True, agg_col=['Артикул']):
    # Выгружаем остатки FBO и FBS
    df_reminders_fbo = getStockReminders_fbo_v2(headers, df_products, to_save=False)
    df_reminders_fbs = getStockReminders_fbs(headers, to_save=False)
    # Переименовываем колонки для удобства
    df_reminders_fbo = df_reminders_fbo.rename(columns={
        'Доступный к продаже товар': 'Остатки FBO',
    })
    df_reminders_fbs = df_reminders_fbs.rename(columns={
        'Доступно на моем складе, шт': 'Остатки FBS',
    })
    # Объединяем остатки FBO и FBS в один df
    df_reminders_all = pd.concat([
        df_reminders_fbo.loc[:, ['Артикул', 'Остатки FBO']],
        df_reminders_fbs.loc[:, ['Артикул', 'Остатки FBS']],
    ])

    # Переводим колонку с артикулом в строку
    df_reminders_all['Артикул'] = df_reminders_all['Артикул'].astype(str)

    # Если стоит флаг агрегации, то считаем сумму остатков
    if to_agg:
        # Считаем остатки FBO + FBS
        df_reminders_total = (
            df_reminders_all
            .groupby(agg_col, as_index=False)
            .sum()
            .assign(**{
                'Остаток': lambda df:df[['Остатки FBO', 'Остатки FBS']].sum(axis=1)
            })
            # .reset_index()
        )
        return df_reminders_total
    # Если флаг не стоит, то возвращаем просто остатки
    else:
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

    # Создаем копию для избежания изменений в оригинальном df
    df_products_processed = df_products.copy()
    # Переименовываем колонку с артикулом, если называется по-другому
    df_products_processed = df_products_processed.rename(columns={
        'Артикул продавца': 'Артикул'
    })

    if client_name in ['SENS_IP']:
        # GPT START----
        # Разделяем с конца через rsplit, берем суффикс
        last = df_products_processed['Артикул'].str.rsplit('_', n=1, expand=True)
        # (если нет "_", во второй части будет NaN)
        if last.shape[1] < 2:
            last = last.reindex(columns=[0,1], fill_value=None)
        # Явно приводим второй столбец к строке
        last[1] = last[1].astype(str).fillna('')
        # Проверка: есть ли буквы (русские/латинские) во второй части
        mask = last[1].str.contains(r'[A-Za-zА-Яа-я]', na=False)
        # Если есть буквы — берем оригинал, иначе — часть до последнего '_'
        df_products_processed['Артикул размерного ряда'] = (
            df_products_processed['Артикул']
            .where(mask, last[0])
        )
    elif client_name in ['SENS']:
        pattern = (
            r'^(?:'
            r'([^/\s]+?)(?=[/\s]\d)'             # 1) короткий артикул типа УБ24 перед / или пробелом+число
            r'|'
            r'(.+?)(?:(?<=\S)(?=[_\-\s]\d))'      # 2) длинный артикул до разделителя перед размером
            r')'
        )
        df_products_processed['Артикул размерного ряда'] = (
            df_products_processed['Артикул']
            .str.extract(pattern, expand=False)   # даст DF из двух колонок
            .bfill(axis=1)                        # заполняет NaN справа на основе непустой
            .iloc[:, 0]                           # выбираем полученный артикул
            .fillna(df_products_processed['Артикул'])  # если не совпало — возвращаем оригинал

        )
    else:
        df_products_processed['Артикул размерного ряда'] = df_products_processed['Артикул']

    return df_products_processed['Артикул размерного ряда']
