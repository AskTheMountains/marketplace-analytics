import argparse
import warnings
from datetime import date

import numpy as np
import pandas as pd
from loguru import logger
from rich_argparse import RichHelpFormatter

warnings.simplefilter(action='ignore', category=FutureWarning)
pd.set_option('future.no_silent_downcasting', True)
# Файл с константами
from ozon.scripts.constants import client_name, marketplace_dir_name

# Создаем директорию, куда будут сохраняться метрики
metricsdir = marketplace_dir_name / "Clients" / client_name / "Metrics"
metricsdir.mkdir(parents=True, exist_ok=True)

from ozon.scripts.command_args_parser import cli_reference_date

# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Calculate metrics for Ozon marketplace",
    usage='python.exe -m ozon.scripts.calcMetrics [YYYY-MM-DD] [--ref_dt YYYY-MM-DD]',
    epilog='Example: python.exe -m ozon.scripts.calcMetrics 2024-11-13',
    formatter_class=RichHelpFormatter
)


# date_upload_files = '2024-11-13'

def calcMetrics(date_upload_files = str(date.today())):
    # Директории с загруженными файлами из апи
    uploaddir = (
        marketplace_dir_name
        / "Clients"
        / client_name
        / "UploadFiles"
        / f"UploadFiles_{date_upload_files}"
    )
    dates_from_to = pd.read_csv(uploaddir / f"{date_upload_files}_dates_from_to.csv", sep=';')
    logger.info(f"Calculation metrics for client {client_name} for dates {dates_from_to['date_start'][0]} - {dates_from_to['date_end'][0]}")

    # Загрузка данных из CSV файлов
    catalog = pd.read_csv(uploaddir / f"{date_upload_files}_Товары.csv", sep=';')
    order_fbo = pd.read_csv(uploaddir / f"{date_upload_files}_Заказы_fbo.csv", encoding='utf-8', sep=';')
    order_fbs = pd.read_csv(uploaddir / f"{date_upload_files}_Заказы_fbs.csv", encoding='utf-8', sep=';')
    sales_fbo = pd.read_csv(uploaddir / f"{date_upload_files}_Продажи_fbo.csv", encoding='utf-8', sep=';')
    sales_fbs = pd.read_csv(uploaddir / f"{date_upload_files}_Продажи_fbs.csv", encoding='utf-8', sep=';')
    warehouses_fbo = pd.read_csv(uploaddir / f"{date_upload_files}_Остатки_fbo.csv", encoding='utf-8', sep=';')
    warehouses_fbs = pd.read_csv(uploaddir / f"{date_upload_files}_Остатки_fbs.csv", encoding='utf-8', sep=';')
    supply_order_items = pd.read_csv(uploaddir / f"{date_upload_files}_Поставки.csv", encoding='utf-8', sep=';')

    # Объединение данных заказов FBO и FBS
    order_fbo = order_fbo.loc[:, ~order_fbo.columns.isin(['Объемный вес товаров, кг'])]
    order_fbo['Схема доставки'] = 'FBO'
    order_fbs['Схема доставки'] = 'FBS'
    if not order_fbs.empty:
        order = pd.concat([order_fbo, order_fbs], axis=0)
    else:
        order = order_fbo

    # Обработка данных заказов
    order = order[['Артикул', 'Кластер доставки', 'Статус', 'Схема доставки', 'Сумма отправления']]
    order = order.rename(columns={'Кластер доставки': 'Кластер'})
    order['Заказы'] = np.where(order['Статус'].notnull(), 1, 0)
    order['Заказы, руб.'] = np.where(order['Заказы'] == 1, order['Сумма отправления'], 0)
    pd.set_option('display.max_columns', None)

    # Объединение данных продаж FBO и FBS
    sales = pd.concat([sales_fbo, sales_fbs], axis=0)
    sales['Продажи'] = sales['Количество']
    sales = sales.rename(columns={'Кластер_доставки': 'Кластер'})
    sales = sales.loc[:, ['Артикул', 'Кластер', 'Продажи']]

    # Удаление ненужных колонок и фильтрация складских данных
    # warehouses = warehouses[warehouses['Доступный к продаже товар'] != 0]
    warehouses_fbo = warehouses_fbo.loc[:, warehouses_fbo.columns.isin([
        'Артикул',
        'Кластер',
        'Доступный к продаже товар'
    ])]
    warehouses_fbo = warehouses_fbo.rename(columns={
        # 'Артикул продавца': 'Артикул',
        'Доступный к продаже товар': 'Остатки'
    })

    # Обработка данных поставок
    # Если есть заявки на поставку, добавляем их в отчет
    if supply_order_items.shape[0] > 0:
        supply_order_items = supply_order_items.rename(columns={'Наименование товара': 'Название товара',
                                                                'quantity': 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ'})
        supply_order_items = supply_order_items.loc[:, ['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ']]
    # Если нет, то создаем пустой df, чтобы потом не было ошибок в concat
    else:
        supply_order_items = pd.DataFrame(columns=['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ'])

    # Обработка данных остатков на складах продавца
    # Если есть остатки на FBS складах, добавляем их в отчет
    warehouses_fbs = warehouses_fbs.rename(columns={'Доступно на моем складе, шт': 'Остатки_fbs'})
    if warehouses_fbs['Остатки_fbs'].sum() >= 1:
        warehouses_fbs = warehouses_fbs.loc[:, ['Артикул', 'Остатки_fbs']]
    # Если нет, то создаем пустой df, чтобы потом не было ошибок в concat
    else:
        warehouses_fbs = pd.DataFrame(columns=['Артикул', 'Остатки_fbs'])

    # Обработка данных каталога
    catalog = catalog.rename(columns={
        'Название товара': 'Наименование товара',
        'Barcode': 'Штрихкод',
    })
    # Убираем лишние символы
    for col in ['Артикул', 'Объемный вес, кг']:
        catalog[col] = catalog[col].str.replace("'", "", regex=False)
    # Переводим артикул в строку
    catalog['Артикул'] = catalog['Артикул'].astype(str)
    # Выбираем нужные колонки
    catalog = catalog[[
        'Артикул',
        'SKU',
        'Наименование товара',
        'Штрихкод',
        'Ozon Product ID',
    ]]

    # Объединение данных заказов и складских остатков
    orderware = pd.concat([order, sales, supply_order_items, warehouses_fbo, warehouses_fbs])
    # Заполняем пропуски
    orderware = orderware.fillna(0)
    # Переводим артикул в строку
    orderware['Артикул'] = orderware['Артикул'].astype(str)
    # Считаем сумму по артикулам и кластерам
    orderware_by_cluster = orderware.groupby(['Артикул', 'Кластер']).agg(
        Продажи=('Продажи', 'sum'),
        Заказы=('Заказы', 'sum'),
        Заказы_руб=('Заказы, руб.', 'sum'),
        # Товары_в_пути=('Товары в пути', 'sum'),
        ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
        Остатки=('Остатки', 'sum'),
        # У остатков fbs нет разбивки по кластерам
        Остатки_fbs=('Остатки_fbs', 'sum')
    ).reset_index()
    # Считаем сумму по артикулам
    summary = orderware.groupby(['Артикул']).agg(
        Продажи=('Продажи', 'sum'),
        Заказы=('Заказы', 'sum'),
        Заказы_руб=('Заказы, руб.', 'sum'),
        # Товары_в_пути=('Товары в пути', 'sum'),
        ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
        Остатки=('Остатки', 'sum'),
        # У остатков fbs нет разбивки по кластерам
        Остатки_fbs=('Остатки_fbs', 'sum')
    ).reset_index()


    # Объединение данных каталога с данными по кластерам
    catalog_with_orderware = catalog.merge(orderware_by_cluster, on='Артикул', how='left')
    catalog_with_orderware = catalog_with_orderware.fillna(0)
    # Сортировка для наглядности
    # catalog_with_orderware = catalog_with_orderware.sort_values(by=['Заказы', 'Остатки'], ascending=False)
    catalog_with_orderware = catalog_with_orderware.sort_values(by=['Артикул', 'Кластер'], ascending=True)

    # Объединение данных каталога с данными по артикулам
    catalog_with_summary = catalog.merge(summary, on='Артикул', how='left')
    catalog_with_summary = catalog_with_summary.fillna(0)
    # Сортировка для наглядности
    # catalog_with_orderware = catalog_with_orderware.sort_values(by=['Заказы', 'Остатки'], ascending=False)
    catalog_with_summary = catalog_with_summary.sort_values(by=['Артикул'], ascending=True)


    # Фильтрация данных по наличию кластеров
    claster_report = catalog_with_orderware[catalog_with_orderware['Кластер'] != 0].reset_index(drop=True)
    # Сохранение итоговых данных в файл Excel с несколькими листами
    metrics_file_name = metricsdir / f"{date_upload_files}_МетрикиИтоги.xlsx"
    writer = pd.ExcelWriter(metrics_file_name, engine='xlsxwriter')

    catalog_with_summary.to_excel(writer, sheet_name='summary', index=False)
    claster_report.to_excel(writer, sheet_name='claster_report', index=False)

    # Закрытие записи в файл Excel
    writer.close()

    logger.info(f"Metrics saved in: {metrics_file_name}")
    # os.remove('orderware.xlsx')
    logger.info('Done')


if __name__ == '__main__':
    date_upload_files = cli_reference_date(parser)
    calcMetrics(date_upload_files)
