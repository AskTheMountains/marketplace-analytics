# %%
import argparse
import warnings
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from rich_argparse.contrib import ParagraphRichHelpFormatter

from ozon.scripts.ozon_logger import ozon_client_logger as logger
from ozon.scripts.constants import client_name, marketplace_dir_path
from ozon.scripts.command_args_parser import cli_reference_date

warnings.simplefilter(action='ignore', category=FutureWarning)
pd.set_option('future.no_silent_downcasting', True)

# Директория для сохранения метрик
metricsdir = marketplace_dir_path / "Clients" / client_name / "Metrics"
metricsdir.mkdir(parents=True, exist_ok=True)

# Парсер аргументов
parser = argparse.ArgumentParser(
    description="Calculate metrics for Ozon marketplace",
    usage=(
        "python.exe -m ozon.scripts.calcMetrics "
        "[YYYY-MM-DD] [--ref_dt YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.calcMetrics 2024-11-13\n\n"
        "  python.exe -m ozon.scripts.calcMetrics --ref_dt 2024-11-13 -c 2\n\n"
        "  python.exe -m ozon.scripts.calcMetrics 2024-11-13 --client 3"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


def calcMetrics(date_upload_files: str = str(date.today())):
    """
    Расчёт метрик по выгруженным файлам Ozon.
    """

    # --------------------- 1. Пути и даты ---------------------

    uploaddir = (
        marketplace_dir_path
        / "Clients"
        / client_name
        / "UploadFiles"
        / f"UploadFiles_{date_upload_files}"
    )

    dates_from_to = pd.read_csv(
        uploaddir / f"{date_upload_files}_dates_from_to.csv",
        sep=';'
    )
    date_start = dates_from_to['date_start'][0]
    date_end = dates_from_to['date_end'][0]
    logger.info(
        f"Calculation metrics for client {client_name} "
        f"for dates {date_start} - {date_end}"
    )

    # --------------------- 2. Загрузка CSV ---------------------

    catalog = pd.read_csv(
        uploaddir / f"{date_upload_files}_Товары.csv",
        sep=';'
    )

    order_fbo = pd.read_csv(
        uploaddir / f"{date_upload_files}_Заказы_fbo.csv",
        encoding='utf-8',
        sep=';'
    )
    order_fbs = pd.read_csv(
        uploaddir / f"{date_upload_files}_Заказы_fbs.csv",
        encoding='utf-8',
        sep=';'
    )

    sales_fbo = pd.read_csv(
        uploaddir / f"{date_upload_files}_Продажи_fbo.csv",
        encoding='utf-8',
        sep=';'
    )
    sales_fbs = pd.read_csv(
        uploaddir / f"{date_upload_files}_Продажи_fbs.csv",
        encoding='utf-8',
        sep=';'
    )

    warehouses_fbo = pd.read_csv(
        uploaddir / f"{date_upload_files}_Остатки_fbo.csv",
        encoding='utf-8',
        sep=';'
    )
    warehouses_fbs = pd.read_csv(
        uploaddir / f"{date_upload_files}_Остатки_fbs.csv",
        encoding='utf-8',
        sep=';'
    )

    # Текущие поставки
    supply_order_items = pd.read_csv(
        uploaddir / f"{date_upload_files}_Поставки.csv",
        encoding='utf-8',
        sep=';'
    )

    # Прошлые поставки (если файл есть). Имя при необходимости подстрой под свой пайплайн.
    past_supplies_file = (
        uploaddir / f"{date_upload_files}_Поставки_прошлые.csv"
    )
    if past_supplies_file.exists():
        past_supply_order_items = pd.read_csv(
            past_supplies_file,
            encoding='utf-8',
            sep=';'
        )
        logger.info("Файл с прошлыми поставками найден и будет учтён в метриках.")
    else:
        past_supply_order_items = pd.DataFrame()
        logger.info("Файл с прошлыми поставками не найден. Будут использованы только текущие поставки.")

    # --------------------- 3. Заказы (FBO/FBS) ---------------------

    # Удаляем ненужные колонки из FBO
    if 'Объемный вес товаров, кг' in order_fbo.columns:
        order_fbo = order_fbo.drop(columns=['Объемный вес товаров, кг'])

    order_fbo['Схема доставки'] = 'FBO'
    order_fbs['Схема доставки'] = 'FBS'

    # Объединяем заказы FBO и FBS
    if not order_fbs.empty:
        order = pd.concat([order_fbo, order_fbs], axis=0, ignore_index=True)
    else:
        order = order_fbo.copy()

    order = order[
        ['Артикул', 'Кластер доставки', 'Статус', 'Схема доставки', 'Количество', 'Сумма отправления']
    ].rename(columns={'Кластер доставки': 'Кластер'})

    # Число заказов и сумма по заказам (учитываем только строки со статусом)
    order['Заказы'] = np.where(order['Статус'].notnull(), order['Количество'], 0)
    order['Заказы, руб.'] = np.where(order['Статус'].notnull(), order['Сумма отправления'], 0)

    # --------------------- 4. Продажи (FBO/FBS) ---------------------

    sales = pd.concat([sales_fbo, sales_fbs], axis=0, ignore_index=True)
    sales['Продажи'] = sales['Количество']
    sales = sales.rename(columns={'Кластер доставки': 'Кластер'})
    sales = sales[['Артикул', 'Кластер', 'Продажи']]

    # --------------------- 5. Остатки FBO ---------------------

    warehouses_fbo = warehouses_fbo.loc[
        :,
        warehouses_fbo.columns.isin(['Артикул', 'Кластер', 'Доступный к продаже товар'])
    ].rename(columns={'Доступный к продаже товар': 'Остатки'})

    # --------------------- 6. Поставки (текущие + прошлые) ---------------------

    # текущие → ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ
    if not supply_order_items.empty:
        current_supplies = supply_order_items.rename(columns={
            'Наименование товара': 'Название товара',
            'Количество': 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ'
        })[['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ']]
    else:
        current_supplies = pd.DataFrame(
            columns=['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ']
        )

    # прошлые → ПРОШЛЫЕ_ПОСТАВКИ
    if not past_supply_order_items.empty:
        past_supplies = past_supply_order_items.rename(columns={
            'Наименование товара': 'Название товара',
            'Количество': 'ПРОШЛЫЕ_ПОСТАВКИ'
        })[['Артикул', 'Кластер', 'ПРОШЛЫЕ_ПОСТАВКИ']]
    else:
        past_supplies = pd.DataFrame(
            columns=['Артикул', 'Кластер', 'ПРОШЛЫЕ_ПОСТАВКИ']
        )


    # --------------------- 7. Остатки FBS ---------------------

    warehouses_fbs = warehouses_fbs.rename(
        columns={'Доступно на моем складе, шт': 'Остатки_fbs'}
    )

    if warehouses_fbs.get('Остатки_fbs', pd.Series(dtype=float)).sum() >= 1:
        warehouses_fbs = warehouses_fbs[['Артикул', 'Остатки_fbs']]
    else:
        warehouses_fbs = pd.DataFrame(columns=['Артикул', 'Остатки_fbs'])

    # --------------------- 8. Каталог ---------------------

    catalog = catalog.rename(columns={
        'Название товара': 'Наименование товара',
        'Barcode': 'Штрихкод',
    })

    # Убираем лишние символы и приводим артикул к строке
    if 'Объемный вес, кг' in catalog.columns:
        # если это строковый столбец — чистим, иначе пропустим
        if pd.api.types.is_string_dtype(catalog['Объемный вес, кг']):
            catalog['Объемный вес, кг'] = catalog['Объемный вес, кг'].str.replace("'", "", regex=False)

    if 'Артикул' in catalog.columns:
        catalog['Артикул'] = catalog['Артикул'].astype(str).str.replace("'", "", regex=False)

    catalog = catalog[[
        'Артикул',
        'SKU',
        'Наименование товара',
        'Штрихкод',
        'Ozon Product ID',
    ]]

    # --------------------- 9. Сбор всех данных по товарам ---------------------

    # Собираем единый датафрейм: заказы, продажи, поставки, остатки FBO и FBS
    orderware = pd.concat(
        [
            order,
            sales,
            current_supplies,   # ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ
            past_supplies,      # ПРОШЛЫЕ_ПОСТАВКИ
            warehouses_fbo,
            warehouses_fbs
        ],
        axis=0,
        ignore_index=True
    )

    orderware = orderware.fillna(0)
    orderware['Артикул'] = orderware['Артикул'].astype(str)

    # --------------------- 10. Агрегация по кластеру и по артикулу ---------------------

    # По артикулам и кластерам
    orderware_by_cluster = (
        orderware
        .groupby(['Артикул', 'Кластер'], as_index=False)
        .agg(
            Продажи=('Продажи', 'sum'),
            Заказы=('Заказы', 'sum'),
            Заказы_руб=('Заказы, руб.', 'sum'),
            ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
            ПРОШЛЫЕ_ПОСТАВКИ=('ПРОШЛЫЕ_ПОСТАВКИ', 'sum'),
            Остатки=('Остатки', 'sum'),
            Остатки_fbs=('Остатки_fbs', 'sum')  # FBS без кластеров, но агрегируем по всем
        )
    )

    # По артикулам (без кластеров)
    summary = (
        orderware
        .groupby(['Артикул'], as_index=False)
        .agg(
            Продажи=('Продажи', 'sum'),
            Заказы=('Заказы', 'sum'),
            Заказы_руб=('Заказы, руб.', 'sum'),
            ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
            ПРОШЛЫЕ_ПОСТАВКИ=('ПРОШЛЫЕ_ПОСТАВКИ', 'sum'),
            Остатки=('Остатки', 'sum'),
            Остатки_fbs=('Остатки_fbs', 'sum')
        )
    )

    # --------------------- 11. Join с каталогом ---------------------

    catalog_with_orderware = catalog.merge(
        orderware_by_cluster,
        on='Артикул',
        how='left'
    ).fillna(0)

    catalog_with_orderware = catalog_with_orderware.sort_values(
        by=['Артикул', 'Кластер'],
        ascending=True
    )

    catalog_with_summary = catalog.merge(
        summary,
        on='Артикул',
        how='left'
    ).fillna(0)

    catalog_with_summary = catalog_with_summary.sort_values(
        by=['Артикул'],
        ascending=True
    )

    # Только строки с кластером (для кластерного отчёта)
    claster_report = catalog_with_orderware[
        catalog_with_orderware['Кластер'] != 0
    ].reset_index(drop=True)

    # --------------------- 12. Сохранение отчёта ---------------------

    metrics_file_name = metricsdir / f"{date_upload_files}_МетрикиИтоги.xlsx"

    with pd.ExcelWriter(metrics_file_name, engine='xlsxwriter') as writer:
        catalog_with_summary.to_excel(writer, sheet_name='summary', index=False)
        claster_report.to_excel(writer, sheet_name='claster_report', index=False)

    logger.info(f"Metrics saved in: {metrics_file_name}")
    logger.info('Done')

# %%
if __name__ == '__main__':
    date_upload_files = cli_reference_date(parser)
    calcMetrics(date_upload_files)
