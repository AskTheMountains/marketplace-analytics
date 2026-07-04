# %%
import argparse
import warnings
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from rich_argparse.contrib import ParagraphRichHelpFormatter

from ozon.scripts.ozon_logger import ozon_client_logger as logger
from ozon.scripts.constants import (
    client_name,
    marketplace_dir_path,
    article_replacements
)
from ozon.scripts.ozon_generic_functions import (
    clean_ozon_article
)

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

    Заказы считаются из файлов Заказы_fbo/fbs.
    Продажи (шт) считаются из файла _Транзакции.csv по типам начислений:
        - "Продажи шт"
        - "Возвраты шт" (вычитаются из продаж).

    Дополнительно:
      - для каждого SKU считаются:
          * 'Продажи'              — чистые продажи в штуках за основной период
                                    (date_start..date_end) = Продажи шт - Возвраты шт
          * 'Продажи (факт) за 2 месяца' — чистые продажи в штуках за последние 2 месяца
                                    (последние 61 день до date_end включительно).
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

    date_start_raw = dates_from_to['date_start'][0]
    date_end_raw = dates_from_to['date_end'][0]

    # Преобразуем в date без времени
    period_start = pd.to_datetime(date_start_raw).date()
    period_end = pd.to_datetime(date_end_raw).date()
    # Хвост 2 месяцев (61 день до конца)
    tail_start = period_end - pd.Timedelta(days=61)

    logger.info(
        f"Calculation metrics for client {client_name} "
        f"for dates {period_start} - {period_end} "
        f"(2-month tail from {tail_start})"
    )

    # --------------------- 2. Загрузка каталога и его обработка ---------------------
    catalog = pd.read_csv(
        uploaddir / f"{date_upload_files}_Товары.csv",
        sep=';'
    )

    catalog = catalog.rename(columns={
        'Название товара': 'Наименование товара',
        'Barcode': 'Штрихкод',
        'Штрихкод (Серийный номер / EAN)': 'Штрихкод'
    })

    if 'Объемный вес, кг' in catalog.columns:
        if pd.api.types.is_string_dtype(catalog['Объемный вес, кг']):
            catalog['Объемный вес, кг'] = catalog['Объемный вес, кг'].str.replace("'", "", regex=False)

    if 'Артикул' in catalog.columns:
        catalog["Артикул"] = catalog["Артикул"].map(clean_ozon_article)

    # Удаляем дубликаты колонок, оставляем первую
    catalog = catalog.loc[:, ~catalog.columns.duplicated()]

    catalog = catalog[[
        'Артикул',
        'SKU',
        'Наименование товара',
        'Штрихкод',
        'Ozon Product ID',
    ]]

    # заранее создаём справочник SKU → Артикул
    sku_to_article = (
        catalog[['SKU', 'Артикул']]
        .dropna()
        .drop_duplicates()
    )

    # --------------------- 3. Загрузка остальных CSV ---------------------
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
    transactions = pd.read_csv(
        uploaddir / f"{date_upload_files}_Транзакции.csv",
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
    supply_order_items = pd.read_csv(
        uploaddir / f"{date_upload_files}_Поставки.csv",
        encoding='utf-8',
        sep=';'
    )

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
        logger.info(
            "Файл с прошлыми поставками не найден. "
            "Будут использованы только текущие поставки."
        )

    # --------------------- 4. Заказы (FBO/FBS) ---------------------
    if 'Объемный вес товаров, кг' in order_fbo.columns:
        order_fbo = order_fbo.drop(columns=['Объемный вес товаров, кг'])
    order_fbo['Схема доставки'] = 'FBO'
    order_fbs['Схема доставки'] = 'FBS'

    if not order_fbs.empty:
        order = pd.concat([order_fbo, order_fbs], axis=0, ignore_index=True)
    else:
        order = order_fbo.copy()

    order = order[
        ['Артикул', 'Кластер доставки', 'Статус', 'Схема доставки', 'Количество', 'Сумма отправления']
    ].rename(columns={'Кластер доставки': 'Кластер'})

    # нормализуем артикул в заказах через ту же функцию
    if 'Артикул' in order.columns:
        order['Артикул'] = order['Артикул'].map(clean_ozon_article)

    order['Заказы'] = np.where(order['Статус'].notnull(), order['Количество'], 0)
    order['Заказы, руб.'] = np.where(order['Статус'].notnull(), order['Сумма отправления'], 0)

    # --------------------- 5. Продажи из _Транзакции.csv ---------------------
    if transactions is None or transactions.empty:
        logger.warning("Файл _Транзакции.csv пуст или не найден, продажи будут нулевыми.")
        sales = pd.DataFrame(columns=['Артикул', 'Кластер', 'Продажи', 'Продажи (факт) за 2 месяца'])
    else:
        tr = transactions.copy()

        # Кластер: подстрой, если в файле другое имя
        if "Кластер доставки" in tr.columns:
            tr = tr.rename(columns={"Кластер доставки": "Кластер"})
        elif "Кластер" not in tr.columns:
            tr["Кластер"] = 0

        # SKU → Int64
        if "SKU" in tr.columns:
            tr["SKU"] = pd.to_numeric(tr["SKU"], errors="coerce").astype("Int64")

        # Дата транзакции
        date_col_tr = "Дата" if "Дата" in tr.columns else "operation_date"
        tr[date_col_tr] = pd.to_datetime(tr[date_col_tr], errors="coerce")
        tr["Дата"] = tr[date_col_tr].dt.tz_localize(None).dt.date

        type_col = "Тип начисления"
        value_col = "Итого"

        mask_sales_qty = tr[type_col] == "Продажи шт"
        mask_returns_qty = tr[type_col] == "Возвраты шт"

        tr_qty = tr.loc[mask_sales_qty | mask_returns_qty, :].copy()

        if tr_qty.empty:
            logger.warning(
                "В _Транзакции.csv нет строк с типами 'Продажи шт' / 'Возвраты шт'."
            )
            sales = pd.DataFrame(columns=['Артикул', 'Кластер', 'Продажи', 'Продажи (факт) за 2 месяца'])
        else:
            # Продажи шт (плюс) и Возвраты шт (плюс — нужно вычесть)
            tr_qty["Продажи_шт"] = np.where(
                tr_qty[type_col] == "Продажи шт",
                tr_qty[value_col],
                0,
            )
            tr_qty["Возвраты_шт"] = np.where(
                tr_qty[type_col] == "Возвраты шт",
                tr_qty[value_col],
                0,
            )

            # Чистые продажи шт по строке
            tr_qty["Продажи_чистые_шт"] = tr_qty["Продажи_шт"] - tr_qty["Возвраты_шт"]

            # Агрегация по SKU + Кластер + Дата
            qty_agg = (
                tr_qty
                .groupby(["SKU", "Кластер", "Дата"], as_index=False)
                .agg(
                    Продажи_чистые_шт=("Продажи_чистые_шт", "sum"),
                )
            )

            # Делим по периодам:
            #  - Основной период: period_start <= Дата <= period_end
            #  - Хвост 2 месяцев:  tail_start  <= Дата <= period_end
            qty_agg["Продажи_шт_период"] = np.where(
                (qty_agg["Дата"] >= period_start) & (qty_agg["Дата"] <= period_end),
                qty_agg["Продажи_чистые_шт"],
                0,
            )
            qty_agg["Продажи_шт_2м"] = np.where(
                (qty_agg["Дата"] >= tail_start) & (qty_agg["Дата"] <= period_end),
                qty_agg["Продажи_чистые_шт"],
                0,
            )

            # Теперь суммируем по SKU + Кластеру за весь период и за 2 месяца
            qty_agg_final = (
                qty_agg
                .groupby(["SKU", "Кластер"], as_index=False)
                .agg(
                    Продажи=('Продажи_шт_период', 'sum'),
                    Продажи_факт_2м=('Продажи_шт_2м', 'sum'),
                )
            )

            # Привязываем к Артикулу через каталог
            sales = qty_agg_final.merge(
                sku_to_article,
                on="SKU",
                how="left"
            )
            sales['Артикул'] = sales['Артикул'].map(clean_ozon_article)

            sales = sales[['Артикул', 'Кластер', 'Продажи', 'Продажи_факт_2м']]
            sales = sales.rename(columns={'Продажи_факт_2м': 'Продажи (факт) за 2 месяца'})

    # --------------------- 6. Остатки FBO ---------------------
    warehouses_fbo = warehouses_fbo.loc[
        :,
        warehouses_fbo.columns.isin(['Артикул', 'Кластер', 'Доступный к продаже товар'])
    ].rename(columns={'Доступный к продаже товар': 'Остатки'})

    if 'Артикул' in warehouses_fbo.columns:
        warehouses_fbo['Артикул'] = warehouses_fbo['Артикул'].map(clean_ozon_article)

    # --------------------- 7. Поставки (текущие + прошлые) ---------------------
    if not supply_order_items.empty:
        current_supplies = supply_order_items.copy()
        current_supplies = current_supplies.rename(columns={
            'Наименование товара': 'Название товара',
            'Количество': 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ',
        })
        current_supplies = current_supplies[['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ']]
        current_supplies['Артикул'] = current_supplies['Артикул'].map(clean_ozon_article)
    else:
        current_supplies = pd.DataFrame(
            columns=['Артикул', 'Кластер', 'ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ']
        )

    if not past_supply_order_items.empty:
        past_supplies = past_supply_order_items.copy()
        past_supplies = past_supplies.rename(columns={
            'Наименование товара': 'Название товара',
            'Количество': 'ПРОШЛЫЕ_ПОСТАВКИ',
        })
        past_supplies = past_supplies[['Артикул', 'Кластер', 'ПРОШЛЫЕ_ПОСТАВКИ']]
        past_supplies['Артикул'] = past_supplies['Артикул'].map(clean_ozon_article)
    else:
        past_supplies = pd.DataFrame(
            columns=['Артикул', 'Кластер', 'ПРОШЛЫЕ_ПОСТАВКИ']
        )

    # --------------------- 8. Остатки FBS ---------------------
    warehouses_fbs = warehouses_fbs.rename(
        columns={'Доступно на моем складе, шт': 'Остатки_fbs'}
    )
    if warehouses_fbs.get('Остатки_fbs', pd.Series(dtype=float)).sum() >= 1:
        warehouses_fbs = warehouses_fbs[['Артикул', 'Остатки_fbs']]
        warehouses_fbs['Артикул'] = warehouses_fbs['Артикул'].map(clean_ozon_article)
    else:
        warehouses_fbs = pd.DataFrame(columns=['Артикул', 'Остатки_fbs'])

    # --------------------- 9. Сбор всех данных по товарам ---------------------
    orderware = pd.concat(
        [
            order,
            sales,
            current_supplies,
            past_supplies,
            warehouses_fbo,
            warehouses_fbs,
        ],
        axis=0,
        ignore_index=True
    ).fillna(0)

    # --------------------- 10. Агрегация по кластеру и по артикулу ---------------------
    orderware_by_cluster = (
        orderware
        .groupby(['Артикул', 'Кластер'], as_index=False)
        .agg(
            Продажи=('Продажи', 'sum'),
            Продажи_факт_2м=('Продажи (факт) за 2 месяца', 'sum'),
            Заказы=('Заказы', 'sum'),
            Заказы_руб=('Заказы, руб.', 'sum'),
            ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
            ПРОШЛЫЕ_ПОСТАВКИ=('ПРОШЛЫЕ_ПОСТАВКИ', 'sum'),
            Остатки=('Остатки', 'sum'),
            Остатки_fbs=('Остатки_fbs', 'sum'),
        )
    )

    summary = (
        orderware
        .groupby(['Артикул'], as_index=False)
        .agg(
            Продажи=('Продажи', 'sum'),
            Продажи_факт_2м=('Продажи (факт) за 2 месяца', 'sum'),
            Заказы=('Заказы', 'sum'),
            Заказы_руб=('Заказы, руб.', 'sum'),
            ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ=('ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ', 'sum'),
            ПРОШЛЫЕ_ПОСТАВКИ=('ПРОШЛЫЕ_ПОСТАВКИ', 'sum'),
            Остатки=('Остатки', 'sum'),
            Остатки_fbs=('Остатки_fbs', 'sum'),
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
