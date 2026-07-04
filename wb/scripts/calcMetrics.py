# %%
import argparse
import warnings
from datetime import date

import numpy as np
import pandas as pd
# from loguru import logger
from wb.scripts.wb_logger import wb_client_logger as logger
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter

warnings.simplefilter(action='ignore', category=FutureWarning)
from wb.scripts.command_args_parser import cli_reference_date
# Файл с константами
from wb.scripts.constants import (
    cargo_type, client_name,
    article_replacement,
    ignore_virtual_warehouses,
    marketplace_dir_path
)

# Добавляем описание скрипта в парсер аргументов
parser = argparse.ArgumentParser(
    description="Calculate metrics for WB marketplace",
    usage=(
        "python.exe -m wb.scripts.calcMetrics [YYYY-MM-DD] [--ref_dt YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.calcMetrics 2025-05-28\n\n"
        "  python.exe -m wb.scripts.calcMetrics --ref_dt 2025-05-28 --client 2\n\n"
        "  python.exe -m wb.scripts.calcMetrics 2025-05-28 -c 3"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


def read_clusters_mapping(cargo_type_param: str) -> pd.DataFrame:
    """
    Читает файл соответствия кластеров WB.
    Если в файле есть лист с именем client_name, используем его.
    Иначе используем лист по cargo_type_param ('Монопалеты' или 'Короба').
    """
    mapping_path = marketplace_dir_path / "scripts" / "wb_warehouses_mapping.xlsx"

    # Открываем файл
    xls = pd.ExcelFile(mapping_path)

    # Определяем из какого листа считывать маппинг складов
    if client_name in xls.sheet_names:
        sheet_to_read = client_name
        logger.info(
            f"Найден лист '{client_name}' в wb_warehouses_mapping.xlsx. "
            f"Используем его для группировки складов."
        )
    else:
        sheet_mapping = {
            "monopallets": "Монопалеты",
            "boxes": "Короба",
        }
        sheet_to_read = sheet_mapping.get(cargo_type_param, "Короба")
        logger.info(
            f"Лист '{client_name}' не найден в wb_warehouses_mapping.xlsx. "
            f"Используем лист '{sheet_to_read}'."
        )

    clusters_mapping = pd.read_excel(xls, sheet_name=sheet_to_read)
    # На всякий случай удаляем дубликаты
    clusters_mapping = clusters_mapping.drop_duplicates(subset=["Склад"])
    # Убираем колонку с номером
    clusters_mapping = clusters_mapping.loc[:, ~clusters_mapping.columns.isin(['№'])]

    return clusters_mapping

def calcMetrics(date_upload_files: str = str(date.today())):
    """
    Расчёт метрик по выгруженным файлам Wildberries.

    Заказы считаются из файла _Заказы.csv.
    Продажи (шт) считаются из файла _Продажи.csv:
        - базовая метрика 'Продажи' — количество продаж за основной период
          (date_start..date_end);
        - 'Продажи (факт) за 2 месяца' — количество продаж за последние 2 месяца
          (хвост 61 день до date_end включительно).

    Агрегация:
      - по складам (лист 'claster_report');
      - по Артикул_Размер (лист 'summary').

    Структура выходных файлов максимально сохраняет текущий формат WB‑отчёта.
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
        sep=";",
    )

    date_start_raw = dates_from_to["date_start"][0]
    date_end_raw = dates_from_to["date_end"][0]

    # Преобразуем в date без времени
    period_start = pd.to_datetime(date_start_raw).date()
    period_end = pd.to_datetime(date_end_raw).date()
    # Хвост 2 месяцев (61 день до конца)
    tail_start = period_end - pd.Timedelta(days=61)

    logger.info(
        "Calculation metrics for client {} for dates {} - {} "
        "(2-month tail from {})",
        client_name,
        period_start,
        period_end,
        tail_start,
    )

    # --------------------- 2. Загрузка CSV ---------------------
    catalog = pd.read_csv(uploaddir / f"{date_upload_files}_Товары.csv", sep=";")
    order = pd.read_csv(uploaddir / f"{date_upload_files}_Заказы.csv", sep=";")
    sales = pd.read_csv(uploaddir / f"{date_upload_files}_Продажи.csv", sep=";")
    reminders = pd.read_csv(uploaddir / f"{date_upload_files}_Остатки.csv", sep=";")
    reminders_fbs = pd.read_csv(
        uploaddir / f"{date_upload_files}_Остатки_fbs.csv", sep=";"
    )
    supply_orders = pd.read_csv(
        uploaddir / f"{date_upload_files}_Поставки.csv", sep=";"
    )

    # --------------------- 3. Общая предварительная обработка ---------------------
    df_list = [catalog, order, sales, reminders, reminders_fbs, supply_orders]

    for i, df in enumerate(df_list):
        tmp_df = df_list[i]

        # Для некоторых клиентов: замена запятой на точку в размере и перевод в число
        if client_name in ["KU_And_KU", "Soyuz"]:
            if "Размер" in tmp_df.columns:
                tmp_df["Размер"] = (
                    tmp_df["Размер"]
                    .astype(str)
                    .str.replace(",", ".", regex=False)
                )
                tmp_df["Размер"] = tmp_df["Размер"].apply(
                    lambda x: pd.to_numeric(x, errors="coerce")
                )
                tmp_df["Размер"] = tmp_df["Размер"].fillna(0)

        # barcode как строка (длинные штрихкоды)
        if "barcode" in tmp_df.columns:
            tmp_df["barcode"] = tmp_df["barcode"].astype(str)

        # Удаляем ненужную колонку
        if "Unnamed: 0" in tmp_df.columns:
            tmp_df.drop(columns="Unnamed: 0", inplace=True)

        # Замена артикулов по словарю, если есть
        if article_replacement and "Артикул продавца" in tmp_df.columns:
            tmp_df["Артикул продавца"] = tmp_df["Артикул продавца"].replace(
                article_replacement
            )

        # Создаем колонку Артикул_Размер
        if {"Артикул продавца", "Размер"}.issubset(tmp_df.columns):
            tmp_df["Артикул_Размер"] = tmp_df[
                ["Артикул продавца", "Размер"]
            ].apply(
                lambda row: "_size_".join(row.values.astype(str)), axis=1
            )

    # --------------------- 4. Каталог ---------------------
    catalog = catalog.rename(
        columns={
            "nmID": "Артикул WB",
            "barcode": "barcode_list",
            "last_barcode": "Штрихкод",
            "discount_price": "РРЦ",
        }
    )
    # Удаляем дубликаты по Артикул_Размер
    catalog = catalog.drop_duplicates(subset=["Артикул_Размер"])
    # Удаляем Артикул продавца и Размер (мерджим по Артикул_Размер)
    catalog = catalog.drop(columns=["Артикул продавца", "Размер"])
    # Нумерация товаров
    catalog["№ товара"] = np.arange(1, catalog.shape[0] + 1)

    # --------------------- 5. Заказы ---------------------
    order["Заказы"] = 1
    order["Заказы_руб"] = np.where(order["Заказы"] == 1, order["priceWithDisc"], 0)

    # --------------------- 6. Продажи и Продажи (факт) за 2 месяца ---------------------
    # Базовая логика WB: Продажи = 1, если finishedPrice > 0
    sales["Продажи"] = np.where(sales["finishedPrice"] > 0, 1, 0)
    sales["Продажи_руб"] = sales["finishedPrice"]

    # Пытаемся определить дату продажи для деления по периодам
    # Приоритет полей даты: saleDate -> date -> lastChangeDate (как fallback)
    date_cols_priority = ["saleDate", "date", "lastChangeDate"]
    sales_date_col = None
    for col in date_cols_priority:
        if col in sales.columns:
            sales_date_col = col
            break

    if sales_date_col:
        sales["_ДатаПродажи"] = pd.to_datetime(
            sales[sales_date_col], errors="coerce"
        ).dt.date

        # Флаг продаж в основном периоде
        mask_period = (
            (sales["_ДатаПродажи"] >= period_start)
            & (sales["_ДатаПродажи"] <= period_end)
        )
        # Флаг продаж в 2-месячном хвосте
        mask_tail = (
            (sales["_ДатаПродажи"] >= tail_start)
            & (sales["_ДатаПродажи"] <= period_end)
        )

        # На всякий случай, если _ДатаПродажи пустая — считаем такие строки только в основном периоде
        mask_period = mask_period | sales["_ДатаПродажи"].isna()
        mask_tail = mask_tail | sales["_ДатаПродажи"].isna()

        # Продажи за период (сохраняем твою старую логику: 1 продажа = 1 шт)
        sales["Продажи_период"] = np.where(mask_period, sales["Продажи"], 0)
        # Продажи за 2 месяца
        sales["Продажи_факт_2м"] = np.where(mask_tail, sales["Продажи"], 0)
    else:
        logger.warning(
            "В таблице продаж WB не найдено ни одного поля даты "
            f"из {date_cols_priority}. 'Продажи (факт) за 2 месяца' "
            "будут равны нулю."
        )
        sales["Продажи_период"] = sales["Продажи"]
        sales["Продажи_факт_2м"] = 0

    # --------------------- 7. Остатки и поставки ---------------------
    reminders["Остатки"] = reminders["quantityFull"]
    reminders_fbs["Остатки_fbs"] = reminders_fbs["amount"]

    supply_orders = supply_orders.rename(
        columns={
            "Количество товара": "Ожидаемое поступление",
            "Склад": "warehouseName",
        }
    )

    # --------------------- 8. Выбор нужных колонок ---------------------
    order = order.loc[:, ["Артикул_Размер", "warehouseName", "Заказы", "Заказы_руб"]]

    # в sales оставляем и Продажи (за период), и Продажи_факт_2м
    sales = sales.loc[
        :,
        [
            "Артикул_Размер",
            "warehouseName",
            "Продажи_период",
            "Продажи_факт_2м",
            "Продажи_руб",
        ],
    ].rename(columns={"Продажи_период": "Продажи"})

    reminders = reminders.loc[:, ["Артикул_Размер", "warehouseName", "Остатки"]]
    reminders_fbs = reminders_fbs.loc[:, ["Артикул_Размер", "Остатки_fbs"]]
    supply_orders = supply_orders.loc[
        :, ["Артикул_Размер", "warehouseName", "Ожидаемое поступление"]
    ]

    # --------------------- 9. Объединение всего в одну таблицу ---------------------
    orderware = pd.concat(
        [order, sales, reminders, reminders_fbs, supply_orders],
        axis=0,
        ignore_index=True,
    )

    # TODO: Убираем виртуальные склады из списка складов
    if ignore_virtual_warehouses:
        orderware = orderware.loc[
            ~orderware["warehouseName"].str.contains("Виртуальный", na=False), :
        ]

    # --------------------- 10. Агрегация по складам ---------------------
    orderware_by_warehouses = (
        orderware.groupby(["Артикул_Размер", "warehouseName"])
        .agg(**{
            'Заказы': ("Заказы", "sum"),
            'Заказы_руб': ("Заказы_руб", "sum"),
            'Продажи': ("Продажи", "sum"),
            'Продажи_факт_2м': ("Продажи_факт_2м", "sum"),
            'Продажи_руб': ("Продажи_руб", "sum"),
            'Ожидаемое поступление': ("Ожидаемое поступление", "sum"),
            'Остатки': ("Остатки", "sum"),
            'Остатки_fbs': ("Остатки_fbs", "sum"),
        })
        .reset_index()
    )

    # --------------------- 11. Агрегация по Артикул_Размер ---------------------
    orderware_by_sku = (
        orderware.groupby(["Артикул_Размер"])
        .agg(**{
            'Заказы': ("Заказы", "sum"),
            'Заказы_руб': ("Заказы_руб", "sum"),
            'Продажи': ("Продажи", "sum"),
            'Продажи_факт_2м': ("Продажи_факт_2м", "sum"),
            'Продажи_руб': ("Продажи_руб", "sum"),
            'Ожидаемое поступление': ("Ожидаемое поступление", "sum"),
            'Остатки': ("Остатки", "sum"),
            'Остатки_fbs': ("Остатки_fbs", "sum"),
        })
        .reset_index()
    )

    # --------------------- 12. claster_report (по складам) ---------------------
    claster_report = catalog.merge(
        orderware_by_warehouses, on=["Артикул_Размер"], how="left"
    )

    # Достаем артикул и размер из Артикул_Размер
    claster_report = claster_report.assign(
        **{
            "Артикул продавца": claster_report["Артикул_Размер"]
            .str.split("_size_", expand=True)[0],
            "Размер": claster_report["Артикул_Размер"]
            .str.split("_size_", expand=True)[1],
        }
    )

    # Удаляем те товары, которые не были ни на одном складе
    claster_report = claster_report.dropna(subset=["warehouseName"])

    # Заполнение NA
    claster_report = claster_report.fillna(0)

    # Переименование
    claster_report = claster_report.rename(
        columns={
            "warehouseName": "Склад",
        }
    )

    claster_report = claster_report.sort_values(
        by=["Артикул продавца", "Размер", "Склад"], ignore_index=True
    )

    # Выборка нужных колонок
    claster_report = claster_report.loc[
        :,
        [
            "№ товара",
            "Артикул_Размер",
            "Артикул продавца",
            "Артикул WB",
            "Наименование товара",
            "Предмет",
            "Размер",
            "Цвет",
            "Штрихкод",
            "Склад",
            "Заказы",
            "Заказы_руб",
            "Продажи",
            "Продажи_факт_2м",
            "Продажи_руб",
            "Ожидаемое поступление",
            "Остатки",
            "Остатки_fbs",
        ],
    ]

    # --------------------- 13. summary (по Артикул_Размер) ---------------------
    summary = catalog.merge(orderware_by_sku, on=["Артикул_Размер"], how="left")

    summary = summary.assign(
        **{
            "Артикул продавца": summary["Артикул_Размер"]
            .str.split("_size_", expand=True)[0],
            "Размер": summary["Артикул_Размер"]
            .str.split("_size_", expand=True)[1],
        }
    )

    summary = summary.fillna(0)

    summary = summary.loc[
        :,
        [
            "№ товара",
            "Артикул_Размер",
            "Артикул продавца",
            "Артикул WB",
            "Наименование товара",
            "Предмет",
            "Размер",
            "Цвет",
            "Штрихкод",
            "Заказы",
            "Заказы_руб",
            "Продажи",
            "Продажи_факт_2м",
            "Продажи_руб",
            "Ожидаемое поступление",
            "Остатки",
            "Остатки_fbs",
        ],
    ]

    # --------------------- 14. Сохранение в Excel ---------------------
    metrics_file_name = metricsdir / f"{date_upload_files}_МетрикиИтоги.xlsx"

    with pd.ExcelWriter(metrics_file_name) as w:
        summary.to_excel(w, sheet_name="summary", index=False)
        claster_report.to_excel(w, sheet_name="claster_report", index=False)

    logger.info("Done")
    logger.info(f"Metrics saved in: {metrics_file_name}")

# %%
if __name__ == '__main__':
    metricsdir = marketplace_dir_path / "Clients" / client_name / "Metrics"
    if not metricsdir.exists():
        metricsdir.mkdir(parents=True, exist_ok=True)
    date_upload_files = cli_reference_date(parser)
    calcMetrics(date_upload_files)

# %%
