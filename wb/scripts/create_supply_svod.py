"""
Модуль для расчета отчета по поставкам для маркетплейса Wildberries.
Скрипт создает Excel-файл с расчетами потребности в товарах по складам,
анализируя продажи, заказы и остатки за определенный период.
"""

import argparse
from datetime import date, datetime
from itertools import product
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
from rich_argparse.contrib import ParagraphRichHelpFormatter

from wb.scripts.wb_logger import wb_client_logger as logger
import wb.scripts.format_supply_svod as format_supply_svod
from wb.scripts.command_args_parser import cli_reference_date
from wb.scripts.constants import (
    cargo_type,
    catalog_supply_columns,
    cargo_multiplicity,
    DEMAND_PERIOD_1,
    DEMAND_PERIOD_2,
    client_name,
    is_heavy_cargo,
    marketplace_dir_path,
    shop_type,
    supply_svod_clusters_columns,
    supply_svod_total_columns,
)

# Парсер аргументов командной строки
parser = argparse.ArgumentParser(
    description="Расчет отчета по поставкам (WB)",
    usage=(
        "python.exe -m wb.scripts.create_supply_svod [YYYY-MM-DD] "
        "[--ref_dt YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m wb.scripts.create_supply_svod 2025-01-01\n\n"
        "  python.exe -m wb.scripts.create_supply_svod --ref_dt 2025-01-01 --client 2\n\n"
        "  python.exe -m wb.scripts.create_supply_svod 2025-01-01 -c 3"
    ),
    formatter_class=ParagraphRichHelpFormatter,
)


# ============================================================================
# ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ
# ============================================================================

DAYS_IN_PERIOD = 30

APPLY_BOX_ROUNDING = cargo_multiplicity["flag_multiplicity"]
DEFAULT_BOX_MULTIPLICITY = cargo_multiplicity["multiplicity"]

UPLOAD_DIR = marketplace_dir_path / "Clients" / client_name / "UploadFiles"
UPLOAD_DIR_TODAY = UPLOAD_DIR / f"UploadFiles_{date.today()}"
METRICS_DIR = marketplace_dir_path / "Clients" / client_name / "Metrics"
CATALOG_DIR = marketplace_dir_path / "Clients" / client_name / "catalog"
SUPPLY_SVOD_DIR = marketplace_dir_path / "Clients" / client_name / "SupplySvod"

# ============================================================================
# ФУНКЦИИ ЧТЕНИЯ ДАННЫХ
# ============================================================================


def read_dates_file(date_metrics: str) -> pd.DataFrame:
    """
    Читает файл с датами формирования отчета.
    """
    file_path = (
        UPLOAD_DIR
        / f"UploadFiles_{date_metrics}"
        / f"{date_metrics}_dates_from_to.csv"
    )
    report_dates = pd.read_csv(file_path, sep=";")
    for col in report_dates.columns:
        report_dates[col] = pd.to_datetime(report_dates[col])
    return report_dates


def read_metrics_file(
    date_metrics: str,
    sheet_name: str = "summary",
) -> pd.DataFrame:
    """
    Читает файл с метриками.
    """
    file_path = METRICS_DIR / f"{date_metrics}_МетрикиИтоги.xlsx"
    metrics = pd.read_excel(file_path, sheet_name=sheet_name)
    return metrics.copy()


def read_catalog() -> pd.DataFrame:
    """
    Читает и обрабатывает справочную таблицу WB.
    """
    catalog_path = CATALOG_DIR / f"Справочная_таблица_{client_name}_WB.xlsx"
    catalog = pd.read_excel(catalog_path)
    catalog_processed = catalog.copy()

    if "Размер" not in catalog_processed.columns:
        catalog_processed["Размер"] = 0

    if "Кратность коробов" not in catalog_processed.columns:
        logger.info(
            "Колонка 'Кратность коробов' не найдена в справочной "
            f"таблице. Добавляем с дефолтным значением: {DEFAULT_BOX_MULTIPLICITY}."
        )
        catalog_processed["Кратность коробов"] = DEFAULT_BOX_MULTIPLICITY

    catalog_processed["Артикул_Размер"] = (
        catalog_processed[["Артикул продавца", "Размер"]]
        .apply(lambda row: "_size_".join(row.values.astype(str)), axis=1)
    )

    catalog_processed = catalog_processed.dropna(subset=["Размер"])

    for col in catalog_supply_columns:
        if col not in catalog_processed.columns:
            catalog_processed[col] = np.nan

    if client_name in ["KU_And_KU", "Soyuz"]:
        catalog_processed["Размер"] = catalog_processed["Размер"].apply(
            lambda x: pd.to_numeric(x, errors="coerce")
        )
        catalog_processed["Размер"] = catalog_processed["Размер"].fillna(0)

    return catalog_processed


def read_clusters_mapping(cargo_type_param: str) -> pd.DataFrame:
    """
    Читает файл соответствия кластеров (складов).
    """
    sheet_mapping = {
        "monopallets": "Монопалеты",
        "boxes": "Короба",
    }
    sheet_name = sheet_mapping.get(cargo_type_param, "Короба")
    mapping_path = marketplace_dir_path / "scripts" / "wb_warehouses_mapping.xlsx"
    clusters_mapping = pd.read_excel(mapping_path, sheet_name=sheet_name)
    clusters_mapping = clusters_mapping.drop_duplicates(subset=["Склад"])
    return clusters_mapping


# ============================================================================
# ФУНКЦИИ ОБРАБОТКИ КЛАСТЕРОВ
# ============================================================================


def add_missing_clusters(date_metrics: str) -> pd.DataFrame:
    """
    Добавляет недостающие комбинации Артикул_Размер–Склад
    (аналог add_clusters в Ozon).
    """
    metrics_clusters = read_metrics_file(
        date_metrics,
        sheet_name="claster_report",
    )
    metrics_summary = read_metrics_file(
        date_metrics,
        sheet_name="summary",
    )

    sku_clusters = list(
        product(
            metrics_summary["Артикул_Размер"].unique(),
            metrics_clusters["Склад"].unique(),
        )
    )
    sku_clusters_df = pd.DataFrame(
        sku_clusters,
        columns=["Артикул_Размер", "Склад"],
    )

    metrics_merged = sku_clusters_df.merge(
        metrics_clusters[
            [
                "Артикул_Размер",
                "Склад",
                "Продажи",
                "Заказы",
                "Ожидаемое поступление",
                "Остатки",
            ]
        ],
        on=["Артикул_Размер", "Склад"],
        how="left",
    )

    fill_columns = ["Продажи", "Заказы", "Ожидаемое поступление", "Остатки"]
    metrics_merged[fill_columns] = metrics_merged[fill_columns].fillna(0)

    metrics_merged = metrics_merged.merge(
        metrics_summary.drop(
            columns=[
                "Продажи",
                "Заказы",
                "Ожидаемое поступление",
                "Остатки",
                "Остатки_fbs",
            ]
        ),
        on="Артикул_Размер",
        how="left",
    )
    return metrics_merged


def get_cluster_column_from_mapping(
    is_heavy_cargo_flag: bool,
    specific_mapping: Optional[str] = None,
) -> str:
    """
    Определяет нужную колонку для группировки по складам.
    """
    if specific_mapping:
        return "Группировка ювелирка"
    return "Группировка СГТ" if is_heavy_cargo_flag else "Группировка"


def replace_clusters(
    clusters_mapping_df: pd.DataFrame,
    cluster_col: str,
    df_with_clusters: pd.DataFrame,
) -> pd.DataFrame:
    """
    Объединяет несколько складов в один кластер согласно маппингу.
    Похоже на ozon.replace_clusters, но использует таблицу Excel с маппингом.
    """
    df_merged = df_with_clusters.copy()

    df_merged = df_merged.merge(clusters_mapping_df, how="left", on="Склад")

    for col in clusters_mapping_df.columns:
        if col in df_merged.columns:
            df_merged[col] = df_merged[col].fillna("Неизвестный склад")

    df_merged = df_merged.drop(columns="Склад").rename(
        columns={cluster_col: "Склад"}
    )

    group_columns = [
        "Артикул_Размер",
        "№ товара",
        "Склад",
        "Предмет",
        "Наименование товара",
        "Цвет",
        "Штрихкод",
    ]
    sum_columns = ["Продажи", "Заказы", "Ожидаемое поступление", "Остатки"]

    df_aggregated = (
        df_merged[group_columns + sum_columns]
        .groupby(group_columns)
        .sum()
        .reset_index()
    )
    return df_aggregated


# ============================================================================
# ФУНКЦИИ РАСЧЕТА ПОКАЗАТЕЛЕЙ
# ============================================================================


def add_demand_columns(
    df: pd.DataFrame,
    stock: pd.Series,
    days: int,
    label: str = "",
) -> pd.DataFrame:
    """
    Добавляет колонки потребности и дефицита/избытка (как в Ozon).
    Создает:
      - 'Потребность на {days} дней{label_suffix}'
      - 'Дефицит/Избыток {days} дней{label_suffix}'
      - 'Потребность на {days} дней{label_suffix} (округл.)'
    """
    suffix = f" {label}" if label else ""
    need_col = f"Потребность на {days} дней{suffix}"
    deficit_col = f"Дефицит/Избыток {days} дней{suffix}"
    need_rounded_col = f"{need_col} (округл.)"

    df[need_col] = df["Оборачиваемость"] * days
    df[deficit_col] = stock - df[need_col]

    df[need_rounded_col] = np.where(
        df[deficit_col] > 0,
        0,
        np.ceil(df[deficit_col].abs()),
    )
    return df


def calc_svod_for_clusters(
    metrics_df_with_clusters: pd.DataFrame,
) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости по кластерам (лист 'По кластерам').
    Логика аналогична Ozon: оборачиваемость, потребность на два периода,
    корректировка по второму периоду.
    """
    df = metrics_df_with_clusters.copy()

    df["Оборачиваемость"] = (
        (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
    )

    stock = df["Остатки"] + df["Ожидаемое поступление"]

    df = add_demand_columns(df, stock, DEMAND_PERIOD_1)
    df = add_demand_columns(df, stock, DEMAND_PERIOD_2)

    df["Корректировка с учетом нулевых остатков"] = df[
        f"Потребность на {DEMAND_PERIOD_2} дней (округл.)"
    ]

    df["№"] = np.arange(1, len(df) + 1)
    return df


def create_pivot_clusters(
    svod_for_clusters: pd.DataFrame,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Создание сводной таблицы по кластерам (аналог Ozon):
    - 'Потребность <Склад>' по 'Корректировка с учетом нулевых остатков'
    - 'Остаток <Склад>' по 'Остатки'

    index: Артикул_Размер
    columns: Склад
    """
    df_pivot = pd.pivot_table(
        svod_for_clusters,
        columns="Склад",
        index="Артикул_Размер",
        values=["Корректировка с учетом нулевых остатков", "Остатки"],
        aggfunc="sum",
        fill_value=0,
    )

    df_pivot.columns = [
        (
            f"Потребность {col[1]}"
            if col[0] == "Корректировка с учетом нулевых остатков"
            else f"Остаток {col[1]}"
            if col[0] == "Остатки"
            else col[0]
        )
        for col in df_pivot.columns.values
    ]

    df_pivot = df_pivot.reset_index()

    clusters = svod_for_clusters["Склад"].unique().tolist()
    cols_order = ["Артикул_Размер"] + [
        val
        for cluster in clusters
        for val in (f"Потребность {cluster}", f"Остаток {cluster}")
    ]
    cols_order = [c for c in cols_order if c in df_pivot.columns]

    df_pivot = df_pivot[cols_order]
    return df_pivot, clusters


def calc_svod_by_sku(metrics_df_with_clusters: pd.DataFrame) -> pd.DataFrame:
    """
    Расчет агрегированного свода по SKU (без разреза по складам).
    """
    metrics_df_by_sku = (
        metrics_df_with_clusters.groupby(
            ["Артикул продавца", "Размер", "Наименование товара", "Предмет", "РРЦ"]
        )
        .agg(
            Продажи=("Продажи", "sum"),
            Заказы=("Заказы", "sum"),
            Остатки=("Остатки", "sum"),
            Ожидаемое_поступление=("Ожидаемое поступление", "sum"),
        )
        .reset_index()
        .rename(columns={"Ожидаемое_поступление": "Ожидаемое поступление"})
    )
    return metrics_df_by_sku


def calc_svod_for_sku(metrics_df_by_sku: pd.DataFrame) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости для товаров по SKU (лист 'Всего').
    """
    df = metrics_df_by_sku.copy()

    df["Оборачиваемость"] = (
        (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
    )

    stock_with_fbs = (
        df["Остатки"] + df["Остатки_fbs"] + df["Ожидаемое поступление"]
    )

    df = add_demand_columns(df, stock_with_fbs, DEMAND_PERIOD_1)
    df = add_demand_columns(
        df,
        stock_with_fbs,
        DEMAND_PERIOD_1,
        "с учетом FBS",
    )
    df = add_demand_columns(df, stock_with_fbs, DEMAND_PERIOD_2)

    df["№"] = np.arange(1, len(df) + 1)
    return df


# ============================================================================
# ФУНКЦИИ ОБЪЕДИНЕНИЯ ДАННЫХ
# ============================================================================


def add_demand_for_cluster_columns(
    df_pivot_by_clusters: pd.DataFrame,
    metrics_df_by_sku_with_stats: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в таблицу по SKU колонки с потребностью/остатком по кластерам
    через merge с pivot-таблицей по 'Артикул_Размер' (аналог Ozon).
    """
    df_result = metrics_df_by_sku_with_stats.copy()
    df_result = (
        df_result.merge(
            df_pivot_by_clusters,
            how="left",
            on="Артикул_Размер",
        )
        .fillna(0)
    )
    return df_result


def add_columns_from_catalog(
    metrics_df_by_sku_with_stats: pd.DataFrame,
    catalog: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет колонки из справочной таблицы (по Артикул_Размер).
    """
    df_merged = metrics_df_by_sku_with_stats.copy()

    df_merged = df_merged.merge(
        catalog[catalog_supply_columns],
        how="left",
        on=["Артикул_Размер"],
    )

    fill_values = {"Кратность коробов": DEFAULT_BOX_MULTIPLICITY}
    df_merged = df_merged.fillna(value=fill_values)
    return df_merged


def add_columns_for_excel(
    df_with_stats: pd.DataFrame,
    cluster_list: List[str],
    date_metrics: str,
    day_start: datetime,
    day_end: datetime,
    svod_type: str = "Всего",
) -> pd.DataFrame:
    """
    Подготавливает DataFrame для записи в Excel (как в Ozon).
    """
    df = df_with_stats.copy()

    df["Артикул продавца"] = df["Артикул_Размер"].str.split("_size_").str[0]
    df["Размер"] = df["Артикул_Размер"].str.split("_size_").str[1]

    if svod_type == "Всего":
        svod_columns = supply_svod_total_columns.copy()
        cluster_columns = [
            val
            for cluster in cluster_list
            for val in (f"Потребность {cluster}", f"Остаток {cluster}")
        ]
        svod_columns.extend(cluster_columns)
    else:
        svod_columns = supply_svod_clusters_columns.copy()

    for col in svod_columns:
        if col not in df.columns:
            df[col] = np.nan

    df = df[svod_columns]
    df = df.sort_values(by=["Артикул продавца", "Размер"])

    date_metrics_dt = datetime.strptime(date_metrics, "%Y-%m-%d")
    rename_mapping = {
        "Наименование товара": "Наименование",
        "Заказы": f"ЗАКАЗЫ с {day_start.strftime('%d.%m')} по {day_end.strftime('%d.%m')}, шт.",
        "Продажи": f"ПРОДАЖИ с {day_start.strftime('%d.%m')} по {day_end.strftime('%d.%m')}, шт.",
        "Остатки": f"ОСТАТОК {date_metrics_dt.strftime('%d.%m')}",
        "Остатки_fbs": f"ОСТАТОК FBS {date_metrics_dt.strftime('%d.%m')}",
    }
    df = df.rename(columns=rename_mapping)
    return df


# ============================================================================
# ФУНКЦИИ ДЛЯ СПЕЦИАЛЬНЫХ СЛУЧАЕВ (ЮВЕЛИРКА)
# ============================================================================


def create_jewelry_clusters_svod(
    clusters_mapping_df: pd.DataFrame,
    metrics_df_all_clusters: pd.DataFrame,
    cluster_column: str,
    date_metrics: str,
    catalog: pd.DataFrame,
) -> Optional[pd.DataFrame]:
    """
    Создает отчет 'Всего по ювелирным складам' (WB),
    по аналогии с общим подходом:
      - replace_clusters по ювелирному маппингу
      - calc_svod_for_clusters
      - create_pivot_clusters
      - merge pivot в свод по SKU.
    """
    if shop_type != "jewelry":
        return None

    metrics_df_union_clusters = replace_clusters(
        clusters_mapping_df,
        cluster_column,
        metrics_df_all_clusters,
    )

    metrics_df_with_stats = calc_svod_for_clusters(
        metrics_df_union_clusters,
    )

    df_pivot_by_clusters, cluster_list = create_pivot_clusters(
        metrics_df_with_stats
    )

    metrics_df_by_sku = read_metrics_file(date_metrics)
    metrics_df_by_sku_stats = calc_svod_for_sku(metrics_df_by_sku)

    svod = add_demand_for_cluster_columns(
        df_pivot_by_clusters,
        metrics_df_by_sku_stats,
    )

    svod_with_catalog = add_columns_from_catalog(svod, catalog)

    report_dates = read_dates_file(date_metrics)

    svod_excel = add_columns_for_excel(
        svod_with_catalog,
        cluster_list,
        date_metrics,
        report_dates["date_start_file"][0],
        report_dates["date_end_file"][0],
        svod_type="Всего",
    )
    return svod_excel


# ============================================================================
# ФУНКЦИИ СОХРАНЕНИЯ И ФОРМАТИРОВАНИЯ
# ============================================================================


def save_sheets_to_excel(
    svod_by_sku: pd.DataFrame,
    svod_excel_by_sku_jewelry: Optional[pd.DataFrame],
    svod_for_clusters: pd.DataFrame,
    clusters_mapping_df: pd.DataFrame,
    date_metrics: str,
) -> None:
    """
    Сохраняет данные в Excel-файл (сырые данные, без форматирования).
    """
    SUPPLY_SVOD_DIR.mkdir(parents=True, exist_ok=True)

    filepath = (
        SUPPLY_SVOD_DIR
        / f"{date_metrics}_Расчет_поставок_{client_name}_WB.xlsx"
    )

    with pd.ExcelWriter(
        filepath,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": True}},
    ) as writer:
        svod_by_sku.to_excel(writer, sheet_name="Всего", index=False)
        if svod_excel_by_sku_jewelry is not None:
            svod_excel_by_sku_jewelry.to_excel(
                writer,
                sheet_name="Всего по ювелирным складам",
                index=False,
            )
        svod_for_clusters.to_excel(
            writer,
            sheet_name="По кластерам",
            index=False,
        )
        clusters_mapping_df.to_excel(
            writer,
            sheet_name="Таблица соотв. складов",
            index=False,
        )

    logger.info(f"Данные сохранены в: {filepath}")


def format_excel_file(
    client_name: str,
    date_metrics: str,
    clusters_mapping_df: pd.DataFrame,
    cluster_column: str,
    cluster_column_jewelry: str,
    svod_excel_sku: pd.DataFrame,
    svod_excel_by_sku_jewelry: Optional[pd.DataFrame],
    svod_excel_clusters: pd.DataFrame,
) -> None:
    """
    Форматирует Excel-файл (формулы, стили) через wb.scripts.format_supply_svod.
    """
    format_supply_svod.copy_supply_svod_file(
        SUPPLY_SVOD_DIR,
        client_name,
        date_metrics,
    )

    format_supply_svod.format_sheet_clusters(
        path_supply_svod=SUPPLY_SVOD_DIR,
        client_name=client_name,
        date_report_created=date_metrics,
        svod_excel=svod_excel_clusters,
        sheet_name="По кластерам",
        demand_period_1=DEMAND_PERIOD_1,
        demand_period_2=DEMAND_PERIOD_2,
        days_in_period=DAYS_IN_PERIOD,
        apply_box_rounding=APPLY_BOX_ROUNDING,
    )

    format_supply_svod.format_sheet_total(
        path_supply_svod=SUPPLY_SVOD_DIR,
        client_name=client_name,
        date_report_created=date_metrics,
        svod_excel=svod_excel_sku,
        svod_excel_clusters=svod_excel_clusters,
        clusters_mapping_df=clusters_mapping_df,
        cluster_column=cluster_column,
        sheet_name="Всего",
        demand_period_1=DEMAND_PERIOD_1,
        demand_period_2=DEMAND_PERIOD_2,
        days_in_period=DAYS_IN_PERIOD,
        apply_box_rounding=APPLY_BOX_ROUNDING,
    )

    if svod_excel_by_sku_jewelry is not None:
        format_supply_svod.format_sheet_total(
            path_supply_svod=SUPPLY_SVOD_DIR,
            client_name=client_name,
            date_report_created=date_metrics,
            svod_excel=svod_excel_by_sku_jewelry,
            svod_excel_clusters=svod_excel_clusters,
            clusters_mapping_df=clusters_mapping_df,
            cluster_column=cluster_column_jewelry,
            sheet_name="Всего по ювелирным складам",
            demand_period_1=DEMAND_PERIOD_1,
            demand_period_2=DEMAND_PERIOD_2,
            days_in_period=DAYS_IN_PERIOD,
            apply_box_rounding=APPLY_BOX_ROUNDING,
        )

    filename = (
        SUPPLY_SVOD_DIR
        / f"{date_metrics}_Расчет_поставок_{client_name}_WB_formatted.xlsx"
    )
    logger.info(f"Отформатированный отчет сохранен: {filename}")


# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================


def create_supply_svod_report(date_metrics: str) -> None:
    """
    Основная функция создания отчета по поставкам WB.
    """
    logger.info(
        f"Создание отчета поставок для клиента {client_name} на дату {date_metrics}"
    )
    logger.info(f"Период оборачиваемости: {DAYS_IN_PERIOD} дней")
    logger.info(
        f"Периоды потребности: {DEMAND_PERIOD_1} дней и {DEMAND_PERIOD_2} дней"
    )
    logger.info(
        "Округление по кратности коробов в Excel: "
        f"{'включено' if APPLY_BOX_ROUNDING else 'отключено'}"
    )

    catalog = read_catalog()

    # Данные по всем кластерам (сырые)
    metrics_df_all_clusters = read_metrics_file(
        date_metrics,
        sheet_name="claster_report",
    )

    # Лист "По кластерам" (без объединения складов)
    svod_for_clusters = calc_svod_for_clusters(metrics_df_all_clusters)

    # Маппинг складов
    clusters_mapping_df = read_clusters_mapping(cargo_type)
    cluster_column = get_cluster_column_from_mapping(is_heavy_cargo)
    cluster_column_jewelry = get_cluster_column_from_mapping(
        is_heavy_cargo,
        specific_mapping=True,
    )

    # Объединяем склады в кластеры
    metrics_df_union_clusters = replace_clusters(
        clusters_mapping_df,
        cluster_column,
        metrics_df_all_clusters,
    )

    metrics_df_with_stats = calc_svod_for_clusters(metrics_df_union_clusters)

    # Pivot по кластерам (потребность + остаток), как на Ozon
    df_pivot_by_clusters, cluster_list = create_pivot_clusters(
        metrics_df_with_stats
    )

    # Обработка данных по SKU
    metrics_df_by_sku = read_metrics_file(date_metrics)
    metrics_df_by_sku_stats = calc_svod_for_sku(metrics_df_by_sku)

    # Добавляем потребность/остатки по кластерам в свод по SKU
    svod = add_demand_for_cluster_columns(
        df_pivot_by_clusters,
        metrics_df_by_sku_stats,
    )

    # Добавляем данные из каталога
    svod_with_catalog_by_sku = add_columns_from_catalog(svod, catalog)
    svod_with_catalog_for_clusters = add_columns_from_catalog(
        svod_for_clusters,
        catalog,
    )

    report_dates = read_dates_file(date_metrics)

    svod_excel_by_sku = add_columns_for_excel(
        svod_with_catalog_by_sku,
        cluster_list,
        date_metrics,
        report_dates["date_start_file"][0],
        report_dates["date_end_file"][0],
        svod_type="Всего",
    )

    svod_excel_by_sku_jewelry = create_jewelry_clusters_svod(
        clusters_mapping_df,
        metrics_df_all_clusters,
        cluster_column_jewelry,
        date_metrics,
        catalog,
    )

    svod_excel_for_clusters = add_columns_for_excel(
        svod_with_catalog_for_clusters,
        cluster_list,
        date_metrics,
        report_dates["date_start_file"][0],
        report_dates["date_end_file"][0],
        svod_type="По кластерам",
    )

    save_sheets_to_excel(
        svod_excel_by_sku,
        svod_excel_by_sku_jewelry,
        svod_excel_for_clusters,
        clusters_mapping_df,
        date_metrics,
    )

    format_excel_file(
        client_name,
        date_metrics,
        clusters_mapping_df,
        cluster_column,
        cluster_column_jewelry,
        svod_excel_by_sku,
        svod_excel_by_sku_jewelry,
        svod_excel_for_clusters,
    )

    logger.info(
        f"Отчет успешно создан для {client_name} на дату {date_metrics}"
    )


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================

if __name__ == "__main__":
    date_metrics = cli_reference_date(parser)
    create_supply_svod_report(date_metrics)
