# %%
import argparse
from datetime import datetime
from itertools import product
from pathlib import Path
from re import I
from typing import List, Optional

import numpy as np
import pandas as pd
from rich_argparse.contrib import ParagraphRichHelpFormatter

from generic_functions import add_element_to_list
from ozon.scripts.ozon_logger import ozon_client_logger as logger
import ozon.scripts.format_supply_svod_RiverBoats as format_supply_svod
from ozon.scripts.command_args_parser import cli_reference_date
from ozon.scripts.constants import (
    client_name,
    default_clusters_mapping,
    client_clusters_overrides,
    default_cluster_planning_days,
    client_cluster_planning_overrides,
    clusters_mapping_jewelry,
    marketplace_dir_path,
    shop_type,
    catalog_supply_columns,
    cargo_multiplicity,
    DEMAND_PERIOD_1,
    DEMAND_PERIOD_2,
    demand_period_for_total,
    include_previous_supplies_flag,
    supply_svod_clusters_columns,
    supply_svod_total_columns,
)

# ============================================================================
# ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ
# ============================================================================

# Базовый период для расчёта оборачиваемости
DAYS_IN_PERIOD = 30

# Применять ли округление по кратности коробов (в Excel-формулах)
APPLY_BOX_ROUNDING = cargo_multiplicity['flag_multiplicity']
# Дефолтная кратность коробов для клиента (если не найдена в справочнике)
DEFAULT_BOX_MULTIPLICITY = cargo_multiplicity['multiplicity']
# Включать ли предыдущие поставки в расчёт потребности
INCLUDE_PREVIOUS_SUPPLIES = False
if include_previous_supplies_flag:
    INCLUDE_PREVIOUS_SUPPLIES = include_previous_supplies_flag
    logger.info(f"Для клиента {client_name} включено добавление предыдущих поставок")

# Директории
MARKETPLACE_PATH = Path(marketplace_dir_path)
UPLOAD_DIR = MARKETPLACE_PATH / "Clients" / client_name / "UploadFiles"
METRICS_DIR = MARKETPLACE_PATH / "Clients" / client_name / "Metrics"
CATALOG_DIR = MARKETPLACE_PATH / "Clients" / client_name / "catalog"
SUPPLY_SVOD_DIR = MARKETPLACE_PATH / "Clients" / client_name / "SupplySvod"
special_clusters_mapping_ = {
    # --- по таблице ---
    # СПб и СЗО (ключ из таблицы) → кластер из исходного словаря
    'СПб и СЗО': ['Санкт-Петербург и СЗО'],

    # Москва, МО и ДР → Москва, МО и Дальние регионы + Беларусь
    'Москва, МО и ДР': [
        'Москва, МО и Дальние регионы',
        'Беларусь',
    ],

    # Ярославль → Ярославль + Тверь
    'Ярославль': ['Ярославль', 'Тверь'],

    # Казань
    'Казань': ['Казань'],

    # Уфа
    'Уфа': ['Уфа'],

    # Воронеж → Воронеж + Саратов
    'Воронеж': ['Воронеж', 'Саратов'],

    # Красноярск → Красноярск + Дальний Восток
    'Красноярск': ['Красноярск', 'Дальний Восток'],

    # Самара
    'Самара': ['Самара'],

    # Тюмень → Тюмень + Астана + Алматы
    'Тюмень': ['Тюмень', 'Астана', 'Алматы'],

    # Ростов → Ростов + Краснодар + Невинномысск + Махачкала
    'Ростов': ['Ростов', 'Краснодар', 'Невинномысск', 'Махачкала', 'Армения'],

    # Кавказ (если добавишь в исходный clusters_mapping)
    # 'Кавказ': ['Кавказ'],

    # Пермь → Пермь + Екатеринбург
    'Пермь': ['Пермь', 'Екатеринбург'],

    # Калининград
    'Калининград': ['Калининград'],

    # Омск → Омск + Новосибирск
    'Омск': ['Омск', 'Новосибирск'],

    # Оренбург
    'Оренбург': ['Оренбург'],

    # --- остальные из исходного clusters_mapping, не упомянутые в таблице ---
    # Кластеры остаются "как были" (без объединения):

    # Тверь уже включена в 'Ярославль', но если хочешь,
    # можно либо не добавлять её как отдельный агрегат,
    # либо оставить — зависит от того, нужен ли отдельный кластер на листе "Всего".
    # Здесь я предполагаю, что агрегаты определяются только таблицей,
    # поэтому отдельного ключа 'Тверь' не добавляю.

    # Санкт-Петербург и СЗО уже участвует в 'СПб и СЗО' как "подкластер",
    # сам отдельным агрегатом не нужен.

    # Москва, МО и Дальние регионы и Беларусь уже посчитаны в 'Москва, МО и ДР';

    # Из исходного словаря остались только те, кто нигде не фигурировали в таблице:
    # Армения, и (опционально) Казахстан, Кавказ.
    # 'Армения': ['Армения'],
    # Если раскомментируешь в исходном словаре:
    # 'Казахстан': ['Казахстан'],
    # 'Кавказ': ['Кавказ'],
}

# ============================================================================
# ПАРСЕР АРГУМЕНТОВ
# ============================================================================

parser = argparse.ArgumentParser(
    description="Calculate supply report (Ozon)",
    usage=(
        "python.exe -m ozon.scripts.create_supply_svod [YYYY-MM-DD] "
        "[--ref_dt YYYY-MM-DD] [-c N | --client N]"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.create_supply_svod 2024-11-13\n\n"
        "  python.exe -m ozon.scripts.create_supply_svod --ref_dt 2024-11-13 --client 2\n\n"
        "  python.exe -m ozon.scripts.create_supply_svod 2024-11-13 -c 3"
    ),
    formatter_class=ParagraphRichHelpFormatter,
)

# ============================================================================
# ФУНКЦИИ ЧТЕНИЯ И ПОДГОТОВКИ ДАННЫХ
# ============================================================================


def read_products_file(date_metrics: str) -> pd.DataFrame:
    """
    Чтение файла с товарами Ozon.
    """
    filepath = (
        UPLOAD_DIR
        / f"UploadFiles_{date_metrics}"
        / f"{date_metrics}_Товары.csv"
    )
    df_products = pd.read_csv(filepath, sep=";")
    return df_products


def process_product_list(df_products: pd.DataFrame) -> pd.DataFrame:
    """
    Обработка списка товаров: определение сортируемости, очистка артикула.
    """
    df = df_products.copy()

    df = df.loc[
        :,
        df.columns.isin(
            [
                "Артикул",
                "Категория",
                "Super-товар",
                "Высота",
                "Глубина",
                "Ширина",
                "Единицы измерения габаритов",
                "Вес",
                "Единицы измерения веса",
            ]
        ),
    ]

    df["Артикул"] = (
        df["Артикул"]
        .astype(str)
        .str.replace("'", "", regex=False)
    )

    # sortable = {
    #     "длина": 450,
    #     "ширина": 320,
    #     "высота": 320,
    #     "вес": 10000,
    # }
    # unsortable = {
    #     "длина_min": 450,
    #     "длина_max": 1200,
    #     "ширина_min": 320,
    #     "ширина_max": 800,
    #     "высота_min": 320,
    #     "высота_max": 600,
    #     "вес_max": 35000,
    # }

    # conditions = [
    #     (df["Вес"] <= sortable["вес"])
    #     & (df["Глубина"] <= sortable["длина"])
    #     & (df["Ширина"] <= sortable["ширина"])
    #     & (df["Высота"] <= sortable["высота"]),
    #     (df["Вес"] <= unsortable["вес_max"])
    #     & df["Глубина"].between(unsortable["длина_min"], unsortable["длина_max"])
    #     & df["Ширина"].between(unsortable["ширина_min"], unsortable["ширина_max"])
    #     & df["Высота"].between(unsortable["высота_min"], unsortable["высота_max"]),
    # ]
    # choices = ["Сортируемый товар", "Несортируемый товар"]

    # df["Сортируемость товара"] = np.select(
    #     conditions,
    #     choices,
    #     default="Не определена",
    # )

    df = df.drop(
        columns=[
            "Вес",
            "Глубина",
            "Ширина",
            "Высота",
            "Единицы измерения габаритов",
            "Единицы измерения веса",
        ]
    )
    return df


def read_dates_file(date_metrics: str) -> pd.DataFrame:
    """
    Чтение файла с датами формирования выгрузок.
    """
    filepath = (
        UPLOAD_DIR
        / f"UploadFiles_{date_metrics}"
        / f"{date_metrics}_dates_from_to.csv"
    )
    report_dates = pd.read_csv(filepath, sep=";")
    for col in report_dates:
        report_dates[col] = pd.to_datetime(report_dates[col])
    return report_dates


def read_metrics_file(date_metrics: str, sheet_name: str = "summary") -> pd.DataFrame:
    """
    Чтение файла с метриками.
    Удаляем колонки 'Размер', 'Цвет', 'РРЦ', если они есть.
    """
    filepath = METRICS_DIR / f"{date_metrics}_МетрикиИтоги.xlsx"
    if not filepath.exists():
        logger.error(f"Metrics file not found: {filepath}")

    metrics = pd.read_excel(filepath, sheet_name=sheet_name)

    for col in ["Размер", "Цвет", "РРЦ"]:
        if col in metrics.columns:
            metrics.drop(columns=col, inplace=True)
    return metrics


def read_catalog() -> pd.DataFrame:
    """
    Чтение и подготовка справочной таблицы Ozon.
    """
    catalog_path = (
        CATALOG_DIR
        / f"Справочная_таблица_{client_name}_Ozon.xlsx"
    )
    catalog = pd.read_excel(catalog_path)

    # Делаем первую букву в названиях столбцов заглавной, остальные - строчными
    catalog.columns = [col.capitalize() for col in catalog.columns]

    # Переименовываем ещё некоторые колонки
    catalog = catalog.rename(columns={
        'Sku': 'SKU',
        "Артикул продавца": "Артикул",
        "Артикул озон": "Артикул",
        "Сортируемость": "Сортируемость товара",
        "Статус товара - переносим в отчет": "Статус",
        'Ррц': 'РРЦ',
        'Кратность короба': 'Кратность коробов',
    })
    # Удаляем дубликаты колонок, если есть. Оставляем первую встречающуюся.
    catalog = catalog.loc[:, ~catalog.columns.duplicated()]

    # Переводим артикул в строку
    catalog["Артикул"] = catalog["Артикул"].astype(str)

    # Добавляем колонку с кратностью коробов
    if 'Кратность коробов' not in catalog.columns:
        logger.info(f"Колонка 'Кратность коробов' не найдена в справочной таблице. Добавляем с дефолтным значением: {DEFAULT_BOX_MULTIPLICITY}.")
        catalog['Кратность коробов'] = DEFAULT_BOX_MULTIPLICITY


    missing_columns = []
    for col in catalog_supply_columns:
        if col not in catalog.columns:
            catalog[col] = np.nan
            missing_columns.append(col)

    if missing_columns:
        logger.warning(
            f"Отсутствуют колонки в справочной таблице: {missing_columns}"
        )

    return catalog


# ============================================================================
# РАБОТА С КЛАСТЕРАМИ
# ============================================================================
def get_clusters_mapping_for_client(client_name: str) -> dict:
    """
    Возвращает полный маппинг кластеров для клиента:
    - берёт default_clusters_mapping как базу;
    - поверх накладывает client_clusters_overrides[client_name], если есть.
    """
    mapping = default_clusters_mapping.copy()
    overrides = client_clusters_overrides.get(client_name)
    if overrides:
        mapping.update(overrides)
    return mapping


def get_cluster_planning_days_for_client(client_name: str) -> dict:
    """
    Возвращает словарь 'кластер → дни планирования' для клиента:
    - берёт default_cluster_planning_days как базу;
    - поверх накладывает client_cluster_planning_overrides[client_name], если есть.
    """
    planning = default_cluster_planning_days.copy()
    overrides = client_cluster_planning_overrides.get(client_name)
    if overrides:
        planning.update(overrides)
    return planning


def add_clusters(date_metrics: str) -> pd.DataFrame:
    """
    Добавление недостающих комбинаций SKU–Кластер:
    - берем все SKU из summary
    - берем все кластеры из справочника clusters_and_warehouses.csv
    - формируем декартово произведение и мержим метрики.
    """
    metrics_clusters = read_metrics_file(date_metrics, sheet_name="claster_report")
    metrics_summary = read_metrics_file(date_metrics, sheet_name="summary")

    clusters = pd.read_csv(
        MARKETPLACE_PATH / "scripts" / "clusters_and_warehouses.csv",
        sep=";",
    )

    sku_clusters = list(
        product(
            metrics_summary["SKU"].unique(),
            clusters["cluster_name"].unique(),
        )
    )
    sku_clusters_df = pd.DataFrame(
        sku_clusters,
        columns=["SKU", "Кластер"],
    )

    sales_columns = [
        "Заказы",
        "Заказы_руб",
        "Продажи",
        "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ",
        "Остатки",
        "Остатки_fbs",
    ]
    if INCLUDE_PREVIOUS_SUPPLIES:
        sales_columns = add_element_to_list(sales_columns, "ПРОШЛЫЕ_ПОСТАВКИ", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ")

    metrics_df_all_clusters = sku_clusters_df.merge(
        metrics_clusters[["SKU", "Кластер"] + sales_columns],
        on=["SKU", "Кластер"],
        how="left",
    )

    for col in sales_columns:
        metrics_df_all_clusters[col] = metrics_df_all_clusters[col].fillna(0)

    metrics_df_all_clusters = metrics_df_all_clusters.merge(
        metrics_summary.drop(columns=sales_columns),
        on="SKU",
        how="left",
    )
    return metrics_df_all_clusters


def replace_clusters(
    clusters_mapping_dict: dict,
    df: pd.DataFrame,
    cluster_col: str = "Кластер",
) -> pd.DataFrame:
    """
    Объединение нескольких кластеров в один по словарю clusters_mapping.
    clusters_mapping_dict: { 'Кластер_итоговый': [subcluster1, subcluster2, ...], ... }
    """
    cluster_by_sub = {}
    for cluster, subclusters in clusters_mapping_dict.items():
        for sub in subclusters:
            cluster_by_sub[sub] = cluster

    df_res = df.copy()
    df_res[cluster_col] = df_res[cluster_col].map(cluster_by_sub).fillna(
        df_res[cluster_col]
    )
    return df_res


# ============================================================================
# РАСЧЕТ ПОКАЗАТЕЛЕЙ
# ============================================================================


def add_demand_columns(
    df: pd.DataFrame,
    stock: pd.Series,
    days: int,
    label: str = "",
) -> pd.DataFrame:
    """
    Универсальное добавление колонок потребности и дефицита/избытка.

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


def calc_svod_for_clusters(metrics_df_all_clusters: pd.DataFrame, cluster_planning_days: dict) -> pd.DataFrame:
    """
    Специфический расчет для листа 'По кластерам':
    - Оборачиваемость считается только по заказам: Заказы / DAYS_IN_PERIOD
    - Для каждого кластера задается свое количество дней планирования:
        в колонке 'Дни планирования'
        (по словарю cluster_planning_days, если кластера нет — 40).
    - Считаются универсальные колонки:
        'Потребность',
        'Дефицит/Избыток',
        'Дефицит/Избыток с учетом FBS',
        'Потребность (округл.)',
        'Потребность с учетом FBS (округл.)'
    - 'Корректировка с учетом нулевых остатков' берется по той же
      округленной потребности.
    """
    df = metrics_df_all_clusters.copy()

    # Оборачиваемость: ТОЛЬКО заказы
    df["Оборачиваемость"] = df["Заказы"] / DAYS_IN_PERIOD

    # Заполняем NaN нулями для остатков/поставок/FBS на всякий случай
    for col in ["Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
        if col in df.columns:
            df[col] = df[col].fillna(0)
        else:
            df[col] = 0

    # Нормализуем названия кластеров
    clusters_normalized = df["Кластер"].astype(str).str.strip()

    # Векторно ставим дни планирования: если кластера нет в словаре — 40
    df["Дни планирования"] = clusters_normalized.map(cluster_planning_days).fillna(40).astype(int)

    # Векторный расчёт с учётом "Дни планирования"
    days = df["Дни планирования"]
    turnover = df["Оборачиваемость"]

    demand = turnover * days
    stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
    stock_with_fbs = df["Остатки"] + df["Остатки_fbs"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]

    deficit = stock - demand
    deficit_fbs = stock_with_fbs - demand

    df["Потребность"] = demand
    df["Дефицит/Избыток"] = deficit
    df["Дефицит/Избыток с учетом FBS"] = deficit_fbs

    df["Потребность (округл.)"] = np.where(
        deficit > 0,
        0,
        np.ceil(np.abs(deficit))
    )
    df["Потребность с учетом FBS (округл.)"] = np.where(
        deficit_fbs > 0,
        0,
        np.ceil(np.abs(deficit_fbs))
    )
    # Корректировка — по округленной потребности
    df["Корректировка с учетом нулевых остатков"] = df["Потребность (округл.)"]

    # Нумерация строк
    df["№"] = np.arange(1, len(df) + 1)

    return df


def create_pivot_clusters(svod_for_clusters: pd.DataFrame) -> pd.DataFrame:
    """
    Создание сводной таблицы по кластерам:
    - 'Потребность <Кластер>' по Корректировке
    - 'Остаток <Кластер>' по Остаткам
    """
    pivot_values = [
        "Корректировка с учетом нулевых остатков",
        "Остатки",
    ]
    if INCLUDE_PREVIOUS_SUPPLIES:
        pivot_values.append("ПРОШЛЫЕ_ПОСТАВКИ")

    df_pivot = pd.pivot_table(
        svod_for_clusters,
        columns="Кластер",
        index="Артикул",
        values=pivot_values,
        aggfunc="sum",
        fill_value=0,
    )

    df_pivot.columns = [
        (
            f"Потребность {col[1]}"
            if col[0] == "Корректировка с учетом нулевых остатков"
            else f"Остаток {col[1]}"
            if col[0] == "Остатки"
            else f"Прошлые поставки {col[1]}"
            if col[0] == "ПРОШЛЫЕ_ПОСТАВКИ"
            else f"{col[0]} {col[1]}"
        )
        for col in df_pivot.columns.values
    ]

    df_pivot = df_pivot.reset_index()

    clusters = svod_for_clusters["Кластер"].unique()
    cols_order = ["Артикул"] + [
        val
        for cluster in clusters
        for val in (f"Потребность {cluster}", f"Остаток {cluster}")
    ]
    df_pivot = df_pivot[cols_order]

    return df_pivot



def calc_svod_for_sku(metrics_df_by_sku: pd.DataFrame) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости для товаров, агрегированных по SKU
    (лист 'Всего').

    Специфика:
    - Оборачиваемость считается только по заказам: Заказы / DAYS_IN_PERIOD.
    - Потребность считается на два периода (DEMAND_PERIOD_1/2), с и без FBS.
    - Добавляется колонка 'ABC по заказам' по накопленной доле заказов.
    """
    df = metrics_df_by_sku.copy()

    # Оборачиваемость: ТОЛЬКО заказы, без продаж
    df["Оборачиваемость"] = df["Заказы"] / DAYS_IN_PERIOD

    # Защита от NaN в остатках/поставках/FBS
    for col in ["Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
        if col in df.columns:
            df[col] = df[col].fillna(0)
        else:
            df[col] = 0

    stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
    stock_with_fbs = (
        df["Остатки"]
        + df["Остатки_fbs"]
        + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
    )

    # Потребность на DEMAND_PERIOD_1 дней (без FBS и с учетом FBS)
    df = add_demand_columns(df, stock, DEMAND_PERIOD_1)
    df = add_demand_columns(
        df,
        stock_with_fbs,
        DEMAND_PERIOD_1,
        "с учетом FBS",
    )

    # Потребность на DEMAND_PERIOD_2 дней
    df = add_demand_columns(df, stock, DEMAND_PERIOD_2)

    # ABC-анализ по заказам
    # Считаем только по положительным заказам, чтобы не накапливать шум
    df["Заказы_tmp"] = df["Заказы"].fillna(0)
    total_orders = df["Заказы_tmp"].sum()

    if total_orders > 0:
        # Сортируем по заказам (по убыванию)
        df = df.sort_values(by="Заказы_tmp", ascending=False)

        # Доля в общем объеме заказов
        df["Доля_заказов"] = df["Заказы_tmp"] / total_orders

        # Накопленная доля
        df["Накопленная_доля_заказов"] = df["Доля_заказов"].cumsum()

        # ABC-классификация
        def classify_abc(cum_share: float) -> str:
            if cum_share <= 0.8:
                return "A"
            elif cum_share <= 0.95:
                return "B"
            else:
                return "C"

        df["ABC по заказам"] = df["Накопленная_доля_заказов"].apply(classify_abc)

        # Можно вернуть исходный порядок, если важно
        df = df.sort_index()
    else:
        # Если заказов нет совсем — все C
        df["ABC по заказам"] = "C"

    # Нумерация строк
    df["№"] = np.arange(1, len(df) + 1)

    # Чистим временные колонки
    df.drop(columns=["Заказы_tmp", "Доля_заказов", "Накопленная_доля_заказов"], inplace=True, errors="ignore")

    return df

# ============================================================================
# ОБЪЕДИНЕНИЕ ДАННЫХ
# ============================================================================


def add_demand_for_cluster_columns(
    df_pivot_by_clusters: pd.DataFrame,
    metrics_df_by_sku_stats: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в таблицу по SKU колонки с потребностью/остатком по кластерам
    через merge с pivot-таблицей по 'Артикул', а также считает
    суммарную потребность по всем кластерам:
      'Потребность по кластерам всего'.
    """
    svod_for_sku = (
        metrics_df_by_sku_stats.merge(
            df_pivot_by_clusters,
            how="left",
            on="Артикул",
        )
        .fillna(0)
    )

    # Находим все колонки с потребностью по кластерам:
    # 'Потребность <Кластер>' (исключаем колонки типа 'Потребность на N дней')
    demand_cluster_cols = [
        col
        for col in svod_for_sku.columns
        if col.startswith("Потребность ")
        and not col.startswith("Потребность на ")
    ]

    if demand_cluster_cols:
        svod_for_sku["Потребность по кластерам всего"] = (
            svod_for_sku[demand_cluster_cols].sum(axis=1)
        )
    else:
        # На всякий случай, если по какой-то причине колонок нет
        svod_for_sku["Потребность по кластерам всего"] = 0

    return svod_for_sku


def add_columns_from_products(
    svod: pd.DataFrame,
    df_products_processed: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в свод доп. колонки из списка товаров (по 'Артикул').
    """
    svod_with_additional = svod.merge(
        df_products_processed,
        how="left",
        on="Артикул",
    )
    return svod_with_additional


def add_columns_from_catalog(
    df_with_stats: pd.DataFrame,
    catalog: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в свод колонки из справочной таблицы (по 'Артикул').
    """
    df = df_with_stats.copy()
    df["Артикул"] = df["Артикул"].astype(str)

    catalog_use = catalog.copy()
    catalog_use["Артикул"] = catalog_use["Артикул"].astype(str)

    df = df.merge(
        catalog_use[catalog_supply_columns],
        how="left",
        on="Артикул",
    )
    # Заполнение значений по умолчанию
    fill_values = {'Кратность коробов': DEFAULT_BOX_MULTIPLICITY}
    df = df.fillna(value=fill_values)

    return df


def add_columns_for_excel(
    df_with_stats: pd.DataFrame,
    cluster_list: List[str],
    date_metrics: str,
    day_start: datetime,
    day_end: datetime,
    svod_type: str = "Всего",
) -> pd.DataFrame:
    """
    Переименование и упорядочивание колонок под шаблон Excel.
    """
    svod_excel = df_with_stats.copy()
    date_metrics_dt = datetime.strptime(date_metrics, "%Y-%m-%d")

    if svod_type == "Всего":
        svod_columns = list(supply_svod_total_columns)

        cluster_columns = []
        for cluster in cluster_list:
            # Порядок внутри кластера
            cluster_columns.append(f"Потребность {cluster}")
            cluster_columns.append(f"Остаток {cluster}")
            if include_previous_supplies_flag:
                cluster_columns.append(f"Прошлые поставки {cluster}")

        svod_columns.extend(cluster_columns)
    else:
        svod_columns = list(supply_svod_clusters_columns)

    for col in svod_columns:
        if col not in svod_excel.columns:
            svod_excel[col] = np.nan

    svod_excel = svod_excel[svod_columns]
    svod_excel = svod_excel.sort_values(by=["Артикул"])

    svod_excel.rename(
        columns={
            "SKU": "Ozon SKU ID",
            "Наименование товара": "Наименование",
            "Заказы": f"ЗАКАЗЫ с {day_start.strftime('%d.%m')} по {day_end.strftime('%d.%m')}, шт.",
            "Продажи": f"ПРОДАЖИ с {day_start.strftime('%d.%m')} по {day_end.strftime('%d.%m')}, шт.",
            "Остатки": f"ОСТАТОК {date_metrics_dt.strftime('%d.%m')}",
            "Остатки_fbs": f"ОСТАТОК FBS {date_metrics_dt.strftime('%d.%m')}",
        },
        inplace=True,
    )

    return svod_excel


# ============================================================================
# ЮВЕЛИРНЫЕ КЛАСТЕРЫ
# ============================================================================


def create_jewelry_clusters_svod(
    metrics_df_all_clusters: pd.DataFrame,
    date_metrics: str,
    df_products_processed: pd.DataFrame,
    catalog: pd.DataFrame,
) -> Optional[pd.DataFrame]:
    """
    Создает свод "Всего по ювелирным кластерам", если тип магазина — ювелирка.
    Логика по кластерам сохраняет ozon-подход:
      - replace_clusters с clusters_mapping_jewelry
      - pivot по кластерам
      - merge по 'Артикул'.
    """
    if shop_type != "jewelry":
        return None

    metrics_df_jewelry = replace_clusters(
        clusters_mapping_jewelry,
        metrics_df_all_clusters,
    )

    svod_for_jewelry_clusters = calc_svod_for_clusters(metrics_df_jewelry)
    df_pivot_by_jewelry_clusters = create_pivot_clusters(
        svod_for_jewelry_clusters
    )

    metrics_df_by_sku = read_metrics_file(date_metrics)
    metrics_df_by_sku_stats = calc_svod_for_sku(metrics_df_by_sku)

    svod_for_sku_jewelry = add_demand_for_cluster_columns(
        df_pivot_by_jewelry_clusters,
        metrics_df_by_sku_stats,
    )

    svod_with_products = add_columns_from_products(
        svod_for_sku_jewelry,
        df_products_processed,
    )
    svod_with_catalog = add_columns_from_catalog(
        svod_with_products,
        catalog,
    )

    report_dates = read_dates_file(date_metrics)
    cluster_list = list(clusters_mapping_jewelry.keys())

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
# СОХРАНЕНИЕ И ФОРМАТИРОВАНИЕ EXCEL
# ============================================================================


def save_sheets_to_excel(
    svod_excel_for_sku: pd.DataFrame,
    svod_excel_for_clusters: pd.DataFrame,
    df_pivot_by_clusters: pd.DataFrame,
    svod_excel_for_sku_jewelry: Optional[pd.DataFrame],
    date_metrics: str,
) -> Path:
    """
    Сохранение листов в Excel-файл (без форматирования).
    """
    SUPPLY_SVOD_DIR.mkdir(parents=True, exist_ok=True)

    file_path = (
        SUPPLY_SVOD_DIR
        / f"{date_metrics}_Расчет_поставок_{client_name}_Ozon.xlsx"
    )

    with pd.ExcelWriter(file_path) as writer:
        svod_excel_for_sku.to_excel(
            writer,
            sheet_name="Всего",
            index=False,
        )
        svod_excel_for_clusters.to_excel(
            writer,
            sheet_name="По кластерам",
            index=False,
        )
        # при необходимости можно писать df_pivot_by_clusters на отдельный лист
        # df_pivot_by_clusters.to_excel(writer, sheet_name='Сводная по кластерам', index=False)

        if svod_excel_for_sku_jewelry is not None:
            svod_excel_for_sku_jewelry.to_excel(
                writer,
                sheet_name="Всего по ювелирным кластерам",
                index=False,
            )

    logger.info(f"Supply svod (raw) saved in: {file_path}")
    return file_path


def format_excel_file(
    client_name: str,
    date_metrics: str,
    svod_excel_for_sku: pd.DataFrame,
    svod_excel_for_sku_jewelry: Optional[pd.DataFrame],
    svod_excel_for_clusters: pd.DataFrame,
    clusters_mapping_dict: dict,
    clusters_mapping_jewelry_dict: dict,
) -> None:
    """
    Форматирование Excel (формулы, стили и т.п.) через модуль format_supply_svod.
    """
    path_supply_svod = SUPPLY_SVOD_DIR

    # format_supply_svod.copy_supply_svod_file(
    #     path_supply_svod,
    #     client_name,
    #     date_metrics,
    # )

    # Лист "Всего"
    format_supply_svod.format_sheet_total(
        path_supply_svod=path_supply_svod,
        client_name=client_name,
        date_report_created=date_metrics,
        svod_excel=svod_excel_for_sku,
        svod_excel_clusters=svod_excel_for_clusters,
        clusters_mapping_dict=clusters_mapping_dict,
        sheet_name="Всего",
        demand_period_1=DEMAND_PERIOD_1,
        demand_period_2=DEMAND_PERIOD_2,
        days_in_period=DAYS_IN_PERIOD,
        apply_box_rounding=APPLY_BOX_ROUNDING,
    )

    # Лист "Всего по ювелирным кластерам"
    if svod_excel_for_sku_jewelry is not None and shop_type == "jewelry":
        format_supply_svod.format_sheet_total(
            path_supply_svod=path_supply_svod,
            client_name=client_name,
            date_report_created=date_metrics,
            svod_excel=svod_excel_for_sku_jewelry,
            svod_excel_clusters=svod_excel_for_clusters,
            clusters_mapping_dict=clusters_mapping_jewelry_dict,
            sheet_name="Всего по ювелирным кластерам",
            demand_period_1=DEMAND_PERIOD_1,
            demand_period_2=DEMAND_PERIOD_2,
            days_in_period=DAYS_IN_PERIOD,
            apply_box_rounding=APPLY_BOX_ROUNDING,
        )

    # Лист "По кластерам"
    format_supply_svod.format_sheet_clusters(
        path_supply_svod=path_supply_svod,
        client_name=client_name,
        date_report_created=date_metrics,
        svod_excel=svod_excel_for_clusters,
        sheet_name="По кластерам",
        # demand_period_1=DEMAND_PERIOD_1,
        # demand_period_2=DEMAND_PERIOD_2,
        days_in_period=DAYS_IN_PERIOD,
        apply_box_rounding=APPLY_BOX_ROUNDING,
    )

    filename_formatted = (
        path_supply_svod
        / f"{date_metrics}_Расчет_поставок_{client_name}_Ozon_formatted.xlsx"
    )
    logger.info(f"Расчет поставок сохранен в: {filename_formatted}")

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================


def create_supply_svod_report(date_metrics: str) -> None:
    """
    Основная функция формирования отчета по поставкам Ozon.
    """
    logger.info(
        f"Создание отчета поставок для клиента {client_name} на дату {date_metrics}"
    )
    logger.info(f"Период оборачиваемости: {DAYS_IN_PERIOD} дней")
    logger.info(
        f"Периоды потребности: {DEMAND_PERIOD_1} дней и {DEMAND_PERIOD_2} дней"
    )
    # Получаем маппинг кластеров для клиента и дни планирования на кластеры
    clusters_mapping = get_clusters_mapping_for_client(client_name)
    cluster_planning_days = get_cluster_planning_days_for_client(client_name)


    # Список товаров
    df_products = read_products_file(date_metrics)
    df_products_processed = process_product_list(df_products)

    # Справочник
    catalog = read_catalog()

    # Метрики по всем кластерам (с добавлением недостающих)
    metrics_df_all_clusters = add_clusters(date_metrics)

    # Отчет "По кластерам" без объединения кластеров
    svod_for_clusters = calc_svod_for_clusters(metrics_df_all_clusters, cluster_planning_days)
    df_pivot_by_clusters = create_pivot_clusters(svod_for_clusters)

    # Объединение кластеров по общему маппингу
    metrics_df_union_clusters = replace_clusters(
        clusters_mapping,
        metrics_df_all_clusters,
    )
    svod_for_union_clusters = calc_svod_for_clusters(metrics_df_union_clusters, cluster_planning_days)
    df_pivot_by_union_clusters = create_pivot_clusters(
        svod_for_union_clusters
    )

    # Метрики по SKU (лист summary)
    metrics_df_by_sku = read_metrics_file(date_metrics)
    metrics_df_by_sku_stats = calc_svod_for_sku(metrics_df_by_sku)

    # Добавляем потребность/остатки по кластерам в свод по SKU
    svod_for_sku = add_demand_for_cluster_columns(
        df_pivot_by_union_clusters,
        metrics_df_by_sku_stats,
    )

    # Добавляем данные из списка товаров
    svod_sku_with_products = add_columns_from_products(
        svod_for_sku,
        df_products_processed,
    )
    svod_clusters_with_products = add_columns_from_products(
        svod_for_clusters,
        df_products_processed,
    )

    # Добавляем данные из справочника
    svod_sku_with_catalog = add_columns_from_catalog(
        svod_sku_with_products,
        catalog,
    )
    svod_clusters_with_catalog = add_columns_from_catalog(
        svod_clusters_with_products,
        catalog,
    )

    # Даты отчета
    report_dates = read_dates_file(date_metrics)
    cluster_list = list(clusters_mapping.keys())

    # Лист "Всего"
    svod_excel_for_sku = add_columns_for_excel(
        svod_sku_with_catalog,
        cluster_list,
        date_metrics,
        report_dates["date_start_file"][0],
        report_dates["date_end_file"][0],
        svod_type="Всего",
    )

    # Лист "Всего по ювелирным кластерам"
    svod_excel_for_sku_jewelry = create_jewelry_clusters_svod(
        metrics_df_all_clusters,
        date_metrics,
        df_products_processed,
        catalog,
    )

    # Лист "По кластерам"
    svod_excel_for_clusters = add_columns_for_excel(
        svod_clusters_with_catalog,
        cluster_list,
        date_metrics,
        report_dates["date_start_file"][0],
        report_dates["date_end_file"][0],
        svod_type="По кластерам",
    )

    # Сохранение
    save_sheets_to_excel(
        svod_excel_for_sku,
        svod_excel_for_clusters,
        df_pivot_by_clusters,
        svod_excel_for_sku_jewelry,
        date_metrics,
    )

    # Форматирование
    format_excel_file(
        client_name,
        date_metrics,
        svod_excel_for_sku,
        svod_excel_for_sku_jewelry,
        svod_excel_for_clusters,
        clusters_mapping,
        clusters_mapping_jewelry,
    )

    logger.info(
        f"Отчет поставок Ozon успешно создан для {client_name} на дату {date_metrics}"
    )


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================
# %%
if __name__ == "__main__":
    date_metrics = cli_reference_date(parser)
    create_supply_svod_report(date_metrics)

# %%
