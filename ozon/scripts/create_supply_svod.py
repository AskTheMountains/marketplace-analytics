# %%
import argparse
from datetime import date, datetime
from itertools import product
from pathlib import Path
from re import I
from typing import List, Optional

import numpy as np
import pandas as pd
from rich_argparse.contrib import ParagraphRichHelpFormatter

from generic_functions import add_element_to_list
from ozon.scripts.ozon_generic_functions import clean_ozon_article
from ozon.scripts.ozon_logger import ozon_client_logger as logger
from ozon.scripts.command_args_parser import cli_reference_date
from ozon.scripts.constants import (
    client_name,
    project_root,
    marketplace_name,
    special_clients_supply_report,
    article_replacements,
    default_clusters_mapping,
    client_clusters_overrides,
    default_cluster_planning_days,
    client_cluster_planning_overrides,
    DEFAULT_PLANNING_DAYS,
    clusters_mapping_jewelry,
    marketplace_dir_path,
    shop_type,
    catalog_supply_columns,
    cargo_multiplicity,
    DEMAND_PERIOD_1,
    DEMAND_PERIOD_2,
    demand_period_for_total,
    include_previous_supplies_flag,
    only_reference_table_products_flag,
    supply_svod_clusters_columns,
    supply_svod_total_columns,
)

# Функция получения справочной таблицы (Яндекс.Диск или локально)
from ozon.scripts.get_reference_table import get_reference_table

# Функция форматирования Excel-файла с отчётом (формулы, стили и т.п.)
import ozon.scripts.format_supply_svod as format_supply_svod

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
    Обработка списка товаров:
    - очистка артикула,
    - определение сортируемости (WB-логика),
    - определение габаритов товара для Ozon (КГТ / Обычный).
    """
    df = df_products.copy()

    # Оставляем только нужные колонки
    df = df.loc[
        :,
        df.columns.isin(
            [
                "Артикул",
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

    # Очистка артикула (твоя функция)
    df["Артикул"] = df["Артикул"].map(clean_ozon_article)

    # ---- Сортируемость товара (как было) ----
    sortable = {
        "длина": 450,
        "ширина": 320,
        "высота": 320,
        "вес": 10000,
    }
    unsortable = {
        "длина_min": 450,
        "длина_max": 1200,
        "ширина_min": 320,
        "ширина_max": 800,
        "высота_min": 320,
        "высота_max": 600,
        "вес_max": 35000,
    }

    # Приводим к числам, NaN → 0, чтобы сравнения не падали
    for col in ["Вес", "Глубина", "Ширина", "Высота"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    conditions = [
        (df["Вес"] <= sortable["вес"])
        & (df["Глубина"] <= sortable["длина"])
        & (df["Ширина"] <= sortable["ширина"])
        & (df["Высота"] <= sortable["высота"]),
        (df["Вес"] <= unsortable["вес_max"])
        & df["Глубина"].between(unsortable["длина_min"], unsortable["длина_max"])
        & df["Ширина"].between(unsortable["ширина_min"], unsortable["ширина_max"])
        & df["Высота"].between(unsortable["высота_min"], unsortable["высота_max"]),
    ]
    choices = ["Сортируемый товар", "Несортируемый товар"]

    df["Сортируемость товара (расчет)"] = np.select(
        conditions,
        choices,
        default="Не определена",
    )

    # ---- Габариты товара (логика Ozon КГТ) ----
    # Переводим габариты в см, вес в кг
    length_cm = df["Глубина"] / 10.0   # мм → см
    width_cm = df["Ширина"] / 10.0
    height_cm = df["Высота"] / 10.0
    weight_kg = df["Вес"] / 1000.0     # г → кг

    # Условия крупногабаритного товара на Ozon:
    # Длина, см от 45 до 220
    cond_len = length_cm.between(45, 220)
    # Ширина, см от 32 до 120
    cond_wid = width_cm.between(32, 120)
    # Высота, см от 32 до 80
    cond_hgt = height_cm.between(32, 80)
    # Вес, кг до 120
    cond_wgt = weight_kg <= 120

    mask_kgt_ozon = cond_len & cond_wid & cond_hgt & cond_wgt

    df["Габариты товара (расчет)"] = np.where(mask_kgt_ozon, "КГТ", "Стандарт")

    # Убираем колонку с исходными габаритами и весом,
    # если они дальше не нужны (как было у тебя)
    df = df.drop(
        columns=[
            "Вес",
            "Глубина",
            "Ширина",
            "Высота",
            "Единицы измерения габаритов",
            "Единицы измерения веса",
        ],
        errors="ignore",
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


def process_catalog(catalog_df: pd.DataFrame) -> pd.DataFrame:
    """
    Обработка справочной таблицы Ozon.
    """
    # Создаем копию для обработки, чтобы не изменять оригинал
    catalog_processed = catalog_df.copy()
    # Делаем первую букву в названиях столбцов заглавной, остальные - строчными
    catalog_processed.columns = [col.capitalize() for col in catalog_processed.columns]

    # Переименовываем некоторые колонки
    catalog_processed = catalog_processed.rename(columns={
        "Артикул продавца": "Артикул",
        "Артикул 1с": "Артикул 1С",
        'Sku': 'SKU',
        "Артикул озон": "Артикул",
        # "Сортируемость товара": "Сортируемость",
        'Сортируемость': 'Сортируемость товара',
        "Статус товара - переносим в отчет": "Статус",
        'Статус продажи fbo/fbs': 'Статус продажи FBO/FBS',
        'Габариты товара': 'Габариты товара (справочно)',
        'Ррц': 'РРЦ',
        'Товары в пути': 'В производстве',
        "Остатки на складе 1с": "Остатки на складе 1С",
    })
    if client_name in ('BIO', 'ARK'):
        catalog_processed = catalog_processed.rename(columns={
            "Статус": "Статус товара",
        })
    # Удаляем дубликаты колонок, если есть. Оставляем первую встречающуюся.
    catalog_processed = catalog_processed.loc[:, ~catalog_processed.columns.duplicated()]

    # Переводим SKU в int
    if 'SKU' in catalog_processed.columns:
        catalog_processed['SKU'] = catalog_processed['SKU'].astype('Int64')

    # Переводим артикул в строку и зменяем некоторые значения
    catalog_processed["Артикул"] = catalog_processed["Артикул"].map(clean_ozon_article)

    # Добавляем колонку с кратностью коробов
    if not {"Кратность коробов", "Кратность короба"} & set(catalog_processed.columns):
        logger.info(f"Колонка 'Кратность коробов' не найдена в справочной таблице. Добавляем с дефолтным значением: {DEFAULT_BOX_MULTIPLICITY}.")
        catalog_processed['Кратность коробов'] = DEFAULT_BOX_MULTIPLICITY


    missing_columns = []
    for col in catalog_supply_columns:
        if col not in catalog_processed.columns:
            catalog_processed[col] = np.nan
            missing_columns.append(col)

    if missing_columns:
        logger.warning(
            f"Отсутствуют колонки в справочной таблице: {missing_columns}"
        )

    return catalog_processed


# ============================================================================
# РАБОТА С КЛАСТЕРАМИ
# ============================================================================


def get_all_clusters_from_ozon_and_report(
    date_metrics: str,
    verbose: bool = False,
) -> List[str]:
    """
    Возвращает объединённый список кластеров:
      - из файла clusters_and_warehouses.csv (столбец 'cluster_name');
      - из отчёта claster_report (столбец 'Кластер').
    Используется для того, чтобы не потерять кластеры, которые есть в старых
    отчётах, но уже отсутствуют в актуальном справочнике Ozon.
    """
    # --- 1. Кластеры из справочника Ozon ---
    clusters_path = MARKETPLACE_PATH / "scripts" / "clusters_and_warehouses.csv"
    ozon_clusters: set[str] = set()
    try:
        df_clusters = pd.read_csv(clusters_path, sep=";")
        if "cluster_name" in df_clusters.columns:
            ozon_clusters = set(
                df_clusters["cluster_name"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )
        else:
            logger.warning(
                f"В файле {clusters_path} нет колонки 'cluster_name'. "
                f"Кластеры из справочника Ozon не будут учтены."
            )
    except FileNotFoundError:
        logger.warning(
            f"Файл с кластерами Ozon не найден: {clusters_path}. "
            f"Кластеры из справочника Ozon не будут ��чтены."
        )

    # --- 2. Кластеры из отчёта claster_report ---
    report_clusters: set[str] = set()
    try:
        metrics_clusters = read_metrics_file(date_metrics, sheet_name="claster_report")
        if "Кластер" in metrics_clusters.columns:
            report_clusters = set(
                metrics_clusters["Кластер"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )
        else:
            logger.warning(
                f"В claster_report за {date_metrics} нет колонки 'Кластер'. "
                f"Кластеры из отчёта не будут учтены."
            )
    except Exception as e:
        logger.warning(
            f"Не удалось прочитать кластеры из claster_report на дату {date_metrics}: {e}"
        )

    # --- 3. Объединяем ---
    all_clusters = sorted(ozon_clusters | report_clusters)

    if verbose:
        logger.info(
            f"Собран объединённый список кластеров Ozon и отчёта за {date_metrics}: "
            f"{all_clusters}"
        )
    else:
        logger.info(
            f"Собран объединённый список кластеров за {date_metrics}: "
            f"{len(all_clusters)} кластеров "
            f"(Ozon: {len(ozon_clusters)}, report: {len(report_clusters)})"
        )

    return all_clusters

def get_clusters_mapping_for_client(
    client_name: str,
    use_all_clusters: bool = False,
    date_metrics: str | None = None,
    verbose: bool = False,
) -> dict:
    """
    Возвращает полный маппинг кластеров для клиента:
    - берёт default_clusters_mapping как базу;
    - удаляет из базы те ключи, которые клиент явно использует как подкластер;
    - поверх накладывает client_clusters_overrides[client_name], если есть (приоритет клиента);
    - при use_all_clusters=True дополнительно может добавить кластеры
      из справочника/отчёта (get_all_clusters_from_ozon_and_report), которых нет
      в словаре и которые не используются как подкластер.
    """
    logger.info(f"Формирование маппинга кластеров для клиента '{client_name}'")

    # 1. Базовый маппинг (дефолт)
    mapping = default_clusters_mapping.copy()

    # 2. Оверрайды клиента
    overrides = client_clusters_overrides.get(client_name) or {}

    # 3. Подкластеры, которые клиент ЯВНО использует в overrides
    subs_in_overrides: set[str] = set()
    for subs in overrides.values():
        subs_in_overrides.update(subs)

    if subs_in_overrides and verbose:
        logger.info(
            f"Для клиента '{client_name}' в overrides используются подкластеры: "
            f"{sorted(subs_in_overrides)}"
        )
    else:
        logger.info(
            f"Для клиента '{client_name}' overrides используют "
            f"{len(subs_in_overrides)} различных подкластеров"
        )

    # 4. Удаляем из базового маппинга те ключи, которые клиент использует как подкластер
    keys_before = set(mapping.keys())
    for k in list(mapping.keys()):
        if k in subs_in_overrides:
            mapping.pop(k, None)
    keys_after = set(mapping.keys())
    removed_keys = sorted(keys_before - keys_after)

    if removed_keys and verbose:
        logger.info(
            f"Для клиента '{client_name}' из default_clusters_mapping удалены ключи, "
            f"используемые как подкластеры: {removed_keys}"
        )
    else:
        logger.info(
            f"Для клиента '{client_name}' из default_clusters_mapping удалено ключей: "
            f"{len(removed_keys)}"
        )

    # 5. Накладываем overrides поверх очищенного дефолта (приоритет клиента)
    mapping.update(overrides)

    # 6. Собираем множество ПОДКЛАСТЕРОВ в итоговом маппинге (после overrides)
    subs_in_final_mapping: set[str] = set()
    for subs in mapping.values():
        subs_in_final_mapping.update(subs)

    # 7. Опционально дополняем маппинг кластерами из справочника/отчёта
    if use_all_clusters and date_metrics is not None:
        all_clusters = get_all_clusters_from_ozon_and_report(
            date_metrics,
            verbose=verbose,
        )
        if all_clusters:
            to_add = []
            for cl in all_clusters:
                # Не добавляем, если:
                #  - кластер уже есть как ключ в маппинге,
                #  - или уже используется как подкластер в каком‑то агрегате.
                if cl in mapping or cl in subs_in_final_mapping:
                    continue
                mapping[cl] = [cl]
                to_add.append(cl)
            if to_add:
                if verbose:
                    logger.info(
                        f"Для клиента '{client_name}' в маппинг добавлены отсутствующие кластеры "
                        f"(сам-на-себя) из справочника/отчёта, не используемые как подкластеры: "
                        f"{sorted(to_add)}"
                    )
                else:
                    logger.info(
                        f"Для клиента '{client_name}' в маппинг добавлены кластеры "
                        f"{to_add} из справочника/отчёта"
                    )

    # Итоговые логи
    keys_sorted = sorted(mapping.keys())
    logger.info(
        f"Итоговый маппинг кластеров для клиента '{client_name}': "
        f"{len(keys_sorted)} агрегирующих кластеров"
    )
    if verbose:
        logger.debug(
            f"Ключи итогового маппинга для '{client_name}': {keys_sorted}"
        )
        logger.debug(
            f"Полный итоговый маппинг для '{client_name}': {mapping}"
        )

    return mapping

def get_cluster_planning_days_for_client(
    client_name: str,
    use_all_clusters: bool = False,
    date_metrics: str | None = None,
    verbose: bool = False,
) -> dict:
    """
    Возвращает словарь 'кластер → дни планирования' для клиента:
    - берёт default_cluster_planning_days как базу;
    - поверх накладывает client_cluster_planning_overrides[client_name], если есть
      (приоритет клиента);
    - при use_all_clusters=True дополняет словарь днями планирования для всех
      кластеров (из справочника/отчёта), которых нет в словаре
      (берётся DEFAULT_PLANNING_DAYS).
    """
    logger.info(
        f"Формирование планирования дней по кластерам для клиента '{client_name}'"
    )

    # 1. Базовое планирование (дефолт)
    planning = default_cluster_planning_days.copy()

    # 2. Оверрайды клиента (приоритет)
    overrides = client_cluster_planning_overrides.get(client_name) or {}
    if overrides:
        planning.update(overrides)
        if verbose:
            logger.info(
                f"Для клиента '{client_name}' применены overrides по планированию дней: "
                f"{overrides}"
            )
        else:
            logger.info(
                f"Для клиента '{client_name}' применены overrides по планированию дней, "
                f"кол-во кластеров: {len(overrides)}"
            )

    # 3. Дополняем днями кластеры (из справочника/отчёта), если включено
    if use_all_clusters and date_metrics is not None:
        all_clusters = get_all_clusters_from_ozon_and_report(
            date_metrics,
            verbose=verbose,
        )
        if all_clusters:
            planning_keys = set(planning.keys())
            missing_clusters = sorted(set(all_clusters) - planning_keys)
            for cl in missing_clusters:
                planning[cl] = DEFAULT_PLANNING_DAYS
            if missing_clusters:
                if verbose:
                    logger.info(
                        f"Для клиента '{client_name}' в планировании дней добавлены кластеры "
                        f"из справочника/отчёта с дефолтом {DEFAULT_PLANNING_DAYS} дней: "
                        f"{missing_clusters}"
                    )
                else:
                    logger.info(
                        f"Для клиента '{client_name}' в планировании дней добавлены кластеры "
                        f"{missing_clusters}с дефолтом {DEFAULT_PLANNING_DAYS} дней"
                    )

    if verbose:
        logger.debug(
            f"Итоговый словарь планирования дней для клиента '{client_name}': {planning}"
        )
    else:
        logger.info(
            f"Итоговый словарь планирования дней для клиента '{client_name}': "
            f"{len(planning)} кластеров"
        )

    return planning

def add_clusters(date_metrics: str) -> pd.DataFrame:
    """
    Добавление недостающих комбинаций SKU–Кластер:
    - берём все SKU из summary;
    - берём все кластеры из справочника + отчёта
      (get_all_clusters_from_ozon_and_report);
    - формируем декартово произведение и мержим метрики.
    """
    # Метрики по кластерам и по SKU
    metrics_clusters = read_metrics_file(date_metrics, sheet_name="claster_report")
    metrics_summary = read_metrics_file(date_metrics, sheet_name="summary")

    # Все кластеры (справочник + claster_report)
    all_clusters = get_all_clusters_from_ozon_and_report(date_metrics)

    # Декартово произведение SKU × Кластер
    sku_clusters = list(
        product(
            metrics_summary["SKU"].unique(),
            all_clusters,
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
        "Продажи_факт_2м",
        "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ",
        "Остатки",
        "Остатки_fbs",
    ]
    if INCLUDE_PREVIOUS_SUPPLIES:
        sales_columns = add_element_to_list(
            sales_columns,
            "ПРОШЛЫЕ_ПОСТАВКИ",
            "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ",
        )

    # Мержим метрики по кластерам
    metrics_df_all_clusters = sku_clusters_df.merge(
        metrics_clusters[["SKU", "Кластер"] + sales_columns],
        on=["SKU", "Кластер"],
        how="left",
    )

    # Заполняем NaN нулями по продажам/остаткам/поставкам
    for col in sales_columns:
        if col in metrics_df_all_clusters.columns:
            metrics_df_all_clusters[col] = metrics_df_all_clusters[col].fillna(0)

    # Подтягиваем общие поля по SKU из summary
    metrics_df_all_clusters = metrics_df_all_clusters.merge(
        metrics_summary.drop(columns=sales_columns, errors="ignore"),
        on="SKU",
        how="left",
    )

    return metrics_df_all_clusters

def add_clusters_light(date_metrics: str) -> pd.DataFrame:
    """
    Вариант без декартова произведения.
    Просто возвращаем claster_report, обогащённый summary.
    """
    metrics_clusters = read_metrics_file(date_metrics, sheet_name="claster_report")
    metrics_summary = read_metrics_file(date_metrics, sheet_name="summary")
    return metrics_clusters

    sales_columns = [
        "Заказы",
        "Заказы_руб",
        "Продажи",
        "Продажи_факт_2м",
        "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ",
        "Остатки",
        "Остатки_fbs",
    ]
    if INCLUDE_PREVIOUS_SUPPLIES:
        sales_columns = add_element_to_list(
            sales_columns,
            "ПРОШЛЫЕ_ПОСТАВКИ",
            "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ",
        )

    # Заполняем NaN нулями в метриках по кластерам
    for col in sales_columns:
        if col in metrics_clusters.columns:
            metrics_clusters[col] = metrics_clusters[col].fillna(0)

    # Добавляем общие поля по SKU из summary
    metrics_df_all_clusters = metrics_clusters.merge(
        metrics_summary.drop(columns=sales_columns, errors="ignore"),
        on="SKU",
        how="left",
    )
    return metrics_df_all_clusters

def replace_clusters(
    clusters_mapping_dict: dict,
    df: pd.DataFrame,
    cluster_col: str = "Кластер",
    new_col: str = "Кластер для группировки",
    replace_original: bool = False,
) -> pd.DataFrame:
    """
    Объединяет кластеры по словарю clusters_mapping_dict, создавая столбец
    'Кластер для группировки' и (по желанию) заменяя исходный 'Кластер'.

    Параметры
    ---------
    clusters_mapping_dict : dict
        Словарь вида:
          {
            'Кластер_итоговый_1': ['subcluster1', 'subcluster2', ...],
            'Кластер_итоговый_2': [...],
            ...
          }
        где ключ — название итогового кластера (для группировки),
        а список — список «подкластеров» (оригинальных значений кластера).

    df : pd.DataFrame
        Исходный датафрейм с колонкой cluster_col (по умолчанию 'Кластер').

    cluster_col : str, optional
        Имя колонки с исходным кластером в df. По умолчанию 'Кластер'.

    new_col : str, optional
        Имя новой колонки, куда будет записан «кластер для группировки».
        По умолчанию 'Кластер для группировки'.

    replace_original : bool, optional
        Если True — сразу перезаписывает df[cluster_col] значением new_col
        (после построения), оставляя при этом new_col в датафрейме.
        Если False (по умолчанию) — исходный кластер остаётся нетронутым.

    Возвращает
    ----------
    pd.DataFrame
        Копия df с добавленным столбцом new_col и, при необходимости,
        обновлённым cluster_col.
    """
    # строим обратный маппинг: подкластер -> итоговый кластер
    cluster_by_sub = {}
    for cluster_name, subclusters in clusters_mapping_dict.items():
        for sub in subclusters:
            cluster_by_sub[sub] = cluster_name

    df_res = df.copy()

    # нормализуем исходный кластер в строку
    df_res[cluster_col] = df_res[cluster_col].astype(str).str.strip()

    # создаём новый столбец для группировки:
    # если подкластер есть в словаре — берём итоговый кластер,
    # иначе оставляем исходное значение
    df_res[new_col] = df_res[cluster_col].map(cluster_by_sub).fillna(
        df_res[cluster_col]
    )

    # при необходимости перезаписываем исходный кластер
    if replace_original:
        df_res[cluster_col] = df_res[new_col]

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



def calc_svod_for_clusters(
    metrics_df_all_clusters: pd.DataFrame,
    cluster_planning_days: dict,
) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости по кластерам (лист 'По кластерам').

    RiverBoats:
      - Оборачиваемость = Заказы / DAYS_IN_PERIOD
      - Дни планирования из cluster_planning_days (если нет — 40)
      - Потребность = Оборачиваемость * Дни планирования
      - Дефицит/Избыток, с учетом FBS, округления
      - 'Корректировка с учетом нулевых остатков' = 'Потребность (округл.)'

    BIO, ARK:
      - Для КГТ: Оборачиваемость = Заказы / DAYS_IN_PERIOD
      - Для остальных: (Заказы + Продажи) / 2 / DAYS_IN_PERIOD
      - Дни планирования по кластерам (cluster_planning_days, если нет — 40)
      - Потребность = Оборачиваемость * Дни планирования
      - Далее как у RiverBoats.

    Neva_Metall:
      - Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD (как общий случай)
      - Дни планирования по кластерам (cluster_planning_days, если нет — 40)
      - Потребность = Оборачиваемость * Дни планирования
      - Далее как у RiverBoats/BIO/ARK.

    Остальные клиенты:
      - Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD
      - Потребность на DEMAND_PERIOD_1/2 через add_demand_columns
      - 'Корректировка с учетом нулевых остатков' =
        'Потребность на {demand_period_for_total} дней (округл.)'
    """
    df = metrics_df_all_clusters.copy()

    # ----------------- Специальная логика для RiverBoats -----------------
    if client_name == "RiverBoats":
        # Оборачиваемость: только заказы
        df["Оборачиваемость"] = df["Заказы"] / DAYS_IN_PERIOD

        # Защита от NaN в остатках и поставках
        for col in ["Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = 0

        # Дни планирования по кластерам (если кластера нет в словаре — 40)
        clusters_norm = df["Кластер"].astype(str).str.strip()
        df["Дни планирования"] = (
            clusters_norm.map(cluster_planning_days).fillna(DEFAULT_PLANNING_DAYS).astype(int)
        )

        days = df["Дни планирования"]
        turnover = df["Оборачиваемость"]

        # Потребность для каждого кластера по своим дням
        demand = turnover * days
        stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        stock_with_fbs = (
            df["Остатки"] + df["Остатки_fbs"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        )

        deficit = stock - demand
        deficit_fbs = stock_with_fbs - demand

        df["Потребность"] = demand
        df["Дефицит/Избыток"] = deficit
        df["Дефицит/Избыток с учетом FBS"] = deficit_fbs
        df["Потребность (округл.)"] = np.where(
            deficit > 0, 0, np.ceil(np.abs(deficit))
        )
        df["Потребность с учетом FBS (округл.)"] = np.where(
            deficit_fbs > 0, 0, np.ceil(np.abs(deficit_fbs))
        )
        # Корректировка — по округленной потребности
        df["Корректировка с учетом нулевых остатков"] = df["Потребность (округл.)"]

        df["№"] = np.arange(1, len(df) + 1)
        return df

    # ----------------- Логика для клиентов BIO и ARK -----------------
    if client_name in ("BIO", "ARK"):
        # Приводим числовые колонки к числу и заполняем NaN нулями
        for col in ["Заказы", "Продажи", "Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = 0

        # Базовая оборачиваемость и оборачиваемость только по заказам
        base_turnover = (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        orders_turnover = df["Заказы"] / DAYS_IN_PERIOD

        # Определяем КГТ
        if "Габариты товара" in df.columns:
            is_kgt = df["Габариты товара"].astype(str).str.strip() == "КГТ"
        else:
            is_kgt = df.get("Статус", "").astype(str).str.strip() == "КГТ"

        # Оборачиваемость: КГТ — только заказы, остальное — базовая формула
        df["Оборачиваемость"] = np.where(is_kgt, orders_turnover, base_turnover)

        # Дни планирования по кластерам
        clusters_norm = df["Кластер"].astype(str).str.strip()
        df["Дни планирования"] = (
            clusters_norm.map(cluster_planning_days).fillna(DEFAULT_PLANNING_DAYS).astype(int)
        )
        days = df["Дни планирования"]
        turnover = df["Оборачиваемость"]

        # Потребность и дефициты
        demand = turnover * days
        stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        stock_with_fbs = (
            df["Остатки"] + df["Остатки_fbs"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        )

        deficit = stock - demand
        deficit_fbs = stock_with_fbs - demand

        df["Потребность"] = demand
        df["Дефицит/Избыток"] = deficit
        df["Дефицит/Избыток с учетом FBS"] = deficit_fbs
        df["Потребность (округл.)"] = np.where(
            deficit > 0, 0, np.ceil(np.abs(deficit))
        )
        df["Потребность с учетом FBS (округл.)"] = np.where(
            deficit_fbs > 0, 0, np.ceil(np.abs(deficit_fbs))
        )
        df["Корректировка с учетом нулевых остатков"] = df["Потребность (округл.)"]

        df["№"] = np.arange(1, len(df) + 1)
        return df

    # ----------------- Логика для клиента Neva_Metall -----------------
    if client_name == "Neva_Metall":
        # Приводим числовые колонки к числу и заполняем NaN нулями
        for col in ["Заказы", "Продажи", "Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = 0

        # Оборачиваемость как в общем случае
        df["Оборачиваемость"] = (
            (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        )

        # Дни планирования по кластерам
        clusters_norm = df["Кластер"].astype(str).str.strip()
        df["Дни планирования"] = (
            clusters_norm.map(cluster_planning_days).fillna(DEFAULT_PLANNING_DAYS).astype(int)
        )
        days = df["Дни планирования"]
        turnover = df["Оборачиваемость"]

        # Потребность и дефициты
        demand = turnover * days
        stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        stock_with_fbs = (
            df["Остатки"] + df["Остатки_fbs"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]
        )

        deficit = stock - demand
        deficit_fbs = stock_with_fbs - demand

        df["Потребность"] = demand
        df["Дефицит/Избыток"] = deficit
        df["Дефицит/Избыток с учетом FBS"] = deficit_fbs
        df["Потребность (округл.)"] = np.where(
            deficit > 0, 0, np.ceil(np.abs(deficit))
        )
        df["Потребность с учетом FBS (округл.)"] = np.where(
            deficit_fbs > 0, 0, np.ceil(np.abs(deficit_fbs))
        )
        df["Корректировка с учетом нулевых остатков"] = df["Потребность (округл.)"]

        df["№"] = np.arange(1, len(df) + 1)
        return df

    # ----------------- Общая логика для остальных клиентов -----------------
    df["Оборачиваемость"] = (
        (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
    )
    stock = df["Остатки"] + df["ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]

    df = add_demand_columns(df, stock, DEMAND_PERIOD_1)
    df = add_demand_columns(df, stock, DEMAND_PERIOD_2)

    df["Корректировка с учетом нулевых остатков"] = df[
        f"Потребность на {demand_period_for_total} дней (округл.)"
    ]
    df["№"] = np.arange(1, len(df) + 1)
    return df

def create_pivot_clusters(svod_for_clusters: pd.DataFrame) -> pd.DataFrame:
    """
    Создание сводной таблицы по кластерам:
    - 'Потребность <Кластер>' по 'Корректировке'
    - 'Остаток <Кластер>' по 'Остатки'

    Если есть колонка 'Кластер для группировки', используем её как измерение,
    иначе — исходный 'Кластер'.
    """
    pivot_col = "Кластер для группировки" if "Кластер для группировки" in svod_for_clusters.columns else "Кластер"

    pivot_values = [
        "Корректировка с учетом нулевых остатков",
        "Остатки",
    ]
    if INCLUDE_PREVIOUS_SUPPLIES:
        pivot_values.append("ПРОШЛЫЕ_ПОСТАВКИ")

    df_pivot = pd.pivot_table(
        svod_for_clusters,
        columns=pivot_col,
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

    clusters = svod_for_clusters[pivot_col].unique()
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

    Для RiverBoats:
      - Оборачиваемость считается только по заказам: Заказы / DAYS_IN_PERIOD.
      - Потребность на DEMAND_PERIOD_1/2 считается от этой оборачиваемости.
      - Добавляется колонка 'ABC по заказам'.

    Для клиентов BIO и ARK:
      - Для товаров со статусом 'КГТ' оборачиваемость считается только по заказам:
          Заказы / DAYS_IN_PERIOD.
      - Для остальных товаров (в т.ч. 'Обычный товар') оборачиваемость считается как раньше:
          (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.

    Для остальных клиентов:
      - Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.
      - Потребность на DEMAND_PERIOD_1/2 дней через add_demand_columns.
    """
    df = metrics_df_by_sku.copy()

    # ----------------- Спец-логика для RiverBoats -----------------
    if client_name == "RiverBoats":
        # Оборачиваемость: только заказы
        df["Оборачиваемость"] = df["Заказы"] / DAYS_IN_PERIOD

        # Защита от NaN
        for col in ["Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
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

        # --- ABC по заказам ---
        df["Заказы_tmp"] = df["Заказы"].fillna(0)
        total_orders = df["Заказы_tmp"].sum()
        if total_orders > 0:
            # Сортировка по заказам (по убыванию)
            df = df.sort_values(by="Заказы_tmp", ascending=False)
            # Доля и накопленная доля
            df["Доля_заказов"] = df["Заказы_tmp"] / total_orders
            df["Накопленная_доля_заказов"] = df["Доля_заказов"].cumsum()

            def classify_abc(cum_share: float) -> str:
                if cum_share <= 0.8:
                    return "A"
                elif cum_share <= 0.95:
                    return "B"
                else:
                    return "C"

            df["ABC по заказам"] = df["Накопленная_доля_заказов"].apply(classify_abc)
            # Возвращаем исходный порядок строк
            df = df.sort_index()
        else:
            df["ABC по заказам"] = "C"

        # Нумерация строк
        df["№"] = np.arange(1, len(df) + 1)

        # Убираем временные колонки
        df.drop(
            columns=["Заказы_tmp", "Доля_заказов", "Накопленная_доля_заказов"],
            inplace=True,
            errors="ignore",
        )
        return df

    # ----------------- Логика для клиентов BIO и ARK -----------------
    if client_name in ("BIO", "ARK"):
        # Подстрахуемся по NaN в Остатках и Поставках
        for col in ["Остатки", "Остатки_fbs", "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ", "Заказы", "Продажи"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Базовая оборачиваемость как "общая" формула
        base_turnover = (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        # Оборачиваемость только по заказам
        orders_turnover = df["Заказы"] / DAYS_IN_PERIOD

        # Статус КГТ считается отдельной логикой:
        # предполагаем, что в справочнике/своде уже есть колонка "Габариты товара"
        # или "Габариты товара (справочно)" / "Габариты товара (расчет)".
        # Здесь ориентируемся именно на статус товара, как ты описал:
        # если статус == "КГТ", используем orders_turnover, иначе — base_turnover.
        if "Габариты товара" in df.columns:
            is_kgt = df["Габариты товара"].astype(str).str.strip() == "КГТ"
        else:
            # Если нет, пытаемся по колонке статуса товара
            is_kgt = df.get("Статус", "").astype(str).str.strip() == "КГТ"

        df["Оборачиваемость"] = np.where(is_kgt, orders_turnover, base_turnover)

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

        df["№"] = np.arange(1, len(df) + 1)
        return df

    # ----------------- Общая логика для остальных клиентов -----------------
    df["Оборачиваемость"] = (
        (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
    )

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

    df["№"] = np.arange(1, len(df) + 1)
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
    через merge с pivot-таблицей по ключу (по умолчанию 'Артикул').
    """
    key_col = "Артикул"

    svod_for_sku = metrics_df_by_sku_stats.merge(
        df_pivot_by_clusters,
        how="left",
        on=key_col,
    )

    # Колонки, которые пришли из pivot-таблицы (кроме ключа)
    new_cols = [c for c in df_pivot_by_clusters.columns if c != key_col]
    svod_for_sku[new_cols] = svod_for_sku[new_cols].fillna(0)

    # Для некоторых клиентов добавляем доп. столбцы
    if client_name in ['RiverBoats']:
        # Находим все колонки с потребностью по кластерам:
        # 'Потребность <Кластер>' (исключаем колонки типа 'Потребность на N дней'
        # и саму итоговую колонку 'Потребность по кластерам всего', если она уже есть)
        demand_cluster_cols = [
            col
            for col in svod_for_sku.columns
            if col.startswith("Потребность ")
            and not col.startswith("Потребность на ")
            and col != "Потребность по кластерам всего"
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
    catalog_processed: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в свод колонки из справочной таблицы (по 'Артикул') и
    формирует итоговую колонку 'Габариты товара':
      1) используем 'Габариты товара (справочно)' из справочника;
      2) если пусто — 'Габариты товара (расчет)' из списка товаров;
      3) если всё ещё пусто — 'Стандарт'.

    Если стоит флаг, то фильтрует только товары, которые есть в справочнике.
    """
    df = df_with_stats.copy()

    # Приведение типов артикула к строке
    df["Артикул"] = df["Артикул"].astype(str)
    catalog_use = catalog_processed.copy()
    catalog_use["Артикул"] = catalog_use["Артикул"].astype(str)

    # Мерж справочника по Артикулу
    df = df.merge(
        catalog_use[catalog_supply_columns],
        how="left",
        on="Артикул",
        indicator='Наличие товара в справочнике',
    )

    # Заполнение значений по умолчанию для кратности
    fill_values = {"Кратность коробов": DEFAULT_BOX_MULTIPLICITY}
    df = df.fillna(value=fill_values)

    # Итоговая колонка "Габариты товара":
    # 1) справочно, 2) расчет, 3) "Стандарт"
    has_ref = "Габариты товара (справочно)" in df.columns
    has_calc = "Габариты товара (расчет)" in df.columns

    if has_ref and has_calc:
        df["Габариты товара"] = (
            df["Габариты товара (справочно)"]
            .fillna(df["Габариты товара (расчет)"])
            .fillna("Стандарт")
        )
    elif has_ref:
        df["Габариты товара"] = df["Габариты товара (справочно)"].fillna("Стандарт")
    elif has_calc:
        df["Габариты товара"] = df["Габариты товара (расчет)"].fillna("Стандарт")
    else:
        # На крайний случай, если ни одной колонки нет
        df["Габариты товара"] = "Стандарт"

    # Если стотит флаг "Включить товары только из справочника", фильтруем по индикатору merge
    if only_reference_table_products_flag:
        df = df[df['Наличие товара в справочнике'] == 'both'].reset_index(drop=True).copy()
        df.drop(columns=['Наличие товара в справочнике'], inplace=True)

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
            "Продажи_факт_2м": "Продажи (факт) за 2 месяца",
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
    catalog_processed: pd.DataFrame,
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
        catalog_processed,
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
    logger.info(f"Сохранение расчета поставок в Excel для клиента {client_name}")
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

    logger.success(f"Расчет поставок без форматирования сохранен в: {file_path}")
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
    logger.info(f"Начало форматирования расчета поставок для клиента {client_name}")

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
        demand_period_1=DEMAND_PERIOD_1,
        demand_period_2=DEMAND_PERIOD_2,
        days_in_period=DAYS_IN_PERIOD,
        apply_box_rounding=APPLY_BOX_ROUNDING,
    )

    filename_formatted = (
        path_supply_svod
        / f"{date_metrics}_Расчет_поставок_{client_name}_Ozon_formatted.xlsx"
    )
    logger.success(f"Расчет поставок сохранен в: {filename_formatted}")

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

# %%
def create_supply_svod_report(date_metrics: str = str(date.today())) -> None:
    """
    Основная функция формирования отчета по поставкам Ozon.
    Логика порядка шагов:
      1) товары → габариты (расчет);
      2) справочник → габариты (справочно), кратность и пр.;
      3) метрики по кластерам + добавление недостающих кластеров;
      4) расчёт свода по кластерам с учётом BIO/ARK/RiverBoats;
      5) объединение кластеров и pivot;
      6) метрики по SKU, расчёт свода по SKU;
      7) добавление pivot по кластерам к SKU;
      8) подготовка Excel и форматирование.
    """
    logger.info(
        f"Создание отчета поставок для клиента {client_name} на дату {date_metrics}"
    )
    logger.info(f"Период оборачиваемости: {DAYS_IN_PERIOD} дней")
    logger.info(
        f"Периоды потребности: {DEMAND_PERIOD_1} дней и {DEMAND_PERIOD_2} дней"
    )

    # ---------- 1. Товары (список) ----------
    products_raw = read_products_file(date_metrics)
    products_prepared = process_product_list(products_raw)

    # ---------- 2. Справочник ----------
    catalog_raw = get_reference_table()
    catalog_prepared = process_catalog(catalog_raw)

    # ---------- 3. Метрики по кластерам (claster_report + добавление недостающих) ----------
    metrics_clusters_full = add_clusters(date_metrics)

    # 3.1. Добавляем габариты из списка товаров (Габариты товара (расчет))
    metrics_clusters_with_products = add_columns_from_products(
        metrics_clusters_full,
        products_prepared,
    )

    # 3.2. Добавляем данные из справочника (итоговая 'Габариты товара', кратность)
    metrics_clusters_enriched = add_columns_from_catalog(
        metrics_clusters_with_products,
        catalog_prepared,
    )

    # ---------- 4. Расчет свода по кластерам (без объединения) ----------
    cluster_planning_days = get_cluster_planning_days_for_client(
        client_name,
        date_metrics=date_metrics,
        use_all_clusters=True
    )
    svod_clusters_initial = calc_svod_for_clusters(
        metrics_clusters_enriched,
        cluster_planning_days,
    )

    # Pivot по "сырым" кластерам (для листа "По кластерам")
    pivot_clusters_raw = create_pivot_clusters(svod_clusters_initial)

    # ---------- 5. Объединение кластеров и расчет по объединённым ----------
    clusters_mapping = get_clusters_mapping_for_client(
        client_name,
        date_metrics=date_metrics,
        use_all_clusters=True
    )

    # ДОБАВЛЯЕМ КОЛОНКУ "Кластер для группировки" В СВОД ПО КЛАСТЕРАМ
    # (это нужно, чтобы она попала на лист "По кластерам")
    svod_clusters_initial = replace_clusters(
        clusters_mapping,
        svod_clusters_initial,
        cluster_col="Кластер",
        new_col="Кластер для группировки",
        replace_original=False,
    )
    metrics_clusters_grouped = replace_clusters(
        clusters_mapping,
        metrics_clusters_enriched,
    )

    svod_clusters_grouped = calc_svod_for_clusters(
        metrics_clusters_grouped,
        cluster_planning_days,
    )

    pivot_clusters_grouped = create_pivot_clusters(svod_clusters_grouped)

    # ---------- 6. Метрики по SKU (summary) ----------
    metrics_sku_raw = read_metrics_file(date_metrics, sheet_name="summary")

    # 6.1. Обогащаем SKU данными из списка товаров
    metrics_sku_with_products = add_columns_from_products(
        metrics_sku_raw,
        products_prepared,
    )

    # 6.2. Обогащаем SKU данными из справочника (итоговая 'Габариты товара', кратность)
    metrics_sku_enriched = add_columns_from_catalog(
        metrics_sku_with_products,
        catalog_prepared,
    )

    # 6.3. Считаем свод по SKU (BIO/ARK/RiverBoats — со своей логикой)
    svod_sku_with_stats = calc_svod_for_sku(metrics_sku_enriched)

    # ---------- 7. Добавляем pivot по кластерам в свод по SKU ----------
    svod_sku_with_clusters = add_demand_for_cluster_columns(
        pivot_clusters_grouped,
        svod_sku_with_stats,
    )

    # ---------- 8. Подготовка данных для Excel ----------
    report_dates_df = read_dates_file(date_metrics)
    date_start_period = report_dates_df["date_start_file"][0]
    date_end_period = report_dates_df["date_end_file"][0]
    cluster_list = list(clusters_mapping.keys())

    # Лист "Всего"
    svod_excel_total = add_columns_for_excel(
        svod_sku_with_clusters,
        cluster_list,
        date_metrics,
        date_start_period,
        date_end_period,
        svod_type="Всего",
    )

    # Лист "Всего по ювелирным кластерам"
    svod_excel_jewelry = create_jewelry_clusters_svod(
        metrics_clusters_enriched,
        date_metrics,
        products_prepared,
        catalog_prepared,
    )

    # Лист "По кластерам"
    svod_excel_clusters = add_columns_for_excel(
        svod_clusters_initial,
        cluster_list,
        date_metrics,
        date_start_period,
        date_end_period,
        svod_type="По кластерам",
    )

    # ---------- 9. Сохранение и форматирование Excel ----------
    save_sheets_to_excel(
        svod_excel_total,
        svod_excel_clusters,
        pivot_clusters_raw,
        svod_excel_jewelry,
        date_metrics,
    )
    # %%
    format_excel_file(
        client_name,
        date_metrics,
        svod_excel_total,
        svod_excel_jewelry,
        svod_excel_clusters,
        clusters_mapping,
        clusters_mapping_jewelry,
    )
    # %%

    logger.success(
        f"Отчет поставок Ozon успешно создан для {client_name} на дату {date_metrics}"
    )

# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================

# %%
if __name__ == "__main__":

    date_metrics = cli_reference_date(parser)
    # date_metrics = '2026-07-03'
    create_supply_svod_report(date_metrics)

# %%
