"""
Модуль для расчета отчета по поставкам для маркетплейса Wildberries.
Скрипт создает Excel-файл с расчетами потребности в товарах по складам,
анализируя продажи, заказы и остатки за определенный период.
"""
# %%
import argparse
from datetime import date, datetime
from itertools import product
from pathlib import Path
from typing import List, Optional, Tuple
from functools import lru_cache


import numpy as np
import pandas as pd
import difflib
from rich_argparse.contrib import ParagraphRichHelpFormatter

from wb.scripts.wb_logger import wb_client_logger as logger
# from wb.scripts.uploadDataFromWB import get_warehouses_load_full
import wb.scripts.format_supply_svod as format_supply_svod
from wb.scripts.get_reference_table import get_reference_table
from wb.scripts.command_args_parser import cli_reference_date
from wb.scripts.constants import (
    headers,
    article_replacement,
    WB_MERGE_KEY,
    cargo_type,
    catalog_supply_columns,
    cargo_multiplicity,
    only_reference_table_products_flag,
    default_warehouse_planning_days,
    client_warehouse_planning_overrides,
    warehouse_days_default,
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
WB_WAREHOUSES_MAPPING_FILE = marketplace_dir_path / "scripts" / "wb_warehouses_mapping.xlsx"


# ============================================================================
# ФУНКЦИИ НАХОЖДЕНИЯ ПОХОЖИХ СКЛАДОВ
# ============================================================================

def _find_similar_warehouses_single_source(
    warehouse: str,
    df_source: pd.DataFrame,
    max_matches: int = 5,
    cutoff: float = 0.6,
) -> list[str]:
    """
    Ищет похожие склады в одном DataFrame с колонкой 'Склад'.
    Возвращает список строк вида '<склад> (строки: 5, 7, ...)'.
    Если ничего не найдено — возвращает пустой список.
    """
    if df_source is None or df_source.empty or "Склад" not in df_source.columns:
        return []

    known = df_source["Склад"].astype(str)
    known_list = known.tolist()

    similar = difflib.get_close_matches(
        warehouse,
        known_list,
        n=max_matches,
        cutoff=cutoff,
    )
    if not similar:
        return []

    parts: list[str] = []
    for sim in similar:
        idx_list = df_source.index[known == sim].tolist()
        excel_rows = [i + 2 for i in idx_list]  # +2: индекс с 0, + заголовок
        parts.append(f"{sim} (строки: {', '.join(map(str, excel_rows))})")

    return parts

@lru_cache(maxsize=None)
def load_wb_warehouses_mapping() -> dict[str, pd.DataFrame]:
    """
    Читает wb_warehouses_mapping.xlsx и возвращает словарь {sheet_name: df},
    где в df есть колонка 'Склад'.
    """
    sheets: dict[str, pd.DataFrame] = {}
    try:
        xls = pd.ExcelFile(WB_WAREHOUSES_MAPPING_FILE)
    except FileNotFoundError:
        logger.warning(
            f"Файл маппинга складов WB не найден: {WB_WAREHOUSES_MAPPING_FILE}"
        )
        return sheets

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception as e:
            logger.warning(
                f"Не удалось прочитать лист '{sheet_name}' "
                f"из {WB_WAREHOUSES_MAPPING_FILE}: {e}"
            )
            continue
        if "Склад" not in df.columns:
            continue
        df = df.copy()
        df["Склад"] = df["Склад"].astype(str).str.strip()
        sheets[sheet_name] = df

    return sheets



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
    # Переименовываем некоторые колонки для унификации
    # metrics = metrics.rename(columns={
    #     'Артикул_размер': 'Артикул_Размер',
    # })
    return metrics.copy()


def process_catalog(catalog: pd.DataFrame) -> pd.DataFrame:
    """
    Обрабатывает справочную таблицу WB.
    """
    # catalog_path = CATALOG_DIR / f"Справочная_таблица_{client_name}_WB.xlsx"
    # catalog = pd.read_excel(catalog_path)
    catalog_processed = catalog.copy()

    if "Размер" not in catalog_processed.columns:
        catalog_processed["Размер"] = '0'

    catalog_processed["Размер"] = catalog_processed["Размер"].fillna('0')

    # Делаем первую букву в названиях столбцов заглавной, остальные - строчными
    catalog_processed.columns = [col.capitalize() for col in catalog_processed.columns]

    # Переименовываем ещё некоторые колонки
    catalog_processed = catalog_processed.rename(columns={
        # "Артикул продавца": "Артикул",
        'Артикул wb': 'Артикул WB',
        "Артикул 1с": "Артикул 1С",
        "Сортируемость товара": "Сортируемость",
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

    if not {"Кратность коробов", "Кратность короба"} & set(catalog_processed.columns):
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
        catalog_processed["Размер"] = catalog_processed["Размер"].fillna('0')

    return catalog_processed


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

def load_wb_warehouses_load(date_metrics: str) -> pd.DataFrame:
    """
    Читает файл с загруженностью складов WB за указанную дату выгрузки
    и переименовывает ключевые столбцы:
      warehouseName -> Склад
      boxAcceptanceType -> Ограничения

    Дополнительно:
      - если название склада заканчивается на «СГТ» (с пробелом или без),
        оно преобразуется в «СГТ <имя_склада_без_СГТ>».

    Ожидаемый путь:
      <marketplace_dir_path>/Clients/<client_name>/UploadFiles/UploadFiles_<date_metrics>/
        <date_metrics>_Склады_WB.csv
    """
    uploaddir = (
        Path(marketplace_dir_path)
        / "Clients"
        / client_name
        / "UploadFiles"
    )
    uploaddir_metrics = uploaddir / f"UploadFiles_{date_metrics}"
    file_path = uploaddir_metrics / f"{date_metrics}_Склады_WB.csv"

    try:
        df_warehouses = pd.read_csv(file_path, sep=";")

    except FileNotFoundError:
        logger.error(
            "Файл с загруженностью складов WB не найден: {}",
            file_path,
        )
        return pd.DataFrame()

    except pd.errors.EmptyDataError:
        logger.warning(
            "Файл с загруженностью складов WB пустой: {}",
            file_path,
        )
        return pd.DataFrame()

    # Переименование столбцов
    rename_map = {
        "warehouseName": "Склад",
        "boxAcceptanceType": "Ограничения",
    }
    existing_rename = {
        col_src: col_dst
        for col_src, col_dst in rename_map.items()
        if col_src in df_warehouses.columns
    }

    if existing_rename:
        df_warehouses = df_warehouses.rename(columns=existing_rename)

    # Переименование складов: «... СГТ» / «...СГТ» -> «СГТ ...»
    if "Склад" in df_warehouses.columns:

        def _normalize_sgt(name: str) -> str:
            if not isinstance(name, str):
                return name

            s = name.strip()

            if s.endswith("СГТ"):
                base = s[:-3].rstrip()

                if base:
                    return f"СГТ {base}"

                return "СГТ"

            return s

        # map без astype(str), чтобы не превращать NaN в строку "nan"
        df_warehouses["Склад"] = df_warehouses["Склад"].map(_normalize_sgt)

    logger.info(
        "Считан файл с загруженностью складов WB: {}, строк: {}, колонок: {}",
        file_path.name,
        len(df_warehouses),
        len(df_warehouses.columns),
    )

    return df_warehouses

def read_products(date_metrics: str) -> pd.DataFrame:
    """
    Читает файл со списком товаров без дополнительной обработки.
    """
    file_path = (
        UPLOAD_DIR
        / f"UploadFiles_{date_metrics}"
        / f"{date_metrics}_Товары.csv"
    )

    try:
        df_products = pd.read_csv(file_path, sep=";")
    except FileNotFoundError:
        logger.error(f"Файл со списком товаров не найден: {file_path}")
        return pd.DataFrame()

    logger.info(
        f"Считан файл со списком товаров: {file_path.name}, "
        f"строк: {len(df_products)}, колонок: {len(df_products.columns)}"
    )
    return df_products


def process_products(
    df_products: pd.DataFrame,
    article_replacement: Optional[dict] = None,
) -> pd.DataFrame:
    """
    Обработка списка товаров:
    - нормализует 'Артикул продавца';
    - нормализует/создает 'Размер' (строковый, '0' при отсутствии);
    - приводит штрихкоды к строке;
    - при необходимости заменяет артикулы по словарю article_replacement и удаляет дубликаты;
    - формирует колонку 'Артикул_Размер';
    - добавляет признак 'Габариты товара' по габаритам (СГТ/КГТ/Обычный/...).
    """
    df = df_products.copy()

    # 1. Артикул продавца
    if "Артикул продавца" not in df.columns:
        logger.warning(
            "В списке товаров не найдена колонка 'Артикул продавца'. "
            "Дальнейшая обработка может быть некорректной."
        )
    else:
        df["Артикул продавца"] = (
            df["Артикул продавца"]
            .astype(str)
            .str.replace("'", "", regex=False)
            .str.strip()
        )

    # 2. Размер
    if "Размер" in df.columns:
        if client_name in ["KU_And_KU", "Soyuz"]:
            df["Размер"] = (
                df["Размер"]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .apply(lambda x: pd.to_numeric(x, errors="coerce"))
                .fillna(0)
            )
        df["Размер"] = df["Размер"].astype(str)
    else:
        # Если колонки 'Размер' нет — создаём её и заполняем строковым '0'
        df["Размер"] = "0"

    # 3. Штрихкоды → строка (если есть)
    for bc_col in ["Штрихкод", "barcode", "last_barcode"]:
        if bc_col in df.columns:
            df[bc_col] = df[bc_col].astype(str)

    # 4. Замена артикулов по словарю и удаление дубликатов после замены
    if article_replacement and "Артикул продавца" in df.columns:
        df["Артикул продавца"] = df["Артикул продавца"].replace(article_replacement)

    if "Артикул продавца" in df.columns:
        df = df.drop_duplicates(subset=["Артикул продавца", "Размер"])

    # 5. Колонка Артикул_Размер
    if "Артикул продавца" in df.columns:
        df["Артикул_Размер"] = (
            df[["Артикул продавца", "Размер"]]
            .apply(lambda row: "_size_".join(row.values.astype(str)), axis=1)
        )

    # 6. Признак крупногабаритного товара (СГТ/КГТ)
    # Приводим габариты и вес к числу, NaN -> 0
    for col in ["Длина, см", "Ширина, см", "Высота, см", "Вес, кг"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0.0

    # Сумма трёх сторон
    df["Сумма_3_сторон"] = df["Длина, см"] + df["Ширина, см"] + df["Высота, см"]

    # Базовый признак
    df["Габариты товара (расчет)"] = "Стандарт"

    # Опционально: некорректные габариты
    if "Корректность габаритов" in df.columns:
        mask_bad_dims = df["Корректность габаритов"].astype(str).str.contains(
            "некоррект", case=False, na=False
        )
        df.loc[mask_bad_dims, "Габариты товара (расчет)"] = "Некорректные габариты"

    # --- Условия СГТ ---
    # Вес одной коробки — от 25 до 100 кг включительно
    cond_sgt_weight = (df["Вес, кг"] >= 25) & (df["Вес, кг"] <= 100)
    # Длина хотя бы одной стороны упаковки — от 121 до 400 см включительно
    cond_sgt_side = (
        df["Длина, см"].between(121, 400)
        | df["Ширина, см"].between(121, 400)
        | df["Высота, см"].between(121, 400)
    )
    # Сумма трех сторон упаковки — от 201 до 1 200 см включительно
    cond_sgt_sum = df["Сумма_3_сторон"].between(201, 1200)
    # Если выполняется хотя бы одно из условий — СГТ
    mask_sgt = cond_sgt_weight | cond_sgt_side | cond_sgt_sum
    df.loc[mask_sgt, "Габариты товара (расчет)"] = "СГТ"

    # --- Условия КГТ ---
    mask_not_sgt = ~mask_sgt
    # Вес менее 25 кг
    cond_kgt_weight = df["Вес, кг"] < 25
    # Одна сторона от 121 до 200 см
    cond_kgt_side = (
        df["Длина, см"].between(121, 200)
        | df["Ширина, см"].between(121, 200)
        | df["Высота, см"].between(121, 200)
    )
    # Сумма трех сторон от 201 до 280 см
    cond_kgt_sum = df["Сумма_3_сторон"].between(201, 280)
    # КГТ: вес < 25 кг и (одна сторона в диапазоне или сумма сторон в диапазоне)
    mask_kgt = mask_not_sgt & cond_kgt_weight & (cond_kgt_side | cond_kgt_sum)
    df.loc[mask_kgt, "Габариты товара (расчет)"] = "КГТ"

    # Вспомогательный столбец можно убрать, если не нужен в отчёте
    df.drop(columns=["Сумма_3_сторон"], inplace=True)

    return df

# ============================================================================
# ФУНКЦИИ ОБРАБОТКИ КЛАСТЕРОВ
# ============================================================================

def get_warehouse_planning_days_for_client(client_name: str) -> dict:
    """
    Возвращает словарь 'Склад -> дни планирования' для клиента WB:
    - берёт default_warehouse_planning_days как базу;
    - поверх накладывает client_warehouse_planning_overrides[client_name], если есть.
    """
    planning = default_warehouse_planning_days.copy()
    overrides = client_warehouse_planning_overrides.get(client_name)
    if overrides:
        planning.update(overrides)
    return planning



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
    df_with_clusters: pd.DataFrame,
    sgt_labels: tuple[str, ...] = ("СГТ", "КГТ"),
    aggregate: bool = True,
    log_unmapped: bool = True,
) -> pd.DataFrame:
    """
    Объединяет склады в кластеры согласно маппингу, с учётом габаритов товара.

    Параметры
    ----------
    clusters_mapping_df : pd.DataFrame
        Таблица маппинга складов WB. Ожидаемые колонки:
          - 'Склад' — исходное имя склада WB;
          - 'Группировка' — целевой кластер для обычных товаров;
          - 'Группировка СГТ' — целевой кластер для СГТ/КГТ‑товаров.
        Перед использованием дубликаты по 'Склад' удаляются, чтобы
        исключить размытие данных при merge.

    df_with_clusters : pd.DataFrame
        Исходные данные по складам, как минимум с колонками:
          - 'Склад' — имя склада WB (будет нормализовано strip());
          - 'Габариты товара' или 'Габариты товара (расчет)' — тип
            габаритов для определения СГТ/КГТ.
        Все остальные колонки (продажи, заказы, остатки и т.п.) сохраняются.

    sgt_labels : tuple[str, ...], optional
        Набор значений, по которым определяется тяжёлый товар.
        Если значение в 'Габариты товара' / 'Габариты товара (расчет)'
        принадлежит этому набору, склад относится к СГТ/КГТ и для него
        используется 'Группировка СГТ'. По умолчанию: ('СГТ', 'КГТ').

    aggregate : bool, optional
        Управляет агрегацией:
          - True (по умолчанию): после добавления 'Склад для группировки'
            выполняется groupby по колонкам
              ['WB_MERGE_KEY', '№ товара', 'Склад', 'Склад для группировки']
            с суммированием ['Продажи', 'Заказы', 'Ожидаемое поступление',
            'Остатки'] (если эти колонки присутствуют).
          - False: никаких groupby не выполняется, возвращается
            df_with_clusters с добавленным 'Склад для группировки'.

    log_unmapped : bool, optional
        Флаг логирования складов без группировки:
          - True (по умолчанию): для всех складов, у которых и 'Группировка',
            и 'Группировка СГТ' равны NaN, формируется warning с
            перечислением таких складов и списком похожих складов из
            маппинга с номерами строк (Excel‑нумерация).
          - False: логирование пропускается.

    Возвращает
    ----------
    pd.DataFrame
        Если aggregate=False:
          - исходный df_with_clusters с добавленным столбцом
            'Склад для группировки' (и временными колонками маппинга).
        Если aggregate=True:
          - агрегированный DataFrame по складам:
              * группирующие колонки — ['WB_MERGE_KEY', '№ товара',
                'Склад', 'Склад для группировки'] (только те, что есть в df);
              * числовые колонки ['Продажи', 'Заказы',
                'Ожидаемое поступление', 'Остатки'] суммируются.
    """
    df_merged = df_with_clusters.copy()

    # Нормализуем названия складов
    df_merged["Склад"] = df_merged["Склад"].astype(str).str.strip()

    clusters_mapping_norm = clusters_mapping_df.copy()
    clusters_mapping_norm["Склад"] = (
        clusters_mapping_norm["Склад"].astype(str).str.strip()
    )
    # Удаляем дубликаты по "Склад", чтобы при merge не было размытия
    clusters_mapping_norm = clusters_mapping_norm.drop_duplicates(subset=["Склад"])

    # Мержим по 'Склад', забираем обе группировки
    df_merged = df_merged.merge(
        clusters_mapping_norm[["Склад", "Группировка", "Группировка СГТ"]],
        how="left",
        on="Склад",
        suffixes=("", "_map"),
    )

    # ---- Определяем тип габаритов ----
    if "Габариты товара" in df_merged.columns:
        dims = df_merged["Габариты товара"].astype(str).str.strip()
    elif "Габариты товара (расчет)" in df_merged.columns:
        dims = df_merged["Габариты товара (расчет)"].astype(str).str.strip()
    else:
        dims = pd.Series("", index=df_merged.index)

    is_sgt = dims.isin(sgt_labels)

    # Таргет для СГТ/КГТ и обычных
    target_sgt = df_merged["Группировка СГТ"].where(
        df_merged["Группировка СГТ"].notna(), df_merged["Склад"]
    )
    target_norm = df_merged["Группировка"].where(
        df_merged["Группировка"].notna(), df_merged["Склад"]
    )

    # Новый столбец с группировкой для расчёта
    df_merged["Склад для группировки"] = np.where(is_sgt, target_sgt, target_norm)

    # ---- Логирование складов без группировки (опционально) ----
    # ---- Логирование складов без группировки (опционально) ----
    if log_unmapped:
        mask_no_group = (
            df_merged["Группировка"].isna()
            & df_merged["Группировка СГТ"].isna()
        )
        unmapped_warehouses = (
            df_merged.loc[mask_no_group, "Склад"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if unmapped_warehouses:
            unmapped_sorted = sorted(unmapped_warehouses)

            # Короткий warning — список складов без группировки (для копирования в Excel)
            logger.warning(
                "Склады без группировки (можно скопировать в Excel):\n"
                + "\n".join(unmapped_sorted)
            )

            # Загружаем внешние листы маппинга (кэшируются)
            mapping_sheets = load_wb_warehouses_mapping()
            df_boxes = mapping_sheets.get("Короба")
            df_pallets = mapping_sheets.get("Монопалеты")

            # Подробные списки по каждому источнику
            # 1) клиентский маппинг (clusters_mapping_norm)
            client_lines: list[str] = []
            for wh in unmapped_sorted:
                sim_client = _find_similar_warehouses_single_source(
                    warehouse=wh,
                    df_source=clusters_mapping_norm,
                    max_matches=5,
                    cutoff=0.6,
                )
                if sim_client:
                    client_lines.append(f"{wh}: " + "; ".join(sim_client))
                else:
                    client_lines.append(f"{wh}: похожие склады не найдены")

            # 2) Короба
            boxes_lines: list[str] = []
            for wh in unmapped_sorted:
                sim_boxes = _find_similar_warehouses_single_source(
                    warehouse=wh,
                    df_source=df_boxes,
                    max_matches=5,
                    cutoff=0.6,
                )
                if sim_boxes:
                    boxes_lines.append(f"{wh}: " + "; ".join(sim_boxes))
                else:
                    boxes_lines.append(f"{wh}: похожие склады не найдены")

            # 3) Монопаллеты
            pallets_lines: list[str] = []
            for wh in unmapped_sorted:
                sim_pallets = _find_similar_warehouses_single_source(
                    warehouse=wh,
                    df_source=df_pallets,
                    max_matches=5,
                    cutoff=0.6,
                )
                if sim_pallets:
                    pallets_lines.append(f"{wh}: " + "; ".join(sim_pallets))
                else:
                    pallets_lines.append(f"{wh}: похожие склады не найдены")

            # Выводим отдельно по клиенту, коробам и монопаллетам (debug, чтобы не засорять обычные логи)
            logger.debug(
                "Похожие склады по КЛИЕНТСКОМУ маппингу:\n  "
                + "\n  ".join(client_lines)
            )
            logger.debug(
                "Похожие склады по листу 'Короба':\n  "
                + "\n  ".join(boxes_lines)
            )
            logger.debug(
                "Похожие склады по листу 'Монопаллеты':\n  "
                + "\n  ".join(pallets_lines)
            )

    # ---- Если агрегация не нужна — возвращаем df_merged как есть ----
    if not aggregate:
        return df_merged

    # ---- Агрегация: используем 'Склад для группировки', исходный 'Склад' сохраняем ----
    group_columns = [
        WB_MERGE_KEY,
        "№ товара",
        "Склад",              # исходный склад
        "Склад для группировки",  # агрегирующий кластер
    ]
    sum_columns = ["Продажи", "Заказы", "Ожидаемое поступление", "Остатки"]

    # Оставляем только существующие колонки из group_columns + sum_columns
    existing_group_cols = [c for c in group_columns if c in df_merged.columns]
    existing_sum_cols = [c for c in sum_columns if c in df_merged.columns]

    df_aggregated = (
        df_merged[existing_group_cols + existing_sum_cols]
        .groupby(existing_group_cols, as_index=False)
        .sum()
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
    warehouse_planning_days: dict,
    default_planning_days: int,
) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости по складам (лист 'По кластерам').

    Для клиентов BIO и ARK:
      - Оборачиваемость:
          * для КГТ/СГТ: Заказы / DAYS_IN_PERIOD;
          * для остальных товаров: (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.
      - Дни планирования:
          * берутся из warehouse_planning_days по имени склада
            (Склад для группировки, если есть; иначе Склад);
          * если склад не найден в словаре (в т.ч. СГТ-склады) —
            используется default_planning_days (дефолт для клиента).
      - Потребность = Оборачиваемость * Дни планирования.
      - Считаются:
          'Потребность',
          'Дефицит/Избыток',
          'Дефицит/Избыток с учетом FBS',
          'Потребность (округл.)',
          'Потребность с учетом FBS (округл.)',
          'Корректировка с учетом нулевых остатков' = 'Потребность (округл.)'.

    Для остальных клиентов:
      - Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.
      - Потребность на DEMAND_PERIOD_1/2 через add_demand_columns.
      - 'Корректировка с учетом нулевых остатков' =
         'Потребность на {DEMAND_PERIOD_2} дней (округл.)'.
    """
    df = metrics_df_with_clusters.copy()

    # Числа без NaN
    for col in ["Заказы", "Продажи", "Остатки", "Остатки_fbs", "Ожидаемое поступление"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    # ----------------- Логика для BIO и ARK -----------------
    if client_name in ("BIO", "ARK"):
        # Базовая оборачиваемость (как обычно)
        base_turnover = (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        # Только по заказам
        orders_turnover = df["Заказы"] / DAYS_IN_PERIOD

        # Признак КГТ/СГТ по габаритам
        if "Габариты товара" in df.columns:
            dims = df["Габариты товара"].astype(str).str.strip()
        elif "Габариты товара (расчет)" in df.columns:
            dims = df["Габариты товара (расчет)"].astype(str).str.strip()
        else:
            dims = pd.Series("", index=df.index)

        is_heavy = dims.isin(["КГТ", "СГТ"])

        # Оборачиваемость: тяжелые → только по заказам, остальные → базовая
        df["Оборачиваемость"] = np.where(is_heavy, orders_turnover, base_turnover)

        # Имя склада для планирования: Склад для группировки (если есть), иначе Склад
        warehouse_col = (
            "Склад для группировки"
            if "Склад для группировки" in df.columns
            else "Склад"
        )
        warehouses_norm = df[warehouse_col].astype(str).str.strip()

        # Дни планирования: из словаря, если нет — дефолт по клиенту
        df["Дни планирования"] = (
            warehouses_norm.map(warehouse_planning_days)
            .fillna(default_planning_days)
            .astype(int)
        )

        days = df["Дни планирования"]
        turnover = df["Оборачиваемость"]

        demand = turnover * days

        stock = df["Остатки"] + df["Ожидаемое поступление"]
        stock_with_fbs = (
            df["Остатки"] + df["Остатки_fbs"] + df["Ожидаемое поступление"]
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
    df_warehouses_limits: Optional[pd.DataFrame] = None,
    default_limit: str | int = "Неизвестно",
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Создание сводной таблицы по кластерам (аналог Ozon):
      - 'Потребность <Склад>' по 'Корректировка с учетом нулевых остатков'
      - 'Остаток <Склад>' по 'Остатки'
      - 'Ограничения <Склад>' по колонке 'Ограничения' из справочника складов.

    index: Артикул_Размер
    columns: Склад

    :param svod_for_clusters: DataFrame с расчётом по складам.
    :param df_warehouses_limits: DataFrame с колонками 'Склад' и 'Ограничения'
                                 (результат load_wb_warehouses_load с rename).
    :param default_limit: значение ограничения по умолчанию, если для склада
                          ничего не найдено в справочнике.
    """
    pivot_cluster_col = (
        "Склад для группировки"
        if "Склад для группировки" in svod_for_clusters.columns
        else "Склад"
    )

    # Базовый pivot по потребности и остаткам
    df_pivot = pd.pivot_table(
        svod_for_clusters,
        columns=pivot_cluster_col,
        index=WB_MERGE_KEY,
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

    clusters = svod_for_clusters[pivot_cluster_col].astype(str).unique().tolist()

    # Формируем порядок колонок: Потребность / Остаток / Ограничения для каждого склада
    cols_order = [WB_MERGE_KEY]
    for cluster in clusters:
        cols_order.append(f"Потребность {cluster}")
        cols_order.append(f"Остаток {cluster}")
        cols_order.append(f"Ограничения {cluster}")

    # Добавляем колонки Ограничения <Склад>, заполняя default_limit
    if df_warehouses_limits is not None and not df_warehouses_limits.empty:
        # нормализуем имена складов
        limits = df_warehouses_limits.copy()
        limits["Склад"] = limits["Склад"].astype(str).str.strip()
        limits = limits.set_index("Склад")["Ограничения"].to_dict()
    else:
        limits = {}

    # Для каждого склада создаём колонку с одним и тем же значением ограничения
    for cluster in clusters:
        col_name = f"Ограничения {cluster}"
        limit_value = limits.get(cluster, default_limit)
        # Колонка должна быть того же размера, что и df_pivot
        df_pivot[col_name] = limit_value

    # Оставляем только существующие колонки в заданном порядке
    cols_order = [c for c in cols_order if c in df_pivot.columns]
    df_pivot = df_pivot[cols_order]

    return df_pivot, clusters


def calc_svod_for_sku(metrics_df_by_sku: pd.DataFrame) -> pd.DataFrame:
    """
    Расчет потребности и оборачиваемости для товаров по SKU (лист 'Всего').

    Для клиентов BIO и ARK:
      - если товар КГТ, оборачиваемость считается только по заказам:
          Оборачиваемость = Заказы / DAYS_IN_PERIOD;
      - для остальных товаров:
          Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.

    Для остальных клиентов:
      - Оборачиваемость = (Заказы + Продажи) / 2 / DAYS_IN_PERIOD.

    Во всех случаях:
      - потребность считается на DEMAND_PERIOD_1/2 через add_demand_columns
        от запаса с учётом FBS;
      - нумерация строк в колонке '№'.
    """
    df = metrics_df_by_sku.copy()

    # Защита от NaN в числовых полях
    for col in ["Заказы", "Продажи", "Остатки", "Остатки_fbs", "Ожидаемое поступление"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    # ----------------- Логика для BIO и ARK -----------------
    if client_name in ("BIO", "ARK"):
        # Базовая оборачиваемость: как раньше
        base_turnover = (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        # Оборачиваемость только по заказам
        orders_turnover = df["Заказы"] / DAYS_IN_PERIOD

        # Признак КГТ: сначала по итоговой колонке "Габариты товара",
        # если её нет — пробуем "Габариты товара (расчет)".
        if "Габариты товара" in df.columns:
            is_kgt = df["Габариты товара"].astype(str).str.strip() == "КГТ"
        elif "Габариты товара (расчет)" in df.columns:
            is_kgt = df["Габариты товара (расчет)"].astype(str).str.strip() == "КГТ"
        else:
            is_kgt = pd.Series(False, index=df.index)

        # Оборачиваемость: КГТ — только заказы, остальные — базовая формула
        df["Оборачиваемость"] = np.where(is_kgt, orders_turnover, base_turnover)

    # ----------------- Общая логика для остальных клиентов -----------------
    else:
        df["Оборачиваемость"] = (
            (df["Заказы"] + df["Продажи"]) / 2 / DAYS_IN_PERIOD
        )

    # Запас с учётом FBS
    stock_with_fbs = (
        df["Остатки"] + df["Остатки_fbs"] + df["Ожидаемое поступление"]
    )

    # Потребность на DEMAND_PERIOD_1 дней (без и с учётом FBS)
    df = add_demand_columns(df, stock_with_fbs, DEMAND_PERIOD_1)
    df = add_demand_columns(
        df,
        stock_with_fbs,
        DEMAND_PERIOD_1,
        "с учетом FBS",
    )

    # Потребность на DEMAND_PERIOD_2 дней
    df = add_demand_columns(df, stock_with_fbs, DEMAND_PERIOD_2)

    # Нумерация строк
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
    через merge с pivot-таблицей по WB_MERGE_KEY.

    - Числовые колонки (Потребность/Остаток и т.п.) заполняются нулями.
    - Колонки с ограничениями ('Ограничения <склад>') заполняются
      по всему складу единым значением (по колонке), чтобы внутри склада
      не было NaN.
    """
    df_result = metrics_df_by_sku_with_stats.copy()

    # merge по ключу WB_MERGE_KEY
    df_result = df_result.merge(
        df_pivot_by_clusters,
        how="left",
        on=WB_MERGE_KEY,
    )

    # Колонки, пришедшие из pivot (кроме merge-key)
    new_cols = [c for c in df_pivot_by_clusters.columns if c != WB_MERGE_KEY]

    # Разделим новые колонки на:
    #  - ограничения: начинаются с 'Ограничения '
    #  - остальные: числовые (Потребность/Остаток/Прошлые поставки и т.п.)
    limit_cols = [c for c in new_cols if c.startswith("Ограничения ")]
    numeric_cols = [c for c in new_cols if c not in limit_cols]

    # 1) Числовые новые колонки заполняем нулями
    if numeric_cols:
        df_result[numeric_cols] = df_result[numeric_cols].fillna(0)

    # 2) Колонки с ограничениями:
    #    если в колонке есть NaN, берём первое ненулевое значение
    #    и подставляем его во все NaN (ограничение действует на весь склад).
    for col in limit_cols:
        if col not in df_result.columns:
            continue
        series = df_result[col]
        if series.isna().any():
            default_limit = series.dropna().iloc[0] if series.dropna().size > 0 else "Нет ограничений"
            df_result[col] = series.fillna(default_limit)

    return df_result

def add_columns_from_products(
    svod: pd.DataFrame,
    df_products_processed: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет в свод колонку 'Габариты товара (расчет)' из списка товаров WB.
    Мерж по 'Артикул_Размер'.
    """
    # Берём только нужные поля из списка товаров
    cols_for_merge = [WB_MERGE_KEY, "Габариты товара (расчет)"]
    cols_for_merge = [c for c in cols_for_merge if c in df_products_processed.columns]

    df_products_short = df_products_processed[cols_for_merge].copy()

    svod_with_additional = svod.merge(
        df_products_short,
        how="left",
        on=WB_MERGE_KEY,
    )
    return svod_with_additional

def add_columns_from_catalog(
    metrics_df_by_sku_with_stats: pd.DataFrame,
    catalog_processed: pd.DataFrame,
) -> pd.DataFrame:
    """
    Добавляет колонки из справочной таблицы (по Артикул_Размер) и формирует
    итоговую колонку 'Габариты товара' по аналогии с Ozon:
      1) 'Габариты товара (справочно)' из справочника,
      2) если пусто — 'Габариты товара (расчет)' из метрик/списка товаров,
      3) если всё ещё пусто — 'Стандарт'.
    Если стоит флаг, то фильтрует только товары, которые есть в справочнике.

    Также заполняет 'Кратность коробов' дефолтным значением, если оно отсутствует.
    """
    df_merged = metrics_df_by_sku_with_stats.copy()

    # Merge справочника по Артикул_Размер
    df_merged = df_merged.merge(
        catalog_processed[catalog_supply_columns],
        how="left",
        on=WB_MERGE_KEY,
        suffixes=("", "_catalog_reference"),
        indicator='Наличие товара в справочнике',
    )

    # Заполнение кратности коробов дефолтом
    fill_values = {"Кратность коробов": DEFAULT_BOX_MULTIPLICITY}
    df_merged = df_merged.fillna(value=fill_values)

    # Итоговая колонка "Габариты товара":
    # приоритет — справочник, затем расчет, затем "Стандарт"
    has_ref = "Габариты товара (справочно)" in df_merged.columns
    has_calc = "Габариты товара (расчет)" in df_merged.columns

    if has_ref and has_calc:
        df_merged["Габариты товара"] = (
            df_merged["Габариты товара (справочно)"]
            .fillna(df_merged["Габариты товара (расчет)"])
            .fillna("Стандарт")
        )
    elif has_ref:
        df_merged["Габариты товара"] = (
            df_merged["Габариты товара (справочно)"].fillna("Стандарт")
        )
    elif has_calc:
        df_merged["Габариты товара"] = (
            df_merged["Габариты товара (расчет)"].fillna("Стандарт")
        )
    else:
        df_merged["Габариты товара"] = "Стандарт"

        # Если стотит флаг "Включить товары только из справочника", фильтруем по индикатору merge
    if only_reference_table_products_flag:
        df_merged = df_merged[df_merged['Наличие товара в справочнике'] == 'both'].reset_index(drop=True).copy()
        df_merged.drop(columns=['Наличие товара в справочнике'], inplace=True)


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
    if 'Артикул_Размер' not in df.columns:
        df["Артикул продавца"] = df["Артикул_Размер"].str.split("_size_").str[0]
        df["Размер"] = df["Артикул_Размер"].str.split("_size_").str[1]

    if svod_type == "Всего":
        svod_columns = supply_svod_total_columns.copy()
        cluster_columns = [
            val
            for cluster in cluster_list
            for val in (f"Потребность {cluster}", f"Остаток {cluster}", f"Ограничения {cluster}")
        ]
        svod_columns.extend(cluster_columns)
    else:
        svod_columns = supply_svod_clusters_columns.copy()

    # Удаляем дубликаты колонок
    svod_columns = list(dict.fromkeys(svod_columns))

    for col in svod_columns:
        if col not in df.columns:
            df[col] = np.nan

    df = df[svod_columns]
    df = df.sort_values(by=WB_MERGE_KEY, ignore_index=True)

    date_metrics_dt = datetime.strptime(date_metrics, "%Y-%m-%d")
    rename_mapping = {
        "Наименование товара": "Наименование",
        "Продажи_факт_2м": "Продажи (факт) за 2 месяца",
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
    catalog_processed: pd.DataFrame,
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

    svod_with_catalog = add_columns_from_catalog(svod, catalog_processed)

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
    logger.info(f"Начало сохранения расчета поставок в Excel-файл для клиента {client_name} на дату {date_metrics}")
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

    logger.success(f"Данные сохранены в: {filepath}")


def format_excel_file(
    client_name: str,
    date_metrics: str,
    svod_excel_sku: pd.DataFrame,
    svod_excel_by_sku_jewelry: Optional[pd.DataFrame],
    svod_excel_clusters: pd.DataFrame,
    WB_MERGE_KEY: str = WB_MERGE_KEY,
) -> None:
    """
    Форматирует Excel-файл (формулы, стили) через wb.scripts.format_supply_svod.
    """
    # format_supply_svod.copy_supply_svod_file(
    #     SUPPLY_SVOD_DIR,
    #     client_name,
    #     date_metrics,
    # )

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
        WB_MERGE_KEY=WB_MERGE_KEY,
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
            WB_MERGE_KEY=WB_MERGE_KEY,
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
    logger.success(f"Отформатированный отчет сохранен: {filename}")


# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

# %%
def create_supply_svod_report(date_metrics: str = str(date.today())) -> None:
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

    # ---------- 1. Настройки планирования по складам ----------
    warehouse_days_by_warehouse = get_warehouse_planning_days_for_client(client_name)

    # ---------- 2. Справочник и список товаров ----------
    catalog_df = get_reference_table()
    catalog_processed = process_catalog(catalog_df)

    products_raw = read_products(date_metrics)
    products_prepared = process_products(products_raw, article_replacement)

    # ---------- 3. Метрики по складам (claster_report) ----------
    metrics_clusters_raw = read_metrics_file(
        date_metrics,
        sheet_name="claster_report",
    )

    # 3.1. Добавляем габариты из списка товаров (Габариты товара (расчет))
    metrics_clusters_with_products = add_columns_from_products(
        metrics_clusters_raw,
        products_prepared,
    )

    # 3.2. Добавляем данные из справочника (итоговая 'Габариты товара', кратность)
    metrics_clusters_enriched = add_columns_from_catalog(
        metrics_clusters_with_products,
        catalog_processed,
    )

    # ---------- 4. Маппинг складов ----------
    warehouses_mapping = read_clusters_mapping(cargo_type)
    jewelry_group_column = "Группировка ювелирка"

    # 4.1. Для листа "По кластерам":
    #      добавляем 'Склад для группировки', НО НЕ агрегируем строки
    metrics_clusters_with_calc = replace_clusters(
        clusters_mapping_df=warehouses_mapping,
        df_with_clusters=metrics_clusters_enriched,
        aggregate=False,     # без groupby
        log_unmapped=True,   # логируем незамапленные склады
    )

    # 4.2. Рассчитываем свод "По кластерам" до агрегации по складам
    svod_clusters_initial = calc_svod_for_clusters(
        metrics_clusters_with_calc,
        warehouse_days_by_warehouse,
        warehouse_days_default,
    )

    # 4.3. Для pivot и расчёта свода по складам:
    #      та же функция с агрегированием по 'Склад для группировки'
    metrics_clusters_grouped = replace_clusters(
        clusters_mapping_df=warehouses_mapping,
        df_with_clusters=metrics_clusters_enriched,
        aggregate=True,      # агрегируем по WB_MERGE_KEY + Склад для группировки
        log_unmapped=False,  # повторно не логируем
    )

    svod_clusters_grouped = calc_svod_for_clusters(
        metrics_clusters_grouped,
        warehouse_days_by_warehouse,
        warehouse_days_default,
    )

    # ---------- 5. Pivot по кластерам + ограничения ----------
    warehouses_limits = load_wb_warehouses_load(date_metrics)
    pivot_by_warehouses, warehouse_groups_list = create_pivot_clusters(
        svod_clusters_grouped,
        df_warehouses_limits=warehouses_limits,
        default_limit="Неизвестно",
    )

    # ---------- 6. Метрики по SKU (summary) ----------
    metrics_sku_raw = read_metrics_file(
        date_metrics,
        sheet_name="summary",
    )

    # 6.1. Добавляем габариты из списка товаров
    metrics_sku_with_products = add_columns_from_products(
        metrics_sku_raw,
        products_prepared,
    )

    # 6.2. Добавляем справочник (итоговая 'Габариты товара', кратность)
    metrics_sku_enriched = add_columns_from_catalog(
        metrics_sku_with_products,
        catalog_processed,
    )

    # 6.3. Считаем свод по SKU (BIO/ARK: отдельная логика по КГТ/СГТ)
    svod_sku_with_stats = calc_svod_for_sku(metrics_sku_enriched)

    # 6.4. Добавляем потребность/остатки по складам в свод по SKU (через merge по WB_MERGE_KEY)
    svod_sku_with_clusters = add_demand_for_cluster_columns(
        pivot_by_warehouses,
        svod_sku_with_stats,
    )

    # ---------- 7. Подготовка таблиц для Excel ----------
    # 7.1. "По кластерам" — обогащаем справочником (если нужны поля каталога)
    svod_clusters_final = svod_clusters_initial.copy()

    report_dates_df = read_dates_file(date_metrics)
    date_start_period = report_dates_df["date_start_file"][0]
    date_end_period = report_dates_df["date_end_file"][0]

    # Лист "Всего"
    svod_excel_total = add_columns_for_excel(
        svod_sku_with_clusters,
        warehouse_groups_list,
        date_metrics,
        date_start_period,
        date_end_period,
        svod_type="Всего",
    )

    # Лист "Всего по ювелирным складам"
    svod_excel_jewelry = create_jewelry_clusters_svod(
        warehouses_mapping,
        metrics_clusters_enriched,
        jewelry_group_column,
        date_metrics,
        catalog_processed,
    )

    # Лист "По кластерам"
    svod_excel_clusters = add_columns_for_excel(
        svod_clusters_final,
        warehouse_groups_list,
        date_metrics,
        date_start_period,
        date_end_period,
        svod_type="По кластерам",
    )

    # ---------- 8. Сохранение и форматирование Excel ----------
    save_sheets_to_excel(
        svod_excel_total,
        svod_excel_jewelry,
        svod_excel_clusters,
        warehouses_mapping,
        date_metrics,
    )

    format_excel_file(
        client_name,
        date_metrics,
        svod_excel_total,
        svod_excel_jewelry,
        svod_excel_clusters,
    )

    logger.success(
        f"Отчет успешно создан для {client_name} на дату {date_metrics}"
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
