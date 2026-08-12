"""
Модуль форматирования Excel-файла отчета по поставкам для Ozon.
Функции:
- copy_supply_svod_file
- format_sheet_total  (лист 'Всего' и 'Всего по ювелирным кластерам')
- format_sheet_clusters (лист 'По кластерам')
"""

import shutil
from typing import Dict, List

import numpy as np
import openpyxl
import pandas as pd
from openpyxl import formatting
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ozon.scripts.ozon_logger import ozon_client_logger as logger
from ozon.scripts.constants import (
    include_previous_supplies_flag,
    demand_period_for_total,
)

# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================

def _create_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """
    Маппинг между названиями колонок и буквами столбцов Excel.
    """
    mapping = pd.DataFrame({
        "column": df.columns,
        "column_number": np.arange(1, len(df.columns) + 1),
    })
    mapping["excel_column"] = mapping["column_number"].apply(get_column_letter)
    return mapping


def _get_excel_column(
    mapping: pd.DataFrame,
    column_name: str,
    contains: bool = False,
) -> str | None:
    """
    Получить букву столбца Excel по названию колонки.
    Если contains=True — искать по частичному совпадению.
    """
    if contains:
        result = mapping.loc[
            mapping["column"].str.contains(column_name, na=False),
            "excel_column",
        ]
    else:
        result = mapping.loc[mapping["column"] == column_name, "excel_column"]

    return result.values[0] if len(result) > 0 else None


def _apply_header_formatting(header_cells, thin: Side) -> None:
    """
    Форматирование заголовков таблицы.
    """
    for row in header_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                wrap_text=True,
            )
            cell.border = Border(
                top=thin,
                bottom=thin,
                right=thin,
                left=thin,
            )


def _apply_cell_borders(all_cells, thin: Side) -> None:
    """
    Применить границы ко всем ячейкам таблицы.
    """
    for row in all_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = Border(
                top=thin,
                bottom=thin,
                right=thin,
                left=thin,
            )


def _auto_adjust_column_width(ws, barcode_column: str | None = None) -> None:
    """
    Автоподбор ширины столбцов по длине заголовка.
    Для колонки со штрихкодом (если передана) — фиксированная ширина.
    """
    # первая строка — заголовки
    header_row = 1
    for col_cells in ws.columns:
        first_cell = col_cells[0]
        column_letter = first_cell.column_letter
        header_value = first_cell.value

        try:
            header_length = len(str(header_value)) if header_value is not None else 0
        except Exception:
            header_length = 0

        # небольшое поле + коэффициент, чтобы текст не был впритык
        adjusted_width = (header_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    # Для колонки штрихкода — фиксированная ширина (если нужно)
    if barcode_column:
        ws.column_dimensions[barcode_column].width = 22

def _apply_conditional_formatting(
    ws,
    deficit_columns: List[str],
    row_start: int,
    row_end: int,
) -> None:
    """
    Условное форматирование для колонок дефицита/избытка (красный <0, зелёный >0).
    """
    red_font = Font(color="9C0006")
    red_fill = PatternFill(
        start_color="FFC7CE",
        end_color="FFC7CE",
        fill_type="solid",
    )
    green_font = Font(color="006100")
    green_fill = PatternFill(
        start_color="C6EFCE",
        end_color="C6EFCE",
        fill_type="solid",
    )

    for col in deficit_columns:
        cell_range = f"{col}{row_start}:{col}{row_end}"
        ws.conditional_formatting.add(
            cell_range,
            formatting.rule.CellIsRule(
                operator="lessThan",
                formula=["0"],
                fill=red_fill,
                font=red_font,
            ),
        )
        ws.conditional_formatting.add(
            cell_range,
            formatting.rule.CellIsRule(
                operator="greaterThan",
                formula=["0"],
                fill=green_fill,
                font=green_font,
            ),
        )


def _apply_number_format(
    ws,
    columns: List[str],
    row_start: int,
    row_end: int,
    decimals: int,
) -> None:
    """
    Применяет числовой формат к указанным колонкам.
    """
    format_map = {
        0: "0",
        1: "0.0",
        2: "0.00",
    }
    number_format = format_map.get(decimals, "0.00")

    for col in columns:
        if not col:
            continue
        for row_idx in range(row_start, row_end + 1):
            cell = ws[f"{col}{row_idx}"]
            if isinstance(cell.value, str) and "E+" in cell.value:
                cell.value = int(float(cell.value))
                cell.number_format = "0"
            else:
                cell.number_format = number_format


def _apply_alignment(
    ws,
    columns: List[str],
    row_start: int,
    row_end: int,
    horizontal: str,
) -> None:
    """
    Выравнивание содержимого ячеек по горизонтали.
    """
    for col in columns:
        cell_range = f"{col}{row_start}:{col}{row_end}"
        for row in ws[cell_range]:
            for cell in row:
                cell.alignment = Alignment(
                    horizontal=horizontal,
                    vertical="center",
                )


def _create_demand_rounded_formula(
    deficit_cell: str,
    box_multiplicity_cell: str | None = None,
    apply_rounding: bool = False,
) -> str:
    """
    Формула Excel для округлённой потребности с учётом кратности коробов.

    Без учёта кратности:
      =ROUNDUP(IF(deficit>0,0,-deficit),0)

    С учётом кратности (как в WB):
      =IF(deficit>0,0,
           IF(MOD(-deficit, multiplicity) < multiplicity/2,
              -ROUNDUP(deficit,0),
              ROUNDUP(-deficit/multiplicity,0)*multiplicity
           )
        )
    """
    if not apply_rounding or not box_multiplicity_cell:
        return f"= ROUNDUP(IF({deficit_cell} > 0, 0, -{deficit_cell}), 0)"

    formula = (
        f"= IF({deficit_cell} > 0, 0, "
        f"IF(MOD(-{deficit_cell}, {box_multiplicity_cell}) < {box_multiplicity_cell}/2, "
        f"-ROUNDUP({deficit_cell}, 0), "
        f"ROUNDUP(-{deficit_cell}/{box_multiplicity_cell}, 0) * {box_multiplicity_cell}))"
    )
    return formula


# ============================================================================
# ОСНОВНЫЕ ФУНКЦИИ
# ============================================================================

def copy_supply_svod_file(
    path_supply_svod,
    client_name: str,
    date_report_created: str,
) -> None:
    """
    Создает копию оригинального файла для форматирования.
    """
    src_file = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_Ozon.xlsx"
    )
    dst_file = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_Ozon_formatted.xlsx"
    )
    shutil.copy(src_file, dst_file)
    logger.info(
        f"Создана копия файла для форматирования: {dst_file.name}"
    )


def format_sheet_total(
    path_supply_svod,
    client_name: str,
    date_report_created: str,
    svod_excel: pd.DataFrame,
    svod_excel_clusters: pd.DataFrame,
    clusters_mapping_dict: Dict[str, List[str]],
    sheet_name: str = "Всего",
    demand_period_1: int = 40,
    demand_period_2: int = 60,
    days_in_period: int = 30,
    apply_box_rounding: bool = False,
) -> None:
    """
    Форматирование листа 'Всего' или 'Всего по ювелирным кластерам' для Ozon.
    Логика:
    - пересчёт в Excel оборачиваемости, потребности, дефицитов, округлённой потребности;
    - расчёт 'Потребность <Кластер>' и 'Остаток <Кластер>' через SUMIFS
      по листу 'По кластерам' и словарю clusters_mapping_dict;
    - опционально (если include_previous_supplies_flag == True):
        расчёт 'ПРОШЛЫЕ_ПОСТАВКИ <Кластер>' через SUMIFS по листу 'По кластерам';
    - условное форматирование, числовые форматы и выравнивание.
    """
    logger.info(f'Форматирование листа "{sheet_name}"')
    file_path = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_Ozon.xlsx"
    )
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    svod_len = svod_excel.shape[0]
    row_start = 2
    row_end = svod_len + 1

    # Маппинг колонок
    mapping = _create_column_mapping(svod_excel)
    mapping_clusters = _create_column_mapping(svod_excel_clusters)
    min_col = mapping["excel_column"].iloc[0]
    max_col = mapping["excel_column"].iloc[-1]

    # Основные колонки на листе "Всего"
    cols = {
        "article": _get_excel_column(mapping, "Артикул"),
        "barcode": _get_excel_column(mapping, "Штрихкод"),
        "turnover": _get_excel_column(mapping, "Оборачиваемость"),
        "orders": _get_excel_column(mapping, "ЗАКАЗЫ", contains=True),
        "sales": _get_excel_column(mapping, "ПРОДАЖИ", contains=True),
        "stock": _get_excel_column(mapping, "ОСТАТОК", contains=True),
        "stock_fbs": _get_excel_column(mapping, "ОСТАТОК FBS", contains=True),
        "supplies": _get_excel_column(mapping, "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"),
        "box_multiplicity": _get_excel_column(mapping, "Кратность коробов"),
        # Тип габаритов / КГТ (для логики BIO / ARK)
        "size_type": _get_excel_column(mapping, "Габариты товара"),
    }

    # Колонки по периодам
    cols.update(
        {
            f"demand_{demand_period_1}": _get_excel_column(
                mapping, f"Потребность на {demand_period_1} дней"
            ),
            f"demand_{demand_period_2}": _get_excel_column(
                mapping, f"Потребность на {demand_period_2} дней"
            ),
            f"deficit_{demand_period_1}": _get_excel_column(
                mapping,
                f"Дефицит/Избыток {demand_period_1} дней",
            ),
            f"deficit_{demand_period_1}_fbs": _get_excel_column(
                mapping,
                f"Дефицит/Избыток {demand_period_1} дней с учетом FBS",
            ),
            f"deficit_{demand_period_2}": _get_excel_column(
                mapping,
                f"Дефицит/Избыток {demand_period_2} дней",
            ),
            f"demand_{demand_period_1}_rounded": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_1} дней (округл.)",
            ),
            f"demand_{demand_period_1}_rounded_fbs": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_1} дней с учетом FBS (округл.)",
            ),
            f"demand_{demand_period_2}_rounded": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_2} дней (округл.)",
            ),
        }
    )

    # Колонки из листа "По кластерам"
    cols_clusters = {
        "article": _get_excel_column(mapping_clusters, "Артикул"),
        "cluster": _get_excel_column(mapping_clusters, "Кластер"),
        "cluster_group": _get_excel_column(   # НОВОЕ
            mapping_clusters,
            "Кластер для группировки",
        ),
        "correction": _get_excel_column(
            mapping_clusters,
            "Корректировка с учетом нулевых остатков",
        ),
        "stock": _get_excel_column(
            mapping_clusters,
            "ОСТАТОК",
            contains=True,
        ),
    }

    # Если надо учитывать ПРОШЛЫЕ_ПОСТАВКИ — добавляем колонку с листа "По кластерам"
    if include_previous_supplies_flag:
        cols_clusters["prev_supplies"] = _get_excel_column(
            mapping_clusters,
            "ПРОШЛЫЕ_ПОСТАВКИ",
        )

    # Базовое форматирование
    thin = Side(border_style="thin", color="000000")
    header_cells = ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    all_cells = ws[f"{min_col}{row_start}:{max_col}{row_end}"]
    _apply_header_formatting(header_cells, thin)
    ws.auto_filter.ref = ws.dimensions
    _apply_cell_borders(all_cells, thin)
    _auto_adjust_column_width(ws, cols["barcode"])

    # Формулы по строкам (оборачиваемость, потребности, дефициты, округления)
    for i in range(row_start, row_end + 1):
        # Оборачиваемость
        if cols["turnover"] and cols["orders"]:
            # ----------------- Логика для RiverBoats -----------------
            if client_name == "RiverBoats":
                # Оборачиваемость = Заказы / days_in_period
                ws[f"{cols['turnover']}{i}"] = (
                    f"= {cols['orders']}{i} / {days_in_period}"
                )

            # ----------------- Логика для BIO и ARK -----------------
            elif client_name in ("BIO", "ARK"):
                # Если есть колонка габаритов и колонка продаж:
                #   - если 'Габариты товара' = "КГТ" → только заказы
                #   - иначе → (Заказы + Продажи) / 2 / days_in_period
                if cols["size_type"] and cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        "=IF("
                        f'{cols["size_type"]}{i}="КГТ", '
                        f"{cols['orders']}{i} / {days_in_period}, "
                        f"({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                        ")"
                    )
                # Если нет колонки габаритов, но есть продажи — fallback: общая формула
                elif cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                    )
                # Если нет продаж — хотя бы по заказам
                else:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= {cols['orders']}{i} / {days_in_period}"
                    )

            # ----------------- Общая логика для остальных клиентов -----------------
            else:
                if cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= {cols['orders']}{i} / {days_in_period}"
                    )

        # Потребности
        if cols.get(f"demand_{demand_period_1}"):
            ws[f"{cols[f'demand_{demand_period_1}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_1}"
            )
        if cols.get(f"demand_{demand_period_2}"):
            ws[f"{cols[f'demand_{demand_period_2}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_2}"
            )

        # Дефицит / избыток
        if (
            cols.get(f"deficit_{demand_period_1}")
            and cols["stock"]
            and cols["supplies"]
            and cols.get(f"demand_{demand_period_1}")
        ):
            ws[f"{cols[f'deficit_{demand_period_1}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
            )
        if (
            cols.get(f"deficit_{demand_period_1}_fbs")
            and cols["stock"]
            and cols["stock_fbs"]
            and cols["supplies"]
            and cols.get(f"demand_{demand_period_1}")
        ):
            ws[f"{cols[f'deficit_{demand_period_1}_fbs']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['stock_fbs']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
            )
        if (
            cols.get(f"deficit_{demand_period_2}")
            and cols["stock"]
            and cols["supplies"]
            and cols.get(f"demand_{demand_period_2}")
        ):
            ws[f"{cols[f'deficit_{demand_period_2}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_2}']}{i}"
            )

        # Округлённые потребности
        if (
            cols.get(f"demand_{demand_period_1}_rounded")
            and cols.get(f"deficit_{demand_period_1}")
        ):
            ws[f"{cols[f'demand_{demand_period_1}_rounded']}{i}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_1}']}{i}",
                    f"{cols['box_multiplicity']}{i}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )
        if (
            cols.get(f"demand_{demand_period_1}_rounded_fbs")
            and cols.get(f"deficit_{demand_period_1}_fbs")
        ):
            ws[f"{cols[f'demand_{demand_period_1}_rounded_fbs']}{i}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_1}_fbs']}{i}",
                    f"{cols['box_multiplicity']}{i}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )
        if (
            cols.get(f"demand_{demand_period_2}_rounded")
            and cols.get(f"deficit_{demand_period_2}")
        ):
            ws[f"{cols[f'demand_{demand_period_2}_rounded']}{i}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_2}']}{i}",
                    f"{cols['box_multiplicity']}{i}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )

    # Формулы по кластерам (Потребность/Остаток/[Прошлые поставки] <Кластер>)
    # через одну SUMIFS по "Артикул" + "Кластер для группировки"
    svod_len_clusters = svod_excel_clusters.shape[0]
    row_start_clusters = 2

    # Колонки-источники на листе "По кластерам"
    corr_col = cols_clusters.get("correction")
    stock_src_col = cols_clusters.get("stock")
    article_src_col = cols_clusters.get("article")
    cluster_group_src_col = cols_clusters.get("cluster_group")
    prev_supplies_src_col = cols_clusters.get("prev_supplies") if include_previous_supplies_flag else None

    # Если базовых колонок нет — выходим
    if not article_src_col or not cluster_group_src_col or not corr_col:
        pass
    else:
        # Диапазоны на листе "По кластерам"
        corr_range = f"'По кластерам'!${corr_col}${row_start_clusters}:${corr_col}${svod_len_clusters + 1}"
        stock_range = (
            f"'По кластерам'!${stock_src_col}${row_start_clusters}:${stock_src_col}${svod_len_clusters + 1}"
            if stock_src_col
            else None
        )
        article_range = f"'По кластерам'!${article_src_col}${row_start_clusters}:${article_src_col}${svod_len_clusters + 1}"
        cluster_group_range = f"'По кластерам'!${cluster_group_src_col}${row_start_clusters}:${cluster_group_src_col}${svod_len_clusters + 1}"
        prev_supplies_range = (
            f"'По кластерам'!${prev_supplies_src_col}${row_start_clusters}:${prev_supplies_src_col}${svod_len_clusters + 1}"
            if prev_supplies_src_col
            else None
        )

        for cluster in clusters_mapping_dict.keys():
            # Итоговые колонки на листе "Всего"
            demand_cluster_col = _get_excel_column(
                mapping,
                f"Потребность {cluster}",
            )
            stock_cluster_col = _get_excel_column(
                mapping,
                f"Остаток {cluster}",
            )
            if include_previous_supplies_flag:
                prev_supplies_cluster_col = _get_excel_column(
                    mapping,
                    f"Прошлые поставки {cluster}",
                )
            else:
                prev_supplies_cluster_col = None

            # Если нет ни одной целевой колонки — пропускаем этот кластер
            if (
                not demand_cluster_col
                and not stock_cluster_col
                and not prev_supplies_cluster_col
            ):
                continue

            for i in range(row_start, row_end + 1):
                # Потребность: сумма Корректировки по артикулу и Кластеру для группировки = <cluster>
                if demand_cluster_col:
                    ws[f"{demand_cluster_col}{i}"] = (
                        "=SUMIFS("
                        f"{corr_range}, "
                        f"{article_range}, ${cols['article']}{i}, "
                        f"{cluster_group_range}, \"{cluster}\""
                        ")"
                    )
                # Остаток
                if stock_cluster_col and stock_range:
                    ws[f"{stock_cluster_col}{i}"] = (
                        "=SUMIFS("
                        f"{stock_range}, "
                        f"{article_range}, ${cols['article']}{i}, "
                        f"{cluster_group_range}, \"{cluster}\""
                        ")"
                    )
                # ПРОШЛЫЕ_ПОСТАВКИ
                if (
                    include_previous_supplies_flag
                    and prev_supplies_cluster_col
                    and prev_supplies_range
                ):
                    ws[f"{prev_supplies_cluster_col}{i}"] = (
                        "=SUMIFS("
                        f"{prev_supplies_range}, "
                        f"{article_range}, ${cols['article']}{i}, "
                        f"{cluster_group_range}, \"{cluster}\""
                        ")"
                    )
    # ---- Суммарная потребность по всем кластерам ----
    if client_name == "RiverBoats":
        # находим все колонки "Потребность <Кластер>", но не "Потребность на N дней"
        # и не "Потребность по кластерам всего"
        demand_cluster_cols = [
            col
            for col in svod_excel.columns
            if col.startswith("Потребность ")
            and not col.startswith("Потребность на ")
            and col != "Потребность по кластерам всего"
        ]
        # Получаем их буквенные индексы
        demand_cluster_excel_cols = [
            _get_excel_column(mapping, col) for col in demand_cluster_cols
        ]
        demand_cluster_excel_cols = [c for c in demand_cluster_excel_cols if c]

        # Колонка для итоговой суммы
        total_demand_col = _get_excel_column(
            mapping,
            "Потребность по кластерам всего",
        )

        if total_demand_col and demand_cluster_excel_cols:
            for i in range(row_start, row_end + 1):
                # формула вида =SUM(<col1>i, <col2>i, ...)
                sum_range = ", ".join(f"{col}{i}" for col in demand_cluster_excel_cols)
                ws[f"{total_demand_col}{i}"] = f"=SUM({sum_range})"

    # Условное форматирование
    deficit_columns = [
        cols.get(f"deficit_{demand_period_1}"),
        cols.get(f"deficit_{demand_period_1}_fbs"),
        cols.get(f"deficit_{demand_period_2}"),
    ]
    deficit_columns = [c for c in deficit_columns if c]
    if deficit_columns:
        _apply_conditional_formatting(ws, deficit_columns, row_start, row_end)
    # ------------------------------------------------------------------
    # Доп. условное форматирование для BIO:
    # если Статус товара = "Основной ассортимент" и Остаток {кластер} = 5,
    # то подсветить ячейку остатка красным цветом (фон + шрифт).
    # Кластеры: "СПб и СЗО", "Москва, МО и ДР", "Ростов", "Краснодар".
    # ------------------------------------------------------------------
    if client_name in ('BIO', 'ARK'):
        # колонка со статусом товара
        status_col = _get_excel_column(mapping, "Статус товара")
        if status_col:
            # интересующие кластеры
            target_clusters = ["СПб и СЗО", "Москва, МО и ДР", "Ростов", "Краснодар"]

            # стиль для подсветки
            red_font = Font(color="9C0006")
            red_fill = PatternFill(
                start_color="FFC7CE",
                end_color="FFC7CE",
                fill_type="solid",
            )

            for cl in target_clusters:
                stock_col = _get_excel_column(
                    mapping,
                    f"Остаток {cl}",
                )
                if not stock_col:
                    continue

                cell_range = f"{stock_col}{row_start}:{stock_col}{row_end}"
                # Условие: статус = "Основной ассортимент" и значение в этой ячейке = 5
                # Используем формулу типа:
                # =AND($<Статус><row>="Основной ассортимент", <эта_ячейка>=5)
                formula = (
                    f'AND(${status_col}{row_start}="Основной ассортимент",'
                    f'{stock_col}{row_start}=5)'
                )

                ws.conditional_formatting.add(
                    cell_range,
                    formatting.rule.FormulaRule(
                        formula=[formula],
                        stopIfTrue=False,
                        fill=red_fill,
                        font=red_font,
                    ),
                )
    # Формат чисел
    two_decimal_cols = [
        cols.get("turnover"),
        cols.get(f"demand_{demand_period_1}"),
        cols.get(f"demand_{demand_period_2}"),
    ]
    two_decimal_cols = [c for c in two_decimal_cols if c]
    if two_decimal_cols:
        _apply_number_format(ws, two_decimal_cols, row_start, row_end, 2)

    one_decimal_cols = deficit_columns
    if one_decimal_cols:
        _apply_number_format(ws, one_decimal_cols, row_start, row_end, 1)

    zero_decimal_cols = [
        cols.get("barcode"),
        cols.get(f"demand_{demand_period_1}_rounded"),
        cols.get(f"demand_{demand_period_1}_rounded_fbs"),
        cols.get(f"demand_{demand_period_2}_rounded"),
    ]
    zero_decimal_cols = [c for c in zero_decimal_cols if c]
    if zero_decimal_cols:
        _apply_number_format(ws, zero_decimal_cols, row_start, row_end, 0)

    # Выравнивание
    center_cols = [
        "Сезон",
        "Статус",
        "Статус товара",
        "Статус продажи FBO/FBS",
        "Основной артикул",
        "Категория",
        "Предмет",
        "Цвет",
    ]
    center_string = "|".join(center_cols)
    text_mask = mapping["column"].str.contains(
        center_string,
        na=False,
    )
    left_cols = mapping.loc[text_mask, "excel_column"].tolist()
    center_cols = mapping.loc[~text_mask, "excel_column"].tolist()
    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, "center")
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, "left")

    wb.save(file_path)
    logger.info(f'Завершено форматирование листа "{sheet_name}"')

def format_sheet_clusters(
    path_supply_svod,
    client_name: str,
    date_report_created: str,
    svod_excel: pd.DataFrame,
    sheet_name: str = "По кластерам",
    demand_period_1: int = 40,
    demand_period_2: int = 60,
    days_in_period: int = 30,
    apply_box_rounding: bool = False,
) -> None:
    """
    Форматирует лист 'По кластерам' для Ozon.
    Формулы аналогичны листу 'По кластерам' в основном расчёте:
      - для RiverBoats и BIO/ARK используются Дни планирования из колонки
        'Дни планирования' и оборачиваемость по их правилам;
      - для остальных клиентов логика такая же, как была:
        оборачиваемость по (Заказы + Продажи) / 2 / days_in_period
        и потребности/дефициты на фиксированные периоды.
    """
    logger.info(f'Форматирование листа "{sheet_name}"')
    file_path = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_Ozon.xlsx"
    )
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    svod_len = svod_excel.shape[0]
    row_start = 2
    row_end = svod_len + 1
    mapping = _create_column_mapping(svod_excel)
    min_col = mapping["excel_column"].iloc[0]
    max_col = mapping["excel_column"].iloc[-1]
    cols = {
        "barcode": _get_excel_column(mapping, "Штрихкод"),
        "turnover": _get_excel_column(mapping, "Оборачиваемость"),
        "orders": _get_excel_column(mapping, "ЗАКАЗЫ", contains=True),
        "sales": _get_excel_column(mapping, "ПРОДАЖИ", contains=True),
        "stock": _get_excel_column(mapping, "ОСТАТОК", contains=True),
        "stock_fbs": _get_excel_column(mapping, "ОСТАТОК FBS", contains=True),
        "supplies": _get_excel_column(mapping, "ОЖИДАЕТСЯ_ПОСТУПЛЕНИЕ"),
        "correction": _get_excel_column(
            mapping,
            "Корректировка с учетом нулевых остатков",
        ),
        "box_multiplicity": _get_excel_column(mapping, "Кратность коробов"),
        # Габариты товара (для логики КГТ у BIO/ARK)
        "size_type": _get_excel_column(mapping, "Габариты товара"),
        # Дни планирования (для RiverBoats и BIO/ARK)
        "planning_days": _get_excel_column(mapping, "Дни планирования"),
        # Потребность/дефициты без FBS (для общей логики)
    }
    cols.update(
        {
            f"demand_{demand_period_1}": _get_excel_column(
                mapping, f"Потребность на {demand_period_1} дней"
            ),
            f"demand_{demand_period_2}": _get_excel_column(
                mapping, f"Потребность на {demand_period_2} дней"
            ),
            f"deficit_{demand_period_1}": _get_excel_column(
                mapping,
                f"Дефицит/Избыток {demand_period_1} дней",
            ),
            f"deficit_{demand_period_2}": _get_excel_column(
                mapping,
                f"Дефицит/Избыток {demand_period_2} дней",
            ),
            f"demand_{demand_period_1}_rounded": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_1} дней (округл.)",
            ),
            f"demand_{demand_period_2}_rounded": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_2} дней (округл.)",
            ),
            # Для RiverBoats и BIO/ARK — колонки с потребностью по дням планирования
            "demand_custom": _get_excel_column(mapping, "Потребность"),
            "deficit_custom": _get_excel_column(mapping, "Дефицит/Избыток"),
            "deficit_custom_fbs": _get_excel_column(
                mapping, "Дефицит/Избыток с учетом FBS"
            ),
            "demand_custom_rounded": _get_excel_column(
                mapping, "Потребность (округл.)"
            ),
            "demand_custom_rounded_fbs": _get_excel_column(
                mapping, "Потребность с учетом FBS (округл.)"
            ),
        }
    )
    thin = Side(border_style="thin", color="000000")
    header_cells = ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    all_cells = ws[f"{min_col}{row_start}:{max_col}{row_end}"]
    _apply_header_formatting(header_cells, thin)
    ws.auto_filter.ref = ws.dimensions
    _apply_cell_borders(all_cells, thin)
    _auto_adjust_column_width(ws, cols["barcode"])
    # Формулы
    for i in range(row_start, row_end + 1):
        # ----------------- Оборачиваемость -----------------
        if cols["turnover"] and cols["orders"]:
            # RiverBoats: только заказы
            if client_name == "RiverBoats":
                ws[f"{cols['turnover']}{i}"] = (
                    f"= {cols['orders']}{i} / {days_in_period}"
                )
            # BIO и ARK: КГТ — только заказы, иначе среднее по заказам и продажам
            elif client_name in ("BIO", "ARK"):
                if cols["size_type"] and cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        "=IF("
                        f'{cols["size_type"]}{i}="КГТ", '
                        f"{cols['orders']}{i} / {days_in_period}, "
                        f"({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                        ")"
                    )
                elif cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= {cols['orders']}{i} / {days_in_period}"
                    )
            # Остальные клиенты: общая формула
            else:
                if cols["sales"]:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{i}"] = (
                        f"= {cols['orders']}{i} / {days_in_period}"
                    )
        # ----------------- Логика для RiverBoats и BIO/ARK (через Дни планирования) -----------------
        if client_name in ("Neva_Metall", "RiverBoats", "BIO", "ARK") and cols["planning_days"]:
            # Потребность = Оборачиваемость * Дни планирования
            if cols["demand_custom"] and cols["turnover"]:
                ws[f"{cols['demand_custom']}{i}"] = (
                    f"= {cols['turnover']}{i} * {cols['planning_days']}{i}"
                )
            # Дефицит / избыток (без FBS)
            if (
                cols["deficit_custom"]
                and cols["stock"]
                and cols["supplies"]
                and cols["demand_custom"]
            ):
                ws[f"{cols['deficit_custom']}{i}"] = (
                    f"= {cols['stock']}{i} + {cols['supplies']}{i} "
                    f"- {cols['demand_custom']}{i}"
                )
            # Дефицит / избыток с учетом FBS
            if (
                cols["deficit_custom_fbs"]
                and cols["stock"]
                and cols["stock_fbs"]
                and cols["supplies"]
                and cols["demand_custom"]
            ):
                ws[f"{cols['deficit_custom_fbs']}{i}"] = (
                    f"= {cols['stock']}{i} + {cols['stock_fbs']}{i} + {cols['supplies']}{i} "
                    f"- {cols['demand_custom']}{i}"
                )
            # Округлённая потребность (без FBS)
            if cols["demand_custom_rounded"] and cols["deficit_custom"]:
                ws[f"{cols['demand_custom_rounded']}{i}"] = (
                    _create_demand_rounded_formula(
                        f"{cols['deficit_custom']}{i}",
                        f"{cols['box_multiplicity']}{i}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            # Округлённая потребность с учетом FBS
            if (
                cols["demand_custom_rounded_fbs"]
                and cols["deficit_custom_fbs"]
            ):
                ws[f"{cols['demand_custom_rounded_fbs']}{i}"] = (
                    _create_demand_rounded_formula(
                        f"{cols['deficit_custom_fbs']}{i}",
                        f"{cols['box_multiplicity']}{i}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            # Корректировка = Потребность (округл.)
            if cols["correction"] and cols["demand_custom_rounded"]:
                ws[f"{cols['correction']}{i}"] = (
                    f"= {cols['demand_custom_rounded']}{i}"
                )
        # ----------------- Общая логика для остальных клиентов -----------------
        else:
            # Потребности на фиксированные периоды
            if cols.get(f"demand_{demand_period_1}"):
                ws[f"{cols[f'demand_{demand_period_1}']}{i}"] = (
                    f"= {cols['turnover']}{i} * {demand_period_1}"
                )
            if cols.get(f"demand_{demand_period_2}"):
                ws[f"{cols[f'demand_{demand_period_2}']}{i}"] = (
                    f"= {cols['turnover']}{i} * {demand_period_2}"
                )
            # Дефицит / избыток
            if (
                cols.get(f"deficit_{demand_period_1}")
                and cols["stock"]
                and cols["supplies"]
                and cols.get(f"demand_{demand_period_1}")
            ):
                ws[f"{cols[f'deficit_{demand_period_1}']}{i}"] = (
                    f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
                )
            if (
                cols.get(f"deficit_{demand_period_2}")
                and cols["stock"]
                and cols["supplies"]
                and cols.get(f"demand_{demand_period_2}")
            ):
                ws[f"{cols[f'deficit_{demand_period_2}']}{i}"] = (
                    f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_2}']}{i}"
                )
            # Округлённые потребности
            if (
                cols.get(f"demand_{demand_period_1}_rounded")
                and cols.get(f"deficit_{demand_period_1}")
            ):
                ws[f"{cols[f'demand_{demand_period_1}_rounded']}{i}"] = (
                    _create_demand_rounded_formula(
                        f"{cols[f'deficit_{demand_period_1}']}{i}",
                        f"{cols['box_multiplicity']}{i}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            if (
                cols.get(f"demand_{demand_period_2}_rounded")
                and cols.get(f"deficit_{demand_period_2}")
            ):
                ws[f"{cols[f'demand_{demand_period_2}_rounded']}{i}"] = (
                    _create_demand_rounded_formula(
                        f"{cols[f'deficit_{demand_period_2}']}{i}",
                        f"{cols['box_multiplicity']}{i}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            # Корректировка = Потребность на demand_period_for_total (округл.)
            if (
                cols["correction"]
                and cols.get(f"demand_{demand_period_for_total}_rounded")
            ):
                ws[f"{cols['correction']}{i}"] = (
                    f"= {cols[f'demand_{demand_period_for_total}_rounded']}{i}"
                )
    # Условное форматирование
    # Для RiverBoats/BIO/ARK дефициты в custom-колонке,
    # для остальных — в дефолтных по периодам.
    if client_name in ("Neva_Metall", "RiverBoats", "BIO", "ARK") and cols["deficit_custom"]:
        deficit_columns = [cols["deficit_custom"]]
        if cols.get("deficit_custom_fbs"):
            deficit_columns.append(cols["deficit_custom_fbs"])
    else:
        deficit_columns = [
            cols.get(f"deficit_{demand_period_1}"),
            cols.get(f"deficit_{demand_period_2}"),
        ]
        deficit_columns = [c for c in deficit_columns if c]
    if deficit_columns:
        _apply_conditional_formatting(ws, deficit_columns, row_start, row_end)
    # Формат чисел
    two_decimal_cols = [
        cols.get("turnover"),
        cols.get(f"demand_{demand_period_1}"),
        cols.get(f"demand_{demand_period_2}"),
        cols.get("demand_custom"),
    ]
    two_decimal_cols = [c for c in two_decimal_cols if c]
    if two_decimal_cols:
        _apply_number_format(ws, two_decimal_cols, row_start, row_end, 2)
    one_decimal_cols = deficit_columns
    if one_decimal_cols:
        _apply_number_format(ws, one_decimal_cols, row_start, row_end, 1)
    zero_decimal_cols = [
        cols.get("barcode"),
        cols.get(f"demand_{demand_period_1}_rounded"),
        cols.get(f"demand_{demand_period_2}_rounded"),
        cols.get("demand_custom_rounded"),
        cols.get("demand_custom_rounded_fbs"),
        cols.get("correction"),
    ]
    zero_decimal_cols = [c for c in zero_decimal_cols if c]
    if zero_decimal_cols:
        _apply_number_format(ws, zero_decimal_cols, row_start, row_end, 0)
    # Выравнивание
    center_cols = [
        "Сезон",
        "Статус",
        "Статус товара",
        "Статус продажи FBO/FBS",
        "Основной артикул",
        "Категория",
        "Предмет",
        "Цвет",
    ]
    center_string = "|".join(center_cols)
    text_mask = mapping["column"].str.contains(
        center_string,
        na=False,
    )
    left_cols = mapping.loc[text_mask, "excel_column"].tolist()
    center_cols = mapping.loc[~text_mask, "excel_column"].tolist()
    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, "center")
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, "left")
    wb.save(file_path)
    logger.info(f'Завершено форматирование листа "{sheet_name}"')
