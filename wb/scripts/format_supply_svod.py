"""
Модуль форматирования Excel-файла отчета по поставкам для Wildberries.

Содержит функции для форматирования листов 'Всего' и 'По кластерам',
включая применение формул, условного форматирования и стилей.
"""

import shutil
from typing import List

import numpy as np
import openpyxl
import pandas as pd
from openpyxl import formatting
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from wb.scripts.wb_logger import wb_client_logger as logger

from wb.scripts.constants import (
    demand_period_for_total,
)

# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================


def _create_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """
    Создает маппинг между названиями колонок и буквами столбцов Excel.

    Args:
        df: DataFrame с данными

    Returns:
        DataFrame с маппингом колонок
    """
    mapping = pd.DataFrame({
        "column": df.columns,
        "column_number": np.arange(1, len(df.columns) + 1)
    })
    mapping['excel_column'] = mapping['column_number'].apply(get_column_letter)

    return mapping


def _get_excel_column(mapping: pd.DataFrame, column_name: str, contains: bool = False) -> str:
    """
    Получает букву столбца Excel по названию колонки.

    Args:
        mapping: DataFrame с маппингом
        column_name: Название колонки для поиска
        contains: Искать по частичному совпадению

    Returns:
        Буква столбца Excel
    """
    if contains:
        result = mapping.loc[mapping['column'].str.contains(column_name, na=False), 'excel_column']
    else:
        result = mapping.loc[mapping['column'] == column_name, 'excel_column']

    return result.values[0] if len(result) > 0 else None


def _apply_header_formatting(header_cells, thin_border: Side) -> None:
    """
    Применяет форматирование к заголовкам таблицы.

    Args:
        header_cells: Ячейки заголовка
        thin_border: Стиль границы
    """
    for row in header_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
            cell.border = Border(top=thin_border, bottom=thin_border,
                               right=thin_border, left=thin_border)


def _apply_cell_borders(all_cells, thin_border: Side) -> None:
    """
    Применяет границы ко всем ячейкам таблицы.

    Args:
        all_cells: Все ячейки таблицы
        thin_border: Стиль границы
    """
    for row in all_cells:
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = Border(top=thin_border, bottom=thin_border,
                               right=thin_border, left=thin_border)


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

    # Для колонки штрихкода — фиксированная ширина
    if barcode_column:
        ws.column_dimensions[barcode_column].width = 22

def _apply_conditional_formatting(ws, deficit_columns: List[str], row_start: int, row_end: int) -> None:
    """
    Применяет условное форматирование к колонкам дефицита/избытка.

    Args:
        ws: Worksheet
        deficit_columns: Список колонок с дефицитом
        row_start: Начальная строка
        row_end: Конечная строка
    """
    red_font = Font(color='9C0006')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_font = Font(color='006100')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    for col in deficit_columns:
        cell_range = f"{col}{row_start}:{col}{row_end}"

        # Красный для отрицательных (дефицит)
        ws.conditional_formatting.add(
            cell_range,
            formatting.rule.CellIsRule(
                operator='lessThan',
                formula=['0'],
                fill=red_fill,
                font=red_font
            )
        )

        # Зеленый для положительных (избыток)
        ws.conditional_formatting.add(
            cell_range,
            formatting.rule.CellIsRule(
                operator='greaterThan',
                formula=['0'],
                fill=green_fill,
                font=green_font
            )
        )


def _apply_number_format(ws, columns: List[str], row_start: int, row_end: int, decimals: int) -> None:
    """
    Применяет форматирование чисел к колонкам.

    Args:
        ws: Worksheet
        columns: Список колонок
        row_start: Начальная строка
        row_end: Конечная строка
        decimals: Количество знаков после запятой
    """
    format_map = {
        0: "0",
        1: "0.0",
        2: "0.00"
    }
    number_format = format_map.get(decimals, "0.00")

    for col in columns:
        if not col:  # Пропускаем None
            continue

        # Применяем форматирование к каждой ячейке в диапазоне
        for row_idx in range(row_start, row_end + 1):
            cell = ws[f"{col}{row_idx}"]

            # Обработка чисел в экспоненциальной форме
            if isinstance(cell.value, str) and "E+" in cell.value:
                cell.value = int(float(cell.value))
                cell.number_format = '0'
            else:
                # Применяем формат ко всем ячейкам (включая формулы, числа и пустые)
                cell.number_format = number_format

def _apply_alignment(ws, columns: List[str], row_start: int, row_end: int, horizontal: str) -> None:
    """
    Применяет выравнивание к колонкам.

    Args:
        ws: Worksheet
        columns: Список колонок
        row_start: Начальная строка
        row_end: Конечная строка
        horizontal: Тип выравнивания ('left', 'center', 'right')
    """
    for col in columns:
        cell_range = f"{col}{row_start}:{col}{row_end}"

        for row in ws[cell_range]:
            for cell in row:
                cell.alignment = Alignment(horizontal=horizontal, vertical="center")


def _create_demand_rounded_formula(
    deficit_cell: str,
    box_multiplicity_cell: str,
    apply_rounding: bool = False
) -> str:
    """
    Формирует Excel-формулу для расчета округленной потребности по дефициту.

    Логика:
    - если дефицит неотрицательный, возвращается 0;
    - если дефицит отрицательный, берется его модуль (объем потребности);
    - при ``apply_rounding=False`` используется простое округление вверх до целого;
    - при ``apply_rounding=True`` применяется логика округления с учетом
      кратности коробов.

    Args:
        deficit_cell: Ссылка на ячейку Excel с дефицитом/избытком.
        box_multiplicity_cell: Ссылка на ячейку Excel с кратностью коробов.
            Используется только при ``apply_rounding=True``.
        apply_rounding: Признак применения округления по кратности коробов.

    Returns:
        Строка Excel-формулы (начинается с ``=``), готовая для записи в ячейку.
    """
    if not apply_rounding:
        # Базовая формула без сложной логики по кратности
        return f"=ROUNDUP(IF({deficit_cell} > 0, 0, -{deficit_cell}), 0)"

    # Переменные для краткости записи
    d = deficit_cell
    m = box_multiplicity_cell

    # Аналог русской формулы c использованием английских функций Excel
    formula = (
        f"=IF({d}>=0,0,\n"
        f"   IF(-{d}<={m}/2,\n"
        f"      -{d},\n"
        f"      IF(-{d}<={m},\n"
        f"         {m},\n"
        f"         IF((-{d}-ROUNDDOWN(-{d}/{m},0)*{m})>{m}/2,\n"
        f"            (ROUNDDOWN(-{d}/{m},0)+1)*{m},\n"
        f"            ROUNDDOWN(-{d}/{m},0)*{m} + (-{d}-ROUNDDOWN(-{d}/{m},0)*{m})\n"
        f"         )\n"
        f"      )\n"
        f"   )\n"
        f")"
    )
    return formula


# ============================================================================
# ОСНОВНЫЕ ФУНКЦИИ
# ============================================================================


def copy_supply_svod_file(path_supply_svod, client_name: str, date_report_created: str) -> None:
    """
    Создает копию оригинального файла для форматирования.

    Args:
        path_supply_svod: Путь к директории с файлами
        client_name: Имя клиента
        date_report_created: Дата создания отчета
    """
    src_file = path_supply_svod / f"{date_report_created}_Расчет_поставок_{client_name}_WB.xlsx"
    dst_file = path_supply_svod / f"{date_report_created}_Расчет_поставок_{client_name}_WB_formatted.xlsx"

    shutil.copy(src_file, dst_file)
    logger.info(f"Создана копия файла для форматирования: {dst_file.name}")


def format_sheet_total(
    path_supply_svod,
    client_name: str,
    date_report_created: str,
    svod_excel: pd.DataFrame,
    svod_excel_clusters: pd.DataFrame,
    WB_MERGE_KEY: str,
    sheet_name: str = "Всего",
    demand_period_1: int = 40,
    demand_period_2: int = 60,
    days_in_period: int = 30,
    apply_box_rounding: bool = False,
) -> None:
    """
    Форматирует лист 'Всего' или 'Всего по ювелирным складам' для WB.
    Все связи с листом 'По кластерам' выполняются по колонке merge-key (WB_MERGE_KEY),
    а не жёстко по 'Артикул_Размер'.
    """
    logger.info(f'Начало форматирование листа "{sheet_name}"')

    file_path = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_WB.xlsx"
    )
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    # Размер таблицы
    row_start = 2
    row_end = svod_excel.shape[0] + 1

    # Маппинг колонок листа "Всего" и "По кластерам"
    mapping = _create_column_mapping(svod_excel)
    mapping_clusters = _create_column_mapping(svod_excel_clusters)

    min_col = mapping["excel_column"].iloc[0]
    max_col = mapping["excel_column"].iloc[-1]

    # Основные колонки на листе "Всего"
    cols = {
        "barcode": _get_excel_column(mapping, "Штрихкод"),
        "article": _get_excel_column(mapping, "Артикул продавца"),
        "size": _get_excel_column(mapping, "Размер"),
        "merge_key": _get_excel_column(mapping, WB_MERGE_KEY),
        "turnover": _get_excel_column(mapping, "Оборачиваемость"),
        "orders": _get_excel_column(mapping, "ЗАКАЗЫ", contains=True),
        "sales": _get_excel_column(mapping, "ПРОДАЖИ", contains=True),
        "stock": _get_excel_column(mapping, "ОСТАТОК", contains=True),
        "stock_fbs": _get_excel_column(mapping, "ОСТАТОК FBS", contains=True),
        "supplies": _get_excel_column(mapping, "Ожидаемое поступление"),
        "box_multiplicity": _get_excel_column(mapping, "Кратность коробов"),
        "size_type": _get_excel_column(mapping, "Габариты товара"),
    }

    # Колонки по периодам
    cols.update(
        {
            f"demand_{demand_period_1}": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_1} дней",
            ),
            f"demand_{demand_period_2}": _get_excel_column(
                mapping,
                f"Потребность на {demand_period_2} дней",
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
        "merge_key": _get_excel_column(mapping_clusters, WB_MERGE_KEY),
        "warehouse_calc": _get_excel_column(
            mapping_clusters,
            "Склад для группировки",
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

    # Список агрегированных складов (колонки вида "Потребность <Склад для группировки>")
    warehouses_calc = [
        col.replace("Потребность ", "")
        for col in mapping["column"]
        if col.startswith("Потребность ")
        and not col.startswith(f"Потребность на ")
    ]

    # ----------------- Базовое форматирование -----------------
    thin = Side(border_style="thin", color="000000")
    header_cells = ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    all_cells = ws[f"{min_col}{row_start}:{max_col}{row_end}"]

    _apply_header_formatting(header_cells, thin)
    ws.auto_filter.ref = ws.dimensions
    _apply_cell_borders(all_cells, thin)
    _auto_adjust_column_width(ws, cols["barcode"])

    # ----------------- Формулы по складам (через SUMIFS) -----------------
    svod_len_clusters = svod_excel_clusters.shape[0]
    row_start_clusters = 2

    # Диапазон по merge_key и складу для расчёта на листе "По кластерам"
    merge_key_range = (
        f"'По кластерам'!${cols_clusters['merge_key']}${row_start_clusters}:"
        f"${cols_clusters['merge_key']}${svod_len_clusters + 1}"
    )
    warehouse_calc_range = (
        f"'По кластерам'!${cols_clusters['warehouse_calc']}${row_start_clusters}:"
        f"${cols_clusters['warehouse_calc']}${svod_len_clusters + 1}"
    )
    corr_range = (
        f"'По кластерам'!${cols_clusters['correction']}${row_start_clusters}:"
        f"${cols_clusters['correction']}${svod_len_clusters + 1}"
    )
    stock_range = (
        f"'По кластерам'!${cols_clusters['stock']}${row_start_clusters}:"
        f"${cols_clusters['stock']}${svod_len_clusters + 1}"
        if cols_clusters["stock"]
        else None
    )

    for warehouse in warehouses_calc:
        demand_col_total = _get_excel_column(mapping, f"Потребность {warehouse}")
        stock_col_total = _get_excel_column(mapping, f"Остаток {warehouse}")
        if not demand_col_total and not stock_col_total:
            continue

        for row_idx in range(row_start, row_end + 1):
            demand_parts = []
            stock_parts = []

            # Потребность = SUMIFS(Корректировка; merge_key; текущий merge_key; Склад для группировки; warehouse)
            if demand_col_total and cols["merge_key"]:
                demand_formula = (
                    "SUMIFS("
                    f"{corr_range}, "
                    f"{merge_key_range}, ${cols['merge_key']}{row_idx}, "
                    f"{warehouse_calc_range}, \"{warehouse}\""
                    ")"
                )
                demand_parts.append(demand_formula)

            # Остаток = SUMIFS(Остаток; merge_key; текущий merge_key; Склад для группировки; warehouse)
            if stock_col_total and stock_range and cols["merge_key"]:
                stock_formula = (
                    "SUMIFS("
                    f"{stock_range}, "
                    f"{merge_key_range}, ${cols['merge_key']}{row_idx}, "
                    f"{warehouse_calc_range}, \"{warehouse}\""
                    ")"
                )
                stock_parts.append(stock_formula)

            if demand_col_total and demand_parts:
                ws[f"{demand_col_total}{row_idx}"] = "= " + " + ".join(demand_parts)
            if stock_col_total and stock_parts:
                ws[f"{stock_col_total}{row_idx}"] = "= " + " + ".join(stock_parts)

    # ----------------- Основные формулы (оборачиваемость, потребности, дефициты) -----------------
    for row_idx in range(row_start, row_end + 1):
        # Оборачиваемость
        if cols["turnover"] and cols["orders"]:
            if client_name in ("BIO", "ARK"):
                if cols["sales"] and cols["size_type"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        "=IF("
                        f'{cols["size_type"]}{row_idx}="КГТ", '
                        f"{cols['orders']}{row_idx} / {days_in_period}, "
                        f"({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                        ")"
                    )
                elif cols["sales"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= ({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= {cols['orders']}{row_idx} / {days_in_period}"
                    )
            else:
                if cols["sales"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= ({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= {cols['orders']}{row_idx} / {days_in_period}"
                    )

        # Потребности
        if cols.get(f"demand_{demand_period_1}"):
            ws[f"{cols[f'demand_{demand_period_1}']}{row_idx}"] = (
                f"= {cols['turnover']}{row_idx} * {demand_period_1}"
            )
        if cols.get(f"demand_{demand_period_2}"):
            ws[f"{cols[f'demand_{demand_period_2}']}{row_idx}"] = (
                f"= {cols['turnover']}{row_idx} * {demand_period_2}"
            )

        # Дефициты
        if (
            cols.get(f"deficit_{demand_period_1}")
            and cols["stock"]
            and cols["supplies"]
        ):
            ws[f"{cols[f'deficit_{demand_period_1}']}{row_idx}"] = (
                f"= {cols['stock']}{row_idx} + {cols['supplies']}{row_idx} "
                f"- {cols[f'demand_{demand_period_1}']}{row_idx}"
            )
        if (
            cols.get(f"deficit_{demand_period_1}_fbs")
            and cols["stock"]
            and cols["stock_fbs"]
            and cols["supplies"]
        ):
            ws[f"{cols[f'deficit_{demand_period_1}_fbs']}{row_idx}"] = (
                f"= {cols['stock']}{row_idx} + {cols['stock_fbs']}{row_idx} "
                f"+ {cols['supplies']}{row_idx} "
                f"- {cols[f'demand_{demand_period_1}']}{row_idx}"
            )
        if (
            cols.get(f"deficit_{demand_period_2}")
            and cols["stock"]
            and cols["supplies"]
        ):
            ws[f"{cols[f'deficit_{demand_period_2}']}{row_idx}"] = (
                f"= {cols['stock']}{row_idx} + {cols['supplies']}{row_idx} "
                f"- {cols[f'demand_{demand_period_2}']}{row_idx}"
            )

        # Округлённые потребности
        if (
            cols.get(f"demand_{demand_period_1}_rounded")
            and cols.get(f"deficit_{demand_period_1}")
        ):
            ws[f"{cols[f'demand_{demand_period_1}_rounded']}{row_idx}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_1}']}{row_idx}",
                    f"{cols['box_multiplicity']}{row_idx}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )
        if (
            cols.get(f"demand_{demand_period_1}_rounded_fbs")
            and cols.get(f"deficit_{demand_period_1}_fbs")
        ):
            ws[f"{cols[f'demand_{demand_period_1}_rounded_fbs']}{row_idx}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_1}_fbs']}{row_idx}",
                    f"{cols['box_multiplicity']}{row_idx}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )
        if (
            cols.get(f"demand_{demand_period_2}_rounded")
            and cols.get(f"deficit_{demand_period_2}")
        ):
            ws[f"{cols[f'demand_{demand_period_2}_rounded']}{row_idx}"] = (
                _create_demand_rounded_formula(
                    f"{cols[f'deficit_{demand_period_2}']}{row_idx}",
                    f"{cols['box_multiplicity']}{row_idx}"
                    if cols["box_multiplicity"]
                    else None,
                    apply_rounding=apply_box_rounding,
                )
            )

    # ----------------- Условное форматирование (дефициты) -----------------
    deficit_columns = [
        cols.get(f"deficit_{demand_period_1}"),
        cols.get(f"deficit_{demand_period_1}_fbs"),
        cols.get(f"deficit_{demand_period_2}"),
    ]
    deficit_columns = [c for c in deficit_columns if c]
    if deficit_columns:
        _apply_conditional_formatting(ws, deficit_columns, row_start, row_end)
    # ------------------------------------------------------------------
    # Доп. условное форматирование для BIO/ARK (WB):
    # если Статус товара = "Основной ассортимент" и
    # Остаток <Склад> = 5, подсветить ячейку остатка красным.
    # Склады: Коледино, Краснодар, Воронеж, Казань.
    # ------------------------------------------------------------------
    if client_name in ("BIO", "ARK"):
        status_col = _get_excel_column(mapping, "Статус товара")
        if status_col:
            target_warehouses = ["Коледино", "Краснодар", "Воронеж", "Казань"]

            red_font = Font(color="9C0006")
            red_fill = PatternFill(
                start_color="FFC7CE",
                end_color="FFC7CE",
                fill_type="solid",
            )

            for wh in target_warehouses:
                stock_col = _get_excel_column(
                    mapping,
                    f"Остаток {wh}",
                )
                if not stock_col:
                    continue

                cell_range = f"{stock_col}{row_start}:{stock_col}{row_end}"
                # Условие: статус = "Основной ассортимент" и значение в ячейке = 5
                # Формула вида:
                # =AND($<Статус><первая_с��рока>="Основной ассортимент",
                #      <эта_ячейка>=5)
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
    # ----------------- Форматы чисел -----------------
    two_decimal_cols = [
        cols.get("turnover"),
        cols.get(f"demand_{demand_period_1}"),
        cols.get(f"demand_{demand_period_2}"),
    ]
    two_decimal_cols = [c for c in two_decimal_cols if c]
    if two_decimal_cols:
        _apply_number_format(ws, two_decimal_cols, row_start, row_end, decimals=2)

    one_decimal_cols = deficit_columns
    if one_decimal_cols:
        _apply_number_format(ws, one_decimal_cols, row_start, row_end, decimals=1)

    zero_decimal_cols = [
        cols.get("barcode"),
        cols.get(f"demand_{demand_period_1}_rounded"),
        cols.get(f"demand_{demand_period_1}_rounded_fbs"),
        cols.get(f"demand_{demand_period_2}_rounded"),
    ]
    zero_decimal_cols = [c for c in zero_decimal_cols if c]
    if zero_decimal_cols:
        _apply_number_format(ws, zero_decimal_cols, row_start, row_end, decimals=0)

    # ----------------- Выравнивание -----------------
    center_cols = [
        "Сезон",
        "Статус",
        "Статус товара",
        "Статус продажи FBO/FBS",
        "Основной артикул",
        "Артикул продавца",
        "Артикул_Размер",
        "Артикул WB",
        "Штрихкод",
        "Наименование товара",
        "Предмет",
        "Цвет",
    ]
    center_mask = '|'.join(center_cols)
    text_mask = mapping["column"].str.contains(
        center_mask,
        na=False,
    )
    left_cols = mapping.loc[text_mask, "excel_column"].tolist()
    center_cols = mapping.loc[~text_mask, "excel_column"].tolist()

    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, horizontal="center")
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, horizontal="left")

    wb.save(file_path)
    logger.success(f'Завершено форматирование листа "{sheet_name}"')

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
    Форматирует лист 'По кластерам' для WB.

    Логика:
      - для BIO и ARK используются 'Дни планирования' и кастомные поля
        'Потребность', 'Дефицит/Избыток', '... с учетом FBS',
        'Потребность (округл.)', 'Потребность с учетом FBS (округл.)';
      - для остальных клиентов:
          * Оборачиваемость = (Заказы + Продажи) / 2 / days_in_period;
          * Потребность/дефициты считаются на фиксированные периоды
            demand_period_1 и demand_period_2;
          * 'Корректировка с учетом нулевых остатков' =
              'Потребность на {demand_period_for_total} дней (округл.)'.
    """
    logger.info(f'Начало форматирования листа "{sheet_name}"')

    file_path = (
        path_supply_svod
        / f"{date_report_created}_Расчет_поставок_{client_name}_WB.xlsx"
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
        "stock_fbs": _get_excel_column(mapping, "Остатки_fbs", contains=True),
        "supplies": _get_excel_column(mapping, "Ожидаемое поступление"),
        "correction": _get_excel_column(
            mapping, "Корректировка с учетом нулевых остатков"
        ),
        "box_multiplicity": _get_excel_column(mapping, "Кратность коробов"),
        "size_type": _get_excel_column(mapping, "Габариты товара"),
        "planning_days": _get_excel_column(mapping, "Дни планирования"),
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
                mapping, f"Дефицит/Избыток {demand_period_1} дней"
            ),
            f"deficit_{demand_period_2}": _get_excel_column(
                mapping, f"Дефицит/Избыток {demand_period_2} дней"
            ),
            f"demand_{demand_period_1}_rounded": _get_excel_column(
                mapping, f"Потребность на {demand_period_1} дней (округл.)"
            ),
            f"demand_{demand_period_2}_rounded": _get_excel_column(
                mapping, f"Потребность на {demand_period_2} дней (округл.)"
            ),
            # кастомные колонки для BIO/ARK
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

    # ----------------- Формулы по строкам -----------------
    for row_idx in range(row_start, row_end + 1):
        # Оборачиваемость
        if cols["turnover"] and cols["orders"]:
            if client_name in ("BIO", "ARK"):
                if cols["sales"] and cols["size_type"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        "=IF("
                        f'{cols["size_type"]}{row_idx}="КГТ", '
                        f"{cols['orders']}{row_idx} / {days_in_period}, "
                        f"({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                        ")"
                    )
                elif cols["sales"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= ({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= {cols['orders']}{row_idx} / {days_in_period}"
                    )
            else:
                if cols["sales"]:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= ({cols['orders']}{row_idx} + {cols['sales']}{row_idx}) / 2 / {days_in_period}"
                    )
                else:
                    ws[f"{cols['turnover']}{row_idx}"] = (
                        f"= {cols['orders']}{row_idx} / {days_in_period}"
                    )

        # ----------------- BIO / ARK: логика с Днями планирования -----------------
        if client_name in ("BIO", "ARK") and cols["planning_days"]:
            # Потребность = Оборачиваемость * Дни планирования
            if cols["demand_custom"] and cols["turnover"]:
                ws[f"{cols['demand_custom']}{row_idx}"] = (
                    f"= {cols['turnover']}{row_idx} * {cols['planning_days']}{row_idx}"
                )
            # Дефицит / избыток
            if (
                cols["deficit_custom"]
                and cols["stock"]
                and cols["supplies"]
                and cols["demand_custom"]
            ):
                ws[f"{cols['deficit_custom']}{row_idx}"] = (
                    f"= {cols['stock']}{row_idx} + {cols['supplies']}{row_idx} "
                    f"- {cols['demand_custom']}{row_idx}"
                )
            # Дефицит / избыток с учетом FBS
            if (
                cols["deficit_custom_fbs"]
                and cols["stock"]
                and cols["stock_fbs"]
                and cols["supplies"]
                and cols["demand_custom"]
            ):
                ws[f"{cols['deficit_custom_fbs']}{row_idx}"] = (
                    f"= {cols['stock']}{row_idx} + {cols['stock_fbs']}{row_idx} "
                    f"+ {cols['supplies']}{row_idx} - {cols['demand_custom']}{row_idx}"
                )
            # Округлённая потребность
            if cols["demand_custom_rounded"] and cols["deficit_custom"]:
                ws[f"{cols['demand_custom_rounded']}{row_idx}"] = (
                    _create_demand_rounded_formula(
                        f"{cols['deficit_custom']}{row_idx}",
                        f"{cols['box_multiplicity']}{row_idx}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            # Округлённая потребность с FBS
            if cols["demand_custom_rounded_fbs"] and cols["deficit_custom_fbs"]:
                ws[f"{cols['demand_custom_rounded_fbs']}{row_idx}"] = (
                    _create_demand_rounded_formula(
                        f"{cols['deficit_custom_fbs']}{row_idx}",
                        f"{cols['box_multiplicity']}{row_idx}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            # Корректировка = Потребность (округл.)
            if cols["correction"] and cols["demand_custom_rounded"]:
                ws[f"{cols['correction']}{row_idx}"] = (
                    f"= {cols['demand_custom_rounded']}{row_idx}"
                )

        # ----------------- Остальные клиенты: логика по периодам -----------------
        else:
            # Потребность на период 1/2
            if cols.get(f"demand_{demand_period_1}"):
                ws[f"{cols[f'demand_{demand_period_1}']}{row_idx}"] = (
                    f"= {cols['turnover']}{row_idx} * {demand_period_1}"
                )
            if cols.get(f"demand_{demand_period_2}"):
                ws[f"{cols[f'demand_{demand_period_2}']}{row_idx}"] = (
                    f"= {cols['turnover']}{row_idx} * {demand_period_2}"
                )

            # Дефицит / избыток
            if (
                cols.get(f"deficit_{demand_period_1}")
                and cols["stock"]
                and cols["supplies"]
                and cols.get(f"demand_{demand_period_1}")
            ):
                ws[f"{cols[f'deficit_{demand_period_1}']}{row_idx}"] = (
                    f"= {cols['stock']}{row_idx} + {cols['supplies']}{row_idx} "
                    f"- {cols[f'demand_{demand_period_1}']}{row_idx}"
                )
            if (
                cols.get(f"deficit_{demand_period_2}")
                and cols["stock"]
                and cols["supplies"]
                and cols.get(f"demand_{demand_period_2}")
            ):
                ws[f"{cols[f'deficit_{demand_period_2}']}{row_idx}"] = (
                    f"= {cols['stock']}{row_idx} + {cols['supplies']}{row_idx} "
                    f"- {cols[f'demand_{demand_period_2}']}{row_idx}"
                )

            # Округлённые потребности
            if (
                cols.get(f"demand_{demand_period_1}_rounded")
                and cols.get(f"deficit_{demand_period_1}")
            ):
                ws[f"{cols[f'demand_{demand_period_1}_rounded']}{row_idx}"] = (
                    _create_demand_rounded_formula(
                        f"{cols[f'deficit_{demand_period_1}']}{row_idx}",
                        f"{cols['box_multiplicity']}{row_idx}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )
            if (
                cols.get(f"demand_{demand_period_2}_rounded")
                and cols.get(f"deficit_{demand_period_2}")
            ):
                ws[f"{cols[f'demand_{demand_period_2}_rounded']}{row_idx}"] = (
                    _create_demand_rounded_formula(
                        f"{cols[f'deficit_{demand_period_2}']}{row_idx}",
                        f"{cols['box_multiplicity']}{row_idx}"
                        if cols["box_multiplicity"]
                        else None,
                        apply_rounding=apply_box_rounding,
                    )
                )

            # Корректировка = Потребность на demand_period_for_total (округл.)
            rounded_col_for_total = cols.get(
                f"demand_{demand_period_for_total}_rounded"
            )
            if cols["correction"] and rounded_col_for_total:
                ws[f"{cols['correction']}{row_idx}"] = (
                    f"= {rounded_col_for_total}{row_idx}"
                )

    # ----------------- Условное форматирование -----------------
    if client_name in ("BIO", "ARK") and cols.get("deficit_custom"):
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

    # ----------------- Форматы чисел -----------------
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

    # ----------------- Выравнивание -----------------
    center_cols = [
        "Сезон",
        "Статус",
        "Статус товара",
        "Статус продажи FBO/FBS",
        "Основной артикул",
        "Артикул продавца",
        "Артикул_Размер",
        "Артикул WB",
        "Штрихкод",
        "Наименование товара",
        "Предмет",
        "Цвет",
    ]
    center_mask = '|'.join(center_cols)
    text_mask = mapping["column"].str.contains(
        center_mask,
        na=False,
    )
    left_cols = mapping.loc[text_mask, "excel_column"].tolist()
    center_cols = mapping.loc[~text_mask, "excel_column"].tolist()


    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, "center")
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, "left")

    wb.save(file_path)
    logger.success(f'Завершено форматирование листа "{sheet_name}"')
