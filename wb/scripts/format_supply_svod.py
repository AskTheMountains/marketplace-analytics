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


def _auto_adjust_column_width(ws, barcode_column: str = None) -> None:
    """
    Автоматически подбирает ширину столбцов.

    Args:
        ws: Worksheet для настройки
        barcode_column: Буква колонки со штрихкодом (фиксированная ширина)
    """
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass

        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    # Фиксированная ширина для штрихкода
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
    Создает формулу Excel для округленной потребности с учетом кратности коробов.

    Логика (аналог твоей формулы):
    1. Если дефицит/избыток > 0 (избыток) → потребность = 0
    2. Иначе (дефицит):
       a. Если остаток от деления < половины короба → округляем обычным ROUNDUP
       b. Иначе → округляем до кратности короба

    Без округления: =ROUNDUP(IF(deficit>0, 0, -deficit), 0)
    С округлением: =IF(deficit>0, 0, IF(MOD(-deficit, multiplicity)<multiplicity/2, -ROUNDUP(deficit, 0), ROUNDUP(-deficit/multiplicity, 0)*multiplicity))

    Args:
        deficit_cell: Адрес ячейки с дефицитом/избытком (может быть отрицательным)
        box_multiplicity_cell: Адрес ячейки с кратностью коробов
        apply_rounding: Применять ли округление по кратности

    Returns:
        Формула Excel
    """
    if not apply_rounding:
        # Стандартная формула без учета кратности коробов
        return f"= ROUNDUP(IF({deficit_cell} > 0, 0, -{deficit_cell}), 0)"

    # Формула с учетом кратности коробов (аналог твоей)
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
    clusters_mapping_df: pd.DataFrame,
    cluster_column: str,
    sheet_name: str = 'Всего',
    demand_period_1: int = 40,
    demand_period_2: int = 60,
    days_in_period: int = 31,
    apply_box_rounding: bool = False
) -> None:
    """
    Форматирует лист 'Всего' или 'Всего по ювелирным складам'.
    Args:
        path_supply_svod: Путь к директории с файлами
        client_name: Имя клиента
        date_report_created: Дата создания отчета
        svod_excel: DataFrame с данными листа
        svod_excel_clusters: DataFrame с данными по кластерам
        clusters_mapping_df: Маппинг складов
        cluster_column: Колонка для группировки складов (в маппинге)
        sheet_name: Название листа
        demand_period_1: Первый период потребности (дни)
        demand_period_2: Второй период потребности (дни)
        days_in_period: Период оборачиваемости (дни)
        apply_box_rounding: Применять ли округление по кратности коробов
    """
    logger.info(f'Форматирование листа "{sheet_name}"')

    file_path = path_supply_svod / f"{date_report_created}_Расчет_поставок_{client_name}_WB_formatted.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    # Параметры
    svod_len = svod_excel.shape[0]
    row_start = 2
    row_end = svod_len + 1

    # Маппинг колонок
    mapping = _create_column_mapping(svod_excel)
    mapping_clusters = _create_column_mapping(svod_excel_clusters)

    # Границы диапазона по колонкам
    min_col = mapping['excel_column'].iloc[0]
    max_col = mapping['excel_column'].iloc[-1]

    # Основные колонки на листе "Всего"
    cols = {
        'barcode': _get_excel_column(mapping, 'Штрихкод'),
        'article': _get_excel_column(mapping, 'Артикул продавца'),
        'size': _get_excel_column(mapping, 'Размер'),
        'article_size': _get_excel_column(mapping, 'Артикул_Размер'),
        'turnover': _get_excel_column(mapping, 'Оборачиваемость'),
        'orders': _get_excel_column(mapping, 'ЗАКАЗЫ', contains=True),
        'sales': _get_excel_column(mapping, 'ПРОДАЖИ', contains=True),
        'stock': _get_excel_column(mapping, 'ОСТАТОК', contains=True),
        'stock_fbs': _get_excel_column(mapping, 'ОСТАТОК FBS', contains=True),
        'supplies': _get_excel_column(mapping, 'Ожидаемое поступление'),
        'box_multiplicity': _get_excel_column(mapping, 'Кратность коробов'),
    }

    # Динамические колонки по периодам
    cols.update({
        f'demand_{demand_period_1}': _get_excel_column(mapping, f'Потребность на {demand_period_1} дней'),
        f'demand_{demand_period_2}': _get_excel_column(mapping, f'Потребность на {demand_period_2} дней'),
        f'deficit_{demand_period_1}': _get_excel_column(mapping, f'Дефицит/Избыток {demand_period_1} дней'),
        f'deficit_{demand_period_1}_fbs': _get_excel_column(mapping, f'Дефицит/Избыток {demand_period_1} дней с учетом FBS'),
        f'deficit_{demand_period_2}': _get_excel_column(mapping, f'Дефицит/Избыток {demand_period_2} дней'),
        f'demand_{demand_period_1}_rounded': _get_excel_column(mapping, f'Потребность на {demand_period_1} дней (округл.)'),
        f'demand_{demand_period_1}_rounded_fbs': _get_excel_column(mapping, f'Потребность на {demand_period_1} дней с учетом FBS (округл.)'),
        f'demand_{demand_period_2}_rounded': _get_excel_column(mapping, f'Потребность на {demand_period_2} дней (округл.)'),
    })

    # Колонки из листа "По кластерам"
    cols_clusters = {
        'article_size': _get_excel_column(mapping_clusters, 'Артикул_Размер'),
        'cluster': _get_excel_column(mapping_clusters, 'Склад'),
        'correction': _get_excel_column(mapping_clusters, 'Корректировка с учетом нулевых остатков'),
        'stock': _get_excel_column(mapping_clusters, 'ОСТАТОК', contains=True),
    }

    # Список "агрегированных" кластеров (по листу "Всего")
    cluster_list = [
        col.replace('Потребность ', '')
        for col in mapping['column']
        if 'Потребность ' in col and f'Потребность на' not in col
    ]

    # ========================================================================
    # БАЗОВОЕ ФОРМАТИРОВАНИЕ
    # ========================================================================
    thin = Side(border_style="thin", color="000000")

    header_cells = ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    _apply_header_formatting(header_cells, thin)

    ws.auto_filter.ref = ws.dimensions

    all_cells = ws[f"{min_col}{row_start}:{max_col}{row_end}"]
    _apply_cell_borders(all_cells, thin)

    _auto_adjust_column_width(ws, cols['barcode'])

    # ========================================================================
    # ФОРМУЛЫ ПО КЛАСТЕРАМ (Потребность + Остаток, как в Ozon)
    # ========================================================================
    # Готовим маппинг: агрегированный кластер -> список субкластеров (Складов)
    clusters_mapping_processed = (
        clusters_mapping_df[['Склад', cluster_column]]
        .groupby([cluster_column])
        .agg(list)
        .reset_index()
        .rename(columns={'Склад': 'subcluster', cluster_column: 'cluster'})
    )

    clusters_mapping_processed = clusters_mapping_processed[
        clusters_mapping_processed['cluster'].isin(cluster_list)
    ]

    svod_len_clusters = svod_excel_clusters.shape[0]
    row_start_clusters = 2

    for _, cluster_row in clusters_mapping_processed.iterrows():
        cluster = cluster_row['cluster']
        subclusters = cluster_row['subcluster'].copy()
        # Добавляем сам кластер в список субкластеров
        subclusters.insert(0, cluster)
        subclusters = list(dict.fromkeys(subclusters))

        demand_cluster_col = _get_excel_column(mapping, f'Потребность {cluster}')
        stock_cluster_col = _get_excel_column(mapping, f'Остаток {cluster}')

        # Если нет хотя бы одной из колонок — нет смысла считать
        if not demand_cluster_col or not stock_cluster_col:
            continue

        for i in range(row_start, row_end + 1):
            demand_parts = []
            stock_parts = []
            for sub in subclusters:
                # Потребность — сумма "Корректировка с учетом нулевых остатков"
                demand_parts.append(
                    f"SUMIFS('По кластерам'!${cols_clusters['correction']}${row_start_clusters}:${cols_clusters['correction']}${svod_len_clusters + 1}, "
                    f"'По кластерам'!${cols_clusters['article_size']}${row_start_clusters}:${cols_clusters['article_size']}${svod_len_clusters + 1}, ${cols['article_size']}{i}, "
                    f"'По кластерам'!${cols_clusters['cluster']}${row_start_clusters}:${cols_clusters['cluster']}${svod_len_clusters + 1}, \"{sub}\")"
                )
                # Остаток — сумма "ОСТАТОК" по тем же субкластерлам
                stock_parts.append(
                    f"SUMIFS('По кластерам'!${cols_clusters['stock']}${row_start_clusters}:${cols_clusters['stock']}${svod_len_clusters + 1}, "
                    f"'По кластерам'!${cols_clusters['article_size']}${row_start_clusters}:${cols_clusters['article_size']}${svod_len_clusters + 1}, ${cols['article_size']}{i}, "
                    f"'По кластерам'!${cols_clusters['cluster']}${row_start_clusters}:${cols_clusters['cluster']}${svod_len_clusters + 1}, \"{sub}\")"
                )

            ws[f"{demand_cluster_col}{i}"] = "= " + " + ".join(demand_parts)
            ws[f"{stock_cluster_col}{i}"] = "= " + " + ".join(stock_parts)

    # ========================================================================
    # ОСНОВНЫЕ ФОРМУЛЫ
    # ========================================================================
    for i in range(row_start, row_end + 1):
        # Оборачиваемость
        if cols['turnover'] and cols['orders'] and cols['sales']:
            ws[f"{cols['turnover']}{i}"] = (
                f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
            )
        # Потребность на период 1
        if cols[f'demand_{demand_period_1}']:
            ws[f"{cols[f'demand_{demand_period_1}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_1}"
            )
        # Потребность на период 2
        if cols[f'demand_{demand_period_2}']:
            ws[f"{cols[f'demand_{demand_period_2}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_2}"
            )
        # Дефицит/избыток период 1
        if cols[f'deficit_{demand_period_1}'] and cols['stock'] and cols['supplies']:
            ws[f"{cols[f'deficit_{demand_period_1}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
            )
        # Дефицит/избыток период 1 с FBS
        if (
            cols[f'deficit_{demand_period_1}_fbs']
            and cols['stock']
            and cols['stock_fbs']
            and cols['supplies']
        ):
            ws[f"{cols[f'deficit_{demand_period_1}_fbs']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['stock_fbs']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
            )
        # Дефицит/избыток период 2
        if cols[f'deficit_{demand_period_2}'] and cols['stock'] and cols['supplies']:
            ws[f"{cols[f'deficit_{demand_period_2}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_2}']}{i}"
            )
        # Потребность период 1 (округл.)
        if cols[f'demand_{demand_period_1}_rounded'] and cols[f'deficit_{demand_period_1}']:
            ws[f"{cols[f'demand_{demand_period_1}_rounded']}{i}"] = _create_demand_rounded_formula(
                f"{cols[f'deficit_{demand_period_1}']}{i}",
                f"{cols['box_multiplicity']}{i}" if cols['box_multiplicity'] else None,
                apply_rounding=apply_box_rounding
            )
        # Потребность период 1 с FBS (округл.)
        if cols[f'demand_{demand_period_1}_rounded_fbs'] and cols[f'deficit_{demand_period_1}_fbs']:
            ws[f"{cols[f'demand_{demand_period_1}_rounded_fbs']}{i}"] = _create_demand_rounded_formula(
                f"{cols[f'deficit_{demand_period_1}_fbs']}{i}",
                f"{cols['box_multiplicity']}{i}" if cols['box_multiplicity'] else None,
                apply_rounding=apply_box_rounding
            )
        # Потребность период 2 (округл.)
        if cols[f'demand_{demand_period_2}_rounded'] and cols[f'deficit_{demand_period_2}']:
            ws[f"{cols[f'demand_{demand_period_2}_rounded']}{i}"] = _create_demand_rounded_formula(
                f"{cols[f'deficit_{demand_period_2}']}{i}",
                f"{cols['box_multiplicity']}{i}" if cols['box_multiplicity'] else None,
                apply_rounding=apply_box_rounding
            )

    # ========================================================================
    # УСЛОВНОЕ ФОРМАТИРОВАНИЕ И СТИЛИ
    # ========================================================================
    deficit_columns = [
        cols[f'deficit_{demand_period_1}'],
        cols[f'deficit_{demand_period_1}_fbs'],
        cols[f'deficit_{demand_period_2}'],
    ]
    deficit_columns = [col for col in deficit_columns if col]
    if deficit_columns:
        _apply_conditional_formatting(ws, deficit_columns, row_start, row_end)

    two_decimal_cols = [
        cols['turnover'],
        cols[f'demand_{demand_period_1}'],
        cols[f'demand_{demand_period_2}'],
    ]
    two_decimal_cols = [col for col in two_decimal_cols if col]
    if two_decimal_cols:
        _apply_number_format(ws, two_decimal_cols, row_start, row_end, decimals=2)

    one_decimal_cols = deficit_columns
    if one_decimal_cols:
        _apply_number_format(ws, one_decimal_cols, row_start, row_end, decimals=1)

    zero_decimal_cols = [
        cols['barcode'],
        cols[f'demand_{demand_period_1}_rounded'],
        cols[f'demand_{demand_period_1}_rounded_fbs'],
        cols[f'demand_{demand_period_2}_rounded'],
    ]
    zero_decimal_cols = [col for col in zero_decimal_cols if col]
    if zero_decimal_cols:
        _apply_number_format(ws, zero_decimal_cols, row_start, row_end, decimals=0)

    center_cols = [
        col for col in mapping['excel_column']
        if col not in [
            cols['article_size'], cols['barcode'], cols['article'],
            cols['size'], _get_excel_column(mapping, 'Предмет'),
            _get_excel_column(mapping, 'Наименование')
        ]
    ]
    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, 'center')

    left_cols = [
        cols['article_size'], cols['barcode'], cols['article'],
        cols['size'], _get_excel_column(mapping, 'Предмет'),
        _get_excel_column(mapping, 'Наименование')
    ]
    left_cols = [col for col in left_cols if col]
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, 'left')

    wb.save(file_path)
    logger.info(f'Завершено форматирование листа "{sheet_name}"')


def format_sheet_clusters(
    path_supply_svod,
    client_name: str,
    date_report_created: str,
    svod_excel: pd.DataFrame,
    sheet_name: str = 'По кластерам',
    demand_period_1: int = 40,
    demand_period_2: int = 60,
    days_in_period: int = 31,
    apply_box_rounding: bool = False
) -> None:
    """
    Форматирует лист 'По кластерам'.

    Args:
        path_supply_svod: Путь к директории с файлами
        client_name: Имя клиента
        date_report_created: Дата создания отчета
        svod_excel: DataFrame с данными листа
        sheet_name: Название листа
        demand_period_1: Первый период потребности (дни)
        demand_period_2: Второй период потребности (дни)
        days_in_period: Период оборачиваемости (дни)
        apply_box_rounding: Применять ли округление по кратности коробов
    """
    logger.info(f'Форматирование листа "{sheet_name}"')

    # Открываем файл
    file_path = path_supply_svod / f"{date_report_created}_Расчет_поставок_{client_name}_WB_formatted.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    # Параметры
    svod_len = svod_excel.shape[0]
    row_start = 2
    row_end = svod_len + 1

    # Создаем маппинг колонок
    mapping = _create_column_mapping(svod_excel)

    # Получаем буквы колонок
    min_col = mapping['excel_column'].iloc[0]
    max_col = mapping['excel_column'].iloc[-1]

    # Основные колонки
    cols = {
        'barcode': _get_excel_column(mapping, 'Штрихкод'),
        'turnover': _get_excel_column(mapping, 'Оборачиваемость'),
        'orders': _get_excel_column(mapping, 'ЗАКАЗЫ', contains=True),
        'sales': _get_excel_column(mapping, 'ПРОДАЖИ', contains=True),
        'stock': _get_excel_column(mapping, 'ОСТАТОК', contains=True),
        'supplies': _get_excel_column(mapping, 'Ожидаемое поступление'),
        'correction': _get_excel_column(mapping, 'Корректировка с учетом нулевых остатков'),
        'box_multiplicity': _get_excel_column(mapping, 'Кратность коробов'),
    }

    # Динамические названия колонок
    cols.update({
        f'demand_{demand_period_1}': _get_excel_column(mapping, f'Потребность на {demand_period_1} дней'),
        f'demand_{demand_period_2}': _get_excel_column(mapping, f'Потребность на {demand_period_2} дней'),
        f'deficit_{demand_period_1}': _get_excel_column(mapping, f'Дефицит/Избыток {demand_period_1} дней'),
        f'deficit_{demand_period_2}': _get_excel_column(mapping, f'Дефицит/Избыток {demand_period_2} дней'),
        f'demand_{demand_period_1}_rounded': _get_excel_column(mapping, f'Потребность на {demand_period_1} дней (округл.)'),
        f'demand_{demand_period_2}_rounded': _get_excel_column(mapping, f'Потребность на {demand_period_2} дней (округл.)'),
    })

    # ========================================================================
    # БАЗОВОЕ ФОРМАТИРОВАНИЕ
    # ========================================================================

    thin = Side(border_style="thin", color="000000")

    # Заголовки
    header_cells = ws[f"{min_col}{row_start - 1}:{max_col}{row_start - 1}"]
    _apply_header_formatting(header_cells, thin)

    # Фильтр
    ws.auto_filter.ref = ws.dimensions

    # Границы
    all_cells = ws[f"{min_col}{row_start}:{max_col}{row_end}"]
    _apply_cell_borders(all_cells, thin)

    # Ширина колонок
    _auto_adjust_column_width(ws, cols['barcode'])

    # ========================================================================
    # ФОРМУЛЫ
    # ========================================================================

    for i in range(row_start, row_end + 1):
        # Оборачиваемость
        if cols['turnover']:
            ws[f"{cols['turnover']}{i}"] = (
                f"= ({cols['orders']}{i} + {cols['sales']}{i}) / 2 / {days_in_period}"
            )

        # Потребность на период 1
        if cols[f'demand_{demand_period_1}']:
            ws[f"{cols[f'demand_{demand_period_1}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_1}"
            )

        # Потребность на период 2
        if cols[f'demand_{demand_period_2}']:
            ws[f"{cols[f'demand_{demand_period_2}']}{i}"] = (
                f"= {cols['turnover']}{i} * {demand_period_2}"
            )

        # Дефицит/избыток период 1
        if cols[f'deficit_{demand_period_1}']:
            ws[f"{cols[f'deficit_{demand_period_1}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_1}']}{i}"
            )

        # Дефицит/избыток период 2
        if cols[f'deficit_{demand_period_2}']:
            ws[f"{cols[f'deficit_{demand_period_2}']}{i}"] = (
                f"= {cols['stock']}{i} + {cols['supplies']}{i} - {cols[f'demand_{demand_period_2}']}{i}"
            )

        # Потребность период 1 (округл.) с учетом кратности коробов
        if cols[f'demand_{demand_period_1}_rounded'] and cols[f'deficit_{demand_period_1}']:
            ws[f"{cols[f'demand_{demand_period_1}_rounded']}{i}"] = _create_demand_rounded_formula(
                f"{cols[f'deficit_{demand_period_1}']}{i}",
                f"{cols['box_multiplicity']}{i}" if cols['box_multiplicity'] else None,
                apply_rounding=apply_box_rounding
            )

        # Потребность период 2 (округл.) с учетом кратности коробов
        if cols[f'demand_{demand_period_2}_rounded'] and cols[f'deficit_{demand_period_2}']:
            ws[f"{cols[f'demand_{demand_period_2}_rounded']}{i}"] = _create_demand_rounded_formula(
                f"{cols[f'deficit_{demand_period_2}']}{i}",
                f"{cols['box_multiplicity']}{i}" if cols['box_multiplicity'] else None,
                apply_rounding=apply_box_rounding
            )

        # Корректировка (используем потребность на период 2)
        if cols['correction'] and cols[f'demand_{demand_period_2}_rounded']:
            ws[f"{cols['correction']}{i}"] = f"= {cols[f'demand_{demand_period_2}_rounded']}{i}"

    # ========================================================================
    # УСЛОВНОЕ ФОРМАТИРОВАНИЕ И СТИЛИ (ПРИМЕНЯЕМ ПОСЛЕ ФОРМУЛ!)
    # ========================================================================

    # Условное форматирование для дефицита/избытка
    deficit_columns = [
        cols[f'deficit_{demand_period_1}'],
        cols[f'deficit_{demand_period_2}']
    ]
    deficit_columns = [col for col in deficit_columns if col]  # Убираем None
    if deficit_columns:
        _apply_conditional_formatting(ws, deficit_columns, row_start, row_end)

    # Форматирование чисел (2 знака после запятой)
    two_decimal_cols = [
        cols['turnover'],
        cols[f'demand_{demand_period_1}'],
        cols[f'demand_{demand_period_2}']
    ]
    two_decimal_cols = [col for col in two_decimal_cols if col]  # Убираем None
    if two_decimal_cols:
        _apply_number_format(ws, two_decimal_cols, row_start, row_end, decimals=2)

    # Форматирование чисел (1 знак после запятой)
    one_decimal_cols = deficit_columns  # Уже отфильтрованный список
    if one_decimal_cols:
        _apply_number_format(ws, one_decimal_cols, row_start, row_end, decimals=1)

    # Форматирование чисел (0 знаков после запятой) - для округленных колонок и штрихкода
    zero_decimal_cols = [
        cols['barcode'],
        cols[f'demand_{demand_period_1}_rounded'],
        cols[f'demand_{demand_period_2}_rounded'],
        cols['correction']
    ]
    zero_decimal_cols = [col for col in zero_decimal_cols if col]  # Убираем None
    if zero_decimal_cols:
        _apply_number_format(ws, zero_decimal_cols, row_start, row_end, decimals=0)

    # Выравнивание по центру (все числовые колонки)
    center_cols = [
        col for col in mapping['excel_column']
        if not any(keyword in mapping.loc[mapping['excel_column'] == col, 'column'].values[0]
                  for keyword in ['Штрихкод', 'Предмет', 'Артикул продавца', 'Размер', 'Наименование', 'Склад'])
    ]
    if center_cols:
        _apply_alignment(ws, center_cols, row_start, row_end, 'center')

    # Выравнивание слева (текстовые колонки)
    left_cols = [
        col for col in mapping['excel_column']
        if any(keyword in mapping.loc[mapping['excel_column'] == col, 'column'].values[0]
              for keyword in ['Штрихкод', 'Предмет', 'Артикул продавца', 'Размер', 'Наименование', 'Склад'])
    ]
    if left_cols:
        _apply_alignment(ws, left_cols, row_start, row_end, 'left')

    # Сохранение
    wb.save(file_path)
    logger.info(f'Завершено форматирование листа "{sheet_name}"')
