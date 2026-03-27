# %%
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from ozon.scripts.uploadDataFromOzon import get_ozon_product
from ozon.scripts.constants import headers

# Получаем список товаров АПИ
df_products = get_ozon_product(headers, to_save=False)
df_products_ = (
    df_products
    .loc[:, ['Артикул', 'SKU', 'Название товара', 'model_id']]
    .copy()
    .sort_values(by=['model_id', 'Артикул'], ignore_index=True)
    .to_clipboard(excel=True, index=False)
)

# %%
path_to_file = Path(
    'd:\\artem\\Downloads\\'
    '2026-03-10_Сводная_РК_Water_Power_Ozon.xlsx'
)
sheet_name='Список товаров'
df = pd.read_excel(path_to_file, sheet_name=sheet_name)
wb = load_workbook(path_to_file)
ws = wb[sheet_name]

# Границы листа
min_row = ws.min_row
max_row = ws.max_row
min_col = ws.min_column
max_col = ws.max_column

# Определяем стиль границ
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Создаем границы
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border


# 1. Список допустимых имён колонок кампаний
campaign_column_names = [
    "model_id"
]
# 2. Палитра из 15 неярких расцветок
color_palette = [
    "FFB3C6", "FFE6E6", "B6E2D3", "B7D3DF", "DEEAF6",
    "F7E1D7", "F5F8C4", "D4E6A5", "A2D5AB", "F9D6C1",
    "F5B5FC", "D9D9D9", "ABC4FF", "C6F1E7", "EEE3CB"
]

# 3. Находим имена колонок кампании в df
df_campaign_cols = [
    col for col in df.columns if any(col.strip().lower() == name.strip().lower() for name in campaign_column_names)
]
if not df_campaign_cols:
    print('jopa')  # Нет нужных колонок — ничего делать не надо

# 4. Получаем все уникальные значения кампании
unique_campaigns = []
for col in df_campaign_cols:
    unique_campaigns.extend(df[col].dropna().unique())
unique_campaigns = list(dict.fromkeys(unique_campaigns))

# 5. Цвета кампаниям (циклично)
campaign_to_color = {
    val: color_palette[i % len(color_palette)]
    for i, val in enumerate(unique_campaigns)
}

# 6. Индексы колонок кампании в Excel (по строке заголовков листа)
header_cells = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
campaign_col_idxs = [
    idx for idx, header in enumerate(header_cells) if header in df_campaign_cols
]
if not campaign_col_idxs:
    print('jopa')

# 7. Форматируем ячейки
for idx in campaign_col_idxs:
    prev_camp = None
    for excel_row_idx, row in enumerate(
        ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2
    ):
        cell = row[idx]
        camp = cell.value
        # Заливаем цвет
        if camp in campaign_to_color:
            cell.fill = PatternFill(
                start_color=campaign_to_color[camp],
                end_color=campaign_to_color[camp],
                fill_type='solid'
            )
        # Проверяем переход к новой кампании
        is_last_row = (excel_row_idx == ws.max_row)
        # Доступ к значению кампании в следующей строке (если есть)
        next_camp = ws[excel_row_idx + 1][idx].value if not is_last_row else None
        if camp != next_camp:
            # Последняя строка блока — нижняя граница жирная у всей строки
            for c in row:
                c.border = Border(
                    left=c.border.left,
                    right=c.border.right,
                    top=c.border.top,
                    bottom=Side(style='thick')
                )
wb.save(path_to_file)

# %%
