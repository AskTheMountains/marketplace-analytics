
# %% Определение всех функций
import pandas as pd
import numpy as np
from datetime import date,datetime,timedelta
import os
import glob
import re
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

# Файл с настройками и номером клиента
from options import settings, headers, client_number

# Файл с некоторыми константами
from constants import ozon_seller_api_url, client_name

# Функция выгрузки текущего списка товаров
from uploadDataFromOzon import get_ozon_product

# Функция чтения файлов по акциям
def read_action_files(client_name, date_start, date_end):
    # Словарь с названиями файлов
    filenames = {"path": [], "file_name": []}
    # Путь к папке с файлам по акциям
    path_actions = f"Clients/{client_name}/Actions"
    # Считывание файла с путем до него
    for file in glob.glob(f"{path_actions}/*_formatted.xlsx"):
        filenames['path'].append(file)
    # Считывание только имени файла
    filenames['file_name'] = [os.path.basename(x) for x in filenames['path']]

    # Переводим в df
    df_filenames_actions = pd.DataFrame(filenames)
    # Создаем столбец с датой создания файла
    df_filenames_actions['date_created'] = df_filenames_actions['file_name'].str.split('_', expand=True)[0]
    # Переводим в datetime
    df_filenames_actions['date_created'] = pd.to_datetime(df_filenames_actions['date_created'])
    df_filenames_actions['date_created'] = df_filenames_actions['date_created']
    # Сортируем по дате
    df_filenames_actions = df_filenames_actions.sort_values(by=['date_created'], ignore_index=True)

    # Фильтруем список файлов,
    # выбирая только те, которые попадают в диапазон начальной и конечной даты
    dt_start = pd.to_datetime(date_start)
    dt_end = pd.to_datetime(date_end)

    df_filenames_actions = df_filenames_actions.loc[df_filenames_actions['date_created'].between(dt_start, dt_end, inclusive='both'), :]
    df_filenames_actions = df_filenames_actions.reset_index(drop=True)

    if df_filenames_actions.empty:
        logger.warning('No action files for specified date range')

    return df_filenames_actions


# Функция чтения справочной таблицы
def read_catalog(client_name):
    # Загрузка справочной таблицы
    catalog_reference = pd.read_excel(f"Clients/{client_name}/catalog/Справочная_таблица_{client_name}_Ozon.xlsx")
    catalog_reference.rename(columns={"Основной артикул": "Основной Артикул"}, inplace=True)
    # Переименовываем колонки для удобства
    catalog_reference = catalog_reference.rename(columns={
        'Артикул': 'Артикул продавца',
        # 'Маржинальность, %': 'Желаемая маржинальность, %'
        })
    # Принудительно переводим некоторые колонки в числовой тип, чтобы не было ошибок
    for col in ['Себестоимость', 'Минимальная цена расчетная, руб.', 'РРЦ', 'Желаемая маржинальность, %', 'Нижняя граница цены после скидки, руб', 'Верхняя граница цены после скидки, руб']:
        if col in catalog_reference.columns:
            catalog_reference[col] = catalog_reference[col].apply(lambda x: pd.to_numeric(x, errors='coerce'))
    # Переводим артикул продавца в строку
    catalog_reference['Артикул продавца'] = catalog_reference['Артикул продавца'].astype(str)

    return catalog_reference


# Функция добавления информации из справочной таблицы в список товаров
def add_catalog_data_to_product_list(df_products, catalog):
    # Создаем копии для избежания изменений в оригинальных df
    df_products_ = df_products.copy()
    catalog_ = catalog.copy()
    # Переименовываем некоторые колонки в списке товаров для удобства
    df_products_ = (
        df_products_
        .loc[:, df_products_.columns.isin([
            'Артикул',
            'SKU',
            'Barcode',
            'Наименование товара',
            'Текущая цена с учетом скидки, ₽',
            'Цена до скидки (перечеркнутая цена), ₽',
            'Цена Premium, ₽',
            'Рыночная цена, ₽',
            'Цена с учетом всех акций',
            'Цена с учетом акций продавца',
            'Минимальная цена после применения всех скидок',
            'Цена до учета скидок (зачеркнутая)',
            'Цена с учетом скидок (на карточке товара)',
            'Цена поставщика',
        ])]
        .rename(columns={
            'Артикул': 'Артикул продавца'
        })
    )

    # Выбираем нужные столбцы из справочной таблицы
    catalog_ = catalog_.loc[:, catalog_.columns.isin([
        'Артикул продавца',
        'РРЦ',
        'Минимальная цена расчетная, руб.'
    ])]

    # Мерджим список товаров и справочную таблицу
    df_products_with_catalog_and_prices = df_products_.merge(
        catalog_,
        on='Артикул продавца',
        how='left'
    )

    return df_products_with_catalog_and_prices


# Функция создания таблицы со сравнениями данных по акциям
def create_action_compare_svod(df_products_with_catalog_and_prices, df_filenames_actions):
    # Создаем копию для избежания изменений в оригинальном df
    df_action_compare_svod = df_products_with_catalog_and_prices.copy()
    # Столбцы по акциям, которые нужно взять из файлов по акциям
    action_cols = ['Цена по акции']
    # Наименования акций, по которым нужно сравнить
    action_names = [
        'Бустинг х2',
        'Бустинг х3',
        'Бустинг х4',
        'Максимальный бустинг: усиление',
        'Максимальный бустинг',
    ]
    # Создаем словарь, в котором будут датафреймы по каждой акции
    actions_dict = {f"{action_name}":df_action_compare_svod.copy() for action_name in action_names}
    # Объединяем в один список
    action_cols_result = [f"{col} {name}" for name in action_names for col in action_cols]

    # Цикл по каждому файлу с акциями
    for i in range(df_filenames_actions.shape[0]):
        # Создаем дату в нужном формате для добавления в столбцы по акциям
        date_action_file = df_filenames_actions['date_created'][i].strftime('%Y-%m-%d')
        dt_action_file = df_filenames_actions['date_created'][i].strftime('%d.%m.%y')
        # Считываем файл по акции
        tmp_df_action = pd.read_excel(df_filenames_actions['path'][i])
        # df, в куда будем помещать данные по конкретной акции (уже содержит данные по ценам и артикулам)
        # tmp_df_action_compare_svod = df_action_compare_svod.copy()
        # Цикл по каждой из нужных колонок по конкретной акции
        for action_name in actions_dict.keys():
            # Создаем список колонок, которые берем по конкретной акции
            action_df_cols = [f"{col} {action_name}" for col in action_cols]
            # Выбираем нужные колонки
            tmp_df_certain_action = tmp_df_action.loc[:, tmp_df_action.columns.isin(
                ['Артикул продавца'] + action_df_cols
            )]
            # Если были найдены колонки по акциям, добавляем их в итоговый df
            if tmp_df_certain_action.columns.size > 1:
                # Убираем ненужные символы или описание из акций
                replace_dict = {'. Без ограничения срока действия': ''}
                for old, new in replace_dict.items():
                    tmp_df_certain_action.columns = tmp_df_certain_action.columns.str.replace(
                        old, new
                    )
                # Сортируем колонки по названию
                # tmp_df_action = tmp_df_action.loc[:, tmp_df_action.columns.sort_values()]
                # Добавляем дату в названия акций
                exclude_cols = ['Артикул продавца']
                tmp_df_certain_action = tmp_df_certain_action.rename(columns=
                    lambda x: f"{dt_action_file} {x}" if x not in exclude_cols else x
                )
                # Достаем df из предыдущего прохода цикла (из предыдущих файлов)
                df_certain_action = actions_dict[action_name].copy()
                # Добавляем колонки по акциям в итоговый df из нового прохода цикла по новому файлу
                df_certain_action = df_certain_action.merge(
                    tmp_df_certain_action,
                    on='Артикул продавца',
                    how='left'
                )
                # Помещаем df в соответствующий элемент словаря по данной акции
                actions_dict[action_name] = df_certain_action

    return actions_dict


# Функция сохранения excel
def save_and_format_excel(actions_dict):
    # Имя файла, куда будем сохранять df
    path_save_actions_compare_svod = (
        f"Clients/{client_name}/Actions"
        f"/Сравнение/"
    )
    # Создаем директорию, если её нет
    if not os.path.exists(path_save_actions_compare_svod):
        os.makedirs(path_save_actions_compare_svod)

    # Книга, куда будем записывать данные
    file_name = (
        f"{path_save_actions_compare_svod}"
        f"{str(date.today())}_{client_name}"
        f"_Сравнение_данных_по_акциям.xlsx"
    )
    # Книга, куда будем записывать данные
    with pd.ExcelWriter(f"{file_name}") as w:
    # Цикл по каждой акции
        for action_name, df in actions_dict.items():
            # Сокращаем название строки
            # action_name_cut = action_name[:31]
            action_name_cut = action_name.split('.')[0]
            # Удаляем лишние символы из названия акций, т.к. excel не позволяет сохранять названия листов некоторыми символами
            action_name_cut = re.sub('[.:]', '', action_name_cut)
            # Получаем df по конкретной акции
            df = actions_dict[action_name].copy()
            # Сохраняем excel
            df.to_excel(w, sheet_name=action_name_cut, index=False)

    # GPT START ----
    # Форматируем сохраненный файл
    wb = openpyxl.load_workbook(file_name)
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        # 1. Добавить фильтр на всю строку заголовков (первая строка)
        ws.auto_filter.ref = ws.dimensions

        # 2. Автоподбор ширины столбца
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) # * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width

        # 3. Добавить тонкую границу ко всем ячейкам
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border

    wb.save(file_name)
    # GPT END ----


# %% Вызов всех функций
if __name__ == '__main__':
    # Диапазон дат, за которые нужно сравнить данные по акциям
    # (будут считаны все файлы в промежутке между данными датами, включая данные даты)
    date_start = '2025-06-02'
    date_end = str(date.today())
    logger.info(
        f"Comparing action files for client {client_name} "
        f"for dates {date_start} - {date_end}"
    )
    # Получаем список файлов по акциям
    df_filenames_actions = read_action_files(
        client_name,
        date_start,
        date_end
    )
    # Получаем список текущих товаров и цен по АПИ
    df_products = get_ozon_product(headers, to_save=False)

    # Чтение справочной таблицы
    catalog = read_catalog(client_name)
    # Добавляем информацию из справочной таблицы в список товаров
    df_products_with_catalog_and_prices = add_catalog_data_to_product_list(
        df_products,
        catalog
    )
    # Создаем словарь сравнения данных по каждой из акций
    actions_dict = create_action_compare_svod(
        df_products_with_catalog_and_prices,
        df_filenames_actions
    )
    # Сохраняем и форматируем файл excel
    save_and_format_excel(actions_dict)
    logger.info(
        f"Done comparing action files"
    )
