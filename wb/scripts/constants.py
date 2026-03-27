from pathlib import Path

# Файл с настройками и номером клиента
from options import settings, client_number

# ID клиента и АПИ-ключ для заголовков запросов к Seller API
api_key = settings['api_key_wb'][client_number]

# Заголовки запросов к API
headers = {
"Authorization": api_key,
"Content-Type": "application/json",
}

# Путь к текущему скрипту
script_path = Path(__file__).resolve()
# Путь к директории с маркетплейсом
marketplace_dir_path = script_path.parents[1]  # ozon или wb
marketplace_name = 'WB'

# Имя клиента
client_name = settings['client_name'][client_number]

# Маппинг для типа магазина
shop_type_map = {
    'SENS': 'clothes',
    'SENS_IP': 'clothes',
    'Orsk_Combinat': 'food',
    'Neva_Metall': 'dishes'
}
shop_type = shop_type_map.get(client_name, 'unassigned')

# Маппинг для типа коробов
cargo_type_map = {
    'New Client': 'monopallets'
}
cargo_type = cargo_type_map.get(client_name, 'boxes')

# Маппинг для крупногабаритного товара
heavy_cargo_clients = {'New Client'}
is_heavy_cargo = client_name in heavy_cargo_clients

# Маппинг для виртуальных складов
ignore_virtual_warehouses_clients = {'New Client'}
ignore_virtual_warehouses = False if client_name in ignore_virtual_warehouses_clients else True

# Периоды потребности по клиентам (в днях)
client_demand_periods_map = {
    # Клиент: (период 1, период 2)
    "kadiev":      (90, 120),
}

# Дефолтные периоды, если клиента нет в маппинге
default_demand_periods = (40, 60)

DEMAND_PERIOD_1, DEMAND_PERIOD_2 = client_demand_periods_map.get(
    client_name,
    default_demand_periods,
)

# Маппинг кратности коробов
cargo_multiplicity_map = {
    'kadiev': {
        'flag_multiplicity': False,
        'multiplicity': 2
    }
}
cargo_multiplicity = cargo_multiplicity_map.get(
    client_name,
    {'flag_multiplicity': False, 'multiplicity': 1}
)

# Столбцы, которые берем из каталога для расчета поставок
catalog_supply_columns_map = {
    'SENS_IP': ['Артикул_Размер', 'РРЦ', 'Статус', 'Кратность коробов'],
    'kadiev': ['Артикул_Размер', 'РРЦ', 'Статус', 'Кратность коробов'],

}

# Дефолт, если клиент не найден
default_catalog_supply_columns = ['Артикул_Размер', 'РРЦ', 'Статус', 'Кратность коробов']
catalog_supply_columns = catalog_supply_columns_map.get(client_name, default_catalog_supply_columns)

# Периоды потребности по клиентам (в днях)
client_demand_periods_map = {
    'kadiev': (90, 120),

}

# Дефолтные периоды, если клиент не найден
default_demand_periods = (40, 60)

DEMAND_PERIOD_1, DEMAND_PERIOD_2 = client_demand_periods_map.get(
    client_name,
    default_demand_periods,
)

# ======================================================================
# Столбцы для расчета поставок на листе "Всего"
# ======================================================================

supply_svod_total_columns_map = {
    'SENS': [
        '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
        'Наименование товара', 'Предмет', 'Штрихкод', 'РРЦ',
        'Заказы', 'Продажи', 'Остатки', 'Остатки_fbs', 'Ожидаемое поступление',
        'Оборачиваемость',
        f'Потребность на {DEMAND_PERIOD_1} дней', f'Потребность на {DEMAND_PERIOD_2} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней с учетом FBS',
        f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
        'Кратность коробов',
        f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
        f'Потребность на {DEMAND_PERIOD_1} дней с учетом FBS (округл.)',
        f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
    ],
    'SENS_IP': [
        '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
        'Наименование товара', 'Предмет', 'Штрихкод', 'РРЦ',
        'Заказы', 'Продажи', 'Остатки', 'Остатки_fbs', 'Ожидаемое поступление',
        'Оборачиваемость',
        f'Потребность на {DEMAND_PERIOD_1} дней', f'Потребность на {DEMAND_PERIOD_2} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней с учетом FBS',
        f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
        'Кратность коробов',
        f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
        f'Потребность на {DEMAND_PERIOD_1} дней с учетом FBS (округл.)',
        f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
    ],
}

# Дефолтный набор, если клиент не найден:
default_supply_svod_total_columns = [
    '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
    'Наименование товара', 'Предмет', 'Штрихкод', 'РРЦ', 'Статус',
    'Заказы', 'Продажи', 'Остатки', 'Остатки_fbs', 'Ожидаемое поступление',
    'Оборачиваемость',
    f'Потребность на {DEMAND_PERIOD_1} дней', f'Потребность на {DEMAND_PERIOD_2} дней',
    f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
    f'Дефицит/Избыток {DEMAND_PERIOD_1} дней с учетом FBS',
    f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
    'Кратность коробов',
    f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
    f'Потребность на {DEMAND_PERIOD_1} дней с учетом FBS (округл.)',
    f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
]

supply_svod_total_columns = supply_svod_total_columns_map.get(
    client_name,
    default_supply_svod_total_columns,
)

# ======================================================================
# Столбцы для расчета поставок на листе "По кластерам"
# ======================================================================

supply_svod_clusters_columns_map = {
    'SENS': [
        '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
        'Предмет', 'Наименование товара', 'Штрихкод', 'РРЦ',
        'Склад',
        'Заказы', 'Продажи', 'Остатки', 'Ожидаемое поступление',
        'Оборачиваемость',
        'Кратность коробов',
        f'Потребность на {DEMAND_PERIOD_1} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
        f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
        f'Потребность на {DEMAND_PERIOD_2} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
        f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
        'Корректировка с учетом нулевых остатков',
    ],
    'SENS_IP': [
        '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
        'Предмет', 'Наименование товара', 'Штрихкод', 'РРЦ',
        'Склад',
        'Заказы', 'Продажи', 'Остатки', 'Ожидаемое поступление',
        'Оборачиваемость',
        'Кратность коробов',
        f'Потребность на {DEMAND_PERIOD_1} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
        f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
        f'Потребность на {DEMAND_PERIOD_2} дней',
        f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
        f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
        'Корректировка с учетом нулевых остатков',
    ],
}

default_supply_svod_clusters_columns = [
    '№ товара', 'Артикул_Размер', 'Артикул продавца', 'Размер',
    'Предмет', 'Наименование товара', 'Штрихкод', 'РРЦ', 'Статус',
    'Склад',
    'Заказы', 'Продажи', 'Остатки', 'Ожидаемое поступление',
    'Оборачиваемость',
    'Кратность коробов',
    f'Потребность на {DEMAND_PERIOD_1} дней',
    f'Дефицит/Избыток {DEMAND_PERIOD_1} дней',
    f'Потребность на {DEMAND_PERIOD_1} дней (округл.)',
    f'Потребность на {DEMAND_PERIOD_2} дней',
    f'Дефицит/Избыток {DEMAND_PERIOD_2} дней',
    f'Потребность на {DEMAND_PERIOD_2} дней (округл.)',
    'Корректировка с учетом нулевых остатков',
]

supply_svod_clusters_columns = supply_svod_clusters_columns_map.get(
    client_name,
    default_supply_svod_clusters_columns,
)


# ======================================================================
# Константы по акциям
# ======================================================================

# Коэффициент себестоимости для акций
net_cost_koef_map = {
    'KU_And_KU': 0.48,
    'Soyuz': 0.48,
}
default_net_cost_koef = 0.48
net_cost_koef = net_cost_koef_map.get(client_name, default_net_cost_koef)

# Колонки, которые берем из справочной таблицы для свода по акциям
catalog_action_columns_map = {
    'SENS': [
        'Артикул_Размер',
        'Артикул продавца',
        'Размер',
        'Статус',
        'Минимальная цена расчетная, руб.',
        'РРЦ',
        'Себестоимость'
    ],
    'SENS_IP': [
        'Артикул_Размер',
        'Артикул продавца',
        'Размер',
        'Статус',
        'Минимальная цена расчетная, руб.',
        'РРЦ',
        'Себестоимость'
    ],
}

# Дефолтный вариант, если клиента нет в списке
default_catalog_action_columns = [
    'Артикул_Размер',
    'Артикул продавца',
    'Размер',
    'Статус',
    'Минимальная цена расчетная, руб.',
    'РРЦ',
    'Себестоимость'
]
catalog_action_columns = catalog_action_columns_map.get(client_name, default_catalog_action_columns)


# Нужные колонки для свода по акциям
svod_actions_cols_map = {
    'SENS': [
        'Артикул WB', 'Артикул продавца', 'Предмет', 'Наименование товара', 'Статус', 'Баркод', 'Размер', 'Цвет',
        'РРЦ', 'Цена до скидки', 'Цена после скидки', 'Скидка до РРЦ', 'Минимальная цена маржинальная, руб.',
        'Себестоимость',
        'Скидка WB',
        'Остатки', 'Остатки_fbs', 'Заказы', 'Продажи',
        'Ожидаемое количество на складе'
    ],
    'SENS_IP': [
        'Артикул WB', 'Артикул продавца', 'Предмет', 'Наименование товара', 'Статус', 'Баркод', 'Размер', 'Цвет',
        'РРЦ', 'Цена до скидки', 'Цена после скидки', 'Скидка до РРЦ', 'Минимальная цена маржинальная, руб.',
        'Себестоимость',
        'Скидка WB',
        'Остатки', 'Остатки_fbs', 'Заказы', 'Продажи',
        'Ожидаемое количество на складе'
    ],
}
# Дефолтный вариант, если клиента нет в списке
default_svod_actions_cols = [
    'Артикул WB', 'Артикул продавца', 'Предмет', 'Наименование товара', 'Баркод', 'Статус',
    'РРЦ', 'Цена до скидки', 'Цена после скидки', 'Скидка до РРЦ', 'Минимальная цена маржинальная, руб.',
    'Себестоимость',
    'Скидка WB',
    'Остатки', 'Остатки_fbs', 'Заказы', 'Продажи',
    'Ожидаемое количество на складе'
]
svod_actions_cols = svod_actions_cols_map.get(client_name, default_svod_actions_cols)


# Колонка для пересчета продаж/возвратов
finance_report_price_column_map = {
    'Orsk_Combinat': 'Цена розничная с учетом согласованной скидки',
    'Neva_Metall': 'Цена розничная с учетом согласованной скидки',
    'LABYT': 'Цена розничная с учетом согласованной скидки',
    'Meste': 'Цена розничная с учетом согласованной скидки',
    'SENS_IP': 'Цена розничная с учетом согласованной скидки',
}
default_finance_report_price_column = 'Цена розничная с учетом согласованной скидки'

finance_report_price_column = finance_report_price_column_map.get(
    client_name, default_finance_report_price_column)

# Колонки для свода по фин. отчету
catalog_finace_svod_columns_map = {
    'SENS': ['Артикул 1С', 'Статус', 'Цвет', 'Себестоимость'],
    'SENS_IP': ['Артикул 1С', 'Статус', 'Цвет', 'Себестоимость'],
    'SENS_KZ': ['Статус', 'Цвет', 'Себестоимость'],
}
default_catalog_finace_svod_columns = ['Себестоимость']

catalog_finace_svod_columns = catalog_finace_svod_columns_map.get(
    client_name, default_catalog_finace_svod_columns
)

# Коэффициент перевода валюты
k_currency_to_rub_map = {
    'SENS_KZ': (1 / 6.5)
}
k_currency_to_rub = k_currency_to_rub_map.get(client_name, 1)


# ID Рекламных кампаний для отчета РК
PROMOTION_COMPANIES = {
    'SENS_IP': [
        30330864,
        30330817,
        30330839,
        31283152,
        31283257,
        31404762,
        31404844,
        31454995,
        31545645,
        31545715,
        31596769,
        31599516,
    ]
}

# Дефолтный список артикулов для отчета по артикулам
ARTICLE_REPORT_LIST = {
    'SENS_IP': {  # Имя клиента
        30330839: {  # ID рекламной кампании
            'skus': [99767994],  # Артикулы WB для анализа
            'start_date': '2025-11-17',  # Дата начала статистики РК
            'end_date':   '2025-11-24'  # Дата окончания статистики РК
        },
    },
}
