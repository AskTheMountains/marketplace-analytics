"""
Модуль для получения справочной таблицы клиента Ozon.

Обычно достаточно вызвать:
    from ozon.scripts.get_reference_table import get_reference_table

    catalog_df = get_reference_table()
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pandas as pd

from wb.scripts.wb_logger import wb_client_logger as logger
from wb.scripts.constants import (
    client_name as DEFAULT_CLIENT_NAME,
    marketplace_name as DEFAULT_MARKETPLACE_NAME,
    project_root as DEFAULT_PROJECT_ROOT,
)
from yadisk.yadisk_reference_core import fetch_reference_table_from_yadisk

OnError = Literal["raise", "use_local"]
ReturnMode = Literal["file", "dataframe"]


def get_reference_table(
    client_name: str | None = None,
    marketplace_name: str | None = None,
    download_flag: bool = False,
    project_root: str | Path | None = None,
    yadisk_folder: str = "СПРАВОЧНИКИ_КОНСАЛТИНГ",
    on_error: OnError = "use_local",
    return_mode: ReturnMode = "dataframe",
) -> pd.DataFrame | Path:
    """
    Обёртка над fetch_reference_table_from_yadisk для WB.

    Аргументы по названию и смыслу совпадают с fetch_reference_table_from_yadisk,
    за исключением logger, который здесь всегда равен wb_client_logger.

    По умолчанию:
    - client_name, marketplace_name, project_root берутся из wb.scripts.constants;
    - download_flag=False (локальную структуру не обновляем);
    - return_mode="dataframe" (возвращаем DataFrame);
    - on_error="raise" (при ошибке Я.Диска выбрасываем исключение).

    Параметры
    ----------
    client_name : str | None, optional
        Имя клиента. Если None, используется DEFAULT_CLIENT_NAME.
    marketplace_name : str | None, optional
        Название маркетплейса. Если None, используется DEFAULT_MARKETPLACE_NAME.
    download_flag : bool, optional
        Нужно ли обновлять локальную структуру справочной таблицы
        (архивация старой версии и запись новой). По умолчанию False.
    project_root : str | pathlib.Path | None, optional
        Корень проекта. Если None, используется DEFAULT_PROJECT_ROOT.
        Нужен при:
        - download_flag=True (обновление локальной таблицы);
        - on_error="use_local" (fallback на локальную таблицу).
    yadisk_folder : str, optional
        Папка на Яндекс.Диске, где лежат справочные таблицы.
        По умолчанию "СПРАВОЧНИКИ_КОНСАЛТИНГ".
    on_error : {"raise", "use_local"}, optional
        Логика обработки ошибок:
        - "raise"     — пробросить исключение;
        - "use_local" — при ошибке попытаться использовать локальную таблицу.
    return_mode : {"file", "dataframe"}, optional
        Что вернуть:
        - "dataframe" (по умолчанию) — pandas.DataFrame;
        - "file"                     — pathlib.Path к локальному файлу.

    Возвращает
    ----------
    pandas.DataFrame | pathlib.Path
        В зависимости от return_mode:
        - "dataframe" → DataFrame
        - "file"      → Path

    Исключения
    ----------
    Все, которые может выбросить fetch_reference_table_from_yadisk,
    в зависимости от выбранного on_error.
    """
    client_name = client_name or DEFAULT_CLIENT_NAME
    marketplace_name = marketplace_name or DEFAULT_MARKETPLACE_NAME
    project_root = project_root or DEFAULT_PROJECT_ROOT

    logger.info(
        f"Запрос справочной таблицы для клиента '{client_name}' "
        f"и маркетплейса '{marketplace_name}' из папки '{yadisk_folder}' на Яндекс.Диске. "
        f"download_flag={download_flag}, return_mode={return_mode}, on_error='{on_error}'"
    )

    result = fetch_reference_table_from_yadisk(
        client_name=client_name,
        marketplace_name=marketplace_name,
        logger=logger,
        download_flag=download_flag,
        project_root=project_root,
        yadisk_folder=yadisk_folder,
        on_error=on_error,
        return_mode=return_mode,
    )

    return result


if __name__ == "__main__":
    # Пример отладочного использования модуля.
    #
    # Обычно модуль используется через:
    #   from ozon.scripts.get_reference_table import get_reference_table
    #   df = get_reference_table()
    #
    # Здесь можно быстро проверить работоспособность.

    # from loguru import logger as debug_logger
    # debug_logger.add("ozon_get_reference_table_debug.log", encoding="utf-8")
    #
    # try:
    #     # 1. Получить справочник как DataFrame без обновления локальной структуры
    #     df = get_reference_table(return_mode="dataframe")
    #     debug_logger.info(f"Тестовая справочная таблица получена, shape={df.shape}")
    #
    #     # 2. Обновить локальную таблицу и получить путь к файлу
    #     path = get_reference_table(
    #         download_flag=True,
    #         return_mode="file",
    #         on_error="use_local",
    #     )
    #     debug_logger.info(f"Актуальная локальная таблица: {path}")
    #
    # except Exception as exc:
    #     debug_logger.exception(f"Ошибка при тестовом получении справочной таблицы: {exc}")
    pass
