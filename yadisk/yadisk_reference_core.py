from __future__ import annotations

import os
from datetime import date
from pathlib import Path
from typing import Literal, Optional, Any

import requests
import traceback
import pandas as pd
from dotenv import load_dotenv

OnError = Literal["raise", "use_local"]
ReturnMode = Literal["file", "dataframe"]

API_ROOT = "https://cloud-api.yandex.net/v1"
DISK_RESOURCE_URL = f"{API_ROOT}/disk/resources"
DISK_DOWNLOAD_URL = f"{DISK_RESOURCE_URL}/download"


def get_yadisk_oauth_token(logger) -> str:
    """
    Загрузить OAuth-токен Яндекс.Диска из .env (YANDEX_DISK_OAUTH_TOKEN).

    Параметры
    ----------
    logger :
        Логгер (обычно loguru.logger или его обёртка), в который будут
        записаны сообщения об ошибках.

    Возвращает
    ----------
    str
        OAuth-токен Яндекс.Диска.

    Исключения
    ----------
    RuntimeError
        Если переменная окружения YANDEX_DISK_OAUTH_TOKEN не найдена.
    """
    load_dotenv()
    token = os.getenv("YANDEX_DISK_OAUTH_TOKEN")
    if not token:
        logger.error("Не найден YANDEX_DISK_OAUTH_TOKEN в .env")
        raise RuntimeError("YANDEX_DISK_OAUTH_TOKEN not set")
    return token


def normalize_marketplace_name(marketplace_name: str) -> str:
    """
    Нормализовать название маркетплейса.

    Правила:
    - 'ozon' (в любом регистре) -> 'Ozon'
    - 'wb'   (в любом регистре) -> 'WB'
    - иначе                       -> нижний регистр исходного имени

    Параметры
    ----------
    marketplace_name : str
        Исходное название маркетплейса.

    Возвращает
    ----------
    str
        Нормализованное название маркетплейса.
    """
    mp_lower = marketplace_name.lower()
    if mp_lower == "ozon":
        return "Ozon"
    if mp_lower == "wb":
        return "WB"
    return mp_lower


def list_yadisk_folder(path: str, token: str, logger) -> list[dict[str, Any]]:
    """
    Получить список элементов (файлы/папки) в указанной папке на Яндекс.Диске.

    Параметры
    ----------
    path : str
        Путь к папке на Яндекс.Диске. Может быть:
        - 'СПРАВОЧНИКИ_КОНСАЛТИНГ'
        - '/СПРАВОЧНИКИ_КОНСАЛТИНГ'
        - 'disk:/СПРАВОЧНИКИ_КОНСАЛТИНГ'
    token : str
        OAuth-токен Яндекс.Диска.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    list[dict]
        Список элементов (словари), содержащих информацию о файлах/папках.

    Исключения
    ----------
    requests.HTTPError
        Если запрос к API Яндекс.Диска завершился ошибкой.
    """
    headers = {"Authorization": f"OAuth {token}"}
    params = {"path": path, "limit": 1000}

    try:
        resp = requests.get(DISK_RESOURCE_URL, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
    except requests.HTTPError as e:
        logger.error(
            "Ошибка при получении списка файлов с Яндекс.Диска: "
            "url='{}', status_code={}, response='{}'",
            e.request.url if e.request else DISK_RESOURCE_URL,
            e.response.status_code if e.response else "N/A",
            (e.response.text[:1000] if e.response and e.response.text else "NO BODY"),
        )
        raise

    return resp.json().get("_embedded", {}).get("items", [])


def get_download_url(path: str, token: str, logger) -> str:
    """
    Получить прямую ссылку (href) для скачивания файла с Яндекс.Диска.

    Параметры
    ----------
    path : str
        Полный путь к файлу на Яндекс.Диске (например: 'disk:/folder/file.xlsx').
    token : str
        OAuth-токен Яндекс.Диска.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    str
        URL (href), по которому можно скачать файл.

    Исключения
    ----------
    requests.HTTPError
        Если запрос к API Яндекс.Диска завершился ошибкой.
    """
    headers = {"Authorization": f"OAuth {token}"}
    params = {"path": path}

    try:
        resp = requests.get(DISK_DOWNLOAD_URL, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
    except requests.HTTPError as e:
        logger.error(
            "Ошибка при получении ссылки для скачивания с Яндекс.Диска: "
            "url='{}', status_code={}, response='{}'",
            e.request.url if e.request else DISK_DOWNLOAD_URL,
            e.response.status_code if e.response else "N/A",
            (e.response.text[:1000] if e.response and e.response.text else "NO BODY"),
        )
        raise

    return resp.json()["href"]


def download_file_from_yadisk(href: str, local_path: Path, logger) -> None:
    """
    Скачать файл с Яндекс.Диска по прямой ссылке и сохранить на локальный диск.

    Параметры
    ----------
    href : str
        Прямая ссылка для скачивания (href), полученная от API Яндекс.Диска.
    local_path : pathlib.Path
        Локальный путь, по которому будет сохранён файл.
    logger :
        Логгер для записи сообщений.

    Исключения
    ----------
    requests.HTTPError
        Если запрос на скачивание завершился ошибкой.
    IOError / OSError
        Если возникли ошибки при записи файла на диск.
    """
    local_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with requests.get(href, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(local_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
    except requests.HTTPError as e:
        logger.error(
            "Ошибка при скачивании файла с Яндекс.Диска: "
            "url='{}', status_code={}, response='{}'",
            e.request.url if e.request else href,
            e.response.status_code if e.response else "N/A",
            (e.response.text[:1000] if e.response and e.response.text else "NO BODY"),
        )
        raise
    except OSError as e:
        logger.error(
            "Ошибка записи файла на диск при скачивании с Яндекс.Диска: "
            "path='{}', error='{}'",
            local_path,
            e,
        )
        raise


def read_excel_from_yadisk_to_df(href: str, logger) -> pd.DataFrame:
    """
    Считать Excel-файл с Яндекс.Диска напрямую в pandas.DataFrame,
    не сохраняя его на локальный диск.

    Параметры
    ----------
    href : str
        Прямая ссылка для скачивания (href), полученная от API Яндекс.Диска.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    pandas.DataFrame
        DataFrame, считанный из Excel-файла.

    Исключения
    ----------
    requests.HTTPError
        Если запрос на скачивание завершился ошибкой.
    pandas.errors.ExcelFileError / ValueError
        Если файл нельзя корректно прочитать как Excel.
    """
    from io import BytesIO

    try:
        with requests.get(href, stream=True, timeout=60) as r:
            r.raise_for_status()
            buffer = BytesIO(r.content)
            df = pd.read_excel(buffer)
        return df
    except requests.HTTPError as e:
        logger.error(
            "Ошибка при скачивании Excel с Яндекс.Диска для чтения в DataFrame: "
            "url='{}', status_code={}, response='{}'",
            e.request.url if e.request else href,
            e.response.status_code if e.response else "N/A",
            (e.response.text[:1000] if e.response and e.response.text else "NO BODY"),
        )
        raise
    except Exception as e:
        logger.error("Ошибка при чтении Excel из потока: {}", e)
        raise


def find_reference_table_in_yadisk_folder(
    folder_path: str,
    client_name: str,
    marketplace_name: str,
    token: str,
    logger,
) -> Optional[str]:
    """
    Найти в указанной папке на Яндекс.Диске справочную таблицу клиента.

    Имя искомого файла:
    'Справочная_таблица_{client_name}_{marketplace_name}.xlsx'

    Параметры
    ----------
    folder_path : str
        Путь к папке на Яндекс.Диске, где расположены справочные таблицы.
    client_name : str
        Имя клиента.
    marketplace_name : str
        Название маркетплейса (ожидается уже нормализованное: 'Ozon', 'WB', ...).
    token : str
        OAuth-токен Яндекс.Диска.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    str | None
        Полный путь к файлу на Яндекс.Диске ('disk:/.../Справочная_таблица...xlsx'),
        либо None, если файл не найден.

    Исключения
    ----------
    requests.HTTPError
        Если запрос к API Яндекс.Диска завершился ошибкой.
    """
    expected_name = f"Справочная_таблица_{client_name}_{marketplace_name}.xlsx"
    items = list_yadisk_folder(folder_path, token, logger)

    for item in items:
        if item.get("type") == "file" and item.get("name") == expected_name:
            return item.get("path")
    return None


def get_local_reference_path(
    project_root: Path,
    client_name: str,
    marketplace_name: str,
) -> Path:
    """
    Получить путь к локальной актуальной справочной таблице клиента.

    Структура:
        {project_root}/{marketplace_name}/Clients/{client_name}/catalog/
            Справочная_таблица_{client_name}_{marketplace_name}.xlsx

    Параметры
    ----------
    project_root : pathlib.Path
        Корень проекта.
    client_name : str
        Имя клиента.
    marketplace_name : str
        Название маркетплейса (обычно нормализованное: 'Ozon', 'WB', ...).

    Возвращает
    ----------
    pathlib.Path
        Путь к файлу справочной таблицы.
    """
    # Имя директории с маркетплейсом всегда нижнем регистре!
    marketplace_dir_name = marketplace_name.lower()
    # Директория маркетплейса внутри проекта, например: {project_root}/ozon/
    marketplace_dir = project_root / marketplace_dir_name
    # Путь к справочной таблице
    catalog_dir = marketplace_dir / "Clients" / client_name / "catalog"
    # Имя файла: Справочная_таблица_{client_name}_{marketplace_name}.xlsx
    reference_name = f"Справочная_таблица_{client_name}_{marketplace_name}.xlsx"

    return catalog_dir / reference_name


def update_local_reference_table(
    project_root: Path,
    client_name: str,
    marketplace_name: str,
    downloaded_file_path: Path,
    logger,
) -> Path:
    """
    Обновить локальную справочную таблицу указанного клиента.

    Логика:
    - текущий файл:
        {project_root}/{marketplace_name}/Clients/{client_name}/catalog/
            Справочная_таблица_{client_name}_{marketplace_name}.xlsx
    - если файл уже существует:
        * переименовать в
          {date.today()}_Справочная_таблица_{client_name}_{marketplace_name}.xlsx
        * переместить в подпапку 'Предыдущие версии'
        * если в 'Предыдущих версиях' уже есть файл с таким именем — перезаписать его
    - затем перенести загруженный с Яндекс.Диска файл как актуальную справочную таблицу.

    Параметры
    ----------
    project_root : pathlib.Path
        Корень проекта.
    client_name : str
        Имя клиента.
    marketplace_name : str
        Название маркетплейса (нормализованное).
    downloaded_file_path : pathlib.Path
        Путь к временно загруженному файлу справочной таблицы.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    pathlib.Path
        Путь к актуальной локальной справочной таблице.
    """
    reference_path = get_local_reference_path(project_root, client_name, marketplace_name)
    catalog_dir = reference_path.parent
    previous_versions_dir = catalog_dir / "Предыдущие версии"
    previous_versions_dir.mkdir(parents=True, exist_ok=True)

    if reference_path.exists():
        today_str = date.today().strftime("%Y-%m-%d")
        archived_name = (
            f"{today_str}_Справочная_таблица_{client_name}_{marketplace_name}.xlsx"
        )
        archived_path = previous_versions_dir / archived_name

        if archived_path.exists():
            logger.info(
                f"В 'Предыдущих версиях' уже есть файл {archived_name}, "
                f"он будет перезаписан."
            )

        logger.info(
            f"Перемещаем текущую справочную таблицу в архив: {archived_path}"
        )
        reference_path.replace(archived_path)

    logger.info(
        f"Обновляем локальную справочную таблицу клиента '{client_name}' "
        f"для маркетплейса '{marketplace_name}' по пути: {reference_path}"
    )
    reference_path.parent.mkdir(parents=True, exist_ok=True)
    downloaded_file_path.replace(reference_path)

    return reference_path


def _handle_error_with_local_fallback(
    err: Exception,
    on_error: OnError,
    project_root: Optional[Path],
    client_name: str,
    marketplace_name: str,
    return_mode: ReturnMode,
    logger,
):
    """
    Внутренняя функция обработки ошибок с возможностью fallback на локальную таблицу.

    Параметры
    ----------
    err : Exception
        Исходное исключение.
    on_error : {"raise", "use_local"}
        Стратегия обработки ошибок:
        - "raise"     — пробросить исключение дальше;
        - "use_local" — залогировать исходную ошибку и попытаться использовать
                        локальную справочную таблицу.
    project_root : pathlib.Path | None
        Корень проекта. Требуется, если on_error="use_local".
    client_name : str
        Имя клиента.
    marketplace_name : str
        Название маркетплейса (нормализованное).
    return_mode : {"file", "dataframe"}
        Режим возвращаемого результата: путь к файлу или DataFrame.
    logger :
        Логгер для записи сообщений.

    Возвращает
    ----------
    pathlib.Path | pandas.DataFrame
        Если on_error="use_local" и локальный файл найден:
        - при return_mode="file"      → Path к локальной таблице;
        - при return_mode="dataframe" → DataFrame, считанный из локального файла.

    Исключения
    ----------
    Exception
        Если on_error="raise" или fallback на локальный файл невозможен/
        завершился неудачно, исходное исключение (err) пробрасывается.
    """
    # 1. Логируем исходную ошибку ВСЕГДА (как минимум warning)
    if isinstance(err, requests.HTTPError):
        resp = err.response
        logger.error(
            "HTTP ошибка при работе с Яндекс.Диском: "
            "url='{}', status_code={}, response='{}'",
            err.request.url if err.request else "N/A",
            resp.status_code if resp else "N/A",
            (resp.text[:1000] if resp and resp.text else "NO BODY"),
        )
    else:
        logger.error(
            "Ошибка при работе с Яндекс.Диском: {}. Traceback:\n{}",
            err,
            traceback.format_exc(),
        )

    # 2. Если стратегия "raise" — после логирования сразу пробрасываем исключение
    if on_error == "raise":
        raise err

    # 3. on_error == "use_local" — пробуем использовать локальный файл
    if project_root is None:
        logger.error(
            "on_error='use_local', но project_root не задан — "
            "невозможно найти локальную справочную таблицу. "
            "Исключение будет проброшено."
        )
        raise err

    project_root = Path(project_root)
    local_path = get_local_reference_path(project_root, client_name, marketplace_name)

    if not local_path.exists():
        logger.error(
            "on_error='use_local': локальная справочная таблица не найдена по пути "
            f"{local_path}. Исключение будет проброшено."
        )
        raise err

    logger.warning(
        "Возникла ошибка при обращении к Яндекс.Диску, но on_error='use_local'. "
        f"Будет использована локальная справочная таблица: {local_path}"
    )

    if return_mode == "file":
        return local_path
    else:
        try:
            df = pd.read_excel(local_path)
            return df
        except Exception as e:
            logger.error(
                "Не удалось прочитать локальную справочную таблицу в DataFrame: {}. "
                "Исключение исходное будет проброшено.",
                e,
            )
            raise err


def fetch_reference_table_from_yadisk(
    client_name: str,
    marketplace_name: str,
    logger,
    download_flag: bool = False,
    project_root: str | Path | None = None,
    yadisk_folder: str = "СПРАВОЧНИКИ_КОНСАЛТИНГ",
    on_error: OnError = "raise",
    return_mode: ReturnMode = "file",
):
    """
    Главная функция для получения справочной таблицы клиента с Яндекс.Диска.

    Логика:
    1) Ищет файл 'Справочная_таблица_{client_name}_{marketplace_name}.xlsx'
       в папке yadisk_folder на Яндекс.Диске.
    2) Если найден:
       - при return_mode="dataframe" и download_flag=False:
           * читает Excel напрямую в DataFrame (без сохранения на диск);
       - иначе:
           * скачивает файл во временную папку "./tmp_yadisk_reference".
           * если download_flag=True и project_root задан:
               - обновляет локальную справочную таблицу (архивирует предыдущую
                 версию и заменяет на новую).
    3) Возвращает:
       - при return_mode="file"      → Path к локальному файлу (либо временный,
                                       либо окончательный путь в структуре проекта);
       - при return_mode="dataframe" → pandas.DataFrame.

    При ошибках в зависимости от on_error:
    - "raise"     — исключения пробрасываются;
    - "use_local" — при ошибке работы с Я.Диском пытается использовать локальную
                    справочную таблицу, если project_root задан и файл существует.

    Параметры
    ----------
    client_name : str
        Имя клиента.
    marketplace_name : str
        Название маркетплейса (может быть в произвольном регистре, внутри функции
        нормализуется).
    logger :
        Логгер для записи сообщений (все логи идут через него).
    download_flag : bool, optional
        Нужно ли обновлять локальную структуру справочной таблицы в проекте.
        По умолчанию False.
    project_root : str | pathlib.Path | None, optional
        Корень проекта. Нужен, если:
        - download_flag=True (для обновления локальной таблицы);
        - on_error="use_local" (для fallback на локальный файл).
    yadisk_folder : str, optional
        Папка на Яндекс.Диске, в которой лежат справочные таблицы.
        По умолчанию "СПРАВОЧНИКИ_КОНСАЛТИНГ".
    on_error : {"raise", "use_local"}, optional
        Логика обработки ошибок работы с Яндекс.Диском.
    return_mode : {"file", "dataframe"}, optional
        Режим возвращаемого результата.

    Возвращает
    ----------
    pathlib.Path | pandas.DataFrame
        В зависимости от return_mode:
        - "file"      → Path
        - "dataframe" → pandas.DataFrame

    Исключения
    ----------
    Различные исключения (requests, pandas, RuntimeError, FileNotFoundError),
    в зависимости от ситуации и выбранного on_error.
    """
    norm_marketplace = normalize_marketplace_name(marketplace_name)

    if download_flag and project_root is None:
        logger.warning(
            "download_flag=True, но project_root не задан. "
            "Локальная структура не будет обновлена."
        )

    try:
        token = get_yadisk_oauth_token(logger)
        logger.info(
            f"Ищем справочную таблицу на Яндекс.Диске для клиента '{client_name}' "
            f"(маркетплейс: '{norm_marketplace}') в папке '{yadisk_folder}'"
        )

        yadisk_file_path = find_reference_table_in_yadisk_folder(
            yadisk_folder,
            client_name,
            norm_marketplace,
            token,
            logger,
        )

        if not yadisk_file_path:
            raise FileNotFoundError(
                f"На Яндекс.Диске в папке '{yadisk_folder}' НЕ найдена "
                f"справочная таблица 'Справочная_таблица_{client_name}_{norm_marketplace}.xlsx'."
            )

        logger.info(f"Найдена справочная таблица на Яндекс.Диске: {yadisk_file_path}")

        href = get_download_url(yadisk_file_path, token, logger)

        # Вариант 1: сразу читаем в DataFrame, без локального обновления
        if return_mode == "dataframe" and not download_flag:
            logger.info(
                "Считываем справочную таблицу напрямую в DataFrame, "
                "без сохранения на диск и без обновления локальной структуры."
            )
            df = read_excel_from_yadisk_to_df(href, logger)
            return df

        # Вариант 2: скачиваем файл на диск
        tmp_dir = Path("./tmp_yadisk_reference")
        tmp_dir.mkdir(parents=True, exist_ok=True)
        local_tmp_path = tmp_dir / Path(yadisk_file_path).name

        logger.info(f"Скачиваем справочную таблицу во временный файл: {local_tmp_path}")
        download_file_from_yadisk(href, local_tmp_path, logger)
        logger.info(f"Справочная таблица скачана: {local_tmp_path}")

        # Если локальное обновление не требуется
        if not download_flag or project_root is None:
            if return_mode == "file":
                return local_tmp_path
            else:
                df = pd.read_excel(local_tmp_path)
                return df

        # Локальное обновление
        project_root = Path(project_root)
        updated_local_path = update_local_reference_table(
            project_root=project_root,
            client_name=client_name,
            marketplace_name=norm_marketplace,
            downloaded_file_path=local_tmp_path,
            logger=logger,
        )

        logger.info(
            f"Актуальная локальная справочная таблица клиента '{client_name}' "
            f"(маркетплейс: '{norm_marketplace}') расположена по пути: {updated_local_path}"
        )

        if return_mode == "file":
            return updated_local_path
        else:
            df = pd.read_excel(updated_local_path)
            return df

    except Exception as e:
        return _handle_error_with_local_fallback(
            err=e,
            on_error=on_error,
            project_root=Path(project_root) if project_root is not None else None,
            client_name=client_name,
            marketplace_name=norm_marketplace,
            return_mode=return_mode,
            logger=logger,
        )


if __name__ == "__main__":
    """
    Пример отладочного использования модуля.

    Для реального проекта логгер и параметры (client_name, marketplace_name,
    project_root, yadisk_folder) обычно берутся из констант/конфигурации
    конкретного маркетплейса и клиента.
    """

    # from loguru import logger as debug_logger
    #
    # # Простейшая настройка логгера для отладки:
    # debug_logger.add(
    #     "yadisk_debug.log",
    #     encoding="utf-8",
    #     rotation="5 MB",
    #     retention="7 days",
    #     compression="zip",
    # )
    #
    # client = "TEST_CLIENT"
    # marketplace = "ozon"
    # project_root_example = Path.cwd()
    #
    # try:
    #     # Пример: получить справочную таблицу как DataFrame, не обновляя локальную структуру
    #     df_catalog = fetch_reference_table_from_yadisk(
    #         client_name=client,
    #         marketplace_name=marketplace,
    #         logger=debug_logger,
    #         download_flag=False,
    #         project_root=project_root_example,
    #         yadisk_folder="СПРАВОЧНИКИ_КОНСАЛТИНГ",
    #         on_error="raise",
    #         return_mode="dataframe",
    #     )
    #     debug_logger.info(f"Справочная таблица получена, shape={df_catalog.shape}")
    #
    #     # Пример: обновить локальную таблицу и получить путь к файлу
    #     updated_path = fetch_reference_table_from_yadisk(
    #         client_name=client,
    #         marketplace_name=marketplace,
    #         logger=debug_logger,
    #         download_flag=True,
    #         project_root=project_root_example,
    #         yadisk_folder="СПРАВОЧНИКИ_КОНСАЛТИНГ",
    #         on_error="use_local",
    #         return_mode="file",
    #     )
    #     debug_logger.info(f"Актуальная локальная таблица: {updated_path}")
    #
    # except Exception as exc:
    #     debug_logger.exception(f"Ошибка при работе с модулем Яндекс.Диска: {exc}")
    pass
