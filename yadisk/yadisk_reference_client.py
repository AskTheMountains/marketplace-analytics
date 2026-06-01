from typing import Any, Callable

from yadisk.yadisk_reference_core import fetch_reference_table_from_yadisk as _fetch_core


def make_yadisk_client(logger) -> Callable[..., Any]:
    """
    Создаёт «клиента» Яндекс.Диска с уже зашитым логгером.
    Возвращаемая функция имеет ту же сигнатуру, что fetch_reference_table_from_yadisk,
    но без аргумента logger.
    """

    def fetch_reference_table_from_yadisk_for_client(
        client_name: str,
        marketplace_name: str,
        download_flag: bool = False,
        project_root=None,
        yadisk_folder: str = "СПРАВОЧНИКИ_КОНСАЛТИНГ",
        on_error: str = "raise",
        return_mode: str = "file",
    ):
        return _fetch_core(
            client_name=client_name,
            marketplace_name=marketplace_name,
            logger=logger,  # ← логгер зашит здесь
            download_flag=download_flag,
            project_root=project_root,
            yadisk_folder=yadisk_folder,
            on_error=on_error,
            return_mode=return_mode,
        )

    return fetch_reference_table_from_yadisk_for_client
