import shutil
import os
from pathlib import Path

from loguru import logger

# Файл с некоторыми константами
from ozon.scripts.constants import marketplace_dir_name, client_name

# Функция создания директорий (для новых клиентов)
def create_dirs():
    # Директория для клиента
    client_dir = marketplace_dir_name / "Clients" / client_name
    # Список директорий для создания
    dir_names = [
        'UploadFiles',
        'Metrics',
        'catalog',
        'SupplySvod',
        'Actions',
        'UploadFilesPerformance',
        'FinanceReports',
        'SalesOrdersSvod',
        'CampaignsReport',
        'ClientCabinetSvod',
        'ClientRKSvod',
        'OrdersSvod',
    ]
    for dir_name in dir_names:
        dir_path = client_dir / dir_name
        if not dir_path.exists():
            logger.info(f"Creating folder {dir_path} for client {client_name}")
            dir_path.mkdir(parents=True, exist_ok=True)
    logger.info(f"Done creating folders for client {client_name}")


if __name__ == '__main__':
    create_dirs()
