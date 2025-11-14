from custom_loggers import client_logger
# from loguru import logger
from ozon.scripts.constants import client_name, marketplace_dir_path

# Создаем объект логгера
ozon_client_logger = client_logger(marketplace_dir_path, client_name, tz_name="Europe/Moscow")
