import argparse
import shutil
from datetime import datetime
from pathlib import Path
from rich_argparse import RichHelpFormatter
from rich_argparse.contrib import ParagraphRichHelpFormatter
# from loguru import logger
from ozon.scripts.ozon_logger import ozon_client_logger as logger

# Файл с некоторыми константами
from ozon.scripts.constants import client_name, marketplace_dir_path, marketplace_name

# Аргументы командной строки
from ozon.scripts.command_args_parser import cli_update_reference_table

parser = argparse.ArgumentParser(
    description="Обновить справочную таблицу для выбранного клиента.",
    usage=(
        "python.exe -m ozon.scripts.update_reference_table "
        "-c N --file PATH_TO_NEW_TABLE.xlsx"
    ),
    epilog=(
        "Примеры:\n\n"
        "  python.exe -m ozon.scripts.update_reference_table -c 2 --file \"C:\\Users\\user\\Downloads\\Обновлённая справочная таблица.xlsx\"\n\n"
        "  python.exe -m ozon.scripts.update_reference_table --client 1 --file ./таблица_2025.xlsx\n\n"
        "Если путь или имя файла содержит пробелы, обязательно заключайте его в кавычки!\n"
        "(например: --file \"file name with spaces.xlsx\")"
    ),
    formatter_class=ParagraphRichHelpFormatter
)


def update_reference_table(marketplace_dir_path, client_name, new_file_path):
    try:
        marketplace_dir = Path(marketplace_dir_path)
        marketplace_folder = marketplace_dir.name.lower()

        catalog_dir = marketplace_dir / "Clients" / client_name / "catalog"
        previous_versions_dir = catalog_dir / "Предыдущие версии"
        previous_versions_dir.mkdir(parents=True, exist_ok=True)

        reference_name = f"Справочная_таблица_{client_name}_{marketplace_name}.xlsx"
        reference_path = catalog_dir / reference_name

        if reference_path.exists():
            last_mod_time = datetime.fromtimestamp(reference_path.stat().st_mtime)
            date_last_changed = last_mod_time.strftime("%Y-%m-%d_%H-%M-%S")
            archived_name = f"{date_last_changed}_Справочная_таблица_{client_name}_{marketplace_name}.xlsx"
            archived_path = previous_versions_dir / archived_name
            shutil.move(str(reference_path), str(archived_path))

        shutil.copy2(str(new_file_path), str(reference_path))

        logger.info(f"Справочная таблица обновлена для {client_name} на маркетплейсе {marketplace_name}. Новая версия размещена в {reference_path}")
    except Exception as e:
        logger.error(f"Ошибка обновления справочной таблицы для {client_name} на маркетплейсе {marketplace_name}: {e}")


if __name__ == '__main__':

    # Парсим аргументы командной строки
    cli_args = cli_update_reference_table(parser)
    # Получаем путь к новой справочной таблице
    new_file_path = cli_args.get('file_path', '')
    # Обновляем справочную таблицу
    update_reference_table(marketplace_dir_path, client_name, new_file_path)
