from pathlib import Path
import sys
from loguru import logger
from datetime import datetime

def get_time_in_timezone(tz_name="Europe/Moscow"):
    """
    Возвращает текущее время в строке (YYYY-MM-DD HH:MM:SS) для указанной таймзоны.
    Сначала пробует zoneinfo (Python 3.9+), если нет — pytz, если нет — локальное время.
    """
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(tz_name)
        return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
    except ImportError:
        try:
            import pytz
            tz = pytz.timezone(tz_name)
            return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
        except ImportError:
            # Если ни zoneinfo, ни pytz — просто localtime
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def client_logger(marketplace_dir_path, client_name, tz_name="Europe/Moscow"):
    """
    Настраивает loguru:
    - Логи в консоль и файл logs/<marketplace>/<client_name>_<marketplace>_log.txt
    - Отдельный столбец custom_time с временем в нужной таймзоне.
    """

    marketplace_name = Path(marketplace_dir_path).name
    logs_folder = Path("logs") / marketplace_name
    logs_folder.mkdir(parents=True, exist_ok=True)
    log_file = logs_folder / f"{client_name}_{marketplace_name}_log.txt"

    logger.remove()
    # Передаём функцию, чтобы custom_time вычислялся каждый раз при вызове логгера
    patched_logger = logger.patch(
        lambda r: r["extra"].update({"custom_time": get_time_in_timezone(tz_name)})
    )
    format_str = "{time:YYYY-MM-DD HH:mm:ss} | {extra[custom_time]} | {level} | {module}:{function}:{line} | {message}"
    patched_logger.add(
        sys.stdout,
        # format=format_str
    )
    # Формат вывода в файл
    patched_logger.add(
        str(log_file),
        encoding="utf-8",
        rotation="20 MB",
        retention='7 days',
        compression="zip",
        format=format_str,
    )

    return patched_logger
