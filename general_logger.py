from custom_loggers import make_logger

# Генеральный (общий) логгер: действия, не привязанные к клиентам Ozon/WB.
# Логи в logs/general/general_log.txt
general_logger = make_logger("general", "general_log.txt", tz_name="Europe/Moscow")