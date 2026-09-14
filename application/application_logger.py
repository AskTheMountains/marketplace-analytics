from custom_loggers import make_logger

# Изолированный логгер приложения: логи в logs/application/application_log.txt
application_logger = make_logger("application", "application_log.txt", tz_name="Europe/Moscow")