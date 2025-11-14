from datetime import date, datetime, timedelta

def month_date_range():
    """
    Возвращает диапазон дат (строкой), либо за текущий месяц до сегодня,
    либо за весь предыдущий месяц, если сегодня — первое число.
    """
    today = datetime.now()
    if today.day == 1:
        # Весь прошлый месяц
        prev_month_last_day = today - timedelta(days=1)
        start_dt = prev_month_last_day.replace(day=1)
        end_dt = prev_month_last_day
    else:
        # С первого числа этого месяца по сегодня
        start_dt = today.replace(day=1)
        end_dt = today - timedelta(days=1)
    start_date_str = start_dt.strftime('%Y-%m-%d')
    end_date_str = end_dt.strftime('%Y-%m-%d')
    # ISO
    # start_date_iso = start_dt.strftime('%Y-%m-%dT00:00:00Z')
    # end_date_iso = end_dt.strftime('%Y-%m-%dT23:59:59Z')
    return start_date_str, end_date_str  #, start_date_iso, end_date_iso


def prev_week_date_range():
    today = datetime.now()

    # Понедельник предыдущей недели:
    prev_mon = today - timedelta(days=7)
    start_dt = prev_mon - timedelta(days=prev_mon.weekday())

    # Воскресенье предыдущей недели:
    end_dt = start_dt + timedelta(days=6)

    start_date_str = start_dt.strftime('%Y-%m-%d')
    end_date_str = end_dt.strftime('%Y-%m-%d')
    # start_date_iso = start_dt.strftime('%Y-%m-%dT00:00:00Z')
    # end_date_iso = end_dt.strftime('%Y-%m-%dT23:59:59Z')


    return start_date_str, end_date_str


def prev_month_date_range():
    today = date.today()
    # Первый день текущего месяца
    first_day_this_month = today.replace(day=1)
    # Последний день предыдущего месяца — это вчера от первого дня этого месяца
    last_day_prev_month = first_day_this_month - timedelta(days=1)
    # Первый день предыдущего месяца: day=1, месяц/год того что у last_day_prev_month
    first_day_prev_month = last_day_prev_month.replace(day=1)
    start_date_str = first_day_prev_month.strftime('%Y-%m-%d')
    end_date_str = last_day_prev_month.strftime('%Y-%m-%d')
    return start_date_str, end_date_str


def last_n_days(days_ago=30):
    """
    Возвращает (start_date, end_date) в формате %Y-%m-%d
    За days_ago дней до вчерашнего (без включения today).
    """
    yesterday = date.today() - timedelta(days=1)
    start = yesterday - timedelta(days=days_ago - 1)
    start_str = start.strftime('%Y-%m-%d')
    end_str = yesterday.strftime('%Y-%m-%d')
    return start_str, end_str
