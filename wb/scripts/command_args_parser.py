import argparse
from datetime import datetime
from rich_argparse import RichHelpFormatter
RichHelpFormatter.styles["argparse.text"] = "bold"
RichHelpFormatter.group_name_formatter = str.upper


def cli_reference_date(parser):
    """
    Парсит reference_date как позиционный и как именованный аргумент.
    Приоритет: позиционный > именованный > сегодняшняя дата.
    """
    # parser = argparse.ArgumentParser(description="Парсер reference_date")
    parser.add_argument(
        'reference_date_pos',
        nargs='?',
        # help='Опорная дата (позиционный аргумент) в формате YYYY-MM-DD',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        '--ref_dt',
        type=str,
        help='Опорная дата в формате YYYY-MM-DD',
        metavar=''
    )
    # Jupyter fix
    args, unknown = parser.parse_known_args()

    if args.reference_date_pos:
        return args.reference_date_pos
    elif args.ref_dt:
        return args.ref_dt
    else:
        return datetime.now().strftime('%Y-%m-%d')

def cli_dates(parser):
    """
    Парсит date_start и date_end как позиционные и именованные,
    reference_date — только как именованный.
    Возвращает dict с ключами: start_date, end_date, reference_date.
    """
    parser.add_argument(
        'start_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        'end_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        '--start_date',
        type=str,
        help='Начальная дата диапазона в формате ГГГГ-ММ-ДД',
        metavar=''
    )
    parser.add_argument(
        '--end_date',
        type=str,
        help='Конечная дата диапазона в формате ГГГГ-ММ-ДД',
        metavar=''
    )
    parser.add_argument(
        '--reference_date',
        type=str,
        help='Опорная дата (например, для автоматического диапазона) в формате YYYY-MM-DD',
        metavar=''
    )
    args, unknown = parser.parse_known_args()

    # Преобразуем приоритетно:
    # (позиционные > именованные; reference_date только именованный)
    params = {}
    if args.start_date_pos and args.end_date_pos:
        params["start_date"] = args.start_date_pos
        params["end_date"] = args.end_date_pos
    elif args.start_date and args.end_date:
        params["start_date"] = args.start_date
        params["end_date"] = args.end_date
    else:
        params["start_date"] = None
        params["end_date"] = None
    params["reference_date"] = args.reference_date if args.reference_date else None
    return params

def cli_date_range(parser):
    """
    Парсит диапазон дат как позиционные и/или именованные аргументы.
    Приоритет: оба позиционных > оба именованных > None.
    Возвращает dict: {"start_date": ..., "end_date": ...}
    """
    parser.add_argument(
        'start_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        'end_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        '--start_date',
        type=str,
        help='Начальная дата диапазона в формате ГГГГ-ММ-ДД',
        metavar=''
    )
    parser.add_argument(
        '--end_date',
        type=str,
        help='Конечная дата диапазона в формате ГГГГ-ММ-ДД',
        metavar=''
    )
    args, unknown = parser.parse_known_args()

    # Приоритет: оба позиционных > оба именованных > None/None
    if args.start_date_pos and args.end_date_pos:
        return {"start_date": args.start_date_pos, "end_date": args.end_date_pos}
    elif args.start_date and args.end_date:
        return {"start_date": args.start_date, "end_date": args.end_date}
    else:
        return {"start_date": None, "end_date": None}

def cli_days_ago(parser):
    """
    Парсит days_ago как позиционный и как именованный аргумент.
    Приоритет: позиционный > именованный > 0.
    """
    parser.add_argument(
        'days_ago_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    parser.add_argument(
        '--days_ago',
        type=int,
        help='Количество дней назад для обработки',
        metavar=''
    )
    args, unknown = parser.parse_known_args()

    if args.days_ago_pos:
        return int(args.days_ago_pos)
    elif args.days_ago:
        return args.days_ago
    else:
        return 30
