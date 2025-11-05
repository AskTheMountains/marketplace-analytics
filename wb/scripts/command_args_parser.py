import argparse
import re

from pathlib import Path
from datetime import datetime
from rich_argparse import RichHelpFormatter
RichHelpFormatter.styles["argparse.text"] = "bold"
RichHelpFormatter.group_name_formatter = str.upper

# Функция безопасного добавления аргументов командной строки
def safe_add_argument(parser, *args, **kwargs):
    # Ключ по приоритету: или именованный параметр, или dest (для позиционных)
    option_strings = args
    dest = kwargs.get('dest')
    for action in parser._actions:
        # Проверка именованных
        if option_strings and any(opt in action.option_strings for opt in option_strings):
            return
        # Проверка позиционных по 'dest'
        if dest and getattr(action, 'dest', None) == dest:
            return
        # Проверка позиционного по имени первого позиционного параметра
        if not option_strings and args and getattr(action, 'dest', None) == args[0]:
            return
    parser.add_argument(*args, **kwargs)

def clean_jupyter_kernel_args(args):
    """
    Пробегает по всем полям argparse.Namespace и удаляет или обнуляет те,
    что содержат kernel-xxx.json (аргументы от Jupyter).
    Возвращает новый Namespace с очищенными полями.
    """
    kernel_pat = re.compile(r'kernel-[\w\d\-]+\.json$', re.I)
    namespace = vars(args).copy()
    for attr, val in namespace.items():
        if isinstance(val, str):
            if kernel_pat.match(Path(val).name):
                namespace[attr] = None
        elif isinstance(val, list):
            namespace[attr] = [
                v for v in val
                if not (isinstance(v, str) and kernel_pat.match(Path(v).name))
            ]
            # Если список стал пустым — можно сделать None, если нужно:
            # if not namespace[attr]:
            #     namespace[attr] = None
    # Можно вернуть новый argparse.Namespace
    from argparse import Namespace
    return Namespace(**namespace)


def cli_client_number(parser, default=None):
    """
    Добавляет аргумент командной строки --client или -c для выбора номера клиента.
    Приоритет: командная строка > default.
    """
    safe_add_argument(
        parser,
        '--client', '-c',
        type=int,
        help='Номер клиента, для которого формируется отчет (см. файл options.py)',
        metavar=''
    )
    args, unknown = parser.parse_known_args()
    if hasattr(args, 'client') and args.client is not None:
        return args.client
    return default  # Можно передавать client_number дефолтный из options.py

def cli_reference_date(parser):
    """
    Парсит reference_date как позиционный и как именованный аргумент.
    Приоритет: позиционный > именованный > сегодняшняя дата.
    """
    safe_add_argument(
        parser,
        "--client", "-c",
        type=int,
        help="Номер клиента (необязательный, если не передан — берётся из options.py)",
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        'reference_date_pos',
        nargs='?',
        # help='Опорная дата (позиционный аргумент) в формате YYYY-MM-DD',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        '--ref_dt',
        type=str,
        help='Опорная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — берется сегодняшняя дата)',
        metavar='YYYY-MM-DD'
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
    safe_add_argument(
        parser,
        "--client", "-c",
        type=int,
        help="Номер клиента (необязательный, если не передан — берётся из options.py)",
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        'start_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        'end_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        '--start_date',
        type=str,
        help='Начальная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
    )
    safe_add_argument(
        parser,
        '--end_date',
        type=str,
        help='Конечная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
    )
    safe_add_argument(
        parser,
        '--reference_date',
        type=str,
        help='Опорная дата (например, для автоматического диапазона) в формате YYYY-MM-DD (опционально)',
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

    safe_add_argument(
        parser,
        "--client", "-c",
        type=int,
        help="Номер клиента (необязательный, если не передан — берётся из options.py)",
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        'start_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        'end_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        '--start_date',
        type=str,
        help='Начальная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
    )
    safe_add_argument(
        parser,
        '--end_date',
        type=str,
        help='Конечная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
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
    Парсит days_ago (позиционный и именованный) ИЛИ start_date и end_date.
    Приоритет: позиционный days_ago > --days_ago > даты > 30 дней по умолчанию.
    """
    import argparse

    safe_add_argument(
        parser,
        "--client", "-c",
        type=int,
        help="Номер клиента (необязательный, если не передан — берётся из options.py)",
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        'days_ago_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        '--days_ago', '-d',
        type=int,
        help='Количество дней назад для обработки (необязательный, если не передан — 30)',
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        '--start_date',
        type=str,
        help='Дата начала периода (формат YYYY-MM-DD)',
        metavar='[YYYY-MM-DD]'
    )
    safe_add_argument(
        parser,
        '--end_date',
        type=str,
        help='Дата конца периода (формат YYYY-MM-DD)',
        metavar='[YYYY-MM-DD]'
    )

    args, unknown = parser.parse_known_args()

    # Приоритет: позиционный days_ago > --days_ago > даты > 30 дней
    if args.days_ago_pos:
        return int(args.days_ago_pos), None, None
    elif args.days_ago:
        return args.days_ago, None, None
    elif args.start_date or args.end_date:
        # Вернем None для days_ago, но либо дату-строку (или None), если указана хотя бы одна дата
        return None, args.start_date, args.end_date
    else:
        # По умолчанию 30
        return 30, None, None

def cli_companies(parser):
    """
    Парсит date_start и date_end как позиционные и именованные,
    reference_date — только как именованный.
    Возвращает dict с ключами: start_date, end_date, reference_date.
    """
    safe_add_argument(
        parser,
        "--client", "-c",
        type=int,
        help="Номер клиента (необязательный, если не передан — берётся из options.py)",
        metavar='[N]'
    )
    safe_add_argument(
        parser,
        'start_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        'end_date_pos',
        nargs='?',
        help=argparse.SUPPRESS
    )
    safe_add_argument(
        parser,
        '--start_date',
        type=str,
        help='Начальная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
    )
    safe_add_argument(
        parser,
        '--end_date',
        type=str,
        help='Конечная дата в формате ГГГГ-ММ-ДД (необязательный, если не передан — генерируется автоматически)',
        metavar='YYYY-MM-DD'
    )
    safe_add_argument(
        parser,
        '--reference_date',
        type=str,
        help='Опорная дата (например, для автоматического диапазона) в формате YYYY-MM-DD (опционально)',
        metavar=''
    )
    safe_add_argument(
    parser,
    '--promotion_companies', '-comps',
    type=str,
    help='Список ID кампаний через запятую или с новой строки. Примеры: "13687629,14962180" или "13687629\\n14962180".',
    # metavar='IDS'
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

    params["promotion_companies"] = getattr(args, 'promotion_companies', None)

    return params
