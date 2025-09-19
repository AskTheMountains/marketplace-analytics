import argparse
from datetime import datetime

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
