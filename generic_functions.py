import pandas as pd
import warnings
import json
import ast
from openpyxl import load_workbook
# from loguru import logger


def move_columns(df, columns_to_move, position, insert_type='after'):
    """
    Перемещает указанные колонки в DataFrame на новую позицию.

    :param df: pandas DataFrame, в котором нужно переместить колонки
    :param columns_to_move: список названий колонок или одно название (строка) для перемещения
    :param position: номер позиции или название колонки для вставки
    :param insert_type: 'before' или 'after', определяющее, куда вставлять колонки
    :return: DataFrame с перемещёнными колонками
    """

    existing_columns = df.columns.tolist()

    # Приводим columns_to_move к списку (на случай, если передали строку)
    if isinstance(columns_to_move, str):
        columns_list = [columns_to_move]
    else:
        columns_list = list(columns_to_move)  # поддержка для любого итерируемого

    # Определяем, какие из указанных колонок реально существуют в DataFrame
    found_columns = [col for col in columns_list if col in existing_columns]
    not_found_columns = [col for col in columns_list if col not in existing_columns]

    # Если некоторые колонки не найдены — предупреждаем об этом
    if not_found_columns:
        print(f"\nВ DataFrame отсутствуют следующие колонки:\n  {not_found_columns}")

    # Если не найдено ни одной колонки для перемещения — возвращаем копию исходного DataFrame
    if not found_columns:
        return df.copy()

    # Формируем список столбцов, которые остаются на местах
    remaining_columns = [col for col in existing_columns if col not in found_columns]

    # --- ЛОГИКА ОПРЕДЕЛЕНИЯ ПОЗИЦИИ ВСТАВКИ ---

    if isinstance(position, int):
        # Пользователь хочет вставить колонки в позицию с заданным индексом
        # Проверяем, что индекс допустимый
        if position < 0 or position > len(remaining_columns):
            raise ValueError("Позиция должна быть в пределах от 0 до {}".format(len(remaining_columns)))
        # Формируем новый порядок: (до позиции) + (перемещаемые колонки) + (после позиции)
        new_order = remaining_columns[:position] + found_columns + remaining_columns[position:]

    elif isinstance(position, str) and position in existing_columns:
        # Пользователь указал имя колонки для вставки до или после
        # Может быть ситуация, что эта колонка попадает в список перемещаемых (в таком случае она уже не в remaining_columns)
        if position not in remaining_columns:
            # Если колонка-ориентир перемещения совпадает с одной из перемещаемых колонок,
            # вставляем найденные колонки в начало списка (позиция = 0)
            pos_index = 0
        else:
            # Индекс колонки-ориентира в списке оставшихся
            pos_index = remaining_columns.index(position)

        if insert_type == 'before':
            # Вставляем перемещаемые колонки перед найденным индексом
            new_order = remaining_columns[:pos_index] + found_columns + remaining_columns[pos_index:]
        elif insert_type == 'after':
            # Вставляем после найденного индекса
            new_order = remaining_columns[:pos_index + 1] + found_columns + remaining_columns[pos_index + 1:]
        else:
            # На случай ошибки в параметрах
            raise ValueError("insert_type должен быть 'before' или 'after'")
    else:
        # Если позиция указана неправильно (не число и не существующее имя колонки)
        raise ValueError("position должен быть либо индексом, либо названием существующей колонки")

    # Возвращаем DataFrame с новым порядком столбцов, основная логика закончена
    return df[new_order]



def add_element_to_list(lst, new_value, search, after=True, by_index=False):
    """
    Вставляет новый элемент (или элементы) до или после указанного значения или индекса.

    :param lst: исходный список (list)
    :param search: значение или индекс, относительно которого будет вставка
    :param new_value: значение или список значений для вставки
    :param after: если True — после; если False — до
    :param by_index: если True — поиск по индексу, если False — по значению
    :return: новый список
    """
    # Определяем значения для вставки
    if isinstance(new_value, list) and not isinstance(new_value, str):
        values_to_insert = new_value
    else:
        values_to_insert = [new_value]

    # Работа с пустым списком
    if not lst:
        return values_to_insert

    if by_index:
        idx = search
        if not isinstance(idx, int) or not (-len(lst) <= idx < len(lst)):
            print(f"Индекс {idx} вне диапазона списка.")
            return lst
        # Для отрицательных индексов Python сам корректно вставит
    else:
        try:
            idx = lst.index(search)
        except ValueError:
            print(f"{search} не найден в списке.")
            return lst

    insert_at = idx + 1 if after else idx
    return lst[:insert_at] + values_to_insert + lst[insert_at:]


def _to_list_or_dict(val):
    """
    Если в ячейке строка, пытаемся преобразовать её в Python‑объект.
    При неудаче возвращаем пустой список (для list‑case) либо
    пустой словарь (для dict‑case).
    """
    if isinstance(val, str):
        try:
            # JSON‑строка
            parsed = json.loads(val)
        except Exception:
            try:
                # «Python‑строка» (например: "[{'a':1}]" или "{'a':1}")
                parsed = ast.literal_eval(val)
            except Exception:
                parsed = []          # fallback – пустой список
        return parsed
    return val

def custom_pd_explode(df,
                      column,
                      meta_columns=None,
                      drop_original=True,
                      errors="ignore"):
    """
    Разворачивает колонку, содержащую список словарей **или** одиночный словарь.

    Parameters
    ----------
    df : pandas.DataFrame
        Исходный датафрейм. Функция не модифицирует его напрямую – создаёт копию.
    column : str
        Название колонки, которую нужно «разворачивать».
    meta_columns : list | None
        Список колонок‑мета, которые «приклеить» к каждому развернутому словарю.
        Если `None`, то добавляются все оставшиеся колонки из `df`
        (за исключением колонки, которую разворачиваем).
    drop_original : bool, default True
        Удалять ли из результата исходную колонку `column`.
    errors : str, default "ignore"
        Параметр, который передаётся в `pd.json_normalize`.
        Позволяет задать, как обрабатывать ошибки при распаковке.

    Returns
    -------
    pandas.DataFrame
        Развёрнутый датафрейм с колонками из словарей (с префиксом `column.`)
        и мета‑полями.

    Notes
    -----
    1. Если в колонке находятся списки словарей – используется
       `df.explode()` + `pd.json_normalize()`.
    2. Если в колонке находятся одиночные словари – используется только
       `pd.json_normalize()`.
    3. В обоих случаях колонка `column` может быть удалена (по умолчанию).
    4. Логируются два ключевых момента:
       * выбранный режим (`EXPLODE` или `DIRECT NORMALIZE`);
       * имя колонки, которую обрабатываем.
    """
    # Создаём копию df, чтобы не менять оригинал
    df = df.copy()
 # ------------------------------------------------------------------
    # 1. Приводим все строки в колонке к Python‑объекту
    # ------------------------------------------------------------------
    df[column] = df[column].apply(_to_list_or_dict)

    # ------------------------------------------------------------------
    # 2. Определяем тип содержимого колонки
    #    Берём первое НЕ‑NaN значение, чтобы избежать ошибок
    # ------------------------------------------------------------------
    dropped = df[column].dropna()
    first_val = dropped.iloc[0] if not dropped.empty else None

    # Если первое значение – список → нужен explode + normalize
    use_explode = isinstance(first_val, list)

    # Если первое значение – словарь или None → только normalize
    use_direct  = isinstance(first_val, dict) or first_val is None

    # ------------------------------------------------------------------
    # 3. Выбираем режим распаковки и логируем
    # ------------------------------------------------------------------
    if use_explode:
        print(f"[custom_pd_explode] Using EXPLODE mode for column '{column}'.")
    else:
        print(f"[custom_pd_explode] Using DIRECT NORMALIZE mode for column '{column}'.")

    # ------------------------------------------------------------------
    # 4. Реализация выбранного режима
    # ------------------------------------------------------------------
    if use_explode:
        # 4.1 explode – каждая строка списка превращается в отдельную строку
        df_exp = df.explode(column).reset_index(drop=True)

        # 4.2 normalize – раскладываем словарь в отдельные колонки
        prod_df = pd.json_normalize(
            df_exp[column],
            errors=errors,
        )
        prod_df = prod_df.add_prefix(column + ".")

        # 4.3 удаляем исходную колонку, если требуется
        if drop_original:
            df_exp = df_exp.drop(columns=column)

        # 4.4 meta‑поля: если не заданы – берём все остальные колонки
        if meta_columns is None:
            df_result = pd.concat([df_exp, prod_df], axis=1)
        else:
            meta_df = df_exp[meta_columns]
            df_result = pd.concat([df_exp.drop(columns=meta_columns), prod_df, meta_df], axis=1)

    else:  # DIRECT NORMALIZE
        # 4.1 не делаем explode – каждый словарь остаётся в той же строке
        prod_df = pd.json_normalize(
            df[column],
            errors=errors,
        )
        prod_df = prod_df.add_prefix(column + ".")

        if drop_original:
            # 4.2 Удаляем колонку `column` и собираем результат
            df_no_col = df.drop(columns=[column])

            # Если мета‑поля не заданы – берём всё, что осталось
            if meta_columns is None:
                meta_df = df_no_col
            else:
                meta_df = df_no_col[meta_columns]

            # Формируем итоговый DataFrame
            df_result = pd.concat([meta_df, prod_df], axis=1)
        else:
            # Если не удаляем колонку – включаем её в итоговый DataFrame
            if meta_columns is None:
                meta_df = df
            else:
                meta_df = df[meta_columns]
            df_result = pd.concat([df.drop(columns=[column]), prod_df, meta_df], axis=1)

    # ------------------------------------------------------------------
    # 5. Возвращаем результат без индекса (как в оригинале)
    # ------------------------------------------------------------------
    return df_result.reset_index(drop=True)


def get_columns_info(path_to_file: str, sheet_name: str = None) -> dict:
    """
    Возвращает словарь вида:
    {
        "Заголовок1": {
            "length": длина_текста_заголовка,
            "width": ширина_столбца (может быть None),
            "alignment": {
                "horizontal": тип_горизонтального_выравнивания (str или None),
                "vertical": тип_вертикального_выравнивания (str или None),
            }
        },
        ...
    }
    """
    wb = load_workbook(path_to_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    result = {}
    header_row = ws[1]  # первая строка — заголовки

    for cell in header_row:
        header_value = str(cell.value) if cell.value is not None else ""
        col_letter = cell.column_letter

        # Ширина столбца
        col_dim = ws.column_dimensions[col_letter]
        width = col_dim.width  # может быть None

        # Ячейка во второй строке того же столбца
        data_cell = ws[f"{col_letter}2"]

        # Выравнивание (может быть None, тогда атрибуты будут None)
        alignment = data_cell.alignment
        horizontal = alignment.horizontal if alignment is not None else None
        vertical = alignment.vertical if alignment is not None else None

        result[header_value] = {
            "length": len(header_value),
            "width": width,
            "alignment": {
                "horizontal": horizontal,
                "vertical": vertical,
            },
        }

    return result
