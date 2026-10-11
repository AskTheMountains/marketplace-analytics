# %% Модуль интеграции с ИИ-помощником для отчёта по рекламным кампаниям (Ozon)
# Все обращения к ИИ-помощнику вынесены в отдельный модуль,
# чтобы его можно было подключать отдельным шагом-функцией в конце create_rk_svod.py.
# При любой ошибке отчёт формируется без рекомендаций ИИ (модуль не должен ронять отчёт).
import ast
import json
import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from dotenv import load_dotenv

# Логгер клиента озон
from ozon.scripts.ozon_logger import ozon_client_logger as logger
# Имя клиента
from ozon.scripts.constants import client_name

# Базовый адрес сервиса ИИ-помощника
PORTARIUS_URL = "https://portarius.1bitai.ru"
# Модель ассистента (из примера API)
# AI_ASSISTANT_MODEL = "string"

# Файл, где хранятся id диалогов по клиентам (thread_id)
THREADS_FILE = Path(__file__).resolve().parent / ".ai_threads.json"

# Флаг использования отдельного диалога (thread_id) на клиента.
# Пока выключен — сообщение отправляется без thread_id.
# Логика создания/переиспользования диалога остаётся в коде (get_or_create_thread),
# чтобы при необходимости просто переключить флаг на True.
USE_THREADS = False

# Порог запуска РК по умолчанию, если для клиента не задан PROMOTION_ORDERS_THRESHOLD
DEFAULT_CAMPAIGN_LAUNCH_THRESHOLD = 10

# Порог ДРР для комментария по умолчанию, если для клиента не задан DRR_COMMENT_THRESHOLD (в процентах)
DEFAULT_DRR_COMMENT_THRESHOLD = 10

# Строка, которая исключается из campaign_details (сшитые товары)
ARTICLE_TYPE_TO_EXCLUDE = "Товар из сшивки"

# Соответствие колонок отчёта и полей промпта для product_summary.
# Внутри кода считаем в русских колонках, в промпт отдаём английские поля.
PRODUCT_SUMMARY_COLUMN_MAP = {
    'SKU': 'sku',
    'Артикул': 'article',
    'Название товара': 'name',
    'Категория (тип)': 'category',
    'Кол-во заказов': 'orders',
    'Сумма заказов': 'revenue',
    'РК': 'ad_spend',
    'ДРР': 'drr',
    'Остаток': 'stock',
    'Дефицит': 'deficit',
    'Отмены %': 'cancellations_pct',
    'Доля РК %': 'ad_order_share_pct',
}

# Соответствие колонок отчёта и полей промпта для campaign_details (поля под воронку)
CAMPAIGN_DETAILS_COLUMN_MAP = {
    'Номер кампании': 'campaign_id',
    'SKU': 'sku',
    'Артикул': 'article',
    'Название товара': 'name',
    'Показы': 'impressions',
    'Клики': 'clicks',
    'Добавления в корзину': 'add_to_cart',
    'Заказы, шт. (РК)': 'ad_orders',
    'Заказы модели, шт. (РК)': 'ad_model_orders',
    'Расход РК, руб.': 'ad_spend',
}

# Статус товара на листе 2, для которого выводим комментарий ИИ
CAMPAIGN_TARGET_TYPE = "Товар из РК"


def _load_dotenv():
    # .env лежит в корне проекта
    project_root = Path(__file__).resolve().parents[2]
    load_dotenv(project_root / ".env")


def get_assistant_api_key():
    """
    Возвращает ключ X-Assistant-Key из переменной окружения AI_PLATTO_API_KEY
    (файл .env в корне проекта).
    Если переменной нет — возвращает None (ИИ-помощника пропускаем).
    """
    _load_dotenv()
    api_key = os.environ.get("AI_PLATTO_API_KEY")
    if not api_key:
        logger.info("AI_PLATTO_API_KEY не задан в .env — пропускаем ИИ-помощника")
        return None
    return api_key


def _load_threads():
    """Читает словарь {client_name: thread_id} из файла."""
    if not THREADS_FILE.exists():
        return {}
    try:
        with THREADS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        logger.warning(f"Не удалось прочитать файл диалогов {THREADS_FILE}: {e}")
        return {}


def _save_threads(threads):
    """Сохраняет словарь {client_name: thread_id} в файл."""
    try:
        THREADS_FILE.write_text(
            json.dumps(threads, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception as e:
        logger.warning(f"Не удалось сохранить id диалогов: {e}")


def get_or_create_thread(assistant_id, api_key):
    """
    Возвращает thread_id (id диалога) для текущего клиента.
    Если диалог для клиента уже создан — используем его,
    иначе создаём новый через метод /v1/assistants/{assistant_id}/conversation.
    """
    threads = _load_threads()
    existing = threads.get(client_name)
    if existing:
        logger.info(f"Используем существующий диалог для {client_name}: {existing}")
        return existing

    url = f"{PORTARIUS_URL}/v1/assistants/{assistant_id}/conversation"
    headers = {
        "X-Assistant-Key": api_key,
        "Content-Type": "application/json",
    }
    logger.info(f"Создаём новый диалог для {client_name}")
    resp = requests.post(url, headers=headers, json={})
    resp.raise_for_status()
    data = resp.json()
    thread_id = data.get("thread_id")
    if not thread_id:
        raise ValueError(f"В ответе на создание диалога нет thread_id: {data}")

    threads[client_name] = thread_id
    _save_threads(threads)
    logger.info(f"Создан диалог {thread_id} для {client_name}")
    return thread_id


def send_message(assistant_id, api_key, thread_id, message):
    """
    Вступает в диалог с ассистентом через метод /v1/assistants/{assistant_id}/chat.
    thread_id может быть None — тогда thread_id не отправляется.
    Возвращает ответ сервиса (dict).
    """
    url = f"{PORTARIUS_URL}/v1/assistants/{assistant_id}/chat"
    headers = {
        "X-Assistant-Key": api_key,
        "Content-Type": "application/json",
    }
    payload = {
        "message": message,
        # "stream": False,
        # "model": AI_ASSISTANT_MODEL,
        # "file_ids": [],
    }
    if thread_id:
        payload["thread_id"] = thread_id
    logger.info("Отправляем данные ИИ-помощнику")
    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    return resp.json()


# ----------------------------------------------------------------------
# Подготовка данных и разбор ответа
# ----------------------------------------------------------------------
def _to_native(value):
    """Переводит numpy-типы в обычные python-типы для корректного json.dumps.
    Не-конечные числа (nan/inf) превращает в None."""
    if isinstance(value, (np.floating, float)):
        value = float(value)
        if not np.isfinite(value):
            return None
        return value
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, dict):
        return {k: _to_native(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_to_native(v) for v in value]
    return value


def _records(df):
    """Переводит датафрейм в список записей без numpy-типов и NaN/inf."""
    df_clean = df.astype(object)
    return [
        {k: _to_native(v) for k, v in row.items()}
        for row in df_clean.to_dict(orient="records")
    ]


def _normalize_key(value):
    """Нормализует ключ сопоставления (SKU/артикул) в строку без хвоста '.0'."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s == "" or s.lower() == "nan":
        return None
    return s


def _aggregate_campaign_orders(df_campaigns):
    """
    Агрегирует campaign_details по SKU и по артикулу:
    - «Заказы, шт. (РК)» — сумма по строкам кампаний;
    - «Заказы всего, шт.» — первое значение (общий показатель кампании не дублируем).
    Из расчёта исключаются строки с «Товар из сшивки».
    Возвращает (agg_by_sku, agg_by_article).
    """
    df = df_campaigns.copy()
    if 'Тип товара' in df.columns:
        df = df.loc[df['Тип товара'] != ARTICLE_TYPE_TO_EXCLUDE, :]

    for col in ['Заказы, шт. (РК)', 'Заказы всего, шт.']:
        if col not in df.columns:
            df[col] = 0
    df['Заказы, шт. (РК)'] = (
        pd.to_numeric(df['Заказы, шт. (РК)'], errors='coerce').fillna(0)
    )
    df['Заказы всего, шт.'] = (
        pd.to_numeric(df['Заказы всего, шт.'], errors='coerce').fillna(0)
    )

    def aggregate(group_col):
        if group_col not in df.columns:
            return pd.DataFrame(columns=['_key', 'Заказы РК', 'Заказы всего'])
        df_grouped = df.copy()
        df_grouped['_key'] = df_grouped[group_col].map(_normalize_key)
        return (
            df_grouped
            .loc[df_grouped['_key'].notna(), :]
            .groupby('_key', as_index=False)
            .agg(**{
                'Заказы РК': ('Заказы, шт. (РК)', 'sum'),
                'Заказы всего': ('Заказы всего, шт.', 'first'),
            })
        )

    return aggregate('SKU'), aggregate('Артикул')


def _build_product_summary(df_client_svod, agg_by_sku, agg_by_article):
    """
    Формирует product_summary для промпта: берём все товары листа 1,
    добавляем расчётные поля (Дефицит, Отмены %, Доля РК %),
    затем переименовываем колонки в английские поля промпта.
    """
    df = df_client_svod.copy()

    # Колонка остатка может называться 'Остаток' или 'Остаток <дата>'
    stock_col = 'Остаток' if 'Остаток' in df.columns else next(
        (c for c in df.columns if str(c).startswith('Остаток')), None
    )

    df['_sku_key'] = df['SKU'].map(_normalize_key) if 'SKU' in df.columns else None
    df['_article_key'] = df['Артикул'].map(_normalize_key) if 'Артикул' in df.columns else None

    # Заказы РК и всего: сначала по SKU, при отсутствии — по артикулу
    df = (
        df
        .merge(
            agg_by_sku.rename(columns={
                'Заказы РК': 'Заказы РК (SKU)',
                'Заказы всего': 'Заказы всего (SKU)',
            }),
            how='left',
            left_on='_sku_key',
            right_on='_key',
        )
        .merge(
            agg_by_article.rename(columns={
                'Заказы РК': 'Заказы РК (артикул)',
                'Заказы всего': 'Заказы всего (артикул)',
            }),
            how='left',
            left_on='_article_key',
            right_on='_key',
        )
    )
    df['Заказы РК'] = df['Заказы РК (SKU)'].fillna(df['Заказы РК (артикул)'])
    df['Заказы всего'] = df['Заказы всего (SKU)'].fillna(df['Заказы всего (артикул)'])

    # Расчётные поля (русские имена)
    df['Дефицит'] = df[stock_col] - 2 * df['Кол-во заказов']
    df['Отмены %'] = (
        (df['Заказы всего'] - df['Кол-во заказов'])
        / df['Заказы всего'].replace(0, np.nan) * 100
    )
    df['Доля РК %'] = (
        df['Заказы РК'] / df['Заказы всего'].replace(0, np.nan) * 100
    )
    # Подставляем фактическое имя колонки остатка, чтобы попало в map
    df = df.rename(columns={stock_col: 'Остаток'})

    # Отбор нужных колонок и переименование в английские поля промпта
    existing_map = {
        ru: en for ru, en in PRODUCT_SUMMARY_COLUMN_MAP.items()
        if ru in df.columns
    }
    product_summary = df.loc[:, list(existing_map.keys())].rename(columns=existing_map)

    # Сортировка по выручке (топ-30 ассистент берёт сам, но так нагляднее)
    if 'revenue' in product_summary.columns:
        product_summary = product_summary.sort_values(
            by='revenue', ascending=False, ignore_index=True
        )

    return product_summary


def _build_categories(df_products):
    """
    Строит словари категорий (по SKU и по артикулу) из колонки «Категория (тип)»
    листа 1, чтобы передать категорию в campaign_details, а не выдумывать её.
    """
    by_sku = {}
    by_article = {}
    if 'Категория (тип)' not in df_products.columns:
        return by_sku, by_article

    for _, row in df_products.iterrows():
        category = row.get('Категория (тип)')
        if category is None or (isinstance(category, float) and pd.isna(category)):
            continue
        category = str(category).strip()
        if not category:
            continue
        sku = _normalize_key(row.get('SKU'))
        article = _normalize_key(row.get('Артикул'))
        if sku:
            by_sku.setdefault(sku, category)
        if article:
            by_article.setdefault(article, category)

    return by_sku, by_article


def _build_campaign_details(df_campaigns, valid_sku, valid_article, category_by_sku, category_by_article):
    """
    Формирует campaign_details для промпта: исключает «Товар из сшивки»
    и SKU/артикулы, которых нет в product_summary,
    оставляет только поля под воронку, добавляет категорию (из листа 1)
    и переименовывает колонки в английские поля промпта.
    """
    df = df_campaigns.copy()
    if 'Тип товара' in df.columns:
        df = df.loc[df['Тип товара'] != ARTICLE_TYPE_TO_EXCLUDE, :]

    sku_key = df['SKU'].map(_normalize_key) if 'SKU' in df.columns else pd.Series(None, index=df.index)
    article_key = df['Артикул'].map(_normalize_key) if 'Артикул' in df.columns else pd.Series(None, index=df.index)
    mask = sku_key.isin(valid_sku) | article_key.isin(valid_article)
    df = df.loc[mask, :].copy()

    # Категория берётся из листа 1: сначала по SKU, при отсутствии — по артикулу
    sku_key = df['SKU'].map(_normalize_key) if 'SKU' in df.columns else pd.Series(None, index=df.index)
    article_key = df['Артикул'].map(_normalize_key) if 'Артикул' in df.columns else pd.Series(None, index=df.index)
    df['category'] = sku_key.map(category_by_sku)
    df['category'] = df['category'].fillna(article_key.map(category_by_article))

    existing_map = {
        ru: en for ru, en in CAMPAIGN_DETAILS_COLUMN_MAP.items()
        if ru in df.columns
    }
    columns = list(existing_map.keys()) + ['category']
    return df.loc[:, columns].rename(columns=existing_map)


def _try_parse_json(text):
    """Пытается распарсить JSON из строки (в т.ч. обёрнутый в ```json ... ```).
    Сначала json.loads, затем ast.literal_eval на случай почти-JSON от модели
    (одинарные кавычки, None/True вместо null/true и т.п.)."""
    if not isinstance(text, str):
        return None
    text = text.strip()
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        pass
    # Пробуем вытащить JSON из текста (например, обёрнут в ```json ... ```)
    match = re.search(r"\{.*\}", text, re.DOTALL)
    candidate = match.group(0) if match else text
    try:
        return json.loads(candidate)
    except (json.JSONDecodeError, TypeError):
        pass
    try:
        return ast.literal_eval(candidate)
    except (ValueError, SyntaxError, TypeError):
        pass
    return None


def _extract_recommendations(payload):
    """
    Ищет в ответе сервиса словарь с ключами campaigns/products.
    Обрабатывает разные возможные структуры ответа.
    """
    if isinstance(payload, dict):
        if "campaigns" in payload or "products" in payload:
            return payload
        for key in ("message", "content", "answer", "text", "response", "result", "data", "output", "reply"):
            sub = payload.get(key)
            if isinstance(sub, dict):
                res = _extract_recommendations(sub)
                if res:
                    return res
            if isinstance(sub, str):
                parsed = _try_parse_json(sub)
                if parsed:
                    res = _extract_recommendations(parsed)
                    if res:
                        return res
    if isinstance(payload, str):
        parsed = _try_parse_json(payload)
        if parsed:
            return _extract_recommendations(parsed)
    return None


# ----------------------------------------------------------------------
# Применение рекомендаций к датафреймам
# ----------------------------------------------------------------------
def _key(value):
    """Нормализует ключ сопоставления (SKU/артикул/номер кампании) в строку."""
    return _normalize_key(value)


def _format_recommendation(item):
    """Формирует текст рекомендации (+ причина) из записи ответа ИИ."""
    rec = str(item.get("recommendation") or "").strip()
    reason = str(item.get("reason") or "").strip()
    if reason:
        return f"{rec}\nПричина: {reason}" if rec else f"Причина: {reason}"
    return rec


def _format_issues(item):
    """Формирует текст проблем анализа из поля 'issues' записи ответа ИИ."""
    issues = item.get("issues")
    if not issues:
        return ""
    if isinstance(issues, str):
        return issues.strip()
    return "\n".join(str(i).strip() for i in issues if str(i).strip())


def _apply_products_recommendations(df, products):
    """Заполняет столбцы 'Комментарий ИИ' и 'Проблемы ИИ' по SKU/артикулу товара."""
    if not products:
        return df

    by_sku = {}
    by_article = {}
    for p in products:
        entry = (_format_recommendation(p), _format_issues(p))
        sku = _key(p.get("sku"))
        article = _key(p.get("article"))
        if sku:
            by_sku.setdefault(sku, entry)
        if article:
            by_article.setdefault(article, entry)

    if not by_sku and not by_article:
        return df

    def lookup(row):
        sku = _key(row.get("SKU"))
        if sku and sku in by_sku:
            return by_sku[sku]
        article = _key(row.get("Артикул"))
        if article and article in by_article:
            return by_article[article]
        return ("", "")

    df = df.copy()
    applied = df.apply(lookup, axis=1)
    df["Комментарий ИИ"] = applied.map(lambda x: x[0])
    df["Проблемы ИИ"] = applied.map(lambda x: x[1])
    return df


def _apply_campaigns_recommendations(df, campaigns):
    """
    Дополняет столбцы 'Комментарий ИИ' и 'Проблемы ИИ' рекомендациями по кампаниям
    (сопоставление по столбцу 'Номер кампании').
    """
    if not campaigns:
        return df

    by_campaign = {}
    for c in campaigns:
        cid = _key(c.get("campaign_id"))
        if cid:
            by_campaign.setdefault(cid, (_format_recommendation(c), _format_issues(c)))

    if not by_campaign:
        return df

    def lookup(row):
        existing_rec = row.get("Комментарий ИИ") or ""
        existing_issues = row.get("Проблемы ИИ") or ""
        # Комментарий по кампании выводим только для строк со статусом «Товар из РК»
        if str(row.get("Тип товара", "")).strip() != CAMPAIGN_TARGET_TYPE:
            return (existing_rec, existing_issues)
        cid = _key(row.get("Номер кампании"))
        entry = by_campaign.get(cid) if cid else None
        if entry is None:
            return (existing_rec, existing_issues)
        camp_rec, camp_issues = entry
        if existing_rec and camp_rec:
            new_rec = f"{existing_rec}\n---\n{camp_rec}"
        else:
            new_rec = existing_rec or camp_rec
        if existing_issues and camp_issues:
            new_issues = f"{existing_issues}\n{camp_issues}"
        else:
            new_issues = existing_issues or camp_issues
        return (new_rec, new_issues)

    df = df.copy()
    applied = df.apply(lookup, axis=1)
    df["Комментарий ИИ"] = applied.map(lambda x: x[0])
    df["Проблемы ИИ"] = applied.map(lambda x: x[1])
    return df


# ----------------------------------------------------------------------
# Главная функция-шаг
# ----------------------------------------------------------------------
def add_ai_recommendations(df_client_excel_svod, df_campaigns_report_processed):
    """
    Добавляет рекомендации ИИ-помощника в столбцы 'Комментарий ИИ' и 'Проблемы ИИ'
    на главной странице (df_client_excel_svod) и странице РК (df_campaigns_report_processed).

    Возвращает (df_client_excel_svod, df_campaigns_report_processed, ai_data_issues),
    где ai_data_issues — глобальные проблемы от ИИ (для отдельного листа отчёта).

    При любой ошибке (нет ключа, нет id ассистента, сбой API) возвращает
    датафреймы без изменений, чтобы отчёт сформировался без рекомендаций.
    """
    from ozon.scripts.constants import (
        AI_ASSISTANT_ID, PROMOTION_ORDERS_THRESHOLD, DRR_COMMENT_THRESHOLD
    )

    ai_data_issues = []

    # Инициализируем столбцы заранее
    df_client_excel_svod = df_client_excel_svod.copy()
    df_campaigns_report_processed = df_campaigns_report_processed.copy()
    df_client_excel_svod["Комментарий ИИ"] = ""
    df_client_excel_svod["Проблемы ИИ"] = ""
    df_campaigns_report_processed["Комментарий ИИ"] = ""
    df_campaigns_report_processed["Проблемы ИИ"] = ""

    try:
        # 1. Ключ ИИ-помощника
        api_key = get_assistant_api_key()
        if not api_key:
            return df_client_excel_svod, df_campaigns_report_processed, ai_data_issues

        # 2. id ассистента
        if not AI_ASSISTANT_ID:
            logger.warning("AI_ASSISTANT_ID не задан в ozon.scripts.constants — пропускаем ИИ-помощника")
            return df_client_excel_svod, df_campaigns_report_processed, ai_data_issues

        # 3. Диалог для клиента (только при включённом USE_THREADS)
        thread_id = None
        if USE_THREADS:
            thread_id = get_or_create_thread(AI_ASSISTANT_ID, api_key)

        # 4. Подготовка данных: product_summary (все товары листа 1) и campaign_details
        threshold = PROMOTION_ORDERS_THRESHOLD.get(
            client_name, DEFAULT_CAMPAIGN_LAUNCH_THRESHOLD
        )
        drr_threshold = DRR_COMMENT_THRESHOLD.get(client_name, DEFAULT_DRR_COMMENT_THRESHOLD)
        agg_by_sku, agg_by_article = _aggregate_campaign_orders(df_campaigns_report_processed)
        product_summary = _build_product_summary(
            df_client_excel_svod, agg_by_sku, agg_by_article
        )

        valid_sku = (
            {_normalize_key(v) for v in product_summary['sku'].dropna()} - {None}
            if 'sku' in product_summary.columns else set()
        )
        valid_article = (
            {_normalize_key(v) for v in product_summary['article'].dropna()} - {None}
            if 'article' in product_summary.columns else set()
        )
        category_by_sku, category_by_article = _build_categories(df_client_excel_svod)
        campaign_details = _build_campaign_details(
            df_campaigns_report_processed, valid_sku, valid_article,
            category_by_sku, category_by_article
        )

        logger.info(
            f"Данные для ИИ-помощника: product_summary={len(product_summary)} строк, "
            f"campaign_details={len(campaign_details)} строк"
        )

        message = json.dumps({
            "client_name": client_name,
            "campaign_launch_threshold": threshold,
            "drr_comment_threshold": drr_threshold,
            "product_summary": _records(product_summary),
            "campaign_details": _records(campaign_details),
        }, ensure_ascii=False, allow_nan=False)

        # 5. Запрос к ИИ
        response = send_message(AI_ASSISTANT_ID, api_key, thread_id, message)

        # 6. Разбор рекомендаций
        recommendations = _extract_recommendations(response)
        if not recommendations:
            logger.warning("ИИ-помощник не вернул рекомендации")
            return df_client_excel_svod, df_campaigns_report_processed, ai_data_issues

        ai_data_issues = recommendations.get("data_issues") or []
        if ai_data_issues:
            issues_text = "\n".join(f"  - {issue}" for issue in ai_data_issues)
            logger.warning(
                f"ИИ-помощник вернул глобальные проблемы ({len(ai_data_issues)}):\n"
                f"{issues_text}"
            )

        # 7. Применяем к датафреймам
        df_client_excel_svod = _apply_products_recommendations(
            df_client_excel_svod,
            recommendations.get("products") or []
        )
        df_campaigns_report_processed = _apply_campaigns_recommendations(
            df_campaigns_report_processed,
            recommendations.get("campaigns") or []
        )

        logger.info("Рекомендации ИИ-помощника добавлены в отчёт")
    except Exception as e:
        logger.exception(f"Ошибка при работе ИИ-помощника: {e}")

    return df_client_excel_svod, df_campaigns_report_processed, ai_data_issues
