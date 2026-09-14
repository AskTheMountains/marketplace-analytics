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


def _target_product_summary(df):
    """
    Целевая выборка: только товары (SKU/артикулы), присутствующие в product_summary
    и у которых в столбце РК что-то есть (не пусто и не 0).
    """
    if "РК" not in df.columns:
        return df
    return df.loc[(df["РК"].notna()) & (df["РК"] != 0), :]


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
    if value is None:
        return None
    return str(value).strip()


def _format_recommendation(item):
    """Формирует текст рекомендации (+ причина) из записи ответа ИИ."""
    rec = str(item.get("recommendation") or "").strip()
    reason = str(item.get("reason") or "").strip()
    if reason:
        return f"{rec}\nПричина: {reason}" if rec else f"Причина: {reason}"
    return rec


def _apply_products_recommendations(df, products):
    """Заполняет столбец 'Комментарий ИИ' по SKU/артикулу товара."""
    if not products:
        return df

    by_sku = {}
    by_article = {}
    for p in products:
        rec = _format_recommendation(p)
        sku = _key(p.get("sku"))
        article = _key(p.get("article"))
        if sku:
            by_sku.setdefault(sku, rec)
        if article:
            by_article.setdefault(article, rec)

    if not by_sku and not by_article:
        return df

    def lookup(row):
        sku = _key(row.get("SKU"))
        if sku and sku in by_sku:
            return by_sku[sku]
        article = _key(row.get("Артикул"))
        if article and article in by_article:
            return by_article[article]
        return ""

    df = df.copy()
    df["Комментарий ИИ"] = df.apply(lookup, axis=1)
    return df


def _apply_campaigns_recommendations(df, campaigns):
    """
    Дополняет столбец 'Комментарий ИИ' рекомендациями по кампаниям
    (сопоставление по столбцу 'Номер кампании').
    """
    if not campaigns:
        return df

    by_campaign = {}
    for c in campaigns:
        cid = _key(c.get("campaign_id"))
        if cid:
            by_campaign.setdefault(cid, _format_recommendation(c))

    if not by_campaign:
        return df

    def lookup(row):
        existing = row.get("Комментарий ИИ") or ""
        cid = _key(row.get("Номер кампании"))
        if cid and cid in by_campaign:
            camp_rec = by_campaign[cid]
            if existing:
                return f"{existing}\n---\n{camp_rec}"
            return camp_rec
        return existing

    df = df.copy()
    df["Комментарий ИИ"] = df.apply(lookup, axis=1)
    return df


# ----------------------------------------------------------------------
# Главная функция-шаг
# ----------------------------------------------------------------------
def add_ai_recommendations(df_client_excel_svod, df_campaigns_report_processed):
    """
    Добавляет рекомендации ИИ-помощника в столбец 'Комментарий ИИ'
    на главной странице (df_client_excel_svod) и странице РК (df_campaigns_report_processed).

    При любой ошибке (нет ключа, нет id ассистента, сбой API) возвращает
    датафреймы без изменений, чтобы отчёт сформировался без рекомендаций.
    """
    from ozon.scripts.constants import AI_ASSISTANT_ID

    # Инициализируем столбцы заранее
    df_client_excel_svod = df_client_excel_svod.copy()
    df_campaigns_report_processed = df_campaigns_report_processed.copy()
    df_client_excel_svod["Комментарий ИИ"] = ""
    df_campaigns_report_processed["Комментарий ИИ"] = ""

    try:
        # 1. Ключ ИИ-помощника
        api_key = get_assistant_api_key()
        if not api_key:
            return df_client_excel_svod, df_campaigns_report_processed

        # 2. id ассистента
        if not AI_ASSISTANT_ID:
            logger.warning("AI_ASSISTANT_ID не задан в ozon.scripts.constants — пропускаем ИИ-помощника")
            return df_client_excel_svod, df_campaigns_report_processed

        # 3. Диалог для клиента (только при включённом USE_THREADS)
        thread_id = None
        if USE_THREADS:
            thread_id = get_or_create_thread(AI_ASSISTANT_ID, api_key)

        # 4. Целевая выборка для ИИ и сборка сообщения
        product_summary = _target_product_summary(df_client_excel_svod)
        if product_summary.empty:
            logger.warning("Нет товаров с расходами по РК — отправляем ИИ только данные по кампаниям")
        campaign_details = df_campaigns_report_processed

        message = json.dumps({
            "client_name": client_name,
            "product_summary": _records(product_summary),
            "campaign_details": _records(campaign_details),
        }, ensure_ascii=False, allow_nan=False)

        # 5. Запрос к ИИ
        response = send_message(AI_ASSISTANT_ID, api_key, thread_id, message)

        # 6. Разбор рекомендаций
        recommendations = _extract_recommendations(response)
        if not recommendations:
            logger.warning("ИИ-помощник не вернул рекомендации")
            return df_client_excel_svod, df_campaigns_report_processed

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

    return df_client_excel_svod, df_campaigns_report_processed
