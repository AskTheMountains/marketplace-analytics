# %%
import re
import time
from pathlib import Path

import requests
import pandas as pd
from loguru import logger


class WildberriesSellerClient:
    def __init__(
        self,
        pause_every=10,          # через сколько запросов делать глобальную паузу
        pause_seconds=2.0,       # длительность глобальной паузы
        timeout=10,              # таймаут HTTP-запроса
        per_request_delay=0.7,   # пауза между КАЖДЫМ запросом
    ):
        self.pause_every = pause_every
        self.pause_seconds = pause_seconds
        self.timeout = timeout
        self.per_request_delay = per_request_delay

        self.session = requests.Session()
        self._request_counter = 0

        logger.info(
            "Инициализирован WildberriesSellerClient: "
            "pause_every={}, pause_seconds={}, timeout={}, per_request_delay={}",
            pause_every, pause_seconds, timeout, per_request_delay
        )

    @staticmethod
    def _extract_supplier_id(seller_url: str) -> str:
        m = re.search(r'/seller/(\d+)', seller_url)
        if not m:
            raise ValueError(f"Не удалось найти ID продавца в ссылке: {seller_url}")
        return m.group(1)

    @staticmethod
    def _normalize_path(path) -> str:
        """
        Нормализуем путь: убираем лишние кавычки и пробелы.
        path может быть и Path, и строкой.
        """
        clean = str(path).strip().strip('\'"')
        clean = clean.replace('"', '').replace("'", "")
        return clean

    def _maybe_sleep_regular(self, rows, progress_path=None, verbose=True):
        """
        Регулярная (глобальная) пауза после каждых `pause_every` запросов.
        Перед паузой сохраняет прогресс, если указан progress_path.
        """
        if self.pause_every and self.pause_every > 0:
            if self._request_counter > 0 and self._request_counter % self.pause_every == 0:
                if verbose:
                    logger.info(
                        "Достигнуто {} запросов, готовим регулярную паузу на {} секунд",
                        self._request_counter,
                        self.pause_seconds,
                    )
                else:
                    logger.info(
                        "Глобальная пауза: {} запросов обработано, пауза {} секунд",
                        self._request_counter,
                        self.pause_seconds,
                    )
                if progress_path is not None:
                    self._save_progress(rows, progress_path, reason="регулярная пауза")
                time.sleep(self.pause_seconds)

    def _save_progress(self, rows, progress_path, reason=""):
        """
        Сохраняет текущий прогресс (список dict) в Excel-файл.
        rows – список словарей (результаты запросов).
        progress_path – Path или строка.
        """
        if not rows:
            logger.warning(
                "Запрошено сохранение прогресса, но данных ещё нет (reason={})",
                reason,
            )
            return

        path_str = self._normalize_path(progress_path)

        df_tmp = pd.DataFrame(rows)
        try:
            df_tmp.to_excel(path_str, index=False)
            logger.info(
                "Промежуточный прогресс сохранён в '{}' ({} записей, причина: {})",
                path_str,
                len(rows),
                reason,
            )
        except Exception as e:
            logger.exception(
                "Не удалось сохранить прогресс в '{}': {}",
                path_str,
                e,
            )

    def get_supplier_info(self, seller_url, verbose=True):
        """
        Возвращает dict с данными о продавце,
        либо None в случае ошибки.

        Делает паузу per_request_delay перед каждым запросом.
        Не занимается сохранением прогресса/глобальными паузами.
        """
        # Пауза между каждым запросом
        if self.per_request_delay and self.per_request_delay > 0:
            time.sleep(self.per_request_delay)

        try:
            supplier_id = self._extract_supplier_id(seller_url)
        except ValueError as e:
            logger.error("Ошибка разбора URL '{}': {}", seller_url, e)
            return None

        api_url = f"https://static-basket-01.wbbasket.ru/vol0/data/supplier-by-id/{supplier_id}.json"

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/144.0.0.0 Safari/537.36"
            ),
            "Accept": "*/*",
            "Referer": seller_url,
            "Origin": "https://www.wildberries.ru",
        }

        try:
            if verbose:
                logger.debug("Запрос к {}, supplier_id={}", api_url, supplier_id)

            resp = self.session.get(api_url, headers=headers, timeout=self.timeout)
            self._request_counter += 1

            # Специальная обработка 429
            if resp.status_code == 429:
                logger.error("Получен статус 429 Too Many Requests для URL '{}'", seller_url)
                # возвращаем маркер ошибки 429, чтобы выше решить, что делать
                return {"_error": "429"}

            resp.raise_for_status()

            data = resp.json()

            if not isinstance(data, dict) or "supplierId" not in data:
                logger.error(
                    "Неверный формат данных для URL '{}': {}",
                    seller_url,
                    data,
                )
                return None

            if verbose:
                logger.info(
                    "Успешно получены данные о продавце supplierId={} для URL {}",
                    data.get("supplierId"),
                    seller_url,
                )
            return data

        except requests.exceptions.HTTPError as e:
            logger.error(
                "HTTP ошибка для URL '{}': {} (код {})",
                seller_url,
                e,
                resp.status_code if 'resp' in locals() else "неизвестен",
            )
        except requests.exceptions.Timeout:
            logger.error("Таймаут для URL '{}'", seller_url)
        except requests.exceptions.RequestException as e:
            logger.error("Сетевая ошибка для URL '{}': {}", seller_url, e)
        except ValueError as e:
            logger.error("Ошибка обработки ответа для URL '{}': {}", seller_url, e)
        except Exception as e:
            logger.exception("Неизвестная ошибка для URL '{}': {}", seller_url, e)

        return None

    def fetch_to_dataframe(
        self,
        seller_urls,
        progress_path=None,
        verbose=True,
    ):
        """
        Принимает список ссылок на продавцов,
        возвращает pandas.DataFrame с данными.
        Неуспешные запросы пропускаются.

        seller_urls   – список URL
        progress_path – Path или строка; если указан, прогресс периодически
                        сохраняется в Excel (при глобальной паузе и при 429).
        verbose       – если False, не логируем подробности по каждому URL,
                        только информацию о пачках и паузах/ошибках.
        """
        rows = []
        failed = []

        if verbose:
            logger.info("Начинаем обработку {} URL(ов)", len(seller_urls))
        else:
            logger.info(
                "Начинаем пачку запросов: {} URL(ов), подробный лог по каждому URL отключён",
                len(seller_urls),
            )

        for url in seller_urls:
            if verbose:
                logger.info("Обработка URL: {}", url)

            data = self.get_supplier_info(url, verbose=verbose)

            # Специальный случай 429
            if isinstance(data, dict) and data.get("_error") == "429":
                logger.warning(
                    "Статус 429 для URL '{}'. Сохраняем прогресс и ждём 60 секунд...",
                    url,
                )
                if progress_path is not None:
                    self._save_progress(rows, progress_path, reason="ошибка 429")
                time.sleep(60)
                failed.append(url)
                continue

            if data is not None:
                rows.append(data)
            else:
                failed.append(url)
                if verbose:
                    logger.warning(
                        "Данные для URL '{}' не получены, продолжаем со следующим",
                        url,
                    )

            # Регулярная (глобальная) пауза + сохранение прогресса
            self._maybe_sleep_regular(rows, progress_path, verbose=verbose)

        df = pd.DataFrame(rows)

        logger.info(
            "Завершено. Успешно: {}, с ошибками: {}",
            len(rows),
            len(failed),
        )
        if failed:
            logger.warning("Список URL с ошибками: {}", failed)

        # Финальное сохранение результата
        if progress_path is not None:
            self._save_progress(rows, progress_path, reason="финальное сохранение")

        return df


def read_seller_urls_from_excel(
    filepath,
    url_column="Профиль на сайте Wildberries",
):
    """
    Считывает Excel-файл и возвращает список ссылок на продавцов
    из столбца `url_column`.
    filepath – Path или строка пути к файлу.
    """
    logger.info("Чтение Excel-файла: {}", filepath)
    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        logger.exception("Не удалось прочитать Excel-файл '{}': {}", filepath, e)
        raise

    if url_column not in df.columns:
        raise ValueError(
            f"В файле '{filepath}' не найден столбец '{url_column}'. "
            f"Найденные столбцы: {list(df.columns)}"
        )

    urls = (
        df[url_column]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    logger.info("Найдено {} ссылок в колонке '{}'", len(urls), url_column)
    return urls


# %%
if __name__ == "__main__":
    # Путь к вашему Excel-файлу с исходными ссылками
    excel_path = Path('d:\\artem\\Downloads\\Telegramm\\Лист Microsoft Excel.xlsx')

    # Путь к файлу, куда будет сохраняться прогресс
    progress_excel_path = Path('d:\\artem\\Downloads\\wb_sellers_progress.xlsx')

    seller_urls = read_seller_urls_from_excel(excel_path)

    client = WildberriesSellerClient(
        pause_every=500,        # глобальная пауза каждые n запросов
        pause_seconds=5.0,     # длина глобальной паузы
        timeout=10,
        per_request_delay=0  # пауза между каждым запросом
    )

    df_sellers = client.fetch_to_dataframe(
        seller_urls,
        progress_path=progress_excel_path,
        verbose=True,  # False — тихий режим; True — подробный лог по каждому URL
    )

    # print(df_sellers.head())

# %%
