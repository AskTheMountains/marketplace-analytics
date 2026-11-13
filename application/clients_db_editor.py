"""Редактор базы данных клиентов (settings.json).

Позволяет редактировать поля клиентов (строки) по маркетплейсам Ozon / WB
(колонки), добавлять и удалять клиентов, а также проверять API-ключи.
Изменения не сохраняются на лету: есть отдельная кнопка «Сохранить изменения»,
а при закрытии окна с несохранёнными правками спрашивается подтверждение.
"""

import base64
import json
import re

from datetime import datetime

import requests

from PySide6.QtCore import Qt, QEvent, QThread, Signal
from PySide6.QtGui import QColor, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QDialog, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMessageBox, QPushButton, QStyledItemDelegate, QTableWidget,
    QTableWidgetItem, QVBoxLayout,
)

from application.paths import SETTINGS_JSON

# Лимитер для запросов к API Ozon (по правилам проекта нельзя использовать
# голый requests; все запросы идут через OzonRateLimiter)
from ozon.scripts.ozon_generic_functions import OzonRateLimiter
from wb.scripts.wb_generic_functions import WbRateLimiter


# --- Заготовки методов проверки API-ключей --------------------------------

def check_wb_api_key(api_key_wb):
    """Проверка API-ключа WB (JWT).

    Отправляет GET на common-api.wildberries.ru/ping — публичный метод,
    который с валидным токеном возвращает 200, а с невалидным — 401.
    Срок окончания токена читается локально из поля `exp` JWT.

    Возвращает (valid: bool|None, expires: datetime|None, message: str).
      valid=True,  expires=<дата> — токен рабочий;
      valid=False, message=<причина> — токен недействителен (401) / не JWT;
      valid=None,  message=<причина> — не удалось выполнить запрос.
    """
    if not api_key_wb or not api_key_wb.strip():
        return False, None, "API-ключ не заполнен"

    # Срок окончания читаем из JWT
    expires = _wb_token_expires(api_key_wb)
    if expires is None:
        return False, None, "Не удалось разобрать срок окончания токена (exp)"

    headers = {"Authorization": api_key_wb.strip()}

    session = requests.Session()
    session.headers.update(headers)
    limiter = WbRateLimiter(period_seconds=60, period_limit=300, max_burst=20)

    try:
        resp, ok = limiter.call(
            session.get,
            url="https://common-api.wildberries.ru/ping",
            timeout=60,
        )
    except Exception as e:
        return None, None, f"Не удалось выполнить запрос: {e}"

    if not ok or resp is None:
        return None, None, "Не удалось проверить ключ из-за ошибки запроса"

    if resp.status_code == 401:
        return False, None, "Токен недействителен (401 Unauthorized)"

    if resp.status_code != 200:
        return None, None, f"Не удалось проверить ключ: ответ API {resp.status_code}"

    return True, expires, ""


def _wb_token_expires(api_key_wb):
    """Возвращает срок окончания WB-токена из поля `exp` JWT, или None."""
    try:
        parts = api_key_wb.strip().split(".")
        if len(parts) != 3:
            return None
        payload_part = parts[1]
        payload_part += "=" * (-len(payload_part) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_part))
        exp = payload.get("exp")
        if exp is None:
            return None
        return datetime.fromtimestamp(exp)
    except Exception:
        return None


def check_ozon_api_key(api_key_ozon, client_id_ozon):
    """Проверка API-ключа Ozon.

    Отправляет POST на /v1/roles и берёт из ответа поле `expires_at` —
    срок окончания действия ключа.

    Возвращает (valid: bool|None, expires: datetime|None, message: str).
      valid=True,  expires=<дата> — ключ действителен;
      valid=False, message=<причина> — ключ недействителен / истёк;
      valid=None,  message=<причина> — не удалось выполнить запрос (ошибка сети/HTTP).
    """
    if not api_key_ozon or not api_key_ozon.strip() \
            or not client_id_ozon or not client_id_ozon.strip():
        return False, None, "Не заполнены API-ключ Ozon или Client ID"

    headers = {
        "Host": "api-seller.ozon.ru",
        "Client-Id": str(client_id_ozon).strip(),
        "Api-Key": api_key_ozon.strip(),
        "Content-Type": "application/json",
    }

    session = requests.Session()
    session.headers.update(headers)
    limiter = OzonRateLimiter(max_retries=3)

    try:
        resp, ok = limiter.call(
            session.post,
            url="https://api-seller.ozon.ru/v1/roles",
            json={},
            timeout=60,
        )
    except Exception as e:
        return None, None, f"Не удалось выполнить запрос: {e}"

    if not ok or resp is None:
        # Запрос не прошёл. Если в теле ответа явно сказано, что ключ истёк
        # (код 7, "Api-key has expired") — считаем ключ недействительным (красный).
        # Прочие ошибки (403 по правам, сеть, лимиты) — жёлтый (не удалось проверить).
        expired = _response_says_key_expired(resp)
        if expired:
            return False, None, expired
        return None, None, "Не удалось проверить ключ из-за ошибки запроса"

    try:
        data = resp.json()
    except Exception:
        return None, None, "Не удалось проверить ключ: неверный ответ API"

    expires_str = data.get("expires_at")
    if not expires_str:
        return None, None, "В ответе API отсутствует срок окончания ключа"

    try:
        expires = datetime.fromisoformat(expires_str.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None, None, "Не удалось разобрать срок окончания ключа"

    return True, expires, ""


def _response_says_key_expired(resp):
    """Проверяет, что тело ответа Ozon явно сообщает об истёкшем API-ключе.

    Возвращает строку-описание, если ключ истёк, иначе пустую строку.
    """
    if resp is None:
        return ""
    try:
        data = resp.json()
    except Exception:
        return ""

    code = data.get("code")
    message = (data.get("message") or "").lower()
    if code == 7 and ("api-key has expired" in message or "api-key is deactivated" in message):
        return "API-ключ недействителен: истёк или деактивирован"
    return ""


# --- Определение структуры таблицы ---------------------------------------

# (ключ в settings.json, заголовок колонки, обязательное поле)
FIELDS = [
    ("client_name", "Имя клиента (client_name)", True),
    ("client_ru_name", "Русское имя (client_ru_name)", False),
    ("api_key_ozon", "API Key Ozon", False),
    ("client_id_ozon", "Client ID Ozon", False),
    ("client_id_performance_ozon", "Client ID Performance Ozon", False),
    ("client_secret_performance_ozon", "Client Secret Performance Ozon", False),
    ("api_key_wb", "API Key WB", False),
]

OZON_REQUIRED_FIELDS = ["api_key_ozon", "client_id_ozon"]
WB_REQUIRED_FIELDS = ["api_key_wb"]

# Цвета подсветки ячеек при проверке ключей
GREEN = QColor("#d4f0d4")
RED = QColor("#f5c6c6")
YELLOW = QColor("#fdf3c8")


def russian_question(parent, title, text, *buttons):
    """QMessageBox с кнопками на русском языке.

    buttons — список кортежей (текст_кнопки, стандартная_кнопка), где
    стандартная_кнопка — одно из значений QMessageBox.StandardButton
    (например QMessageBox.Yes, QMessageBox.No, QMessageBox.Save, ...).
    Возвращает QMessageBox.StandardButton, соответствующий нажатой кнопке.
    """
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(text)
    box.setIcon(QMessageBox.Question)
    for label, standard_button in buttons:
        btn = box.addButton(standard_button)
        btn.setText(label)
        box.setDefaultButton(btn)
    box.exec()
    clicked = box.clickedButton()
    return box.standardButton(clicked)


# --- Очистка вставки из буфера обмена --------------------------------------

def clean_clipboard_text(text):
    """Убирает из вставляемого текста переносы строк, табы и лишние пробелы."""
    text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    text = text.replace("\t", " ")
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


class CleanPasteDelegate(QStyledItemDelegate):
    """Делегат, очищающий вставляемый текст внутри редактора ячейки."""

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.installEventFilter(self)
        return editor

    def eventFilter(self, obj, event):
        if isinstance(obj, QLineEdit) and event.type() == QEvent.KeyPress \
                and event.matches(QKeySequence.Paste):
            text = QApplication.clipboard().text()
            if text:
                obj.insert(clean_clipboard_text(text))
            return True
        return super().eventFilter(obj, event)


class CleanPasteTable(QTableWidget):
    """QTableWidget, очищающий вставку в нередактируемую (неактивную) ячейку."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._clean_delegate = CleanPasteDelegate(self)

    def _apply_to_item(self):
        item = self.currentItem()
        if item is None:
            return
        text = QApplication.clipboard().text()
        if text:
            item.setText(clean_clipboard_text(text))

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Paste) \
                and self.state() != QAbstractItemView.EditingState:
            self._apply_to_item()
            return
        super().keyPressEvent(event)


class KeyCheckThread(QThread):
    """Фоновая проверка API-ключей всех клиентов.

    Сигнал key_result(row, marketplace, valid, expires, message) вызывается для
    каждого проверенного маркетплейса каждого клиента.
    Сигнал finished_all() — после завершения всех проверок.
    """

    key_result = Signal(int, str, object, object, str)
    finished_all = Signal()

    def __init__(self, rows, parent=None):
        super().__init__(parent)
        # rows: список кортежей (реальный_индекс_строки, dict-строка клиента)
        self.rows = rows

    def run(self):
        for row_idx, client in self.rows:
            # WB
            wb_key = client.get("api_key_wb", "").strip()
            if wb_key:
                valid, expires, message = check_wb_api_key(wb_key)
                self.key_result.emit(row_idx, "wb", valid, expires, message)

            # Ozon
            ozon_key = client.get("api_key_ozon", "").strip()
            ozon_cid = client.get("client_id_ozon", "").strip()
            if ozon_key and ozon_cid:
                valid, expires, message = check_ozon_api_key(ozon_key, ozon_cid)
                self.key_result.emit(row_idx, "ozon", valid, expires, message)

        self.finished_all.emit()


class ClientsDbEditor(QDialog):
    def __init__(self, settings, parent=None, on_saved=None):
        super().__init__(parent)
        self.settings = settings
        self.on_saved = on_saved
        self.modified = False
        self._check_thread = None

        self.setWindowTitle("Редактор базы данных клиентов")
        self.resize(1100, 600)

        self.col_by_key = {key: idx for idx, (key, _, _) in enumerate(FIELDS)}
        # Все ключи, которые нужно сохранять (текущие + стандартные поля)
        self.keys_order = list(dict.fromkeys(list(self.settings.keys()) + [f[0] for f in FIELDS]))

        self._build_ui()
        self._load_rows()

    # --- UI ----------------------------------------------------------------

    def _build_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("База данных клиентов (settings.json)")
        title.setStyleSheet("font-weight: 700; font-size: 14px;")
        layout.addWidget(title)

        hint = QLabel(
            "Строки — клиенты, столбцы — поля. <b>*</b> — обязательное поле. "
            "Изменения сохраняются только по кнопке «Сохранить изменения»."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #6b7280;")
        layout.addWidget(hint)

        self.table = CleanPasteTable(0, len(FIELDS))
        self.table.setHorizontalHeaderLabels([h for _, h, _ in FIELDS])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Выделяем отдельные ячейки; строка выделяется по клику на номер строки
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.itemChanged.connect(self._on_item_changed)
        # Клик по номеру строки (вертикальный заголовок) выделяет всю строку
        self.table.verticalHeader().sectionClicked.connect(self._on_row_header_clicked)
        # Делегат очищает вставку при редактировании ячейки
        self.table.setItemDelegate(CleanPasteDelegate(self.table))
        layout.addWidget(self.table)

        # Подсветка обязательных полей в шапке
        for idx, (_, _, required) in enumerate(FIELDS):
            item = self.table.horizontalHeaderItem(idx)
            if required:
                item.setText(item.text() + " *")
                item.setForeground(QColor("#b30000"))

        # --- Панель кнопок ---
        btn_layout = QHBoxLayout()

        self.btn_check = QPushButton("Проверить API-ключи")
        self.btn_check.clicked.connect(self.check_api_keys)
        btn_layout.addWidget(self.btn_check)

        self.btn_add = QPushButton("Добавить клиента")
        self.btn_add.clicked.connect(self.add_client)
        btn_layout.addWidget(self.btn_add)

        self.btn_remove = QPushButton("Удалить клиента")
        self.btn_remove.clicked.connect(self.remove_client)
        btn_layout.addWidget(self.btn_remove)

        btn_layout.addStretch(1)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #6b7280;")
        btn_layout.addWidget(self.status_label)

        self.btn_save = QPushButton("Сохранить изменения")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.save_changes)
        btn_layout.addWidget(self.btn_save)

        self.btn_close = QPushButton("Закрыть")
        self.btn_close.clicked.connect(self.close)
        btn_layout.addWidget(self.btn_close)

        layout.addLayout(btn_layout)

    # --- Данные -------------------------------------------------------------

    def _load_rows(self):
        """Заполняет таблицу из self.settings (столбцы-списки)."""
        n = len(self.settings.get("client_name", []))
        self.table.blockSignals(True)
        self.table.setRowCount(n)
        for row in range(n):
            self._set_row(row, {key: self.settings.get(key, [""] * n)[row] for key in self.keys_order})
        self.table.blockSignals(False)
        self.modified = False
        self.btn_save.setEnabled(False)

    def _set_row(self, row, client):
        for key in self.keys_order:
            col = self.col_by_key.get(key)
            if col is None:
                continue
            item = QTableWidgetItem(str(client.get(key, "")))
            item.setData(Qt.UserRole, key)
            self.table.setItem(row, col, item)

    def _on_row_header_clicked(self, row):
        """Выделяет всю строку при клике на номер строки (как в Excel)."""
        if row < 0:
            return
        self.table.clearSelection()
        self.table.selectRow(row)

    def selected_rows(self):
        """Возвращает список индексов выбранных строк.

        Если выделена отдельная ячейка — возвращает строку этой ячейки.
        """
        rows = set()
        for idx in self.table.selectedIndexes():
            if idx.isValid():
                rows.add(idx.row())
        return sorted(rows)

    def _rows(self):
        """Возвращает список dict-строк клиентов из текущего состояния таблицы."""
        rows = []
        for row in range(self.table.rowCount()):
            client = {}
            for key in self.keys_order:
                col = self.col_by_key.get(key)
                if col is None:
                    continue
                item = self.table.item(row, col)
                client[key] = item.text() if item is not None else ""
            rows.append(client)
        return rows

    def _on_item_changed(self, item):
        if not self.modified:
            self.modified = True
            self.btn_save.setEnabled(True)
            self.status_label.setText("Есть несохранённые изменения")

    # --- Операции с клиентами ----------------------------------------------

    def add_client(self):
        self.table.blockSignals(True)
        row = self.table.rowCount()
        self.table.insertRow(row)
        client = {key: "" for key in self.keys_order}
        self._set_row(row, client)
        self.table.blockSignals(False)
        self._on_item_changed(None)
        self.table.setCurrentCell(row, 0)

    def remove_client(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Удаление клиента", "Выберите клиента для удаления.")
            return
        name = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
        if name:
            name = f" «{name}»"
        answer = russian_question(
            self,
            "Удаление клиента",
            f"Удалить клиента{name}?",
            ("Да", QMessageBox.Yes),
            ("Нет", QMessageBox.No),
        )
        if answer != QMessageBox.Yes:
            return
        self.table.removeRow(row)
        self._on_item_changed(None)

    # --- Проверка API-ключей ------------------------------------------------

    def check_api_keys(self):
        if self._check_thread and self._check_thread.isRunning():
            return

        all_rows = self._rows()
        selected = self.selected_rows()

        # Если выделена строка/ячейка — проверяем только её; иначе всех.
        # Потоку передаём пары (реальный_индекс_строки, dict-строка).
        if selected:
            target_rows = [(r, all_rows[r]) for r in selected]
        else:
            target_rows = list(enumerate(all_rows))

        # Сбрасываем подсветку только для проверяемых клиентов
        rows_to_reset = selected if selected else range(len(all_rows))
        for row in rows_to_reset:
            self._reset_key_cell(row, "ozon")
            self._reset_key_cell(row, "wb")

        self.status_label.setText("Проверка API-ключей…")
        self.btn_check.setEnabled(False)

        self._check_thread = KeyCheckThread(target_rows, parent=self)
        self._check_thread.key_result.connect(self._on_key_result)
        self._check_thread.finished_all.connect(self._on_check_finished)
        self._check_thread.start()

    def _reset_key_cell(self, row, marketplace):
        col = self.col_by_key.get("api_key_ozon" if marketplace == "ozon" else "api_key_wb")
        if col is None:
            return
        item = self.table.item(row, col)
        if item is not None:
            self.table.blockSignals(True)
            item.setBackground(QColor("#ffffff"))
            item.setToolTip("")
            self.table.blockSignals(False)

    def _on_key_result(self, row, marketplace, valid, expires, message):
        col = self.col_by_key.get("api_key_ozon" if marketplace == "ozon" else "api_key_wb")
        if col is None:
            return
        item = self.table.item(row, col)
        if item is None:
            return

        self.table.blockSignals(True)
        if valid is None:
            # Не удалось выполнить проверку (ошибка запроса/сети/API)
            item.setBackground(YELLOW)
            item.setToolTip(message or "Не удалось проверить ключ")
        elif not valid:
            item.setBackground(RED)
            item.setToolTip(f"Проверить API-ключ: {message or 'ключ недействителен'}")
        elif expires is not None and expires < datetime.now():
            # Ключ валиден, но срок истёк
            item.setBackground(RED)
            item.setToolTip(
                f"Проверить API-ключ: срок действия истёк ({expires:%d.%m.%Y %H:%M})"
            )
        else:
            item.setBackground(GREEN)
            if expires is not None:
                item.setToolTip(f"Действует до {expires:%d.%m.%Y %H:%M}")
            else:
                item.setToolTip("Ключ действителен")
        self.table.blockSignals(False)

    def _on_check_finished(self):
        self.btn_check.setEnabled(True)
        if self.modified:
            self.status_label.setText("Проверка завершена. Есть несохранённые изменения")
        else:
            self.status_label.setText("Проверка завершена")

    # --- Сохранение ---------------------------------------------------------

    def save_changes(self):
        rows = self._rows()

        # Проверка обязательных полей
        problems = []
        for idx, client in enumerate(rows):
            name = client.get("client_name", "").strip()
            if not name:
                problems.append(f"Клиент №{idx + 1}: не заполнено обязательное поле «Имя клиента»")
                continue

            has_wb = bool(client.get("api_key_wb", "").strip())

            if has_wb:
                # Клиент только на WB
                missing = [k for k in WB_REQUIRED_FIELDS if not client.get(k, "").strip()]
                if missing:
                    problems.append(f"Клиент «{name}»: не заполнено обязательное поле WB: "
                                    f"{', '.join(missing)}")
            else:
                # Клиент на Ozon (или ещё не заполнен ни один маркетплейс)
                has_ozon = any(client.get(k, "").strip() for k in
                               ("api_key_ozon", "client_id_ozon",
                                "client_id_performance_ozon", "client_secret_performance_ozon"))
                if has_ozon:
                    missing = [k for k in OZON_REQUIRED_FIELDS if not client.get(k, "").strip()]
                    if missing:
                        problems.append(f"Клиент «{name}»: не заполнены обязательные поля Ozon: "
                                        f"{', '.join(missing)}")

        if problems:
            QMessageBox.warning(
                self,
                "Проверка данных",
                "Перед сохранением исправьте обязательные поля:\n\n" + "\n".join(problems),
            )
            return

        # Собираем данные: для каждого ключа список значений по строкам
        new_data = {}
        n = len(rows)
        for key in self.keys_order:
            col = self.col_by_key.get(key)
            if col is not None:
                values = [client.get(key, "") for client in rows]
                # Для совместимости в служебные поля Ozon, если они не заполнены,
                # проставляем 1
                if key in ("client_id_performance_ozon", "client_secret_performance_ozon"):
                    values = [
                        "1" if not str(value).strip() else value
                        for value in values
                    ]
                new_data[key] = values
            else:
                # Ключ без колонки — переносим исходные значения, добивая до нужной длины
                existing = list(self.settings.get(key, []))
                if len(existing) < n:
                    existing = existing + [""] * (n - len(existing))
                new_data[key] = existing

        try:
            with open(SETTINGS_JSON, "w", encoding="utf-8") as f:
                json.dump(new_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка сохранения", f"Не удалось сохранить файл:\n{e}")
            return

        # Обновляем self.settings in-place, чтобы все вкладки увидели изменения
        self.settings.clear()
        self.settings.update(new_data)

        self.modified = False
        self.btn_save.setEnabled(False)
        self.status_label.setText("Изменения сохранены")

        if self.on_saved:
            self.on_saved()

    # --- Закрытие -----------------------------------------------------------

    def closeEvent(self, event):
        if self._check_thread and self._check_thread.isRunning():
            QMessageBox.information(
                self, "Проверка ключей",
                "Дождитесь завершения проверки API-ключей.",
            )
            event.ignore()
            return

        if self.modified:
            answer = russian_question(
                self,
                "Несохранённые изменения",
                "В таблице есть несохранённые изменения. Сохранить их?",
                ("Сохранить", QMessageBox.Save),
                ("Не сохранять", QMessageBox.Discard),
                ("Отмена", QMessageBox.Cancel),
            )
            if answer == QMessageBox.Save:
                self.save_changes()
                if self.modified:  # сохранение отменено/ошибка
                    event.ignore()
                    return
            elif answer == QMessageBox.Cancel:
                event.ignore()
                return
            # Discard — просто закрываем

        super().closeEvent(event)
