# application/reports_ozon.py

import sys
from datetime import date
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QProgressBar
)
from PySide6.QtGui import QIntValidator
from PySide6.QtCore import QDate
from application.report_base import ReportBase
from application.widgets.calendar import DateSelector

class WBSalesOrdersReport(ReportBase):
    report_name = "Продажи и заказы"
    report_dir = "wb/Clients/{client_name}/SalesOrdersSvod"

    def default_params(self):
        return {
            "start_date": "",
            "end_date": "",
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_date_changed(on_params_changed)
            )
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_date_changed(on_params_changed)
            )
            layout.addWidget(self.start_selector)
            layout.addWidget(self.end_selector)

            # --- ВСЕГДА создаём status_label и progress_bar ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            layout.addWidget(status_label)

            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            layout.addWidget(progress_bar)

            # Стандартные кнопки
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)
            # Кастомные (если будут)
            if parent_workspace and key:
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)

        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _on_date_changed(self, cb):
        self.params["start_date"] = self.start_selector.get_date()
        self.params["end_date"] = self.end_selector.get_date()
        cb(self.params)

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        cmd = [
            sys.executable,
            "-m",
            "wb.scripts.create_sales_and_orders_svod",
            "-c", str(client_number)
        ]
        if self.params["start_date"] and self.params["end_date"]:
            cmd += [self.params["start_date"], self.params["end_date"]]
        return cmd


class WBUploadDataReport(ReportBase):
    report_name = "Выгрузка данных WB"
    report_dir = "wb/Clients/{client_name}/UploadFiles"

    def default_params(self):
        return {
            "days_ago": None,        # None — аргумент не будет передан
            "start_date": "",        # строка "YYYY-MM-DD"
            "end_date": "",          # строка "YYYY-MM-DD"
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            vbox = QVBoxLayout(self.panel_widget)
            # --- days_ago ---
            h_box = QHBoxLayout()
            label = QLabel("<b>Количество дней назад:</b>")
            self.days_edit = QLineEdit()
            self.days_edit.setPlaceholderText("30 (по умолчанию)")
            self.days_edit.setValidator(QIntValidator(1, 999))
            self.days_edit.textChanged.connect(lambda _: self._on_params_changed(on_params_changed))
            h_box.addWidget(label)
            h_box.addWidget(self.days_edit)
            vbox.addLayout(h_box)
            # --- календарь start_date ---
            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            vbox.addWidget(self.start_selector)
            # --- календарь end_date ---
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            vbox.addWidget(self.end_selector)

            # --- статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            vbox.addWidget(status_label)

            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            vbox.addWidget(progress_bar)

            # --- стандартные и кастомные кнопки (передаём status_label и progress_bar) ---
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    vbox.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    vbox.addWidget(btn)
        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _on_params_changed(self, cb):
        val = self.days_edit.text().strip()
        self.params["days_ago"] = val if val else None
        self.params["start_date"] = self.start_selector.get_date() or ""
        self.params["end_date"] = self.end_selector.get_date() or ""
        cb(self.params)

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        cmd = [
            sys.executable,
            "-m", "wb.scripts.uploadDataFromWB",
            "-c", str(client_number)
        ]
        days_ago = self.params.get("days_ago")
        start_date = self.params.get("start_date")
        end_date = self.params.get("end_date")
        # Приоритет: days_ago > даты > дефолт 30 внутри скрипта
        if days_ago:
            cmd.append(str(days_ago))
        elif start_date or end_date:
            if start_date:
                cmd += ["--start_date", start_date]
            if end_date:
                cmd += ["--end_date", end_date]
        # если всё пустое — скрипт сам возьмет days_ago=30
        return cmd
