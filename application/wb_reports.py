# application/reports_wb.py
import sys
import os
import importlib
import platform
import subprocess
from pathlib import Path

from application.paths import PYTHON_EXE
from application.report_base import ReportBase
from application.widgets.calendar import DateSelector
from application.widgets.dates_defaults import (
    month_date_range,
    prev_week_date_range,
    prev_month_date_range,
    last_n_days
)
from wb.scripts.constants import PROMOTION_COMPANIES
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QProgressBar,
    QFileDialog, QPushButton, QRadioButton, QButtonGroup, QTextEdit,
    QSizePolicy, QCheckBox, QMessageBox
)
from PySide6.QtGui import QIntValidator

class WBSalesOrdersReport(ReportBase):
    report_name = "Продажи и заказы (WB)"
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

            # --- Значения по умолчанию для дат ---
            start_str, end_str = month_date_range()
            self.start_selector.set_date(start_str)
            self.end_selector.set_date(end_str)
            # Обработчик для синхронизации:
            self._on_date_changed(on_params_changed)

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
            PYTHON_EXE,
            "-m",
            "wb.scripts.create_sales_and_orders_svod",
            "-c", str(client_number)
        ]
        if self.params["start_date"] and self.params["end_date"]:
            cmd += [self.params["start_date"], self.params["end_date"]]
        return cmd

class WBUploadDataReport(ReportBase):
    report_name = "Выгрузка данных (WB)"
    report_dir = "wb/Clients/{client_name}/UploadFiles"

    def default_params(self):
        return {
            "days_ago": None,
            "start_date": "",
            "end_date": "",
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            vbox = QVBoxLayout(self.panel_widget)

            # --- Переключатели режимов ---
            self.mode_group = QButtonGroup(self.panel_widget)
            mode_hbox = QHBoxLayout()
            self.days_radio = QRadioButton("По количеству дней назад")
            self.dates_radio = QRadioButton("По диапазону дат")
            self.mode_group.addButton(self.days_radio, 0)
            self.mode_group.addButton(self.dates_radio, 1)
            self.days_radio.setChecked(True)
            mode_hbox.addWidget(self.days_radio)
            mode_hbox.addWidget(self.dates_radio)
            vbox.addLayout(mode_hbox)

            self.days_radio.toggled.connect(lambda: self._update_mode(on_params_changed))

            # --- days_ago group ---
            self.days_group = QWidget()
            days_layout = QHBoxLayout(self.days_group)
            label = QLabel("<b>Количество дней назад:</b>")
            self.days_edit = QLineEdit()
            self.days_edit.setPlaceholderText("30 (по умолчанию)")
            self.days_edit.setValidator(QIntValidator(1, 999))
            self.days_edit.textChanged.connect(lambda _: self._on_params_changed(on_params_changed))
            days_layout.addWidget(label)
            days_layout.addWidget(self.days_edit)
            vbox.addWidget(self.days_group)

            # --- dates group ---
            self.dates_group = QWidget()
            dates_layout = QVBoxLayout(self.dates_group)
            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(lambda _: self._on_params_changed(on_params_changed))
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(lambda _: self._on_params_changed(on_params_changed))
            dates_layout.addWidget(self.start_selector)
            dates_layout.addWidget(self.end_selector)
            vbox.addWidget(self.dates_group)

            # --- Значения по умолчанию для дат ---
            start_str, end_str = last_n_days()
            self.start_selector.set_date(start_str)
            self.end_selector.set_date(end_str)
            # Обработчик для синхронизации:
            self._on_params_changed(on_params_changed)

            # --- статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            vbox.addWidget(status_label)

            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            vbox.addWidget(progress_bar)

            # --- стандартные и кастомные кнопки ---
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    vbox.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    vbox.addWidget(btn)

            # --- включение-выключение блоков по умолчанию ---
            self._update_mode(on_params_changed)

        already_in_layout = any(
        parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _update_mode(self, cb):
        # Radio с днями
        if self.days_radio.isChecked():
            self.days_group.setEnabled(True)
            self.days_group.setStyleSheet("")
            self.dates_group.setEnabled(False)
            self.dates_group.setStyleSheet("color: grey;")
            # Сброс дат
            self.start_selector.setEnabled(False)
            self.end_selector.setEnabled(False)
            self.start_selector.setStyleSheet("color: grey;")
            self.end_selector.setStyleSheet("color: grey;")
            # Чистим параметры, кроме дней
            self.params["start_date"] = ""
            self.params["end_date"] = ""
        # Radio с датами
        else:
            self.days_group.setEnabled(False)
            self.days_group.setStyleSheet("color: grey;")
            self.dates_group.setEnabled(True)
            self.dates_group.setStyleSheet("")
            self.start_selector.setEnabled(True)
            self.end_selector.setEnabled(True)
            self.start_selector.setStyleSheet("")
            self.end_selector.setStyleSheet("")
            # Чистим параметры, кроме дат
            self.params["days_ago"] = None
        self._on_params_changed(cb)

    def _on_params_changed(self, cb):
        # Обновляем параметры только из активных полей
        if self.days_radio.isChecked():
            val = self.days_edit.text().strip()
            self.params["days_ago"] = val if val else None
            self.params["start_date"] = ""
            self.params["end_date"] = ""
        else:
            self.params["days_ago"] = None
            self.params["start_date"] = self.start_selector.get_date() or ""
            self.params["end_date"] = self.end_selector.get_date() or ""
        cb(self.params)

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        cmd = [
            PYTHON_EXE,
            "-m", "wb.scripts.uploadDataFromWB",
            "-c", str(client_number)
        ]
        days_ago = self.params.get("days_ago")
        start_date = self.params.get("start_date")
        end_date = self.params.get("end_date")
        if days_ago:
            cmd.append(str(days_ago))
        elif start_date or end_date:
            if start_date:
                cmd += ["--start_date", start_date]
            if end_date:
                cmd += ["--end_date", end_date]
        return cmd

class WBSupplyPlanningReport(ReportBase):
    report_name = "Планирование поставок (WB)"
    report_dir = "wb/Clients/{client_name}/SupplySvod"

    def default_params(self):
        return {
            "reference_date": "",  # дата планирования поставки
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            # Календарь выбора даты
            self.date_edit = DateSelector("<b>Дата отчёта:</b>")
            self.date_edit.date_edit.dateChanged.connect(
                lambda _: self._on_date_changed(on_params_changed)
            )
            layout.addWidget(self.date_edit)


            # Статус и прогресс
            status_label = QLabel(getattr(self, 'status_text', ''))
            layout.addWidget(status_label)

            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            layout.addWidget(progress_bar)

            # Дефолтные кнопки
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)

        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _on_date_changed(self, cb):
        self.params["reference_date"] = self.date_edit.get_date()
        cb(self.params)

    def is_params_ready(self):
        return True  # или проверку на обязательность даты, если нужно

    def build_cmd(self, client_number):
        # calcMetrics — первый скрипт (метрики)
        cmd_metrics = [
            PYTHON_EXE, "-u",  # -u для небуферизованного вывода
            "-m", "wb.scripts.calcMetrics",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_metrics += ["--ref_dt", self.params["reference_date"]]

        # create_supply_svod — второй скрипт (сам отчет)
        cmd_supply = [
            PYTHON_EXE, "-u",
            "-m", "wb.scripts.create_supply_svod",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_supply.append(self.params["reference_date"])

        return [cmd_metrics, cmd_supply]

class WBActionsReport(ReportBase):
    report_name = "Отчет по акциям (WB)"
    report_dir = "wb/Clients/{client_name}/Actions"

    def default_params(self):
        return {
            "reference_date": "",
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):

        # if parent_workspace is not None:
        #     self.parent_workspace = parent_workspace

        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            # --- Календарь выбора даты ---
            self.date_edit = DateSelector("<b>Дата отчёта:</b>")
            self.date_edit.date_edit.dateChanged.connect(
                lambda _: self._on_date_changed(on_params_changed)
            )
            layout.addWidget(self.date_edit)

            # --- Кастомная кнопка открытия папки с данными по акциям ---
            btn_open_actions_dir = QPushButton("Открыть папку с данными по акциям")
            btn_open_actions_dir.setStyleSheet("font-weight: bold;")
            btn_open_actions_dir.clicked.connect(self._open_actions_data_dir)
            layout.addWidget(btn_open_actions_dir)

            # --- Статус, прогресс и стандартные/кастомные кнопки ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            layout.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            layout.addWidget(progress_bar)

            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)

        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

        # self.params["reference_date"] = self.date_edit.get_date()
        # или так, если нужно вызывать обработчик (чтобы cb обновился сразу после build_ui)
        self._on_date_changed(lambda x: None)

    def _on_date_changed(self, cb):
        self.params["reference_date"] = self.date_edit.get_date()
        cb(self.params)

    def _open_actions_data_dir(self):
        date_str = self.params.get("reference_date", "")
        if not date_str:
            QMessageBox.warning(self.panel_widget, "Не выбрана дата", "Сначала выберите дату!")
            return

        client_name = ""
        if hasattr(self, "parent_workspace"):
            pw = self.parent_workspace
            if hasattr(pw, "settings") and hasattr(pw, "client_idx"):
                client_name = pw.settings["client_name"][pw.client_idx]
        # Fallback: если не нашли через parent_workspace
        if not client_name:
            client_name = self.params.get("client_name", "")

        directory = Path(f"wb/Clients/{client_name}/Actions/Данные по акциям/{date_str}")

        if not directory.exists():
            msg = QMessageBox(self.panel_widget)
            msg.setWindowTitle("Создать папку?")
            msg.setText(f"Папка\n{directory}\n\nне существует. Создать и открыть?")
            msg.setIcon(QMessageBox.Question)
            yes_btn = msg.addButton("Да", QMessageBox.YesRole)
            no_btn = msg.addButton("Нет", QMessageBox.NoRole)
            msg.setDefaultButton(yes_btn)
            msg.exec()

            if msg.clickedButton() == yes_btn:
                directory.mkdir(parents=True, exist_ok=True)
            else:
                return

        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(str(directory))
            elif system == "Darwin":
                subprocess.run(["open", str(directory)])
            else:
                subprocess.run(["xdg-open", str(directory)])
        except Exception as e:
            QMessageBox.warning(self.panel_widget, "Ошибка открытия", f"Не удалось открыть папку:\n{directory}\n\nОшибка: {e}")

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        cmd_metrics = [
            PYTHON_EXE, "-u",  # -u для небуферизованного вывода
            "-m", "wb.scripts.calcMetrics",
            "-c", str(client_number)
        ]
        cmd_actions = [
            PYTHON_EXE, "-u",
            "-m", "wb.scripts.create_actions_svod",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_actions.append(self.params["reference_date"])
        if self.params.get("actions_file_path"):
            cmd_actions += ["--file", self.params["actions_file_path"]]
        return [cmd_metrics, cmd_actions]

class WBCampaignsReport(ReportBase):
    report_name = "Отчет по рекламным кампаниям (WB)"
    report_dir = "wb/Clients/{client_name}/CampaignsReport"

    def default_params(self):
        return {
            "start_date": "",
            "end_date": "",
            "promotion_companies": "",
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def _get_default_companies(self, client_name):
        """
        Получить дефолтные номера кампаний для клиента из wb.scripts.constants.
        """
        try:
            comp_map = PROMOTION_COMPANIES
            if comp_map:
                companies = comp_map.get(client_name, [])
                return "\n".join(str(c) for c in companies)
        except Exception:
            pass
        return ""

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            vbox = QVBoxLayout(self.panel_widget)

            # --- Два календаря ---
            dates_hbox = QVBoxLayout()
            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            dates_hbox.addWidget(self.start_selector)
            dates_hbox.addWidget(self.end_selector)
            vbox.addLayout(dates_hbox)

            # --- Текстовое поле для кампаний ---
            comp_label = QLabel("<b>ID Рекламных кампаний:</b>")
            vbox.addWidget(comp_label)
            self.txt_companies = QTextEdit()
            self.txt_companies.setPlaceholderText("Вставьте список номеров кампаний (через Enter, пробел, запятую или ;)...")
            self.txt_companies.setFixedHeight(70)
            self.txt_companies.setLineWrapMode(QTextEdit.NoWrap)
            self.txt_companies.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

            # Фильтрация только чисел и допустимых разделителей
            self.txt_companies.textChanged.connect(
                lambda: self._filter_companies_input(on_params_changed)
            )
            vbox.addWidget(self.txt_companies)

            # --- Значения по умолчанию для кампаний ---
            default_text = ""
            if parent_workspace and hasattr(parent_workspace, "settings"):
                cur_idx = getattr(parent_workspace, "client_idx", 0)
                cli_name = parent_workspace.settings["client_name"][cur_idx]
                default_text = self._get_default_companies(cli_name)
            if default_text and not self.params["promotion_companies"]:
                self.txt_companies.setPlainText(default_text)
                self.params["promotion_companies"] = default_text

            # --- Значения по умолчанию для дат ---
            start_str, end_str = prev_week_date_range()
            self.start_selector.set_date(start_str)
            self.end_selector.set_date(end_str)
            # Обработчик для синхронизации:
            self._on_params_changed(on_params_changed)

            # --- статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            vbox.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            vbox.addWidget(progress_bar)
            # --- стандартные и кастомные кнопки ---
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

    def _filter_companies_input(self, cb):
        cur = self.txt_companies.textCursor()
        raw = self.txt_companies.toPlainText()
        # print(self.txt_companies.toPlainText().encode("unicode_escape"))
        # Оставляем только цифры, ; , пробелы и переводы строк
        filtered = ''.join([ch for ch in raw if ch.isdigit() or ch in ",; \n\r"])
        if raw != filtered:
            pos = cur.position()
            self.txt_companies.blockSignals(True)
            self.txt_companies.setPlainText(filtered)
            self.txt_companies.blockSignals(False)
            cur.setPosition(min(pos, len(filtered)))
            self.txt_companies.setTextCursor(cur)
        self._on_params_changed(cb)

    def _on_params_changed(self, cb):
        self.params["start_date"] = self.start_selector.get_date() or ""
        self.params["end_date"] = self.end_selector.get_date() or ""
        companies = self.txt_companies.toPlainText()
        self.params["promotion_companies"] = companies
        cb(self.params)

    def is_params_ready(self):
            # Если нужно обязательное заполнение дат/номеров — добавь здесь проверку!
            return True

    def build_cmd(self, client_number):
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "wb.scripts.create_campaigns_report_common",
            "-c", str(client_number)
        ]
        sdate = self.params.get("start_date")
        edate = self.params.get("end_date")
        if sdate:
            cmd += ["--start_date", sdate]
        if edate:
            cmd += ["--end_date", edate]
        comps = self.params.get("promotion_companies", "")
        if comps.strip():
            cmd += ["--promotion_companies", comps.strip()]
        return cmd

class WBOrdersSvodReport(ReportBase):
    report_name = "Свод по заказам (WB)"
    report_dir = "wb/Clients/{client_name}/OrdersSvod"

    def default_params(self):
        return {
            "use_dates": False,    # включать параметры дат?
            "start_date": "",
            "end_date": "",
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            vbox = QVBoxLayout(self.panel_widget)

            # --- чекбокс включения периода дат ---
            self.dates_checkbox = QCheckBox("Указать диапазон дат")
            self.dates_checkbox.setChecked(False)
            vbox.addWidget(self.dates_checkbox)

            # --- область с календарями ---
            dates_box = QHBoxLayout()
            self.start_selector = DateSelector("Дата начала:")
            self.end_selector = DateSelector("Дата окончания:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_params_changed(on_params_changed)
            )
            # Изначально календарь не активен
            self.start_selector.setEnabled(False)
            self.end_selector.setEnabled(False)
            dates_box.addWidget(self.start_selector)
            dates_box.addWidget(self.end_selector)
            vbox.addLayout(dates_box)

            # --- обработчик включения-выключения календаря ---
            self.dates_checkbox.toggled.connect(lambda state: self._on_dates_checkbox_toggled(state, on_params_changed))

            # --- статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            vbox.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            vbox.addWidget(progress_bar)

            # --- стандартные и кастомные кнопки ---
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

        # Изначально область с календарями неактивна
        self._on_dates_checkbox_toggled(self.dates_checkbox.isChecked(), on_params_changed)

    def _on_dates_checkbox_toggled(self, state, cb):
        enabled = bool(state)
        self.start_selector.setEnabled(enabled)
        self.end_selector.setEnabled(enabled)
        self.params["use_dates"] = enabled

        # Явная подсветка серым для неактивного состояния
        # (можно использовать "grey", но для темных тем лучше #bbbbbb)
        grey_color = "grey"  # или "#bbbbbb"
        if not enabled:
            self.start_selector.setStyleSheet(f"color: {grey_color};")
            self.end_selector.setStyleSheet(f"color: {grey_color};")
            self.params["start_date"] = ""
            self.params["end_date"] = ""
        else:
            self.start_selector.setStyleSheet("")
            self.end_selector.setStyleSheet("")

        self._on_params_changed(cb)


    def _on_params_changed(self, cb):
        if self.dates_checkbox.isChecked():
            self.params["start_date"] = self.start_selector.get_date() or ""
            self.params["end_date"] = self.end_selector.get_date() or ""
        else:
            self.params["start_date"] = ""
            self.params["end_date"] = ""
        cb(self.params)

    def is_params_ready(self):
        # Если добавить проверку на даты — тут!
        return True

    def build_cmd(self, client_number):
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "wb.scripts.create_orders_svod_by_days",
            "-c", str(client_number)
        ]
        if self.params.get("use_dates"):
            sdate = self.params.get("start_date")
            edate = self.params.get("end_date")
            if sdate:
                cmd += ["--start_date", sdate]
            if edate:
                cmd += ["--end_date", edate]

        return cmd

class WBRealizationReport(ReportBase):
    report_name = "Продажи по реализации (WB)"
    report_dir = "wb/Clients/{client_name}/Realization"

    def default_params(self):
        return {
            "start_date": "",
            "end_date": "",
            "save_as_excel": False,
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            # --- Календарь выбора начальной и конечной даты ---
            dates_hbox = QVBoxLayout()
            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_dates_changed(on_params_changed)
            )
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_dates_changed(on_params_changed)
            )
            dates_hbox.addWidget(self.start_selector)
            dates_hbox.addWidget(self.end_selector)
            layout.addLayout(dates_hbox)

            # --- Значения по умолчанию для дат ---
            start_str, end_str = prev_month_date_range()
            self.start_selector.set_date(start_str)
            self.end_selector.set_date(end_str)

            # --- Чекбокс сохранения в файл Excel ---
            self.save_excel_checkbox = QCheckBox("Сохранить результат в файл Excel")
            self.save_excel_checkbox.setChecked(False)
            self.save_excel_checkbox.toggled.connect(
                lambda _: self._on_dates_changed(on_params_changed)
            )
            layout.addWidget(self.save_excel_checkbox)

            # Обработчик для синхронизации:
            self._on_dates_changed(on_params_changed)

            # --- Статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            layout.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            layout.addWidget(progress_bar)

            # --- Дефолтные кнопки ---
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)

        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _on_dates_changed(self, cb):
        self.params["start_date"] = self.start_selector.get_date()
        self.params["end_date"] = self.end_selector.get_date()
        checkbox = getattr(self, 'save_excel_checkbox', None)
        self.params["save_as_excel"] = checkbox.isChecked() if checkbox else self.params.get("save_as_excel", False)
        cb(self.params)

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "wb.scripts.get_realization",
            "-c", str(client_number)
        ]
        if self.params.get("start_date"):
            cmd += ["--start_date", self.params["start_date"]]
        if self.params.get("end_date"):
            cmd += ["--end_date", self.params["end_date"]]
        if self.params.get("save_as_excel"):
            cmd += ["--save_as_excel"]
        return cmd

class WBFinanceReport(ReportBase):
    report_name = "Отчет маржинальности (WB)"
    report_dir = "wb/Clients/{client_name}/FinanceReports"

    def default_params(self):
        return {
            "start_date": "",         # новая пара: начало
            "end_date": "",           # и конец
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            # --- Календарь выбора начальной и конечной даты ---
            dates_hbox = QVBoxLayout()
            self.start_selector = DateSelector("Дата начала:")
            self.start_selector.date_edit.dateChanged.connect(
                lambda _: self._on_dates_changed(on_params_changed)
            )
            self.end_selector = DateSelector("Дата окончания:")
            self.end_selector.date_edit.dateChanged.connect(
                lambda _: self._on_dates_changed(on_params_changed)
            )
            dates_hbox.addWidget(self.start_selector)
            dates_hbox.addWidget(self.end_selector)
            layout.addLayout(dates_hbox)
            # --- Значения по умолчанию для дат ---
            start_str, end_str = prev_month_date_range()
            self.start_selector.set_date(start_str)
            self.end_selector.set_date(end_str)
            # Обработчик для синхронизации:
            self._on_dates_changed(on_params_changed)

            # --- Статус и прогрессбар ---
            status_label = QLabel(getattr(self, 'status_text', ''))
            layout.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            layout.addWidget(progress_bar)

            # --- Дефолтные кнопки ---
            if parent_workspace and key:
                for btn in super().make_standard_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)
                for btn in self.make_custom_buttons(parent_workspace, key, status_label, progress_bar):
                    layout.addWidget(btn)

        already_in_layout = any(
            parent_layout.itemAt(i).widget() is self.panel_widget for i in range(parent_layout.count())
        )
        if not already_in_layout:
            parent_layout.addWidget(self.panel_widget)

    def _on_dates_changed(self, cb):
        self.params["start_date"] = self.start_selector.get_date()
        self.params["end_date"] = self.end_selector.get_date()
        cb(self.params)

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "wb.scripts.calc_svod_finance_report_WB_v3",
            "-c", str(client_number)
        ]
        if self.params.get("start_date"):
            cmd += ["--start_date", self.params["start_date"]]
        if self.params.get("end_date"):
            cmd += ["--end_date", self.params["end_date"]]
        return cmd
