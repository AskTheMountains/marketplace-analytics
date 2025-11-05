# application/reports_ozon.py
import sys
import importlib

from application.paths import PYTHON_EXE
from application.report_base import ReportBase
from application.widgets.calendar import DateSelector
from ozon.scripts.constants import PROMOTION_COMPANIES
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QProgressBar,
    QFileDialog, QPushButton, QRadioButton, QButtonGroup, QTextEdit,
    QSizePolicy, QCheckBox
)
from PySide6.QtGui import QIntValidator

class OzonSalesOrdersReport(ReportBase):
    report_name = "Продажи и заказы"
    report_dir = "ozon/Clients/{client_name}/SalesOrdersSvod"

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
            PYTHON_EXE,
            "-m",
            "ozon.scripts.create_sales_and_orders_svod",
            "-c", str(client_number)
        ]
        if self.params["start_date"] and self.params["end_date"]:
            cmd += [self.params["start_date"], self.params["end_date"]]
        return cmd

class OzonUploadDataReport(ReportBase):
    report_name = "Выгрузка данных Ozon"
    report_dir = "ozon/Clients/{client_name}/UploadFiles"

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
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "ozon.scripts.uploadDataFromOzon",
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

class OzonSupplyPlanningReport(ReportBase):
    report_name = "Планирование поставок Ozon"
    report_dir = "ozon/Clients/{client_name}/SupplySvod"

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
            "-m", "ozon.scripts.calcMetrics",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_metrics += ["--ref_dt", self.params["reference_date"]]

        # create_supply_svod — второй скрипт (сам отчет)
        cmd_supply = [
            PYTHON_EXE, "-u",
            "-m", "ozon.scripts.create_supply_svod",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_supply.append(self.params["reference_date"])

        return [cmd_metrics, cmd_supply]

class OzonActionsReport(ReportBase):
    report_name = "Отчет по акциям (Ozon)"
    report_dir = "ozon/Clients/{client_name}/Actions"

    def default_params(self):
        return {
            "reference_date": "",
            "actions_file_path": ""  # строка! Path будем собирать при запуске
        }

    def __init__(self):
        super().__init__()
        self.panel_widget = None

    def build_ui(self, parent_layout, on_params_changed, parent_workspace=None, key=None):
        if self.panel_widget is None:
            self.panel_widget = QWidget()
            layout = QVBoxLayout(self.panel_widget)

            # --- Календарь выбора даты ---
            # date_label = QLabel("<b>Дата отчёта:</b>")
            self.date_edit = DateSelector("<b>Дата отчёта:</b>")  # свой кастомный календарик!
            self.date_edit.date_edit.dateChanged.connect(
                lambda _: self._on_date_changed(on_params_changed)
            )
            layout.addWidget(self.date_edit)

            # --- Строка для пути к файлу + кнопка выбора ---
            hbox = QHBoxLayout()
            file_label = QLabel("<b>Путь к файлу акций:</b>")
            self.file_line = QLineEdit()
            self.file_line.setPlaceholderText("Укажите путь к .xlsx (не обязательно)")
            self.file_line.textChanged.connect(
                lambda _: self._on_file_path_changed(on_params_changed)
            )
            hbox.addWidget(file_label)
            hbox.addWidget(self.file_line)

            btn_browse = QPushButton("...")
            btn_browse.setFixedWidth(35)
            btn_browse.clicked.connect(self._browse_file)
            hbox.addWidget(btn_browse)
            layout.addLayout(hbox)

            # --- Статус, прогресс и стандартные/кастомные кнопки как у остальных ---
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

    def _on_date_changed(self, cb):
        self.params["reference_date"] = self.date_edit.get_date()
        cb(self.params)

    def _on_file_path_changed(self, cb):
        self.params["actions_file_path"] = self.file_line.text().strip()
        cb(self.params)

    def _browse_file(self):
        # Универсальный файловый диалог с фильтром на xlsx
        file_path, _ = QFileDialog.getOpenFileName(
            self.panel_widget, "Выберите файл с акциями из ЛК", "", "Excel files (*.xlsx);;All files (*)"
        )
        if file_path:
            self.file_line.setText(file_path)

    def is_params_ready(self):
        # Тут можно добавить проверки: нужно ли обязательное наличие даты, например
        return True

    def build_cmd(self, client_number):
        cmd_main = [
            PYTHON_EXE,
            "-u", "-m", "ozon.scripts.create_actions_svod_v2",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_main.append(self.params["reference_date"])
        if self.params.get("actions_file_path"):
            cmd_main += ["--file", self.params["actions_file_path"]]

        cmd_metrics = [
            PYTHON_EXE,
            "-u", "-m", "ozon.scripts.calcMetrics",
            "-c", str(client_number)
        ]
        if self.params.get("reference_date"):
            cmd_metrics += ["--ref_dt", self.params["reference_date"]]

        return [cmd_metrics, cmd_main]  # возвращаем две отдельные команды

class OzonCampaignsReport(ReportBase):
    report_name = "Отчет по рекламным кампаниям Ozon"
    report_dir = "ozon/Clients/{client_name}/ClientRKSvod"

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
        Получить дефолтные номера кампаний для клиента из ozon.scripts.constants.
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
                self.txt_companies.setPlaceholderText("Вставьте список номеров кампаний (через Enter, запятую или ;)...")
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
        print(self.txt_companies.toPlainText().encode("unicode_escape"))
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
            "-m", "ozon.scripts.create_rk_svod",
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

class OzonClientCabinetReport(ReportBase):
    report_name = "Сводная по кабинету клиента Ozon"
    report_dir = "ozon/Clients/{client_name}/ClientCabinetSvod"

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
        Получить дефолтные номера кампаний для клиента из ozon.scripts.constants.
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
                dates_hbox = QHBoxLayout()
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
                self.txt_companies.setPlaceholderText("Вставьте список номеров кампаний (через Enter, запятую или ;)...")
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
        print(self.txt_companies.toPlainText().encode("unicode_escape"))
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
            "-m", "ozon.scripts.create_client_cabinet_svod",
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

class OzonOrdersSvodReport(ReportBase):
    report_name = "Свод по заказам Ozon"
    report_dir = "ozon/Clients/{client_name}/OrdersSvod"

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

    def _on_dates_checkbox_toggled(self, state, cb):
        enabled = bool(state)
        self.start_selector.setEnabled(enabled)
        self.end_selector.setEnabled(enabled)
        self.params["use_dates"] = enabled
        # При отключении диапазона очищаем даты в параметрах
        if not enabled:
            self.params["start_date"] = ""
            self.params["end_date"] = ""
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
            "-m", "ozon.scripts.create_orders_svod_by_days",
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

class OzonFinanceReport(ReportBase):
    report_name = "Отчет маржинальности (Ozon)"
    report_dir = "ozon/Clients/{client_name}/FinanceReports"

    def default_params(self):
        return {
            "start_date": "",         # новая пара: начало
            "end_date": "",           # и конец
            "finance_report_file_path": "",  # путь к файлу отчета
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

            # --- Строка для пути к файлу + кнопка выбора ---
            hbox = QHBoxLayout()
            file_label = QLabel("<b>Путь к файлу отчета:</b>")
            self.file_line = QLineEdit()
            self.file_line.setPlaceholderText("Укажите путь к .xlsx (обязательно)")
            self.file_line.textChanged.connect(
                lambda _: self._on_file_path_changed(on_params_changed)
            )
            hbox.addWidget(file_label)
            hbox.addWidget(self.file_line)

            btn_browse = QPushButton("...")
            btn_browse.setFixedWidth(35)
            btn_browse.clicked.connect(self._browse_file)
            hbox.addWidget(btn_browse)
            layout.addLayout(hbox)

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

    def _on_file_path_changed(self, cb):
        text = self.file_line.text().strip()
        self.params["finance_report_file_path"] = text
        # Подсветка: если путь пустой — красная рамка, иначе обычная
        if not text:
            self.file_line.setStyleSheet("border: 2px solid red;")
        else:
            self.file_line.setStyleSheet("")
        cb(self.params)

    def _browse_file(self):
            file_path, _ = QFileDialog.getOpenFileName(
                self.panel_widget, "Выберите файл финансового отчета", "", "Excel files (*.xlsx);;All files (*)"
            )
            if file_path:
                self.file_line.setText(file_path)

    def is_params_ready(self):
        text = self.params.get("finance_report_file_path", "")
        if not text:
            # Дублируем подсветку на случай ручного запуска
            self.file_line.setStyleSheet(
                "background-color: rgba(255,128,128,0.3);"
            )
            return False
        self.file_line.setStyleSheet("")
        return True

    def build_cmd(self, client_number):
        from application.paths import PYTHON_EXE
        cmd = [
            PYTHON_EXE,
            "-m", "ozon.scripts.calc_svod_finance_report_new",
            "-c", str(client_number)
        ]
        if self.params.get("start_date"):
            cmd += ["--start_date", self.params["start_date"]]
        if self.params.get("end_date"):
            cmd += ["--end_date", self.params["end_date"]]
        if self.params.get("finance_report_file_path"):
            cmd += ["--file", self.params["finance_report_file_path"]]
        return cmd
