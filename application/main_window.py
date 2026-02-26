import sys
import json
import subprocess
import platform
import os
import sys
from pathlib import Path

THIS_DIR = Path(__file__).absolute().parent
BASE_DIR = THIS_DIR.parent

# Добавляем именно корень проекта в sys.path
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from application.paths import (
    BASE_DIR,
    SETTINGS_JSON,
    PYTHON_EXE,
    icon_path
)
# print("BASE_DIR:", BASE_DIR)
# print("sys.path:", sys.path)
# print("application exists:", (BASE_DIR / "application").exists())
# print("BASE_DIR in main_window.py:", BASE_DIR)
# print("sys.path:", sys.path)

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QTabWidget, QSpacerItem, QSizePolicy, QToolTip, QFrame,
    QProgressBar, QMessageBox, QDialog, QTextEdit
)
from PySide6.QtCore import Qt, QSize, QThread, Signal, QTimer
from PySide6.QtGui import QIcon, QPixmap
from application.widgets.update_reference_table_button import UpdateReferenceTableButton

from application.ozon_reports import (
    OzonSalesOrdersReport,
    OzonSupplyPlanningReport,
    OzonUploadDataReport,
    OzonActionsReport,
    OzonCampaignsReport,
    OzonOrdersSvodReport,
    OzonFinanceReport
    )
from application.wb_reports import (
    WBUploadDataReport,
    WBSalesOrdersReport,
    WBSupplyPlanningReport,
    WBActionsReport,
    WBCampaignsReport,
    WBOrdersSvodReport,
    WBFinanceReport,
)

# Список отчетов
REPORTS_REGISTRY = {
    "Ozon": [
        OzonSalesOrdersReport,
        OzonOrdersSvodReport,
        OzonCampaignsReport,
        OzonUploadDataReport,
        OzonSupplyPlanningReport,
        OzonActionsReport,
        OzonFinanceReport,
        # Другие классы Ozon-отчетов...
    ],
    "WB": [
        WBSalesOrdersReport,
        WBOrdersSvodReport,
        WBCampaignsReport,
        WBUploadDataReport,
        WBSupplyPlanningReport,
        WBActionsReport,
        WBFinanceReport
    ],
}

def load_settings():
    with open(SETTINGS_JSON, 'r', encoding='utf-8') as f:
        return json.load(f)

def get_client_display_name(settings, idx):
    ru_name = settings["client_ru_name"][idx].strip()
    name = settings["client_name"][idx]
    return ru_name if ru_name else name

def get_clients_list(settings):
    return [get_client_display_name(settings, idx) for idx in range(len(settings['client_name']))]

def check_ozon_active(settings, idx):
    keys = [
        settings["client_id_ozon"][idx].strip(),
        settings["api_key_ozon"][idx].strip(),
        settings["client_id_performance_ozon"][idx].strip(),
        settings["client_secret_performance_ozon"][idx].strip()
    ]
    return all(keys)

def check_wb_active(settings, idx):
    return bool(settings["api_key_wb"][idx].strip())

def gather_ozon_error(settings, idx):
    fields = {
        "client_id_ozon": "Client ID Ozon",
        "api_key_ozon": "API Key Ozon",
        "client_id_performance_ozon": "Client ID Performance Ozon",
        "client_secret_performance_ozon": "Client Secret Performance Ozon"
    }
    missing = []
    for key, label in fields.items():
        if not settings[key][idx].strip():
            missing.append(label)
    return "Не заполнены поля: " + ", ".join(missing)

def gather_wb_error(settings, idx):
    if not settings["api_key_wb"][idx].strip():
        return "Не заполнен API Key WB"
    return ""


class ReportThread(QThread):

    log_line = Signal(str)
    finished = Signal(bool, str, str)

    def __init__(self, cmds):
        super().__init__()
        self.cmds = cmds

    def run(self):
        all_logs = []

        if sys.platform == "win32":
            CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW
        else:
            CREATE_NO_WINDOW = 0

        for cmd in self.cmds:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=CREATE_NO_WINDOW
            )
            step_logs = []
            for line in proc.stdout:
                self.log_line.emit(line.rstrip('\n'))    # этот сигнал должен работать для КАЖДОЙ строки каждого скрипта
                step_logs.append(line.rstrip('\n'))
            proc.wait()
            all_logs.extend(step_logs)
            if proc.returncode != 0:
                self.finished.emit(False, f"Ошибка: код {proc.returncode}", "\n".join(all_logs))
                return
        self.finished.emit(True, "Готово", "\n".join(all_logs))

class ClientWorkspace(QWidget):
    def __init__(self, settings, clients, parent=None, tab_changed_callback=None):
        super().__init__(parent)
        self.settings = settings
        self.clients = clients
        self.tab_changed_callback = tab_changed_callback

        self.reports_registry = {}
        for group, report_classes in REPORTS_REGISTRY.items():
            self.reports_registry[group] = [cls() for cls in report_classes]

        self.param_panels = {}         # (client_idx, report_obj): panel
        self.report_statuses = {}      # (client_idx, report_obj): text
        self.open_reports = {group: set() for group in self.reports_registry.keys()}
        self.dots_timer = None
        self.dots_count = 0
        self.active_report = None

        self.report_logs = {}  # (client_idx, report_obj): str
        self.log_windows = {}  # (client_idx, report_obj): ссылка на окно (QDialog)

        self.report_threads = {}  #  хранилище потоков
        self.status_labels = {}   # (client_idx, report_obj): QLabel
        self.progress_bars = {}  # (client_idx, report_obj): QProgressBar
        self.run_btns = {}       # (client_idx, report_obj): QPushButton

        self.init_ui()
        self.update_buttons_and_labels()

    @property
    def client_idx(self):
        idx = self.clientCombo.currentIndex()
        if idx <= 0:
            return None  # Не выбран клиент
        return idx - 1

    def init_ui(self):
        layout = QVBoxLayout()

        # --- Кнопка "Темная тема" ---
        self.icon_moon = QIcon(QPixmap("application/src/images/moon.png"))
        self.icon_sun = QIcon(QPixmap("application/src/images/sun.png"))

        self.btn_theme = QPushButton()
        self.btn_theme.setIcon(self.icon_moon)
        self.btn_theme.setIconSize(QSize(24, 24))
        self.btn_theme.setCheckable(True)
        self.btn_theme.toggled.connect(self.switch_theme)
        self.btn_theme.setFixedSize(28, 28)
        self.btn_theme.setStyleSheet("""
            QPushButton {
                border-radius: 5px;
                background: #2b2b2b;
            }
            QPushButton:checked {
                background: #393939;
            }
        """)
        layout.addWidget(self.btn_theme)

        nameLabel = QLabel("<b>Клиент:</b>")
        layout.addWidget(nameLabel)

        self.clientCombo = QComboBox()

        # Добавляем в ComboBox нулевой элемент "Выберите клиента" — но в self.clients его нет!
        self.clientCombo.addItem("— Выберите клиента —")
        self.clientCombo.addItems(self.clients)
        self.clientCombo.setCurrentIndex(0)  # По умолчанию пустой выбор
        self.clientCombo.currentIndexChanged.connect(self.change_client)
        layout.addWidget(self.clientCombo)

        self.report_groups_widgets = {}
        btn_layout = QHBoxLayout()
        group_widgets = []

        report_btn_styles = {
            "Ozon": "background-color: #EAF2F8; color: #21618C; font-weight: 600;",
            "WB": "background-color: #f5e9fe; color: #6c3483; font-weight: 600;",
        }

        for group, reports in self.reports_registry.items():
            group_btn = QPushButton(f"{group}: отчёты")
            color = "#2471A3" if "ozon" in group.lower() else "#8e44ad"
            group_btn.setStyleSheet(f"background-color: {color}; color: white; font-weight: 700;")
            btn_layout.addWidget(group_btn)

            group_reports_widget = QWidget()
            group_reports_vbox = QVBoxLayout()
            group_reports_vbox.setContentsMargins(30, 0, 0, 0)

            for report in reports:
                row_widget = QWidget()
                row_layout = QVBoxLayout(row_widget)
                row_layout.setContentsMargins(0, 0, 0, 0)
                report_btn = QPushButton(report.report_name)
                report_btn.setStyleSheet(report_btn_styles.get(group, "font-weight: 600;"))
                report_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

                # Параметры для аккордеона
                panel_widget = QFrame()
                panel_widget.setFrameShape(QFrame.StyledPanel)
                panel_layout = QVBoxLayout()
                panel_widget.setLayout(panel_layout)
                panel_widget.setVisible(False)
                report_btn.clicked.connect(lambda _, r=report, p=panel_widget: self.toggle_report_panel(r, p))

                row_layout.addWidget(report_btn)
                row_layout.addWidget(panel_widget)
                group_reports_vbox.addWidget(row_widget)

                if not hasattr(self, "report_panels"):
                    self.report_panels = {}
                self.report_panels[(group, report)] = panel_widget

            group_reports_widget.setLayout(group_reports_vbox)
            group_reports_widget.setVisible(False)
            self.report_groups_widgets[group] = (group_btn, group_reports_widget)
            group_widgets.append(group_reports_widget)
            group_btn.clicked.connect(lambda _, g=group: self.show_report_group(g))
            # Кнопка обновления справочной таблицы
            btn_update_reference = UpdateReferenceTableButton(
                group_reports_widget,
                get_client_idx=lambda: self.client_idx,
                get_client_name=lambda: self.settings["client_name"][self.client_idx] if self.client_idx is not None else "",
                get_marketplace_group=lambda g=group: g.lower(),
                get_marketplace_name=lambda g=group: "Ozon" if g.lower()=="ozon" else "WB",
            )
            # Сохраняем для регулировки активности
            if not hasattr(self, "update_reference_buttons"):
                self.update_reference_buttons = {}
            self.update_reference_buttons[group.lower()] = btn_update_reference
            group_reports_vbox.addWidget(btn_update_reference, alignment=Qt.AlignRight | Qt.AlignBottom)


        layout.addLayout(btn_layout)
        for w in group_widgets:
            layout.addWidget(w)

        self.report_params_widget = QFrame()
        self.report_params_widget.setFrameShape(QFrame.StyledPanel)
        self.report_params_layout = QVBoxLayout()
        self.report_params_widget.setLayout(self.report_params_layout)
        layout.addWidget(self.report_params_widget)
        self.report_params_widget.setVisible(False)

        spacer = QSpacerItem(5, 50, QSizePolicy.Minimum, QSizePolicy.Expanding)
        layout.addSpacerItem(spacer)
        self.setLayout(layout)

        spacer = QSpacerItem(5, 50, QSizePolicy.Minimum, QSizePolicy.Expanding)
        layout.addSpacerItem(spacer)
        self.setLayout(layout)

    def switch_theme(self, checked):
        app = QApplication.instance()
        if checked:
            with open("application/themes/dark_theme.qss", "r", encoding="utf-8") as f:
                app.setStyleSheet(f.read())
            self.btn_theme.setIcon(self.icon_sun)   # Показываем солнышко при тёмной теме
            self.btn_theme.setToolTip("Переключить на светлую тему")
        else:
            app.setStyleSheet("")
            self.btn_theme.setIcon(self.icon_moon)  # Показываем луну при светлой теме
            self.btn_theme.setToolTip("Переключить на тёмную тему")

    def get_current_client_name(self):
        idx = self.client_idx
        if idx is None:
            return ""
        return self.clients[idx]

    def change_client(self, idx):
        # idx - индекс выбора в ComboBox
        if idx == 0:  # "— Выберите клиента —"
            self.hide_reports()
            return
        # Везде else работаем с остатком интерфейса через self.client_idx
        self.update_buttons_and_labels()
        self.hide_reports()
        if self.tab_changed_callback:
            self.tab_changed_callback(self)

    def update_buttons_and_labels(self):
        for group, (group_btn, group_reports_widget) in self.report_groups_widgets.items():
            idx = self.client_idx
            is_active = False
            if idx is None:
                is_active = False
                tip = "Сначала выберите клиента"
                color = "#bbb"
            elif group.lower() == "ozon":
                is_active = check_ozon_active(self.settings, idx)
                tip = "" if is_active else gather_ozon_error(self.settings, idx)
                color = "#2471A3"
            elif group.lower() == "wb":
                is_active = check_wb_active(self.settings, idx)
                tip = "" if is_active else gather_wb_error(self.settings, idx)
                color = "#8e44ad"
            else:
                is_active = True
                tip = ""
                color = "#aaa"
            group_btn.setEnabled(is_active)
            group_btn.setToolTip(tip)
            if is_active:
                group_btn.setStyleSheet(f"background-color: {color}; color: white; font-weight: 700;")
            else:
                group_btn.setStyleSheet("background-color: #bbb; color: white; font-weight: 700;")
            if not is_active:
                group_reports_widget.setVisible(False)
        # Кнопка обновления справочной таблицы
        for group in self.update_reference_buttons:
            idx = self.client_idx
            if group == 'ozon':
                self.update_reference_buttons[group].setEnabled(idx is not None and check_ozon_active(self.settings, idx))
            elif group == 'wb':
                self.update_reference_buttons[group].setEnabled(idx is not None and check_wb_active(self.settings, idx))

    def hide_reports(self):
        for _, (_, grp_widget) in self.report_groups_widgets.items():
            grp_widget.setVisible(False)
        for panel in self.param_panels.values():
            panel.setVisible(False)
        self.report_params_widget.setVisible(False)
        self.active_report = None

    def get_report_group(self, report_obj):
        for group, reports in self.reports_registry.items():
            if report_obj in reports:
                return group
        return None

    def open_report_folder(self, report_obj):
        idx = self.client_idx
        if idx is None:
            QMessageBox.warning(self, "Не выбран клиент", "Сначала выберите клиента!")
            return
        client_name = self.settings["client_name"][idx]
        path = Path(report_obj.report_dir.format(client_name=client_name)).resolve()
        if not path.exists():
            QMessageBox.warning(self, "Папка не найдена", f"Папка отчётов не существует:\n{str(path)}")
            return
        system = platform.system()
        if system == "Windows":
            subprocess.run(["explorer", str(path)])
        elif system == "Darwin":
            subprocess.run(["open", str(path)])
        else:
            subprocess.run(["xdg-open", str(path)])

    def append_log_line(self, key, line):
        self.report_logs[key] = self.report_logs.get(key, "") + line + "\n"
        print(f"[LOG][{key}]: {line}")
        if key in self.log_windows and self.log_windows[key] is not None:
            log_window = self.log_windows[key]
            if hasattr(log_window, "live_update_log"):
                log_window.live_update_log()

    def open_log_window(self, client_idx, report_obj):
        idx = client_idx
        if idx is None:
            QMessageBox.warning(self, "Не выбран клиент", "Сначала выберите клиента!")
            return
        key = (client_idx, report_obj)
        client_name = self.settings["client_name"][idx]
        report_name = getattr(report_obj, 'report_name', str(report_obj))
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{client_name}: {report_name}")
        layout = QVBoxLayout(dlg)
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setText(self.report_logs.get(key, ""))
        layout.addWidget(txt)
        dlg.resize(700, 500)
        dlg.show()

        self.log_windows[key] = dlg
        dlg.txt = txt

        def live_update_log():
            content = self.report_logs.get(key, "")
            txt.setPlainText(content)
            txt.verticalScrollBar().setValue(txt.verticalScrollBar().maximum())

        dlg.live_update_log = live_update_log
        live_update_log()

    def toggle_report_panel(self, report_obj, panel_widget):
        group = self.get_report_group(report_obj)
        for r in self.reports_registry[group]:
            p = self.report_panels.get((group, r))
            if p and p is not panel_widget:
                p.setVisible(False)
        panel_widget.setVisible(not panel_widget.isVisible())
        if panel_widget.isVisible():
            if panel_widget.layout().count() == 0:
                key = (self.client_idx, report_obj)
                report_obj.parent_workspace = self
                report_obj.build_ui(
                    panel_widget.layout(),
                    self.on_report_params_changed,
                    parent_workspace=self,
                    key=key
                )

    def show_report_group(self, group):
        for g, (_, widget) in self.report_groups_widgets.items():
            widget.setVisible(g == group)
        self.report_params_widget.setVisible(False)
        for (cli, ro), panel in self.param_panels.items():
            if cli == self.client_idx:
                panel.setVisible(False)
        for report_obj in self.open_reports[group]:
            key = (self.client_idx, report_obj)
            panel = self.param_panels.get(key)
            if panel:
                panel.setVisible(True)
                self.report_params_widget.setVisible(True)

    def show_report_params(self, report_obj):
        group = self.get_report_group(report_obj)
        key = (self.client_idx, report_obj)
        panel = self.param_panels.get(key)

        if panel and panel.isVisible():
            panel.setVisible(False)
            self.open_reports[group].discard(report_obj)
            if not any(
                self.param_panels.get((self.client_idx, r)) and self.param_panels[(self.client_idx, r)].isVisible()
                for r in self.reports_registry[group]
            ):
                self.report_params_widget.setVisible(False)
            return

        to_delete = []
        for (cli, ro), p in list(self.param_panels.items()):
            if ro == report_obj and cli != self.client_idx:
                self.report_params_layout.removeWidget(p)
                p.deleteLater()
                to_delete.append((cli, ro))
        for k in to_delete:
            del self.param_panels[k]

        for (cli, ro), p in self.param_panels.items():
            if cli == self.client_idx and self.get_report_group(ro) == group:
                p.setVisible(False)

        if key not in self.param_panels:
            panel = QFrame()
            panel.setFrameShape(QFrame.StyledPanel)
            panel_layout = QVBoxLayout(panel)
            report_obj.parent_workspace = self
            report_obj.build_ui(
                panel_layout,
                self.on_report_params_changed,
                parent_workspace=self,
                key=key
            )
            status_label = QLabel(getattr(report_obj, 'status_text', ''))
            panel_layout.addWidget(status_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            progress_bar.setVisible(False)
            panel_layout.addWidget(progress_bar)
            self.param_panels[key] = panel
            self.report_params_layout.addWidget(panel)

        self.open_reports[group].add(report_obj)
        self.report_params_widget.setVisible(True)
        self.param_panels[key].setVisible(True)
        self.active_report = report_obj
        self.active_key = key

    def on_report_params_changed(self, params):
        if self.active_report and hasattr(self, 'run_btn'):
            self.run_btn.setEnabled(self.active_report.is_params_ready())

    def animate_dots(self, status_label):
        self.dots_count = (self.dots_count + 1) % 4
        dots = '.' * self.dots_count
        status_label.setText(f"<b>Формируется{dots}</b>")
        if hasattr(self, '_report_key'):
            self.report_statuses[self._report_key] = f"<b>Формируется{dots}</b>"

    def run_report(self, report_obj, key, status_label, progress_bar, btn_run):
        if not report_obj:
            return

        client_number = self.client_idx
        if client_number is None:
            QMessageBox.warning(self, "Не выбран клиент", "Сначала выберите клиента!")
            btn_run.setEnabled(True)
            return

        cmds = report_obj.build_cmd(client_number)
        if isinstance(cmds[0], str):
            cmds = [cmds]

        # --- Управление потоками ---
        # Если поток с этим ключом уже есть, аккуратно удаляем старый
        if not hasattr(self, "report_threads"):
            self.report_threads = {}
        if key in self.report_threads:
            old_thread = self.report_threads.pop(key)
            old_thread.quit()
            old_thread.wait()
            old_thread.deleteLater()

        report_thread = ReportThread(cmds)
        self.report_threads[key] = report_thread
        report_thread.log_line.connect(lambda line, k=key: self.append_log_line(k, line))
        report_thread.finished.connect(
            lambda success, msg, log_text, k=key: self.on_report_finished(success, msg, log_text, k)
        )
        report_thread.start()

        # --- Статусы и UI ---
        if not report_obj.is_params_ready():
            if self.dots_timer:
                self.dots_timer.stop()
                self.dots_timer = None
            progress_bar.setVisible(False)
            msg = "<span style='color:#b30000; font-weight:bold;'>❗️ Заполните все обязательные параметры!</span>"
            status_label.setText(msg)
            btn_run.setEnabled(True)
            self.report_statuses[key] = msg
            return

        msg = "<b>Формируется</b>"
        status_label.setText(msg)
        status_label.setStyleSheet("font-weight: bold;")
        progress_bar.setVisible(True)
        btn_run.setEnabled(False)
        self.report_statuses[key] = msg

        if self.dots_timer:
            self.dots_timer.stop()
            self.dots_timer = None

        self.dots_timer = QTimer(self)
        self.dots_timer.timeout.connect(lambda: self.animate_dots(status_label))
        self.dots_count = 0
        self.dots_timer.start(400)

        self._status_label = status_label
        self._progress_bar = progress_bar
        self._run_btn = btn_run
        self._report_key = key
        self.report_logs[key] = ""


    def on_report_finished(self, success, msg, log_text, key):
        # Отключаем анимацию только если она относится к этому key
        if self.dots_timer:
            self.dots_timer.stop()

        # Прячем progress bar для конкретного отчёта
        if key in self.progress_bars:
            self.progress_bars[key].setVisible(False)

        # Разблокируем кнопку "Сформировать отчет"
        if key in self.run_btns:
            self.run_btns[key].setEnabled(True)

        # Обновляем статус в лейбле — тоже для key!
        if key in self.status_labels:
            contains_warning = "WARNING" in log_text or "ERROR" in log_text
            if success:
                if contains_warning:
                    text = "<span style='color:#e67e22; font-weight:bold;'>⚠️ Отчет сформирован с предупреждениями!</span>"
                else:
                    text = "<span style='color:#157f3b; font-weight:bold;'>✅ Готово!</span>"
            else:
                text = f"<span style='color:#b30000; font-weight:bold;'>❌ Ошибка формирования отчета, см. логи</span>"
            self.status_labels[key].setText(text)
            self.report_statuses[key] = text

        self.report_logs[key] = log_text

        if key in self.log_windows and self.log_windows[key] is not None:
            dlg = self.log_windows[key]
            if hasattr(dlg, "live_update_log"):
                dlg.live_update_log()

        # --- Безопасное завершение потока ---
        if hasattr(self, "report_threads") and key in self.report_threads:
            thread = self.report_threads.pop(key)
            thread.deleteLater()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Отчеты по маркетплейсам")
        self.settings = load_settings()
        self.clients = get_clients_list(self.settings)

        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.tabBar().tabBarClicked.connect(self.on_tab_clicked)

        layout = QVBoxLayout()
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

        self.add_client_tab(activate=True)
        self.add_plus_tab()

    def add_plus_tab(self):
        plus_label = QLabel("")
        self.tab_widget.addTab(plus_label, " + ")

    def add_client_tab(self, activate=False):
        tab = ClientWorkspace(
            self.settings,
            self.clients,
            tab_changed_callback=self.update_tab_name
        )
        tab_label = tab.get_current_client_name()
        plus_idx = self.find_plus_tab_index()
        insert_idx = plus_idx if plus_idx != -1 else self.tab_widget.count()
        self.tab_widget.insertTab(insert_idx, tab, tab_label)
        if activate:
            self.tab_widget.setCurrentWidget(tab)

    def close_tab(self, idx):
        if self.tab_widget.tabText(idx) == " + ":
            return
        work_tabs = [i for i in range(self.tab_widget.count()) if self.tab_widget.tabText(i) != " + "]
        if len(work_tabs) == 1 and work_tabs[0] == idx:
            return
        new_index = idx - 1 if idx - 1 >= 0 else 0
        self.tab_widget.removeTab(idx)
        count = self.tab_widget.count()
        plus_idx = self.find_plus_tab_index()
        if count > 1 and self.tab_widget.currentIndex() == plus_idx:
            if new_index == plus_idx:
                new_index = 0
            self.tab_widget.setCurrentIndex(new_index)

    def find_plus_tab_index(self):
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == " + ":
                return i
        return -1

    def on_tab_clicked(self, idx):
        if self.tab_widget.tabText(idx) == " + ":
            self.add_client_tab(activate=True)
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == " + " and i != self.tab_widget.count() - 1:
                    self.tab_widget.removeTab(i)
                    break
            self.tab_widget.setCurrentIndex(self.tab_widget.count() - 2)

    def update_tab_name(self, workspace_widget):
        idx = self.tab_widget.indexOf(workspace_widget)
        client_name = workspace_widget.get_current_client_name()
        if idx != -1:
            self.tab_widget.setTabText(idx, client_name)

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.setWindowIcon(QIcon(str(icon_path)))
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
