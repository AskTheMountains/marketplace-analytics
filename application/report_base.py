from PySide6.QtWidgets import (
    QPushButton,
)

class ReportBase:
    report_name = ""
    report_dir = ""  # путь к папке с отчетами

    def __init__(self):
        self.status = "collapsed"  # collapsed / expanded / running / success / error
        self.status_text = ""      # текстовое описание для индикатора состояния
        self.result = None
        self.error = None
        self.params = self.default_params()

    def default_params(self):
        return {}

    def build_ui(self, parent_layout, on_params_changed):
        pass

    def make_standard_buttons(self, parent_workspace, key, status_label, progress_bar):
        """
        Возвращает список стандартных кнопок для панели отчёта.
        parent_workspace — ClientWorkspace, key — (client_idx, report_obj)
        status_label, progress_bar нужны для управления UI из run_report.
        """
        btns = []

        # Сформировать отчет
        btn_run = QPushButton("Сформировать отчёт")
        btn_run.setStyleSheet("font-weight: bold;")
        btn_run.clicked.connect(
            lambda: parent_workspace.run_report(self, key, status_label, progress_bar, btn_run)
        )
        btns.append(btn_run)

        # Показать логи
        btn_logs = QPushButton("Показать логи")
        btn_logs.setStyleSheet("font-weight: bold;")
        btn_logs.clicked.connect(
            lambda: parent_workspace.open_log_window(parent_workspace.client_idx, self)
        )
        btns.append(btn_logs)

        # Открыть папку с отчётами
        btn_folder = QPushButton("Открыть папку с отчётами")
        btn_folder.setStyleSheet("font-weight: bold;")
        btn_folder.clicked.connect(
            lambda: parent_workspace.open_report_folder(self)
        )
        btns.append(btn_folder)

        return btns

    def make_custom_buttons(self, parent_workspace, key, status_label=None, progress_bar=None):
        """
        Можно переопределять в потомках для добавления своих кнопок.
        """
        return []

    def is_params_ready(self):
        return True

    def build_cmd(self, client_number):
        return []
