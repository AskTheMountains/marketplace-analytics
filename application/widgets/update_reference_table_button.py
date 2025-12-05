from PySide6.QtWidgets import QPushButton, QFileDialog, QMessageBox, QDialog, QVBoxLayout, QLabel, QTextEdit
import subprocess
import sys

class UpdateReferenceTableButton(QPushButton):
    def __init__(self, parent, get_client_idx, get_client_name, get_marketplace_group, get_marketplace_name):
        super().__init__("Обновить справочную таблицу", parent)
        self.setStyleSheet(
            """
            QPushButton {
                background-color: #b6fcd5;
                color: #20603d;
                font-weight: bold;
                border-radius: 5px;
                border: 1px solid #7be5a5;
            }
            QPushButton:disabled {
                background-color: #e8fbe8;
                color: #aaaaaa;
            }
            QPushButton:hover {
                background-color: #a3f7c3;
            }
            """
        )
        self.setEnabled(False)  # будет включаться только при выбранном клиенте
        self.clicked.connect(self.update_reference_table_dialog)

        # Сохраняем внешние функции для получения нужных данных
        self.get_client_idx = get_client_idx
        self.get_client_name = get_client_name
        self.get_marketplace_group = get_marketplace_group
        self.get_marketplace_name = get_marketplace_name

    def update_reference_table_dialog(self):
        idx = self.get_client_idx()
        if idx is None:
            QMessageBox.warning(self, "Не выбран клиент", "Сначала выберите клиента!")
            return

        client_name = self.get_client_name()
        marketplace_group = self.get_marketplace_group()
        marketplace_name = self.get_marketplace_name()

        if not marketplace_group or not marketplace_name:
            QMessageBox.warning(self, "Ошибка", "Не определён маркетплейс!")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, f"Выберите справочную таблицу для {marketplace_name} (.xlsx)", "",
            "Excel files (*.xlsx);;All files (*)"
        )
        if not file_path:
            return

        try:
            process = subprocess.run(
                [
                    sys.executable, "-m",
                    f"{marketplace_group}.scripts.update_reference_table",
                    "--client", str(idx),
                    "--file", file_path
                ],
                capture_output=True,
                text=True,
            )
            if process.returncode == 0:
                QMessageBox.information(
                    self, "Успех",
                    f"Обновление справочной таблицы для клиента {client_name} на маркетплейсе {marketplace_name} прошло успешно."
                )
            else:
                self.show_error_dialog(process.stderr or process.stdout)
        except Exception as e:
            self.show_error_dialog(str(e))

    def show_error_dialog(self, error_text):
        dlg = QDialog(self)
        dlg.setWindowTitle("Ошибка обновления справочной таблицы")
        layout = QVBoxLayout(dlg)

        short_msg = QLabel("<b>Ошибка обновления справочной таблицы</b>")
        layout.addWidget(short_msg)

        btn_details = QPushButton("Подробнее")
        details_box = QTextEdit()
        details_box.setReadOnly(True)
        details_box.setVisible(False)
        details_box.setText(error_text)
        layout.addWidget(btn_details)
        layout.addWidget(details_box)

        def toggle_details():
            details_box.setVisible(not details_box.isVisible())
            btn_details.setText("Скрыть" if details_box.isVisible() else "Подробнее")
        btn_details.clicked.connect(toggle_details)

        # dlg.resize(600, 300)
        dlg.exec()
