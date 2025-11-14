from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel, QDateEdit
from PySide6.QtCore import QDate

class DateSelector(QWidget):
    def __init__(self, label_text="Дата", default_date=None, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Жирная подпись через HTML-тег или стилем
        label = QLabel(f"<b>{label_text}</b>")
        # Альтернативно: label.setStyleSheet("font-weight: bold;")
        layout.addWidget(label)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        if default_date is None:
            default_date = QDate.currentDate()
        self.date_edit.setDate(default_date)
        layout.addWidget(self.date_edit)
        self.setLayout(layout)

    def get_date(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def set_date(self, date):
        if isinstance(date, QDate):
            self.date_edit.setDate(date)
        elif isinstance(date, str):
            # date: "YYYY-MM-DD"
            y, m, d = map(int, date.split('-'))
            self.date_edit.setDate(QDate(y, m, d))
