from PySide6 import QtWidgets
from PySide6.QtCore import Qt


class TableWidgetItem(QtWidgets.QTableWidgetItem):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
