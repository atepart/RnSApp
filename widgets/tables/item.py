from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt


class TableWidgetItem(QtWidgets.QTableWidgetItem):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
