from PySide6 import QtCore, QtWidgets


class TableWidgetItem(QtWidgets.QTableWidgetItem):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        # Ensure the item is editable unless a read-only delegate is set
        self.setFlags(self.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
