from PySide6 import QtWidgets


class ReadOnlyDelegate(QtWidgets.QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return None


class RoundedDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, rounded: int, parent=None) -> None:
        super().__init__(parent)
        self.rounded = rounded

    def displayText(self, value, locale):
        try:
            return f"{round(float(value), self.rounded)}"
        except Exception:
            return value
