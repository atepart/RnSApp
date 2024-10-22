from PyQt5 import QtWidgets


class ReadOnlyDelegate(QtWidgets.QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return


class RoundedDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, rounded: int = 2, parent=None):
        super().__init__(parent)
        self.rounded = rounded

    def displayText(self, value, locale):
        try:
            return str(round(float(value), self.rounded))
        except ValueError:
            return super().displayText(value, locale)
