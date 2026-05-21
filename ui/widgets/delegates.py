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


class RnSErrorPercentDelegate(RoundedDelegate):
    def __init__(self, rns_column: int, parent=None) -> None:
        super().__init__(rounded=2, parent=parent)
        self.rns_column = rns_column

    def displayText(self, value, locale):
        base = super().displayText(value, locale)
        try:
            error = float(value)
            table = self.parent()
            rns_item = table.item(0, self.rns_column) if table is not None else None
            rns = float(rns_item.text()) if rns_item and rns_item.text() else 0.0
            if not rns:
                return base
            return f"{base} ({abs(error) / abs(rns) * 100:.1f}%)"
        except Exception:
            return base
