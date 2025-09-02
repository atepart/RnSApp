from domain.constants import TableColumns
from ui.widgets.delegates import ReadOnlyDelegate


class TableMixin:
    def get_column_value(self, row: int, column: TableColumns):
        try:
            return column.dtype(self.item(row, column.index).text())
        except (ValueError, AttributeError):
            return None

    def get_column_values(self, column: TableColumns):
        values = []
        for row in range(self.rowCount()):
            value = self.item(row, column.index)
            try:
                values.append(column.dtype(value.text()))
            except (ValueError, AttributeError):
                values.append("")
        return values

    def set_read_only_columns(self, columns):
        for col in columns:
            self.setItemDelegateForColumn(col, ReadOnlyDelegate(self))
