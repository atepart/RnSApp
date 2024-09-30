from constants import TableColumns


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
            if not value:
                continue
            if not value.text():
                continue
            try:
                values.append(column.dtype(value.text()))
            except ValueError:
                continue
        return values
