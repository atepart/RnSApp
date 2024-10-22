from typing import List

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QHeaderView

from constants import DataTableColumns
from store import InitialDataItem
from widgets.delegates import RoundedDelegate
from widgets.tables.item import TableWidgetItem
from widgets.tables.mixins import TableMixin


class DataTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self, rows):
        super(DataTable, self).__init__(rows, len(DataTableColumns.get_all_names()))

        # Set Table headers
        self.setHorizontalHeaderLabels(DataTableColumns.get_all_names())
        self.setColumnWidth(DataTableColumns.NUMBER.index, 30)
        self.setColumnWidth(DataTableColumns.NAME.index, 160)
        self.setColumnWidth(DataTableColumns.RESISTANCE.index, 160)
        self.setColumnWidth(DataTableColumns.RNS.index, 100)
        self.setColumnWidth(DataTableColumns.RN_SQRT.index, 100)
        header = self.horizontalHeader()
        header.setSectionResizeMode(DataTableColumns.DIAMETER.index, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(DataTableColumns.DRIFT.index, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(DataTableColumns.SQUARE.index, QHeaderView.ResizeMode.ResizeToContents)

        # Remove vertical Table headers
        self.verticalHeader().setVisible(False)

        # Set Table grid
        self.setShowGrid(True)
        self.setGridStyle(QtCore.Qt.PenStyle.SolidLine)

        # Set default Numbers
        self.set_default_numbers()

        # Set columns RnS, Rn, Drift, Square as read-only
        self.set_read_only_columns(
            [
                DataTableColumns.RNS.index,
                DataTableColumns.RN_SQRT.index,
                DataTableColumns.DRIFT.index,
                DataTableColumns.SQUARE.index,
            ]
        )

        self.setItemDelegateForColumn(DataTableColumns.DRIFT.index, RoundedDelegate(rounded=3, parent=self))
        self.setItemDelegateForColumn(DataTableColumns.RNS.index, RoundedDelegate(rounded=1, parent=self))
        self.setItemDelegateForColumn(DataTableColumns.SQUARE.index, RoundedDelegate(rounded=2, parent=self))
        self.setItemDelegateForColumn(DataTableColumns.RN_SQRT.index, RoundedDelegate(rounded=2, parent=self))

    def set_default_numbers(self):
        for i in range(self.rowCount()):
            self.setItem(
                i,
                DataTableColumns.NUMBER.index,
                TableWidgetItem(str(i + 1)),
            )

    def keyPressEvent(self, event):
        # На нажатие Enter/Return переход на следующую строку
        if event.key() in (QtCore.Qt.Key.Key_Enter, QtCore.Qt.Key.Key_Return):
            row = self.currentRow()
            col = self.currentColumn()
            if row < self.rowCount() - 1:
                self.setCurrentCell(row + 1, col)

        # На нажатие Delete/Backspace удаление выбранных значений
        elif event.key() in (QtCore.Qt.Key.Key_Delete, QtCore.Qt.Key.Key_Backspace):
            selected_items = self.selectedItems()
            if selected_items:
                for item in selected_items:
                    if item.column() not in [
                        DataTableColumns.RNS.index,
                        DataTableColumns.RN_SQRT.index,
                        DataTableColumns.DRIFT.index,
                        DataTableColumns.SQUARE.index,
                    ]:  # Нельзя изменить Rn, RnS, Drift
                        self.setItem(item.row(), item.column(), TableWidgetItem(""))

        # Ивент вставки ctrl-v
        elif event.matches(QtGui.QKeySequence.Paste):
            self.paste_data()
        else:
            super(DataTable, self).keyPressEvent(event)

    def paste_data(self):
        clipboard = QtWidgets.QApplication.clipboard()
        data = clipboard.text()
        rows = data.split("\n")
        start_row = self.currentRow()
        start_col = self.currentColumn()
        if start_col not in [
            DataTableColumns.DIAMETER.index,
            DataTableColumns.RESISTANCE.index,
            DataTableColumns.NUMBER.index,
            DataTableColumns.NAME.index,
        ]:  # Можно вставлять только в Number, Name, Diameter, Resistance

            return
        for i, row in enumerate(rows):
            values = row.split("\t")
            for j, value in enumerate(values):
                if start_col in [
                    DataTableColumns.DIAMETER.index,
                    DataTableColumns.RESISTANCE.index,
                ]:  # Для данных колонок нужны числа float
                    value = value.replace(",", ".")
                item = TableWidgetItem(value)
                self.setItem(start_row + i, start_col + j, item)

    def get_column_values(self, column: DataTableColumns):
        return super().get_column_values(column)

    def get_column_value(self, row: int, column: DataTableColumns):
        return super().get_column_value(row, column)

    def clear_all(self):
        for row in range(self.rowCount()):
            self.setItem(
                row,
                DataTableColumns.NAME.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Name column
            self.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.setItem(
                row,
                DataTableColumns.DRIFT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Drift column
            self.setItem(
                row,
                DataTableColumns.DIAMETER.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Diameter column
            self.setItem(
                row,
                DataTableColumns.RESISTANCE.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Resistance column
            self.setItem(
                row,
                DataTableColumns.RN_SQRT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
            self.setItem(
                row,
                DataTableColumns.NUMBER.index,
                TableWidgetItem(str(row + 1)),
            )  # Clear Number column
            self.setItem(row, DataTableColumns.SQUARE.index, TableWidgetItem(""))  # Clear Square

    def clear_rn(self):
        for row in range(self.rowCount()):
            self.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.setItem(
                row,
                DataTableColumns.DRIFT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Drift column
            self.setItem(
                row,
                DataTableColumns.RESISTANCE.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Resistance column
            self.setItem(
                row,
                DataTableColumns.RN_SQRT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
            self.setItem(row, DataTableColumns.SQUARE.index, QtWidgets.QTableWidgetItem(""))  # Clear Square

    def clear_calculations(self):
        for row in range(self.rowCount()):
            self.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.setItem(
                row,
                DataTableColumns.DRIFT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Drift column
            self.setItem(
                row,
                DataTableColumns.RN_SQRT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
            self.setItem(row, DataTableColumns.SQUARE.index, QtWidgets.QTableWidgetItem(""))  # Clear Square

    def dump_data(self):
        data = []
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if not item:
                    item = ""
                else:
                    item = item.text()
                data.append(InitialDataItem(value=item, row=row, col=col))
        return data

    def load_data(self, data: List[InitialDataItem]):
        for item in data:
            self.setItem(item["row"], item["col"], TableWidgetItem(f"{item['value']}"))
