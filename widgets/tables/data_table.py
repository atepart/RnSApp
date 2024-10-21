from typing import List

import numpy as np
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QHeaderView

from constants import DataTableColumns
from store import InitialDataItem
from widgets.delegates import ReadOnlyDelegate
from widgets.tables.mixins import TableMixin


class DataTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self, rows):
        super(DataTable, self).__init__(rows, len(DataTableColumns.get_all_names()))

        # Set Table headers
        self.setHorizontalHeaderLabels(DataTableColumns.get_all_names())
        self.setColumnWidth(DataTableColumns.NUMBER.index, 30)
        header = self.horizontalHeader()
        for col in DataTableColumns:
            if col == DataTableColumns.NUMBER:
                continue
            header.setSectionResizeMode(col.index, QHeaderView.Stretch)

        # Remove vertical Table headers
        self.verticalHeader().setVisible(False)

        # Set Table grid
        self.setShowGrid(True)
        self.setGridStyle(QtCore.Qt.PenStyle.SolidLine)

        # Set default Numbers
        self.set_default_numbers()

        # Set columns RnS, Rn as read-only
        self.set_read_only_columns(
            [
                DataTableColumns.RNS.index,
                DataTableColumns.RN.index,
            ]
        )

        # Connect event update_table
        # self.itemChanged.connect(self.update_table)

    def set_default_numbers(self):
        for i in range(self.rowCount()):
            self.setItem(
                i,
                DataTableColumns.NUMBER.index,
                QtWidgets.QTableWidgetItem(str(i + 1)),
            )

    def set_read_only_columns(self, columns):
        for col in columns:
            self.setItemDelegateForColumn(col, ReadOnlyDelegate(self))

    def update_table(self, item):
        self.itemChanged.disconnect(self.update_table)
        row = item.row()
        col = item.column()

        # Для рассчета нужны только колонки Diameter, Resistance
        if col not in (
            DataTableColumns.DIAMETER.index,
            DataTableColumns.RESISTANCE.index,
        ):
            self.itemChanged.connect(self.update_table)
            return

        # Достаем Resistance
        if self.item(row, DataTableColumns.RESISTANCE.index) is None:
            self.itemChanged.connect(self.update_table)
            return

        # Переводим Resistance в float
        try:
            resistance = DataTableColumns.RESISTANCE.dtype(self.item(row, DataTableColumns.RESISTANCE.index).text())
        except ValueError:
            self.itemChanged.connect(self.update_table)
            return

        if resistance != 0:  # Если Resistance != 0, то рассчитываем Rn
            rn_sqrt = 1 / np.sqrt(resistance)
            self.setItem(
                row,
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(str(round(rn_sqrt, 4))),
            )
        else:  # Иначе очищаем RnS, Rn
            self.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.setItem(
                row,
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
        self.itemChanged.connect(self.update_table)

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
                        DataTableColumns.RN.index,
                        DataTableColumns.DRIFT.index,
                    ]:  # Нельзя изменить Rn, RnS, Drift
                        self.setItem(item.row(), item.column(), QtWidgets.QTableWidgetItem(""))

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
                item = QtWidgets.QTableWidgetItem(value)
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
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
            self.setItem(
                row,
                DataTableColumns.NUMBER.index,
                QtWidgets.QTableWidgetItem(str(row + 1)),
            )  # Clear Number column

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
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column

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
            self.setItem(item["row"], item["col"], QtWidgets.QTableWidgetItem(f"{item['value']}"))
