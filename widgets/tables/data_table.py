from typing import List

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtGui import QBrush, QColor
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
        self.setColumnWidth(DataTableColumns.SELECT.index, 30)
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
        self.set_default_checks()

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

        # hide columns DRIFT, RnS Error
        self.setColumnHidden(DataTableColumns.DRIFT.index, True)
        self.setColumnHidden(DataTableColumns.RNS_ERROR.index, True)

    def set_default_numbers(self):
        for i in range(self.rowCount()):
            item = TableWidgetItem(str(i + 1))
            self.setItem(
                i,
                DataTableColumns.NUMBER.index,
                item,
            )

    def set_default_checks(self):
        for i in range(self.rowCount()):
            checkbox = QtWidgets.QCheckBox()
            checkbox.setChecked(True)
            self.setCellWidget(i, DataTableColumns.SELECT.index, checkbox)

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
        elif event.matches(QtGui.QKeySequence.StandardKey.Paste):
            self.paste_data()
        # Ивент копирования ctrl-c
        elif event.matches(QtGui.QKeySequence.StandardKey.Copy):
            self.copy_data()
        else:
            super(DataTable, self).keyPressEvent(event)

    def copy_data(self):
        clipboard = QtWidgets.QApplication.clipboard()
        copied_cells = sorted(self.selectedIndexes())

        copy_text = ""
        max_column = copied_cells[-1].column()
        for c in copied_cells:
            copy_text += self.item(c.row(), c.column()).text()
            if c.column() == max_column:
                copy_text += "\n"
            else:
                copy_text += "\t"
        clipboard.setText(copy_text)

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
        values = []
        for row in range(self.rowCount()):
            if not self.cellWidget(row, DataTableColumns.SELECT.index).isChecked():
                continue
            value = self.item(row, column.index)
            try:
                values.append(column.dtype(value.text()))
            except (ValueError, AttributeError):
                values.append("")
        return values

    def get_column_value(self, row: int, column: DataTableColumns):
        if not self.cellWidget(row, DataTableColumns.SELECT.index).isChecked():
            return None
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
            checkbox = QtWidgets.QCheckBox()
            checkbox.setChecked(True)
            self.setCellWidget(row, DataTableColumns.SELECT.index, checkbox)
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

    def color_row(self, row, background_color, text_color):
        for col in range(self.columnCount()):
            item = self.item(row, col)
            if item:
                item.setBackground(QBrush(QColor(background_color)))
                item.setForeground(QBrush(QColor(text_color)))

    def dump_data(self):
        data = []
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):

                if col == DataTableColumns.SELECT.index:
                    cell_widget = self.cellWidget(row, col)
                    value = "True" if cell_widget.isChecked() else ""
                else:
                    item = self.item(row, col)
                    value = DataTableColumns.get_by_index(col).dtype(item.text()) if item.text() else ""

                data.append(InitialDataItem(value=value, row=row, col=col))
        return data

    def load_data(self, data: List[InitialDataItem]):
        for item in data:
            if item["col"] == DataTableColumns.SELECT.index:
                checkbox = QtWidgets.QCheckBox()
                checkbox.setChecked(item["value"] == "True")
                self.setCellWidget(item["row"], DataTableColumns.SELECT.index, checkbox)
            else:
                self.setItem(item["row"], item["col"], TableWidgetItem(f"{item['value']}"))
