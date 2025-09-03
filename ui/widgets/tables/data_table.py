import contextlib
import logging
from typing import List

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QHeaderView

from domain.constants import DataTableColumns
from domain.models import InitialDataItem
from ui.widgets.delegates import RoundedDelegate
from ui.widgets.tables.item import TableWidgetItem
from ui.widgets.tables.mixins import TableMixin

logger = logging.getLogger(__name__)


class Header(QHeaderView):
    def __init__(self, parent) -> None:
        super(Header, self).__init__(QtCore.Qt.Orientation.Horizontal, parent)
        self.checkbox = QtWidgets.QCheckBox(self)
        self.checkbox.stateChanged.connect(self.on_state_changed)

    def on_state_changed(self, state):
        for row in range(self.parent().rowCount()):
            cb = self.parent().get_row_checkbox(row)
            if cb is not None:
                cb.setChecked(state)

    def paintSection(self, painter, rect, logicalIndex):
        painter.save()
        super().paintSection(painter, rect, logicalIndex)
        painter.restore()
        if logicalIndex == DataTableColumns.SELECT.index:
            self.checkbox.setGeometry(rect)


class DataTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self, rows) -> None:
        super(DataTable, self).__init__(rows, len(DataTableColumns.get_all_names()))
        self.header = Header(self)
        self.setHorizontalHeader(self.header)
        self.setHorizontalHeaderLabels(DataTableColumns.get_all_names())
        self.setColumnWidth(DataTableColumns.NAME.index, 160)
        self.setColumnWidth(DataTableColumns.RESISTANCE.index, 160)
        self.setColumnWidth(DataTableColumns.RNS.index, 100)
        self.setColumnWidth(DataTableColumns.RN_SQRT.index, 100)

        self.header.setSectionResizeMode(DataTableColumns.NUMBER.index, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(DataTableColumns.SELECT.index, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(DataTableColumns.DIAMETER.index, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(DataTableColumns.DRIFT.index, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(DataTableColumns.SQUARE.index, QHeaderView.ResizeMode.ResizeToContents)

        self.verticalHeader().setVisible(False)
        self.setShowGrid(True)
        self.setGridStyle(QtCore.Qt.PenStyle.SolidLine)

        self.set_default_numbers()
        self.set_default_checks()

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

        self.setColumnHidden(DataTableColumns.DRIFT.index, True)
        self.setColumnHidden(DataTableColumns.RNS_ERROR.index, True)

        self.itemChanged.connect(self.on_item_changed)
        self.clear_all()

    def on_item_changed(self, item):
        if item.column() in (DataTableColumns.DIAMETER.index, DataTableColumns.RESISTANCE.index):
            row = item.row()
            cb = self.get_row_checkbox(row)
            if cb:
                cb.setChecked(bool(item.text()))

    def set_default_numbers(self):
        for i in range(self.rowCount()):
            item = TableWidgetItem(str(i + 1))
            self.setItem(i, DataTableColumns.NUMBER.index, item)

    def set_default_checks(self):
        for i in range(self.rowCount()):
            checkbox = QtWidgets.QCheckBox()
            container = QtWidgets.QWidget()
            lay = QtWidgets.QHBoxLayout(container)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            lay.addWidget(checkbox)
            self.setCellWidget(i, DataTableColumns.SELECT.index, container)
            self.setItem(i, DataTableColumns.SELECT.index, QtWidgets.QTableWidgetItem(""))

    def get_row_checkbox(self, row: int) -> QtWidgets.QCheckBox | None:
        w = self.cellWidget(row, DataTableColumns.SELECT.index)
        if isinstance(w, QtWidgets.QCheckBox):
            return w
        if isinstance(w, QtWidgets.QWidget):
            cbs = w.findChildren(QtWidgets.QCheckBox)
            return cbs[0] if cbs else None
        return None

    def keyPressEvent(self, event):
        if event.key() in (QtCore.Qt.Key.Key_Enter, QtCore.Qt.Key.Key_Return):
            row = self.currentRow()
            col = self.currentColumn()
            if row < self.rowCount() - 1:
                self.setCurrentCell(row + 1, col)
        elif event.key() in (QtCore.Qt.Key.Key_Delete, QtCore.Qt.Key.Key_Backspace):
            selected_items = self.selectedItems()
            if selected_items:
                for item in selected_items:
                    if item.column() not in [
                        DataTableColumns.RNS.index,
                        DataTableColumns.RN_SQRT.index,
                        DataTableColumns.DRIFT.index,
                        DataTableColumns.SQUARE.index,
                    ]:
                        self.setItem(item.row(), item.column(), TableWidgetItem(""))
        elif event.matches(QtGui.QKeySequence.StandardKey.Paste):
            self.paste_data()
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
            col_def = DataTableColumns.get_by_index(c.column())
            cell_text = self.item(c.row(), c.column()).text()
            # Replace decimal point with comma only for float columns
            if col_def and col_def.dtype is float and cell_text:
                cell_text = cell_text.replace(".", ",")
            copy_text += cell_text
            if c.column() == max_column:
                copy_text += "\n"
            else:
                copy_text += "\t"
        # Avoid trailing empty line causing extra cleared row on paste elsewhere
        copy_text = copy_text.rstrip("\n")
        clipboard.setText(copy_text)

    def paste_data(self):
        clipboard = QtWidgets.QApplication.clipboard()
        data = clipboard.text()
        rows = data.split("\n")
        # Drop trailing empty line to prevent clearing the next row
        if rows and rows[-1].strip() == "":
            rows = rows[:-1]
        start_row = self.currentRow()
        start_col = self.currentColumn()
        if start_col not in [
            DataTableColumns.DIAMETER.index,
            DataTableColumns.RESISTANCE.index,
            DataTableColumns.NUMBER.index,
            DataTableColumns.NAME.index,
        ]:
            return
        for i, row in enumerate(rows):
            if not row.strip():
                continue
            values = row.split("\t")
            for j, value in enumerate(values):
                if start_col in [
                    DataTableColumns.DIAMETER.index,
                    DataTableColumns.RESISTANCE.index,
                ]:
                    value = value.replace(",", ".")
                item = TableWidgetItem(value)
                self.setItem(start_row + i, start_col + j, item)

    def get_column_values(self, column: DataTableColumns):
        values = []
        for row in range(self.rowCount()):
            cb = self.get_row_checkbox(row)
            if not (cb and cb.isChecked()):
                continue
            value = self.item(row, column.index)
            try:
                values.append(column.dtype(value.text()))
            except (ValueError, AttributeError):
                values.append("")
        return values

    def get_column_value(self, row: int, column: DataTableColumns):
        cb = self.get_row_checkbox(row)
        if not (cb and cb.isChecked()):
            return None
        return super().get_column_value(row, column)

    def clear_all(self):
        with contextlib.suppress(TypeError):
            self.itemChanged.disconnect()

        for row in range(self.rowCount()):
            self.setItem(row, DataTableColumns.NAME.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RNS.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.DRIFT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.DIAMETER.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RESISTANCE.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RN_SQRT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.NUMBER.index, TableWidgetItem(str(row + 1)))
            checkbox = QtWidgets.QCheckBox()
            container = QtWidgets.QWidget()
            lay = QtWidgets.QHBoxLayout(container)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            lay.addWidget(checkbox)
            self.setCellWidget(row, DataTableColumns.SELECT.index, container)
            self.setItem(row, DataTableColumns.SELECT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.SQUARE.index, TableWidgetItem(""))
        self.header.checkbox.setChecked(True)
        self.itemChanged.connect(self.on_item_changed)

    def clear_rn(self):
        with contextlib.suppress(TypeError):
            self.itemChanged.disconnect()
        for row in range(self.rowCount()):
            self.setItem(row, DataTableColumns.RNS.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.DRIFT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RESISTANCE.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RN_SQRT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.SQUARE.index, QtWidgets.QTableWidgetItem(""))
        self.itemChanged.connect(self.on_item_changed)

    def clear_calculations(self):
        for row in range(self.rowCount()):
            self.setItem(row, DataTableColumns.RNS.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.DRIFT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.RN_SQRT.index, QtWidgets.QTableWidgetItem(""))
            self.setItem(row, DataTableColumns.SQUARE.index, QtWidgets.QTableWidgetItem(""))

    def color_row(self, row, background_color, text_color):
        with contextlib.suppress(TypeError):
            self.itemChanged.disconnect()
        for col in range(self.columnCount()):
            item = self.item(row, col)
            if item:
                item.setBackground(QBrush(QColor(background_color)))
                item.setForeground(QBrush(QColor(text_color)))
        self.itemChanged.connect(self.on_item_changed)

    def dump_data(self):
        data = []
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                if col == DataTableColumns.SELECT.index:
                    cb = self.get_row_checkbox(row)
                    value = "True" if (cb and cb.isChecked()) else ""
                else:
                    item = self.item(row, col)
                    if not item:
                        value = ""
                        logger.debug(f"DataTable item row={row} col={col} is None")
                    else:
                        value = DataTableColumns.get_by_index(col).dtype(item.text()) if item.text() else ""
                data.append(InitialDataItem(value=value, row=row, col=col))
        return data

    def load_data(self, data: List[InitialDataItem]):
        self.clear_all()
        with contextlib.suppress(TypeError):
            self.itemChanged.disconnect()
        for item in data:
            if item["col"] == DataTableColumns.SELECT.index:
                checkbox = QtWidgets.QCheckBox()
                checkbox.setChecked(item["value"] == "True")
                container = QtWidgets.QWidget()
                lay = QtWidgets.QHBoxLayout(container)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                lay.addWidget(checkbox)
                self.setCellWidget(item["row"], DataTableColumns.SELECT.index, container)
            else:
                self.setItem(item["row"], item["col"], TableWidgetItem(f"{item['value']}"))
        self.itemChanged.connect(self.on_item_changed)

        if not any(
            (self.get_row_checkbox(row) and self.get_row_checkbox(row).isChecked()) for row in range(self.rowCount())
        ):
            for row in range(self.rowCount()):
                cb = self.get_row_checkbox(row)
                if cb:
                    cb.setChecked(True)
