from PyQt5 import QtWidgets

from constants import ParamTableColumns
from store import Item
from widgets.delegates import RoundedDelegate
from widgets.tables.item import TableWidgetItem
from widgets.tables.mixins import TableMixin


class ParamTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self):
        super(ParamTable, self).__init__(1, len(ParamTableColumns.get_all_names()))

        self.setHorizontalHeaderLabels(ParamTableColumns.get_all_names())
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        self.verticalHeader().setVisible(False)

        # self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Minimum, QtWidgets.QSizePolicy.Preferred)
        self.setFixedHeight(self.rowHeight(0) + self.horizontalHeader().height())

        self.set_read_only_columns(
            [
                ParamTableColumns.SLOPE.index,
                ParamTableColumns.INTERCEPT.index,
                ParamTableColumns.RNS.index,
                ParamTableColumns.DRIFT.index,
                ParamTableColumns.RNS_ERROR.index,
                ParamTableColumns.DRIFT_ERROR.index,
            ]
        )

        self.setItemDelegateForColumn(ParamTableColumns.SLOPE.index, RoundedDelegate(rounded=4, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.INTERCEPT.index, RoundedDelegate(rounded=4, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.DRIFT.index, RoundedDelegate(rounded=3, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.RNS.index, RoundedDelegate(rounded=1, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.DRIFT_ERROR.index, RoundedDelegate(rounded=2, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.RNS_ERROR.index, RoundedDelegate(rounded=2, parent=self))

    def get_column_value(self, row: int, column: ParamTableColumns):
        return super().get_column_value(row, column)

    def clear_all(self):
        for col in range(self.columnCount()):
            self.setItem(
                0,
                col,
                QtWidgets.QTableWidgetItem(""),
            )

    def load_data(self, data: Item):
        self.setItem(0, ParamTableColumns.SLOPE.index, TableWidgetItem(str(data.slope)))
        self.setItem(0, ParamTableColumns.INTERCEPT.index, TableWidgetItem(str(data.intercept)))
        self.setItem(0, ParamTableColumns.DRIFT.index, TableWidgetItem(str(data.drift)))
        self.setItem(0, ParamTableColumns.RNS.index, TableWidgetItem(str(data.rns)))
        self.setItem(0, ParamTableColumns.DRIFT_ERROR.index, TableWidgetItem(str(data.drift_error)))
        self.setItem(0, ParamTableColumns.RNS_ERROR.index, TableWidgetItem(str(data.rns_error)))
