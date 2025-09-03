import contextlib

from PySide6 import QtCore, QtWidgets

from domain.constants import ParamTableColumns
from domain.models import Item
from ui.widgets.delegates import RoundedDelegate
from ui.widgets.tables.item import TableWidgetItem
from ui.widgets.tables.mixins import TableMixin


class ParamTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self) -> None:
        super(ParamTable, self).__init__(1, len(ParamTableColumns.get_all_names()))

        self.setHorizontalHeaderLabels(ParamTableColumns.get_all_names())
        # Enable horizontal scrolling by avoiding stretch and letting content define width
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # Disable vertical scrollbar entirely (single-row table)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.verticalHeader().setVisible(False)
        self.setFixedHeight(self.rowHeight(0) + self.horizontalHeader().height())

        self.set_read_only_columns(
            [
                ParamTableColumns.SLOPE.index,
                ParamTableColumns.INTERCEPT.index,
                ParamTableColumns.RNS.index,
                ParamTableColumns.DRIFT.index,
                ParamTableColumns.RNS_ERROR.index,
                ParamTableColumns.DRIFT_ERROR.index,
                ParamTableColumns.S_REAL_1.index,
                ParamTableColumns.S_REAL_CUSTOM1.index,
                ParamTableColumns.S_REAL_CUSTOM2.index,
            ]
        )

        self.setItemDelegateForColumn(ParamTableColumns.SLOPE.index, RoundedDelegate(rounded=4, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.INTERCEPT.index, RoundedDelegate(rounded=4, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.DRIFT.index, RoundedDelegate(rounded=3, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.RNS.index, RoundedDelegate(rounded=1, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.DRIFT_ERROR.index, RoundedDelegate(rounded=2, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.RNS_ERROR.index, RoundedDelegate(rounded=2, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.S_REAL_1.index, RoundedDelegate(rounded=3, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.S_REAL_CUSTOM1.index, RoundedDelegate(rounded=3, parent=self))
        self.setItemDelegateForColumn(ParamTableColumns.S_REAL_CUSTOM2.index, RoundedDelegate(rounded=3, parent=self))

        self.setColumnHidden(ParamTableColumns.DRIFT_ERROR.index, True)
        self.setColumnHidden(ParamTableColumns.RN_CONSISTENT.index, True)
        self.setColumnHidden(ParamTableColumns.ALLOWED_ERROR.index, True)

        self.clear_all()

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
        self.setItem(0, ParamTableColumns.RN_CONSISTENT.index, TableWidgetItem(str(data.rn_consistent)))
        self.setItem(0, ParamTableColumns.ALLOWED_ERROR.index, TableWidgetItem(str(data.allowed_error)))
        # Optional values may be missing in older saved files
        with contextlib.suppress(Exception):
            self.setItem(0, ParamTableColumns.S_REAL_1.index, TableWidgetItem(str(getattr(data, "s_real_1", ""))))

        with contextlib.suppress(Exception):
            self.setItem(
                0,
                ParamTableColumns.S_REAL_CUSTOM1.index,
                TableWidgetItem(str(getattr(data, "s_real_custom1", ""))),
            )

        with contextlib.suppress(Exception):
            self.setItem(
                0,
                ParamTableColumns.S_REAL_CUSTOM2.index,
                TableWidgetItem(str(getattr(data, "s_real_custom2", ""))),
            )

    def is_empty(self):
        return not all(self.item(0, param.index).text() for param in ParamTableColumns)
