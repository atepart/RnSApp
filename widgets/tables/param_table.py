from PyQt5 import QtWidgets

from constants import ParamTableColumns
from widgets.tables.mixins import TableMixin


class ParamTable(TableMixin, QtWidgets.QTableWidget):
    def __init__(self):
        super(ParamTable, self).__init__(1, len(ParamTableColumns.get_all_names()))

        self.setHorizontalHeaderLabels(ParamTableColumns.get_all_names())
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        self.verticalHeader().setVisible(False)

        # self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Minimum, QtWidgets.QSizePolicy.Preferred)
        self.setFixedHeight(self.rowHeight(0) + self.horizontalHeader().height())

    def get_column_value(self, row: int, column: ParamTableColumns):
        return super().get_column_value(row, column)
