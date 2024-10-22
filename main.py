import sys
import numpy as np
import openpyxl
from PyQt5 import QtWidgets
import pyqtgraph as pg
from PyQt5.QtGui import QIcon
from openpyxl.styles import Side, Border, Font

from constants import DataTableColumns, ParamTableColumns, PLOT_COLORS
from errors import ListsNotSameLength
from store import Store
from utils import (
    linear,
    linear_fit,
    calculate_rns,
    calculate_rns_per_sample,
    calculate_drift,
    calculate_rn_sqrt,
    drop_nans,
    calculate_square,
)
from widgets.cell import CellWidget
from widgets.tables.data_table import DataTable
from widgets.tables.param_table import ParamTable


class Window(QtWidgets.QWidget):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowIcon(QIcon("./assets/rns-logo-sm.png"))
        # Таблица с исходными данными
        self.data_table = DataTable(rows=50)

        # График
        self.plot = pg.PlotWidget()
        self.prepare_plot()

        # Main Layout
        self.layout = QtWidgets.QHBoxLayout()

        # Right Layout (таблица с параметрами, график, экшнс)
        self.right_layout = QtWidgets.QVBoxLayout()

        # Таблица с параметрами
        self.param_table = ParamTable()

        # Экшнс кнопки
        self.actions_group = QtWidgets.QGroupBox("Действия")
        self.actions_layout = QtWidgets.QHBoxLayout()

        self.result_button = QtWidgets.QPushButton("Рассчитать")
        self.result_button.setToolTip("Произвести рассчеты")
        self.result_button.clicked.connect(self.calculate_results)

        self.clean_rn_button = QtWidgets.QPushButton("Очистить столбец Rn")
        self.clean_rn_button.setToolTip("Очистить стоблец Rn и рассчетную таблицу")
        self.clean_rn_button.clicked.connect(self.clean_rn)

        self.clean_all_button = QtWidgets.QPushButton("Очистить все")
        self.clean_all_button.setToolTip("Очистить исходную и рассчетную таблицы и график")
        self.clean_all_button.clicked.connect(self.clean_all)

        self.save_button = QtWidgets.QPushButton("Сохранить все")
        self.save_button.setToolTip("Сохранить входные данные и рассчет")
        self.save_button.clicked.connect(self.save_data)

        self.actions_layout.addWidget(self.result_button)
        self.actions_layout.addWidget(self.clean_rn_button)
        self.actions_layout.addWidget(self.clean_all_button)
        self.actions_layout.addWidget(self.save_button)
        self.actions_group.setLayout(self.actions_layout)

        # Грид с кнопками
        self.cell_group = QtWidgets.QGroupBox("Запись")
        self.cell_layout = QtWidgets.QGridLayout()
        self.cell_widgets = []
        for i in range(4):
            for j in range(4):
                index = i * 4 + j + 1
                cell_widget = CellWidget(self, index, self.param_table)
                self.cell_layout.addWidget(cell_widget, i, j)
                self.cell_widgets.append(cell_widget)
        self.mean_drift = QtWidgets.QLabel("Средний уход: --", self)
        self.cell_layout.addWidget(self.mean_drift, 4, 0)
        self.mean_rns = QtWidgets.QLabel("Средний RnS: --", self)
        self.cell_layout.addWidget(self.mean_rns, 4, 1)
        self.cell_save_button = QtWidgets.QPushButton("Сохранить ячейки")
        self.cell_save_button.setToolTip("Сохранить записанные ячейки с RnS")
        self.cell_save_button.clicked.connect(self.save_cell_data)
        self.cell_layout.addWidget(self.cell_save_button, 4, 3)
        self.cell_group.setLayout(self.cell_layout)

        # Добавляем все виджеты в правый лейаут
        self.right_layout.addWidget(self.param_table)
        self.right_layout.addWidget(self.plot)
        self.right_layout.addWidget(self.actions_group)
        self.right_layout.addWidget(self.cell_group)

        # Добавляет виджеты в основной лейаут
        self.layout.addWidget(self.data_table)
        self.layout.addLayout(self.right_layout)
        self.setLayout(self.layout)

    def calculate_means(self):
        rns_list = []
        drift_list = []
        for cell in self.cell_widgets:
            _, drift, rns = self.parse_cell(cell)
            if drift:
                drift_list.append(drift)
            if rns:
                rns_list.append(rns)

        self.mean_drift.setText(f"Средний уход: {round(np.mean(drift_list), 4)}")
        self.mean_rns.setText(f"Средний RnS: {round(np.mean(rns_list), 4)}")

    def calculate_rn05(self):
        """Рассчет Rn^-0.5 для каждого образца"""
        for row in range(self.data_table.rowCount()):
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            if not resistance or not diameter:
                continue
            rn_sqrt = calculate_rn_sqrt(resistance)
            self.data_table.setItem(row, DataTableColumns.RN.index, QtWidgets.QTableWidgetItem(str(round(rn_sqrt, 4))))

    def calculate_main_params(self):
        """Рассчет Наклона, Пересечения, RnS, Ухода в целом"""
        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN)
        try:
            diameter_list, rn_sqrt_list = drop_nans(diameter_list, rn_sqrt_list)
        except ListsNotSameLength:
            return
        slope, intercept = linear_fit(diameter_list, rn_sqrt_list)

        self.param_table.setItem(
            0,
            ParamTableColumns.SLOPE.index,
            QtWidgets.QTableWidgetItem(str(round(slope, 4))),
        )
        self.param_table.setItem(
            0,
            ParamTableColumns.INTERCEPT.index,
            QtWidgets.QTableWidgetItem(str(round(intercept, 4))),
        )

        drift = -intercept / slope

        self.param_table.setItem(
            0,
            ParamTableColumns.DRIFT.index,
            QtWidgets.QTableWidgetItem(str(round(drift, 4))),
        )

        rns = calculate_rns(slope)
        self.param_table.setItem(
            0,
            ParamTableColumns.RNS.index,
            QtWidgets.QTableWidgetItem(str(round(rns, 4))),
        )

    def calculate_error_params(self):
        """Рассчет ошибок RnS и Ухода"""
        rns_list = np.array([v for v in self.data_table.get_column_values(DataTableColumns.RNS) if v], dtype=float)
        rns_error = np.std(rns_list)
        self.param_table.setItem(
            0,
            ParamTableColumns.RNS_ERROR.index,
            QtWidgets.QTableWidgetItem(str(round(rns_error, 4))),
        )

        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        drift_list = np.array([v for v in self.data_table.get_column_values(DataTableColumns.DRIFT) if v], dtype=float)
        drift_error = np.sqrt(np.sum((drift_list - drift) ** 2))
        self.param_table.setItem(
            0,
            ParamTableColumns.DRIFT_ERROR.index,
            QtWidgets.QTableWidgetItem(str(round(drift_error, 4))),
        )

    def calculate_rns_drift_square_per_sample(self):
        """Рассчет RnS, Ухода и Площади для каждого образца по отдельности"""
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        if not drift:
            return

        rns_mean = self.param_table.get_column_value(0, ParamTableColumns.RNS)
        if not rns_mean:
            return

        for row in range(self.data_table.rowCount()):
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            if not resistance or not diameter:
                continue

            rns_value = calculate_rns_per_sample(resistance=resistance, diameter=diameter, drift=drift)
            self.data_table.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(str(round(rns_value, 4))),
            )

            drift_value = calculate_drift(diameter=diameter, resistance=resistance, rns=rns_mean)
            self.data_table.setItem(
                row, DataTableColumns.DRIFT.index, QtWidgets.QTableWidgetItem(str(round(drift_value, 4)))
            )

            square_value = calculate_square(diameter=diameter, drift=drift_value)
            self.data_table.setItem(
                row, DataTableColumns.SQUARE.index, QtWidgets.QTableWidgetItem(str(round(square_value, 4)))
            )

    def calculate_results(self):
        self.calculate_rn05()
        self.calculate_main_params()
        self.calculate_rns_drift_square_per_sample()
        self.calculate_error_params()
        self.plot_current_data()

    def plot_current_data(self):
        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN)
        try:
            diameter_list, rn_sqrt_list = drop_nans(diameter_list, rn_sqrt_list)
            diameter_list, rn_sqrt_list = np.array(
                sorted(np.array([diameter_list, rn_sqrt_list]).T, key=lambda x: x[0]), dtype=float
            ).T
            diameter_list = diameter_list.tolist()
            rn_sqrt_list = rn_sqrt_list.tolist()
        except ListsNotSameLength:
            return
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        slope = self.param_table.get_column_value(0, ParamTableColumns.SLOPE)
        intercept = self.param_table.get_column_value(0, ParamTableColumns.INTERCEPT)

        plotItem = self.plot.getPlotItem()
        items_data = [item for item in plotItem.items if item.name() == "Data"]

        items_fit = [item for item in plotItem.items if item.name() == "Fit"]

        if items_data:
            items_data[0].setData(diameter_list, rn_sqrt_list)
        else:
            self.plot.plot(diameter_list, rn_sqrt_list, name="Data", symbolSize=6, symbolBrush="#000000")

        if np.min(diameter_list) > drift:
            diameter_list.insert(0, drift)
        if np.max(diameter_list) < drift:
            diameter_list.append(drift)
        y_appr = np.vectorize(lambda x: linear(x, slope, intercept))(diameter_list)

        if items_fit:
            items_fit[0].setData(diameter_list, y_appr)

        else:
            pen2 = pg.mkPen(color="#000000", width=3)
            self.plot.plot(
                diameter_list,
                y_appr,
                name="Fit",
                pen=pen2,
                symbolSize=0,
                symbolBrush=pen2.color(),
            )

    def addCellData(self, cell: int, name: str):
        Store.update_or_create_item(
            cell=cell,
            name=name,
            diameter_list=self.data_table.get_column_values(DataTableColumns.DIAMETER),
            rn_sqrt_list=self.data_table.get_column_values(DataTableColumns.RN),
            drift=self.param_table.get_column_value(0, ParamTableColumns.DRIFT),
            slope=self.param_table.get_column_value(0, ParamTableColumns.SLOPE),
            intercept=self.param_table.get_column_value(0, ParamTableColumns.INTERCEPT),
            initial_data=self.data_table.dump_data(),
        )

    def plot_data(self, cell: int):
        color = PLOT_COLORS[cell]
        item = Store.data.get(cell=cell)
        if not item:
            return
        diameter, rn_sqrt = drop_nans(item.diameter, item.rn_sqrt)
        self.plot.plot(diameter, rn_sqrt, name=f"№{cell}; {item.name}; Data", symbolSize=4, symbolBrush=color)
        diameter_list = diameter.tolist()
        if np.min(diameter_list) > item.drift:
            diameter_list.insert(0, item.drift)
        if np.max(diameter_list) < item.drift:
            diameter_list.append(item.drift)
        y_appr = np.vectorize(lambda x: linear(x, item.slope, item.intercept))(diameter_list)
        pen2 = pg.mkPen(color=color, width=2)
        self.plot.plot(
            diameter_list,
            y_appr,
            name=f"№{cell}; Fit",
            pen=pen2,
            symbolSize=0,
            symbolBrush=pen2.color(),
        )

    def remove_plot(self, cell: int):
        plotItem = self.plot.getPlotItem()
        items_to_remove = [item for item in plotItem.items if item.name().startswith(f"№{cell};")]
        for item in items_to_remove:
            plotItem.removeItem(item)

    def prepare_plot(self):
        y_label = "Rn^-0.5"
        x_label = "Диаметр ACAD (μm)"
        self.plot.setBackground("w")
        styles = {"color": "#413C58", "font-size": "15px"}
        self.plot.setLabel("left", y_label, **styles)
        self.plot.setLabel("bottom", x_label, **styles)
        self.plot.addLegend()
        self.plot.showGrid(x=True, y=True)

    def save_data(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Data",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )
        if not file_name:
            return
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Data"
        headers = [self.data_table.horizontalHeaderItem(i).text() for i in range(self.data_table.columnCount())]
        ws1.append(headers)
        for row in range(self.data_table.rowCount()):
            data = [
                self.data_table.item(row, col).text() if self.data_table.item(row, col) else ""
                for col in range(self.data_table.columnCount())
            ]
            ws1.append(data)
        ws2 = wb.create_sheet("Results")
        headers = [self.param_table.horizontalHeaderItem(i).text() for i in range(self.param_table.columnCount())]
        ws2.append(headers)
        for row in range(self.param_table.rowCount()):
            data = [
                self.param_table.item(row, col).text() if self.param_table.item(row, col) else ""
                for col in range(self.param_table.columnCount())
            ]
            ws2.append(data)

        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        wb.save(filename=file_name)

    def clean_rn(self):
        self.data_table.clear_rn()
        self.param_table.clear_all()

    def clean_all(self):
        self.data_table.clear_all()
        self.param_table.clear_all()
        self.plot.clear()

    @staticmethod
    def parse_cell(cell):
        drift = float(cell.drift.text()) if cell.drift.text() else ""
        rns = float(cell.rns.text()) if cell.rns.text() else ""
        return [cell.name.text(), drift, rns]

    def save_cell_data(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Cell Data",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )
        if not file_name:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Cell Data"

        init_data = [self.parse_cell(cell) for cell in self.cell_widgets]
        output = []

        # Разделение исходного массива на блоки по 4 строки
        blocks = [init_data[i : i + 4] for i in range(0, len(init_data), 4)]

        # Обработка каждого блока
        for block in blocks:
            # Перестановка элементов в блоке
            block_transposed = list(map(list, zip(*block)))
            # Добавление переставленного блока в выходной массив
            output.extend(block_transposed)
        for row_ind, row in enumerate(output, 1):
            for col_ind, coll in enumerate(row, 1):
                ws.cell(row=row_ind, column=col_ind, value=coll)
                if (row_ind - 1) % 3 == 0:  # для клеток с названием серии
                    ws.cell(row=row_ind, column=col_ind).border = Border(
                        right=Side(style="thick"),
                        top=Side("thick"),
                        bottom=Side(style="thick"),
                    )
                    ws.cell(row=row_ind, column=col_ind).font = Font(bold=True)
                else:  # для остальных клеток
                    ws.cell(row=row_ind, column=col_ind).border = Border(right=Side(style="thick"))
        # Save the Excel file
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        wb.save(filename=file_name)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
