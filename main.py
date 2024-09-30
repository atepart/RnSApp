import sys
import numpy as np
import openpyxl
from PyQt5 import QtWidgets
import pyqtgraph as pg
from PyQt5.QtGui import QIcon
from openpyxl.styles import Side, Border, Font

from constants import DataTableColumns, ParamTableColumns
from utils import (
    linear,
    linear_fit,
    calculate_rns,
    calculate_rns_per_sample,
    calculate_drift,
    calculate_rn_sqrt,
)
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

        self.result_button = QtWidgets.QPushButton("Result")
        self.result_button.setToolTip("Произвести рассчет")
        self.result_button.clicked.connect(self.calculate_results)

        self.clean_rn_button = QtWidgets.QPushButton("Clear Rn")
        self.clean_rn_button.setToolTip("Очистить Rn")
        self.clean_rn_button.clicked.connect(self.clean_rn)

        self.clean_all_button = QtWidgets.QPushButton("Clear All")
        self.clean_all_button.setToolTip("Очистить все данные")
        self.clean_all_button.clicked.connect(self.clean_all)

        self.save_button = QtWidgets.QPushButton("Save All data")
        self.save_button.setToolTip("Сохранить входные данные и рассчет")
        self.save_button = QtWidgets.QPushButton("Save")
        self.save_button.clicked.connect(self.save_data)

        self.actions_layout.addWidget(self.result_button)
        self.actions_layout.addWidget(self.clean_rn_button)
        self.actions_layout.addWidget(self.clean_all_button)
        self.actions_layout.addWidget(self.save_button)
        self.actions_group.setLayout(self.actions_layout)

        # Грид с кнопками
        self.cell_group = QtWidgets.QGroupBox("Запись")
        self.cell_layout = QtWidgets.QGridLayout()
        self.cell_buttons = []
        for i in range(4):
            for j in range(4):
                button = QtWidgets.QPushButton(f"Ячейка {i * 4 + j + 1}")
                button.clicked.connect(lambda checked=False, row=i, col=j: self.cell_button_clicked(row, col))
                self.cell_layout.addWidget(button, i, j)
                self.cell_buttons.append(button)
        self.cell_save_button = QtWidgets.QPushButton("Save cells RnS")
        self.cell_save_button.setToolTip("Сохранить выходную таблицу с RnS")
        self.cell_save_button.clicked.connect(self.save_cell_data)
        self.cell_layout.addWidget(self.cell_save_button, 4, 3)
        self.cell_group.setLayout(self.cell_layout)

        # Добавляем все виджеты в правый лейаут
        self.right_layout.addWidget(self.param_table)
        self.right_layout.addWidget(self.plot)
        self.right_layout.addWidget(self.cell_group)
        self.right_layout.addWidget(self.actions_group)

        # Добавляет виджеты в основной лейаут
        self.layout.addWidget(self.data_table)
        self.layout.addLayout(self.right_layout)
        self.setLayout(self.layout)

    def calculate_rn05(self):
        """Рассчет Rn^-0.5 для каждого образца"""
        for row in range(self.data_table.rowCount()):
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            if not resistance:
                continue
            rn_sqrt = calculate_rn_sqrt(resistance)
            self.data_table.setItem(row, DataTableColumns.RN.index, QtWidgets.QTableWidgetItem(str(round(rn_sqrt, 4))))

    def calculate_main_params(self):
        """Рассчет Наклона, Пересечения, RnS, Ухода в целом"""
        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN)
        if not (len(diameter_list) > 1 and len(rn_sqrt_list) > 1):
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
        rns_list = np.array(self.data_table.get_column_values(DataTableColumns.RNS))
        rns_error = np.std(rns_list)
        self.param_table.setItem(
            0,
            ParamTableColumns.RNS_ERROR.index,
            QtWidgets.QTableWidgetItem(str(round(rns_error, 4))),
        )

        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        drift_list = np.array(self.data_table.get_column_values(DataTableColumns.DRIFT))
        drift_error = np.sqrt(np.sum((drift_list - drift) ** 2))
        self.param_table.setItem(
            0,
            ParamTableColumns.DRIFT_ERROR.index,
            QtWidgets.QTableWidgetItem(str(round(drift_error, 4))),
        )

    def calculate_rns_drift_per_sample(self):
        """Рассчет RnS и Ухода для каждого образца по отдельности"""
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        if not drift:
            return

        for row in range(self.data_table.rowCount()):
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            if not resistance or not diameter:
                return

            rns_value = calculate_rns_per_sample(resistance=resistance, diameter=diameter, drift=drift)
            self.data_table.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(str(round(rns_value, 4))),
            )
            drift_value = calculate_drift(diameter=diameter, resistance=resistance, rns=rns_value)
            self.data_table.setItem(
                row, DataTableColumns.DRIFT.index, QtWidgets.QTableWidgetItem(str(round(drift_value, 4)))
            )

    def calculate_results(self):
        self.calculate_rn05()
        self.calculate_main_params()
        self.calculate_rns_drift_per_sample()
        self.calculate_error_params()
        self.plot_current_data()

    def plot_current_data(self):
        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN)
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        slope = self.param_table.get_column_value(0, ParamTableColumns.SLOPE)
        intercept = self.param_table.get_column_value(0, ParamTableColumns.INTERCEPT)

        self.plot.clear()
        self.plot.plot(diameter_list, rn_sqrt_list, name="Data", symbolSize=6, symbolBrush="#088F8F")
        if np.min(diameter_list) > drift:
            diameter_list.insert(0, drift)
        if np.max(diameter_list) < drift:
            diameter_list.append(drift)
        y_appr = np.vectorize(lambda x: linear(x, slope, intercept))(diameter_list)
        pen2 = pg.mkPen(color="#FF0000", width=3)
        self.plot.plot(
            diameter_list,
            y_appr,
            name="Fit",
            pen=pen2,
            symbolSize=0,
            symbolBrush=pen2.color(),
        )

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
        for row in range(self.data_table.rowCount()):
            self.data_table.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.data_table.setItem(
                row,
                DataTableColumns.DRIFT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Drift column
            self.data_table.setItem(
                row,
                DataTableColumns.RESISTANCE.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Resistance column
            self.data_table.setItem(
                row,
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
        self.plot.clear()

    def clean_all(self):
        for row in range(self.data_table.rowCount()):
            self.data_table.setItem(
                row,
                DataTableColumns.NAME.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Name column
            self.data_table.setItem(
                row,
                DataTableColumns.RNS.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear RnS column
            self.data_table.setItem(
                row,
                DataTableColumns.DRIFT.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Drift column
            self.data_table.setItem(
                row,
                DataTableColumns.DIAMETER.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Diameter column
            self.data_table.setItem(
                row,
                DataTableColumns.RESISTANCE.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Resistance column
            self.data_table.setItem(
                row,
                DataTableColumns.RN.index,
                QtWidgets.QTableWidgetItem(""),
            )  # Clear Rn^-0.5 column
            self.data_table.setItem(
                row,
                DataTableColumns.NUMBER.index,
                QtWidgets.QTableWidgetItem(str(row + 1)),
            )  # Clear Number column
        self.plot.clear()

    def cell_button_clicked(self, row, col):
        series, ok = QtWidgets.QInputDialog.getText(self, "Input Dialog", "Enter series:")
        if ok:
            drift = (
                self.param_table.item(0, ParamTableColumns.DRIFT.index).text()
                if self.param_table.item(0, ParamTableColumns.DRIFT.index) is not None
                else ""
            )
            rns = (
                self.param_table.item(0, ParamTableColumns.RNS.index).text()
                if self.param_table.item(0, ParamTableColumns.RNS.index) is not None
                else ""
            )
            button = self.cell_buttons[row * 4 + col]
            button.setText(f"Серия: {series}\nУход: {drift}\nRnS: {rns}")

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

        def parse_btn_text(btn):
            text = btn.text()
            series = text.split("\n")[0].split(": ")[1] if "Серия" in text else text
            care = text.split("\n")[1].split(": ")[1] if "Уход" in text else ""
            rns = text.split("\n")[2].split(": ")[1] if "RnS" in text else ""
            return [series, care, rns]

        init_data = [parse_btn_text(btn) for btn in self.cell_buttons]
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
