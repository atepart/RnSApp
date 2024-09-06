import sys
import numpy as np
import math
import openpyxl
from PyQt5 import QtWidgets, QtCore, QtGui
import pyqtgraph as pg

class Table(QtWidgets.QTableWidget):
    def __init__(self, rows, columns):
        super(Table, self).__init__(rows, columns)
        self.setHorizontalHeaderLabels(['Number', 'Name', 'RnS', 'Diameter (μm)', 'Resistance (Ω)', 'Rn^-0.5'])
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.verticalHeader().setVisible(False)
        self.setItemDelegateForColumn(0, ReadOnlyDelegate(self))
        self.setItemDelegateForColumn(5, ReadOnlyDelegate(self))
        self.itemChanged.connect(self.update_table)
        for i in range(rows):
            self.setItem(i, 0, QtWidgets.QTableWidgetItem(str(i+1)))
        self.setShowGrid(True)
        self.setGridStyle(QtCore.Qt.SolidLine)

    def update_table(self, item):
        self.itemChanged.disconnect(self.update_table)
        row = item.row()
        col = item.column()
        try:
            if col == 3:  # Diameter
                diameter = float(item.text())
                if self.item(row, 4) is not None:
                    resistance = float(self.item(row, 4).text())
                    rn_sqrt = 1 / np.sqrt(resistance)
                    self.setItem(row, 5, QtWidgets.QTableWidgetItem(str(round(rn_sqrt, 4))))
            elif col == 4:  # Resistance
                diameter = float(self.item(row, 3).text())
                resistance = float(item.text())
                rn_sqrt = 1 / np.sqrt(resistance)
                self.setItem(row, 5, QtWidgets.QTableWidgetItem(str(round(rn_sqrt, 4))))
        except ValueError:
            pass
        self.itemChanged.connect(self.update_table)

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            row = self.currentRow()
            col = self.currentColumn()
            if row < self.rowCount() - 1:
                self.setCurrentCell(row + 1, col)
        elif event.key() == QtCore.Qt.Key_Delete:
            self.clearContents()
            self.parent().update_plot()
        elif event.matches(QtGui.QKeySequence.Paste):
            self.paste_data()
        else:
            super(Table, self).keyPressEvent(event)

    def paste_data(self):
        clipboard = QtWidgets.QApplication.clipboard()
        data = clipboard.text()
        rows = data.split('\n')
        start_row = self.currentRow()
        for i, row in enumerate(rows):
            values = row.split('\t')
            for j, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(value)
                self.setItem(start_row + i, j, item)
                self.update_table(item)
        self.parent().update_plot()

class ReadOnlyDelegate(QtWidgets.QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return

def linear_fit(x, y):
    def mean(xs):
        return sum(xs) / len(xs)

    m_x = mean(x)
    m_y = mean(y)

    def std(xs, m):
        normalizer = len(xs) - 1
        return np.sqrt(sum((pow(x1 - m, 2) for x1 in xs)) / normalizer)

    def pearson_r(xs, ys):
        sum_xy = 0
        sum_sq_v_x = 0
        sum_sq_v_y = 0

        for x1, y2 in zip(xs, ys):
            var_x = x1 - m_x
            var_y = y2 - m_y
            sum_xy += var_x * var_y
            sum_sq_v_x += pow(var_x, 2)
            sum_sq_v_y += pow(var_y, 2)
        return sum_xy / np.sqrt(sum_sq_v_x * sum_sq_v_y)

    r = pearson_r(x, y)

    b = r * (std(y, m_y) / std(x, m_x))
    a = m_y - b * m_x

    return b, a

def linear(x, b, a):
    return x * b + a

class Window(QtWidgets.QWidget):
    def __init__(self):
        super(Window, self).__init__()
        self.table = Table(50, 6)
        self.plot = pg.PlotWidget()
        self.y_label = "Rn^-0.5"
        self.x_label = "Diameter"
        self.plot_title = "Rn^-0.5(Diameter)"
        self.prepare_plot()
        self.layout = QtWidgets.QHBoxLayout()
        self.right_layout = QtWidgets.QVBoxLayout()
        self.param_table = QtWidgets.QTableWidget(5, 2)
        self.param_table.setHorizontalHeaderLabels(['Parameter', 'Value'])
        self.param_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.param_table.verticalHeader().setVisible(False)
        self.param_table.setItemDelegateForColumn(0, ReadOnlyDelegate(self))
        self.right_layout.addWidget(self.param_table)
        self.right_layout.addWidget(self.plot)
        self.button_layout = QtWidgets.QHBoxLayout()
        self.result_button = QtWidgets.QPushButton('Result')
        self.result_button.clicked.connect(self.calculate_results)
        self.clean_rn_button = QtWidgets.QPushButton('Clean Rn')
        self.clean_rn_button.clicked.connect(self.clean_rn)
        self.clean_all_button = QtWidgets.QPushButton('Clean All')
        self.clean_all_button.clicked.connect(self.clean_all)
        self.save_button = QtWidgets.QPushButton('Save')
        self.save_button.clicked.connect(self.save_data)
        self.button_layout.addWidget(self.result_button)
        self.button_layout.addWidget(self.clean_rn_button)
        self.button_layout.addWidget(self.clean_all_button)
        self.button_layout.addWidget(self.save_button)
        self.right_layout.addLayout(self.button_layout)
        self.result_table = QtWidgets.QTableWidget(1, 6)
        self.result_table.setHorizontalHeaderLabels(['Number', 'Name', 'RnS', 'Diameter (μm)', 'Resistance (Ω)', 'Rn^-0.5'])
        self.result_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.result_table.verticalHeader().setVisible(False)
        self.cell_layout = QtWidgets.QGridLayout()
        self.cell_buttons = []
        for i in range(4):
            for j in range(4):
                button = QtWidgets.QPushButton(f'Cell {i*4 + j + 1}')
                button.clicked.connect(lambda checked=False, row=i, col=j: self.cell_button_clicked(row, col))
                self.cell_layout.addWidget(button, i, j)
                self.cell_buttons.append(button)
        self.cell_save_button = QtWidgets.QPushButton('S')
        self.cell_save_button.clicked.connect(self.save_cell_data)
        self.cell_layout.addWidget(self.cell_save_button, 4, 3)
        self.right_layout.addLayout(self.cell_layout)
        self.layout.addWidget(self.table)
        self.layout.addLayout(self.right_layout)
        self.setLayout(self.layout)

        # Set the values for the first column of the parameter table
        self.param_table.setItem(0, 0, QtWidgets.QTableWidgetItem('k'))
        self.param_table.setItem(1, 0, QtWidgets.QTableWidgetItem('b'))
        self.param_table.setItem(2, 0, QtWidgets.QTableWidgetItem('Уход'))
        self.param_table.setItem(3, 0, QtWidgets.QTableWidgetItem('RnS'))
        self.param_table.setItem(4, 0, QtWidgets.QTableWidgetItem('Ошибка'))

    def calculate_results(self):
        self.update_plot()

    def update_plot(self):
        x = []
        y = []
        for i in range(self.table.rowCount()):
            item_x = self.table.item(i, 3)
            item_y = self.table.item(i, 5)
            if item_x is not None and item_y is not None and item_x.text() and item_y.text():
                try:
                    x.append(float(item_x.text()))
                    y.append(float(item_y.text()))
                except ValueError:
                    pass
        self.plot.clear()
        self.plot.plot(x, y, name="Data", symbolSize=6, symbolBrush="#088F8F")
        if len(x) > 1 and len(y) > 1:
            b, a = linear_fit(x, y)
            zero_x = -a / b
            if np.min(x) > zero_x:
                x.insert(0, zero_x)
            if np.max(x) < zero_x:
                x.append(zero_x)
            y_appr = np.vectorize(lambda x: linear(x, b, a))(x)
            pen2 = pg.mkPen(color="#FF0000", width=3)
            self.plot.plot(x, y_appr, name="Fit", pen=pen2, symbolSize=0, symbolBrush=pen2.color())
            self.param_table.setItem(0, 1, QtWidgets.QTableWidgetItem(str(round(b, 4))))
            self.param_table.setItem(1, 1, QtWidgets.QTableWidgetItem(str(round(a, 4))))
            self.param_table.setItem(2, 1, QtWidgets.QTableWidgetItem(str(round(zero_x, 4))))
            rns = math.pi * 0.25 / (b ** 2)
            self.param_table.setItem(3, 1, QtWidgets.QTableWidgetItem(str(round(rns, 4))))

            # Calculate the mean deviation from the mean of the approximation
            y_mean = np.mean(y)
            deviation = np.mean(np.abs(y - y_mean))
            self.param_table.setItem(4, 1, QtWidgets.QTableWidgetItem(str(round(deviation, 4))))

            # Update the RnS column in the table
            for i in range(self.table.rowCount()):
                item_diameter = self.table.item(i, 3)
                if item_diameter is not None and item_diameter.text():
                    try:
                        diameter = float(item_diameter.text())
                        rns_value = rns * (diameter - zero_x) ** 2
                        self.table.setItem(i, 2, QtWidgets.QTableWidgetItem(str(round(rns_value, 4))))
                    except ValueError:
                        pass

    def prepare_plot(self):
        self.plot.setBackground("w")
        self.plot.setTitle(self.plot_title, color="#413C58", size="10pt")
        styles = {"color": "#413C58", "font-size": "15px"}
        self.plot.setLabel("left", self.y_label, **styles)
        self.plot.setLabel("bottom", self.x_label, **styles)
        self.plot.addLegend()
        self.plot.showGrid(x=True, y=True)

    def save_data(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Data", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "Data"
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            ws1.append(headers)
            for row in range(self.table.rowCount()):
                data = [self.table.item(row, col).text() for col in range(self.table.columnCount())]
                ws1.append(data)
            ws2 = wb.create_sheet("Results")
            headers = [self.param_table.horizontalHeaderItem(i).text() for i in range(self.param_table.columnCount())]
            ws2.append(headers)
            for row in range(self.param_table.rowCount()):
                data = [self.param_table.item(row, col).text() for col in range(self.param_table.columnCount())]
                ws2.append(data)
            wb.save(filename=file_name)

    def clean_rn(self):
        for row in range(self.table.rowCount()):
            self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(''))
            self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(''))
        self.param_table.clearContents()
        self.plot.clear()

    def clean_all(self):
        self.table.clearContents()
        self.param_table.clearContents()
        self.plot.clear()

    def cell_button_clicked(self, row, col):
        series = self.param_table.item(0, 1).text() if self.param_table.item(0, 1) is not None else ''
        uhod = self.param_table.item(2, 1).text() if self.param_table.item(2, 1) is not None else ''
        rns = self.param_table.item(3, 1).text() if self.param_table.item(3, 1) is not None else ''
        button = self.cell_buttons[row * 4 + col]
        button.setText(f'Серия: {series}\nУход: {uhod}\nRnS: {rns}')

    def save_cell_data(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Cell Data", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Cell Data"
            headers = ['Cell', 'Series', 'Уход', 'RnS']
            ws.append(headers)
            for i, button in enumerate(self.cell_buttons):
                text = button.text()
                if text:
                    series = text.split('\n')[0].split(': ')[1]
                    uhod = text.split('\n')[1].split(': ')[1]
                    rns = text.split('\n')[2].split(': ')[1]
                    data = [f'Cell {i+1}', series, uhod, rns]
                    ws.append(data)
            wb.save(filename=file_name)

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            row = self.table.currentRow()
            col = self.table.currentColumn()
            if row < self.table.rowCount() - 1:
                self.table.setCurrentCell(row + 1, col)
        elif event.key() == QtCore.Qt.Key_Delete:
            self.table.clearContents()
            self.update_plot()
        else:
            super(Window, self).keyPressEvent(event)

app = QtWidgets.QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
