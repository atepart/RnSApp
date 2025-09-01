import logging

import numpy as np
import pyqtgraph as pg
from PySide6 import QtWidgets
from PySide6.QtGui import QIcon

from src.constants import DataTableColumns, ParamTableColumns
from src.services import CalculationService, PlotService, load_cells_from_xlsx, save_cells_to_xlsx
from src.store import Store
from src.widgets import CellWidget, DataTable, ParamTable

logger = logging.getLogger(__name__)


class RnSApp(QtWidgets.QWidget):
    def __init__(self) -> None:
        super(RnSApp, self).__init__()
        self.setGeometry(100, 100, 1200, 900)

        self.setWindowIcon(QIcon("./assets/rns-logo-sm.png"))
        # Таблица с исходными данными
        self.data_table_label = QtWidgets.QLabel("Таблица с данными", self)
        self.data_table = DataTable(rows=50)

        # График
        self.plot = pg.PlotWidget()

        # Main Layout
        self.layout = QtWidgets.QHBoxLayout()

        # Left Layout (таблица с данными)
        self.left_layout = QtWidgets.QVBoxLayout()

        # Right Layout (таблица с параметрами, график, экшнс)
        self.right_layout = QtWidgets.QVBoxLayout()

        # Right layout параметры расчета
        self.actions_params_layout = QtWidgets.QHBoxLayout()

        # Таблица с параметрами
        self.param_table_label = QtWidgets.QLabel("Таблица с расчетом", self)
        self.param_table = ParamTable()

        # Экшнс кнопки
        self.actions_group = QtWidgets.QGroupBox("Действия")
        self.actions_layout = QtWidgets.QHBoxLayout()

        self.result_button = QtWidgets.QPushButton("Расcчитать")
        self.result_button.setToolTip("Произвести расчеты")
        self.result_button.clicked.connect(self.calculate_results)

        self.clean_rn_button = QtWidgets.QPushButton("Очистить Rn")
        self.clean_rn_button.setToolTip("Очистить стоблец Rn и таблицу с расчетом")
        self.clean_rn_button.clicked.connect(self.clean_rn)

        self.clean_all_button = QtWidgets.QPushButton("Очистить таблицы")
        self.clean_all_button.setToolTip("Очистить график и таблицы с даными и расчетом")
        self.clean_all_button.clicked.connect(self.clean_all)

        self.actions_layout.addWidget(self.result_button)
        self.actions_layout.addWidget(self.clean_rn_button)
        self.actions_layout.addWidget(self.clean_all_button)
        self.actions_group.setLayout(self.actions_layout)

        # Параметры для расчета

        self.rn_consistent_group = QtWidgets.QGroupBox("Последовательное Rn (Ом)")
        self.rn_consistent_layout = QtWidgets.QHBoxLayout()
        self.rn_consistent = QtWidgets.QDoubleSpinBox(self)
        self.rn_consistent.setRange(0, 100)
        self.rn_consistent.setDecimals(2)
        self.rn_consistent.setValue(0)
        self.rn_consistent_layout.addWidget(self.rn_consistent)
        self.rn_consistent_group.setLayout(self.rn_consistent_layout)

        self.allowed_error_group = QtWidgets.QGroupBox("Разрешенная ошибка")
        self.allowed_error_layout = QtWidgets.QHBoxLayout()
        self.allowed_error = QtWidgets.QDoubleSpinBox(self)
        self.allowed_error.setRange(0, 1)
        self.allowed_error.setDecimals(2)
        self.allowed_error.setValue(0)
        self.allowed_error_layout.addWidget(self.allowed_error)
        self.allowed_error_group.setLayout(self.allowed_error_layout)

        # Грид с ячейками
        self.cell_group = QtWidgets.QGroupBox("Запись")
        self.cell_v_layout = QtWidgets.QVBoxLayout()
        self.cell_h_layout = QtWidgets.QHBoxLayout()
        self.cell_grid_layout = QtWidgets.QGridLayout()
        self.cell_widgets = []
        for i in range(4):
            for j in range(4):
                index = i * 4 + j + 1
                cell_widget = CellWidget(self, index, self.param_table)
                self.cell_grid_layout.addWidget(cell_widget, i, j)
                self.cell_widgets.append(cell_widget)

        self.mean_drift = QtWidgets.QLabel("Средний уход: --", self)
        self.mean_rns = QtWidgets.QLabel("Средний RnS: --", self)

        self.btn_load_cell_data = QtWidgets.QPushButton("Загрузить из файла")
        self.btn_load_cell_data.setToolTip("Загрузить данные из xlsx файла")
        self.btn_load_cell_data.clicked.connect(self.load_cell_data)

        self.btn_clear_cell_data = QtWidgets.QPushButton("Очистить ячейки")
        self.btn_clear_cell_data.setToolTip("Очистить ячейки с записанными данными")
        self.btn_clear_cell_data.clicked.connect(self.clear_cell_data)

        self.cell_save_button = QtWidgets.QPushButton("Сохранить")
        self.cell_save_button.setToolTip("Сохранить записанные данные с RnS")
        self.cell_save_button.clicked.connect(self.save_cell_data)

        self.cell_h_layout.addWidget(self.mean_drift)
        self.cell_h_layout.addWidget(self.mean_rns)
        self.cell_h_layout.addWidget(self.btn_clear_cell_data)
        self.cell_h_layout.addWidget(self.btn_load_cell_data)
        self.cell_h_layout.addWidget(self.cell_save_button)

        self.cell_v_layout.addLayout(self.cell_grid_layout)
        self.cell_v_layout.addLayout(self.cell_h_layout)
        self.cell_group.setLayout(self.cell_v_layout)

        # Добавляем виджеты в левый лейаут
        self.left_layout.addWidget(self.data_table_label)
        self.left_layout.addWidget(self.data_table)

        # Добавляем все виджеты в правый лейаут
        self.right_layout.addWidget(self.param_table_label)
        self.right_layout.addWidget(self.param_table)
        self.right_layout.addWidget(self.plot)
        self.actions_params_layout.addWidget(self.actions_group)
        self.actions_params_layout.addWidget(self.rn_consistent_group)
        self.actions_params_layout.addWidget(self.allowed_error_group)
        self.right_layout.addLayout(self.actions_params_layout)
        self.right_layout.addWidget(self.cell_group)

        # Добавляет виджеты в основной лейаут
        self.layout.addLayout(self.left_layout)
        self.layout.addLayout(self.right_layout)
        self.setLayout(self.layout)

        # Services init (after widgets are constructed)
        self.calc = CalculationService(
            data_table=self.data_table,
            param_table=self.param_table,
            rn_consistent_widget=self.rn_consistent,
            allowed_error_widget=self.allowed_error,
        )
        self.plot_service = PlotService(self.plot, self.data_table, self.param_table)
        self.plot_service.prepare_plot()

    def calculate_means(self):
        rns_list = []
        drift_list = []
        for cell in self.cell_widgets:
            _, drift, rns = self.parse_cell(cell)
            if drift:
                drift_list.append(drift)
            if rns:
                rns_list.append(rns)

        self.mean_drift.setText(f"Средний уход: {round(np.mean(drift_list), 3)}")
        self.mean_rns.setText(f"Средний RnS: {round(np.mean(rns_list), 1)}")

    def clear_means(self):
        self.mean_drift.setText("Средний уход: --")
        self.mean_rns.setText("Средний RnS: --")

    def calculate_rn05(self):
        """Расчет Rn^-0.5 для каждого образца"""
        return self.calc.calculate_rn05()

    def calculate_main_params(self):
        """Расчет Наклона, Пересечения, RnS, Ухода в целом"""
        return self.calc.calculate_main_params()

    def calculate_error_params(self):
        """Расчет ошибок RnS и Ухода"""
        return self.calc.calculate_error_params()

    def calculate_rns_drift_square_per_sample(self):
        """Расчет RnS, Ухода и Площади для каждого образца по отдельности"""
        return self.calc.calculate_rns_drift_square_per_sample()

    def calculate_results(self):
        if not self.calc.calculate_results():
            return
        self.plot_current_data()

    def plot_current_data(self):
        return self.plot_service.plot_current_data()

    def addCellData(self, cell: int, name: str):
        Store.update_or_create_item(
            cell=cell,
            name=name,
            diameter_list=self.data_table.get_column_values(DataTableColumns.DIAMETER),
            rn_sqrt_list=self.data_table.get_column_values(DataTableColumns.RN_SQRT),
            slope=self.param_table.get_column_value(0, ParamTableColumns.SLOPE),
            intercept=self.param_table.get_column_value(0, ParamTableColumns.INTERCEPT),
            drift=self.param_table.get_column_value(0, ParamTableColumns.DRIFT),
            rns=self.param_table.get_column_value(0, ParamTableColumns.RNS),
            drift_error=self.param_table.get_column_value(0, ParamTableColumns.DRIFT_ERROR),
            rns_error=self.param_table.get_column_value(0, ParamTableColumns.RNS_ERROR),
            initial_data=self.data_table.dump_data(),
            rn_consistent=self.rn_consistent.value(),
            allowed_error=self.allowed_error.value(),
        )
        self.set_active_cell(cell)

    def plot_data(self, cell: int):
        return self.plot_service.plot_cell(cell=cell, store=Store)

    def remove_plot(self, cell: int):
        return self.plot_service.remove_cell_plot(cell=cell, store=Store)

    def prepare_plot(self):
        return self.plot_service.prepare_plot()

    def clean_rn(self):
        self.data_table.clear_rn()
        self.param_table.clear_all()

    def clean_all(self):
        self.data_table.clear_all()
        self.param_table.clear_all()
        self.plot.clear()
        # Re-prepare plot after clearing to restore labels/grid/legend
        self.plot_service.prepare_plot()
        self.set_active_cell(0)

    @staticmethod
    def parse_cell(cell):
        drift = float(cell.drift.text().split("Уход: ")[1]) if cell.drift.text() else ""
        rns = float(cell.rns.text().split("RnS: ")[1]) if cell.rns.text() else ""
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

        init_data = [(cell.name.text(), cell.drift.text(), cell.rns.text()) for cell in self.cell_widgets]
        save_cells_to_xlsx(
            file_name=file_name,
            cell_grid_values=init_data,
            store=Store,
            data_headers=DataTableColumns.get_all_slugs(),
            results_headers=ParamTableColumns.get_all_names(),
        )

    def reload_tables_from_cell_data(self, cell: int):
        cell_data = Store.data.get(cell=cell)
        if not cell_data:
            return
        self.data_table.load_data(data=cell_data.initial_data)
        self.param_table.load_data(data=cell_data)
        self.plot_current_data()
        self.set_active_cell(cell)
        self.rn_consistent.setValue(cell_data.rn_consistent)
        self.allowed_error.setValue(cell_data.allowed_error)

    def set_active_cell(self, cell: int):
        for cw in self.cell_widgets:
            if cw.index == cell:
                cw.set_active(True)
                continue
            cw.set_active(False)

    def clear_cell_data(self):
        reply = QtWidgets.QMessageBox.question(
            self,
            "Очистить записанные данные ячеек",
            "Записанные данные ячеек будут удалены, продолжить?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )

        if reply == QtWidgets.QMessageBox.Yes:
            for cell in self.cell_widgets:
                cell.clear()
            Store.clear()
            self.clear_means()
            self.set_active_cell(0)

    def load_cell_data(self):
        reply = QtWidgets.QMessageBox.question(
            self,
            "Загрузка данных их файла",
            "Все текущие данные будут очищены, продолжить?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )

        if reply != QtWidgets.QMessageBox.Yes:
            return

        for cell in self.cell_widgets:
            cell.clear()
        Store.clear()
        self.clear_means()
        self.clean_all()
        self.set_active_cell(0)

        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            None, "Выберите файл XLSX", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )

        if not fileName:
            return

        is_some_errors = False
        some_errors_text = ""
        try:
            items, errors = load_cells_from_xlsx(fileName)
            for kw in items:
                cell_item = Store.update_or_create_item(**kw)
                cell_widget = self.cell_widgets[cell_item.cell - 1]
                cell_widget.name.setText(cell_item.name)
                cell_widget.drift.setText(f"Уход: {round(cell_item.drift, 3)}")
                cell_widget.rns.setText(f"RnS: {round(cell_item.rns, 1)}")
                cell_widget.updateUI()
                self.calculate_means()
            if errors:
                is_some_errors = True
                some_errors_text = "\n".join(errors)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка чтения", f"Возникли ошибки чтения файла: {str(e)}")
            logger.exception(f"{e}", exc_info=True)
            return

        if is_some_errors:
            QtWidgets.QMessageBox.warning(self, "Файл прочитался с ошибками", some_errors_text)
            return

        QtWidgets.QMessageBox.information(self, "Файл успешно прочитан", "Все данные успешно загружены")
