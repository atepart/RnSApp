import contextlib
import logging

import numpy as np
import pyqtgraph as pg
from PySide6 import QtWidgets
from PySide6.QtCore import QSettings
from PySide6.QtGui import QIcon
from PySide6QtAds import CDockManager, CDockWidget, DockWidgetArea

from application.calculations import CalculationService
from domain.constants import DataTableColumns, ParamTableColumns
from infrastructure.persistence_xlsx import load_cells_from_xlsx, save_cells_to_xlsx
from infrastructure.repository_memory import InMemoryCellRepository
from ui.plotting_service import PlotService
from ui.widgets import CellWidget, DataTable, ParamTable

logger = logging.getLogger(__name__)


class RnSApp(QtWidgets.QMainWindow):
    def __init__(self, repo: InMemoryCellRepository | None = None) -> None:
        super(RnSApp, self).__init__()
        self.setGeometry(100, 100, 1400, 900)

        self.setWindowIcon(QIcon("./assets/rns-logo-sm.png"))
        self.data_table_label = QtWidgets.QLabel("Таблица с данными", self)
        self.data_table = DataTable(rows=50)

        self.plot = pg.PlotWidget()

        self.param_table_label = QtWidgets.QLabel("Таблица с расчетом", self)
        self.param_table = ParamTable()

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

        self.rn_consistent_group = QtWidgets.QGroupBox("Последовательное Rn (Ом)")
        self.rn_consistent_layout = QtWidgets.QHBoxLayout()
        self.rn_consistent = QtWidgets.QDoubleSpinBox(self)
        self.rn_consistent.setRange(0, 100)
        self.rn_consistent.setDecimals(2)
        self.rn_consistent.setValue(0)
        self.rn_consistent_layout.addWidget(self.rn_consistent)
        self.rn_consistent_group.setLayout(self.rn_consistent_layout)

        self.allowed_error_group = QtWidgets.QGroupBox("Допустимое отклонение (%)")
        self.allowed_error_layout = QtWidgets.QHBoxLayout()
        self.allowed_error = QtWidgets.QDoubleSpinBox(self)
        self.allowed_error.setRange(0, 100)
        self.allowed_error.setDecimals(2)
        self.allowed_error.setValue(0)
        self.allowed_error_layout.addWidget(self.allowed_error)
        self.allowed_error_group.setLayout(self.allowed_error_layout)

        self.cell_group = QtWidgets.QGroupBox("Запись")
        self.cell_v_layout = QtWidgets.QVBoxLayout()
        self.cell_h_layout = QtWidgets.QHBoxLayout()
        self.cell_grid_layout = QtWidgets.QGridLayout()
        self.cell_widgets = []
        for i in range(4):
            for j in range(4):
                index = i * 4 + j + 1
                cell_widget = CellWidget(self.cell_group, index, self.param_table, self)
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

        with contextlib.suppress(Exception):
            CDockManager.setConfigFlag(CDockManager.DockManagerFlag.AutoHideFeatureEnabled, True)

        self.dock_manager = CDockManager(self)

        data_container = QtWidgets.QWidget(self)
        data_layout = QtWidgets.QVBoxLayout()
        data_layout.setContentsMargins(6, 6, 6, 6)
        data_layout.setSpacing(6)
        data_layout.addWidget(self.data_table_label)
        data_layout.addWidget(self.data_table)
        data_container.setLayout(data_layout)
        data_dock = CDockWidget("Данные")
        data_dock.setWidget(data_container)

        params_group = QtWidgets.QGroupBox("Параметры")
        params_group_layout = QtWidgets.QHBoxLayout()
        params_group_layout.addWidget(self.rn_consistent_group)
        params_group_layout.addWidget(self.allowed_error_group)
        params_group.setLayout(params_group_layout)
        params_container = QtWidgets.QWidget(self)
        params_container_layout = QtWidgets.QVBoxLayout()
        params_container_layout.setContentsMargins(6, 6, 6, 6)
        params_container_layout.setSpacing(6)
        params_container_layout.addWidget(params_group)
        params_container_layout.addStretch(1)
        params_container.setLayout(params_container_layout)
        params_dock = CDockWidget("Параметры ввода")
        params_dock.setWidget(params_container)

        actions_container = QtWidgets.QWidget(self)
        actions_container_layout = QtWidgets.QVBoxLayout()
        actions_container_layout.setContentsMargins(6, 6, 6, 6)
        actions_container_layout.setSpacing(6)
        actions_container_layout.addWidget(self.actions_group)
        actions_container_layout.addStretch(1)
        actions_container.setLayout(actions_container_layout)
        actions_dock = CDockWidget("Действия")
        actions_dock.setWidget(actions_container)

        calc_container = QtWidgets.QWidget(self)
        calc_layout = QtWidgets.QVBoxLayout()
        calc_layout.setContentsMargins(6, 6, 6, 6)
        calc_layout.setSpacing(6)
        calc_layout.addWidget(self.param_table_label)
        calc_layout.addWidget(self.param_table)
        calc_layout.addStretch(1)
        calc_container.setLayout(calc_layout)
        calc_dock = CDockWidget("Расчет")
        calc_dock.setWidget(calc_container)

        plot_dock = CDockWidget("График")
        plot_dock.setWidget(self.plot)

        cells_container = QtWidgets.QWidget(self)
        cells_container_layout = QtWidgets.QVBoxLayout()
        cells_container_layout.setContentsMargins(6, 6, 6, 6)
        cells_container_layout.setSpacing(6)
        cells_container_layout.addWidget(self.cell_group)
        cells_container_layout.addStretch(1)
        cells_container.setLayout(cells_container_layout)
        cells_dock = CDockWidget("Запись")
        cells_dock.setWidget(cells_container)

        left_area = self.dock_manager.addDockWidget(DockWidgetArea.LeftDockWidgetArea, data_dock)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, params_dock, left_area)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, actions_dock, left_area)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, calc_dock, left_area)

        right_area = self.dock_manager.addDockWidget(DockWidgetArea.RightDockWidgetArea, plot_dock)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, cells_dock, right_area)

        for dock in (data_dock, params_dock, actions_dock, calc_dock, plot_dock, cells_dock):
            try:
                feats = dock.features()
                feats &= ~CDockWidget.DockWidgetFeature.DockWidgetClosable
                feats &= ~CDockWidget.DockWidgetFeature.DockWidgetFloatable
                dock.setFeatures(feats)
            except Exception:
                pass

        self.repo = repo or InMemoryCellRepository()
        self.calc = CalculationService(
            data_table=self.data_table,
            param_table=self.param_table,
            rn_consistent_widget=self.rn_consistent,
            allowed_error_widget=self.allowed_error,
        )
        self.plot_service = PlotService(self.plot, self.data_table, self.param_table)
        self.plot_service.prepare_plot()

        # Restore persisted layout and geometry if available
        with contextlib.suppress(Exception):
            self.restore_settings()

    def save_settings(self):
        settings = QSettings()
        settings.beginGroup("MainWindow")
        settings.setValue("geometry", self.saveGeometry())
        settings.endGroup()

        settings.beginGroup("DockManager")
        with contextlib.suppress(Exception):
            settings.setValue("state", self.dock_manager.saveState())

        settings.endGroup()

    def restore_settings(self):
        settings = QSettings()
        settings.beginGroup("MainWindow")
        geometry = settings.value("geometry")
        if geometry:
            with contextlib.suppress(Exception):
                self.restoreGeometry(geometry)
        settings.endGroup()

        settings.beginGroup("DockManager")
        state = settings.value("state")
        if state:
            with contextlib.suppress(Exception):
                self.dock_manager.restoreState(state)
        settings.endGroup()

    def closeEvent(self, event):
        with contextlib.suppress(Exception):
            self.save_settings()
        super().closeEvent(event)

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
        return self.calc.calculate_rn05()

    def calculate_main_params(self):
        return self.calc.calculate_main_params()

    def calculate_error_params(self):
        return self.calc.calculate_error_params()

    def calculate_rns_drift_square_per_sample(self):
        return self.calc.calculate_rns_drift_square_per_sample()

    def calculate_results(self):
        if not self.calc.calculate_results():
            return
        self.plot_current_data()

    def plot_current_data(self):
        return self.plot_service.plot_current_data()

    def addCellData(self, cell: int, name: str):
        self.repo.update_or_create_item(
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
        return self.plot_service.plot_cell(cell=cell, repo=self.repo)

    def remove_plot(self, cell: int):
        return self.plot_service.remove_cell_plot(cell=cell, store=self.repo)

    def prepare_plot(self):
        return self.plot_service.prepare_plot()

    def clean_rn(self):
        self.data_table.clear_rn()
        self.param_table.clear_all()

    def clean_all(self):
        self.data_table.clear_all()
        self.param_table.clear_all()
        self.plot.clear()
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
            repo=self.repo,
            data_headers=DataTableColumns.get_all_slugs(),
            results_headers=ParamTableColumns.get_all_names(),
        )

    def reload_tables_from_cell_data(self, cell: int):
        cell_data = self.repo.get(cell=cell)
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
            self.repo.clear()
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
        self.repo.clear()
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
                cell_item = self.repo.update_or_create_item(**kw)
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
