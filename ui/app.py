import contextlib
import logging
import os

import numpy as np
import pyqtgraph as pg
from PySide6 import QtCore, QtWidgets
from PySide6.QtCore import QSettings
from PySide6.QtGui import QAction, QDesktopServices, QIcon
from PySide6.QtWidgets import QApplication, QInputDialog
from PySide6QtAds import CDockManager, CDockWidget, DockWidgetArea

from application.calculations import CalculationService
from application.version import REPO_SLUG, __version__
from domain.constants import DataTableColumns, ParamTableColumns
from domain.ports import CellDataIO
from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.template_io import load_template, save_template
from infrastructure.xlsx_io import XlsxCellIO
from ui.plotting_service import PlotService
from ui.update_dialogs import FetchReleasesWorker, ReleasePickerDialog
from ui.widgets import CellWidget, DataTable, ParamTable

logger = logging.getLogger(__name__)


class RnSApp(QtWidgets.QMainWindow):
    def __init__(self, repo: InMemoryCellRepository | None = None, excel_io: CellDataIO | None = None) -> None:
        super(RnSApp, self).__init__()
        self.setGeometry(100, 100, 1400, 900)
        # Keep references to background workers/threads to prevent GC-related crashes
        self._update_fetch_thread = None
        self._update_fetch_worker = None
        self._update_fetch_timer = None
        self._update_spinner = None

        self.setWindowIcon(QIcon("./assets/rns-logo-alt.ico"))
        # Табличка с данными без заголовка-лейбла
        self.data_table = DataTable(rows=50)

        self.plot = pg.PlotWidget()

        # Таблица с расчетом без заголовка-лейбла
        self.param_table = ParamTable()

        # Контейнер действий без рамки группы
        self.actions_group = QtWidgets.QWidget()
        self.actions_layout = QtWidgets.QHBoxLayout()

        self.result_button = QtWidgets.QPushButton("Рассчитать")
        self.result_button.setToolTip("Произвести расчеты")
        self.result_button.clicked.connect(self.calculate_results)

        self.clean_rn_button = QtWidgets.QPushButton("Очистить Rn")
        self.clean_rn_button.setToolTip("Очистить стоблец Rn и таблицу с расчетом")
        self.clean_rn_button.clicked.connect(self.clean_rn)

        self.clean_all_button = QtWidgets.QPushButton("Очистить таблицы")
        self.clean_all_button.setToolTip("Очистить график и таблицы с даными и расчетом")
        self.clean_all_button.clicked.connect(self.clean_all)

        self.open_template_button = QtWidgets.QPushButton("Открыть шаблон")
        self.open_template_button.setToolTip("Загрузить шаблон данных (имена, диаметры, площади)")
        self.open_template_button.clicked.connect(self.open_template)

        # Make buttons expand to fill the bottom row
        for btn in (self.result_button, self.clean_rn_button, self.clean_all_button, self.open_template_button):
            btn.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)

        self.actions_layout.addWidget(self.result_button)
        self.actions_layout.addWidget(self.clean_rn_button)
        self.actions_layout.addWidget(self.clean_all_button)
        self.actions_layout.addWidget(self.open_template_button)
        self.actions_group.setLayout(self.actions_layout)

        # Поля параметров ввода как аккуратные лейблы + спинбоксы
        self.rn_consistent = QtWidgets.QDoubleSpinBox(self)
        self.rn_consistent.setRange(0, 100)
        self.rn_consistent.setDecimals(2)
        self.rn_consistent.setValue(0)

        self.allowed_error = QtWidgets.QDoubleSpinBox(self)
        self.allowed_error.setRange(0, 100)
        self.allowed_error.setDecimals(2)
        self.allowed_error.setValue(2.5)

        # Контейнер без заголовка; сами ячейки остаются группами
        self.cell_group = QtWidgets.QWidget()
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

        self.btn_load_cell_data = QtWidgets.QPushButton("Открыть")
        self.btn_load_cell_data.setToolTip("Открыть данные из xlsx файла")
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
        data_layout.addWidget(self.data_table)
        data_container.setLayout(data_layout)
        self.data_dock = CDockWidget("Данные")
        self.data_dock.setWidget(data_container)

        # Объединенные параметры ввода + действия
        # Inputs for custom nominal areas (um^2)
        self.s_custom1 = QtWidgets.QDoubleSpinBox(self)
        self.s_custom1.setDecimals(3)
        self.s_custom1.setRange(0, 100000)
        self.s_custom1.setValue(1.0)
        self.s_custom2 = QtWidgets.QDoubleSpinBox(self)
        self.s_custom2.setDecimals(3)
        self.s_custom2.setRange(0, 100000)
        self.s_custom2.setValue(1.0)

        inputs_container = QtWidgets.QWidget(self)
        inputs_layout = QtWidgets.QVBoxLayout()
        inputs_layout.setContentsMargins(6, 6, 6, 6)
        inputs_layout.setSpacing(6)
        # 2x2 grid of parameters: each field is label + widget
        inputs_grid = QtWidgets.QGridLayout()
        inputs_grid.setContentsMargins(0, 0, 0, 0)
        inputs_grid.setHorizontalSpacing(12)
        inputs_grid.setVerticalSpacing(6)
        # Row 0
        inputs_grid.addWidget(QtWidgets.QLabel("Последовательное Rn (Ω):"), 0, 0)
        inputs_grid.addWidget(self.rn_consistent, 0, 1)
        inputs_grid.addWidget(QtWidgets.QLabel("Допустимое отклонение (%):"), 0, 2)
        inputs_grid.addWidget(self.allowed_error, 0, 3)
        # Row 1
        inputs_grid.addWidget(QtWidgets.QLabel("Заданная площадь S2 (μm²):"), 1, 0)
        inputs_grid.addWidget(self.s_custom1, 1, 1)
        inputs_grid.addWidget(QtWidgets.QLabel("Заданная площадь S3 (μm²):"), 1, 2)
        inputs_grid.addWidget(self.s_custom2, 1, 3)
        inputs_layout.addLayout(inputs_grid)
        inputs_layout.addStretch(1)
        inputs_layout.addWidget(self.actions_group)
        inputs_container.setLayout(inputs_layout)
        self.inputs_dock = CDockWidget("Действия")
        self.inputs_dock.setWidget(inputs_container)

        calc_container = QtWidgets.QWidget(self)
        calc_layout = QtWidgets.QVBoxLayout()
        calc_layout.setContentsMargins(6, 6, 6, 6)
        calc_layout.setSpacing(6)
        calc_layout.addWidget(self.param_table)
        calc_layout.addStretch(1)
        calc_container.setLayout(calc_layout)
        self.calc_dock = CDockWidget("Расчет")
        self.calc_dock.setWidget(calc_container)

        self.plot_dock = CDockWidget("График")
        self.plot_dock.setWidget(self.plot)

        cells_container = QtWidgets.QWidget(self)
        cells_container_layout = QtWidgets.QVBoxLayout()
        cells_container_layout.setContentsMargins(6, 6, 6, 6)
        cells_container_layout.setSpacing(6)
        cells_container_layout.addWidget(self.cell_group)
        cells_container_layout.addStretch(1)
        cells_container.setLayout(cells_container_layout)
        self.cells_dock = CDockWidget("Запись")
        self.cells_dock.setWidget(cells_container)

        left_area = self.dock_manager.addDockWidget(DockWidgetArea.LeftDockWidgetArea, self.data_dock)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, self.inputs_dock, left_area)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, self.calc_dock, left_area)

        right_area = self.dock_manager.addDockWidget(DockWidgetArea.RightDockWidgetArea, self.plot_dock)
        self.dock_manager.addDockWidget(DockWidgetArea.BottomDockWidgetArea, self.cells_dock, right_area)

        for dock in (self.data_dock, self.inputs_dock, self.calc_dock, self.plot_dock, self.cells_dock):
            try:
                feats = dock.features()
                feats &= ~CDockWidget.DockWidgetFeature.DockWidgetClosable
                feats &= ~CDockWidget.DockWidgetFeature.DockWidgetFloatable
                dock.setFeatures(feats)
            except Exception:
                pass

        # Save default dock layout state for restoring on demand
        with contextlib.suppress(Exception):
            self.default_dock_state = self.dock_manager.saveState()

        # Top toolbar with action to restore default layout
        toolbar = QtWidgets.QToolBar("Вид")
        self.addToolBar(toolbar)
        act_restore = QAction("Восстановить виджеты", self)
        act_restore.setToolTip("Построить расположение панелей по умолчанию")
        act_restore.triggered.connect(self.restore_default_layout)
        toolbar.addAction(act_restore)
        act_create_template = QAction("Создать шаблон", self)
        act_create_template.setToolTip("Сохранить шаблон с именами, диаметрами и площадями")
        act_create_template.triggered.connect(self.create_template)
        toolbar.addAction(act_create_template)

        # Menu: Help
        help_menu = self.menuBar().addMenu("Справка")
        act_about = QAction("О программе", self)
        act_about.triggered.connect(self.show_about)
        help_menu.addAction(act_about)
        act_check_update = QAction("Проверить обновления", self)
        act_check_update.triggered.connect(self.check_updates)
        help_menu.addAction(act_check_update)

        self.repo = repo or InMemoryCellRepository()
        self.excel_io: CellDataIO = excel_io or XlsxCellIO()
        self.active_cell_index: int | None = None
        self.calc = CalculationService(
            data_table=self.data_table,
            param_table=self.param_table,
            rn_consistent_widget=self.rn_consistent,
            allowed_error_widget=self.allowed_error,
            s_custom1_widget=self.s_custom1,
            s_custom2_widget=self.s_custom2,
        )
        self.plot_service = PlotService(self.plot, self.data_table, self.param_table)
        self.plot_service.prepare_plot()

        # Restore persisted layout and geometry if available
        with contextlib.suppress(Exception):
            self.restore_settings()
        # Always ensure all docks are visible on startup
        with contextlib.suppress(Exception):
            for dock in (self.inputs_dock, self.calc_dock, self.data_dock, self.plot_dock, self.cells_dock):
                dock.setVisible(True)
                dock.show()

    # ----- Settings helpers for file dialogs -----
    def _get_initial_directory(self) -> str:
        """Return a starting directory for file dialogs.

        Uses last successful path from settings; otherwise falls back to filesystem root.
        """
        settings = QSettings()
        settings.beginGroup("FileDialog")
        last_dir = settings.value("last_dir", type=str)
        settings.endGroup()
        if isinstance(last_dir, str) and last_dir and os.path.isdir(last_dir):
            return last_dir
        # Fallback: filesystem root to show all volumes (Unix '/')
        try:
            from PySide6.QtCore import QDir

            return QDir.rootPath()
        except Exception:
            return os.path.abspath(os.sep)

    def _remember_path(self, path: str) -> None:
        """Store the directory part of the given path into settings."""
        if not path:
            return
        directory = path if os.path.isdir(path) else os.path.dirname(path)
        if not directory:
            return
        settings = QSettings()
        settings.beginGroup("FileDialog")
        settings.setValue("last_dir", directory)
        settings.endGroup()

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

    def restore_default_layout(self):
        # Restore dock layout to the default captured state
        with contextlib.suppress(Exception):
            if hasattr(self, "default_dock_state") and self.default_dock_state:
                self.dock_manager.restoreState(self.default_dock_state)
        # Ensure all docks are visible
        with contextlib.suppress(Exception):
            for dock in (self.inputs_dock, self.calc_dock, self.data_dock, self.plot_dock, self.cells_dock):
                dock.setVisible(True)
                dock.show()
        # Persist current layout as the new state
        with contextlib.suppress(Exception):
            self.save_settings()

    def show_about(self):
        QtWidgets.QMessageBox.information(
            self,
            "О программе",
            f"RnSApp\nВерсия: {__version__}",
        )

    def check_updates(self):
        # Phase 1: fetch releases in a worker thread
        spinner = QtWidgets.QProgressDialog("Получение списка релизов...", "Отмена", 0, 0, self)
        spinner.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        spinner.setAutoClose(True)
        spinner.show()
        self._update_spinner = spinner

        # Таймер на 10 секунд: по истечении — ошибка и закрытие
        timer = QtCore.QTimer(self)
        timer.setSingleShot(True)
        timer.setInterval(10_000)
        self._update_fetch_timer = timer
        logger.info("Starting release fetch in background thread")

        thread = QtCore.QThread(self)
        worker = FetchReleasesWorker(REPO_SLUG, limit=10)
        worker.moveToThread(thread)
        # Keep Python references to avoid premature GC that can crash Qt threads
        self._update_fetch_thread = thread
        self._update_fetch_worker = worker

        thread.started.connect(worker.run)
        connection_type = QtCore.Qt.QueuedConnection
        worker.finished.connect(self._on_update_fetch_finished, connection_type)
        worker.error.connect(self._on_update_fetch_error, connection_type)
        worker.status.connect(self._on_update_fetch_status, connection_type)
        worker.finished.connect(worker.deleteLater)
        worker.error.connect(worker.deleteLater)
        worker.status.connect(lambda *_: None)
        thread.finished.connect(thread.deleteLater)
        thread.start()
        timer.timeout.connect(self._on_update_fetch_timeout, connection_type)
        timer.start()

    @QtCore.Slot(list)
    def _on_update_fetch_finished(self, releases: list):
        self._stop_update_timer()
        self._cleanup_update_thread(wait=True)
        spinner = self._update_spinner
        self._update_spinner = None

        if not releases:
            if spinner:
                spinner.close()
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось получить список релизов. Попробуйте позже.")
            return

        if spinner:
            spinner.close()

        # Show picker dialog on main thread
        dlg = ReleasePickerDialog(releases, parent=self)
        if dlg.exec() != QtWidgets.QDialog.Accepted or not dlg.selected:
            return
        selected = dlg.selected
        if not getattr(selected, "asset", None) or not selected.asset.download_url:
            QtWidgets.QMessageBox.information(self, "Нет файла", "В выбранном релизе нет файла для вашей платформы.")
            return
        # Только выводим ссылку на скачивание (без автоматической загрузки)
        url = selected.asset.download_url
        logger.info(f"Selected release: {selected.tag}, asset: {url}")
        self._show_download_link(url)

    @QtCore.Slot(str)
    def _on_update_fetch_error(self, msg: str):
        self._stop_update_timer()
        self._cleanup_update_thread(wait=True)
        if self._update_spinner:
            self._update_spinner.close()
        self._update_spinner = None
        QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось получить список релизов: {msg}")

    @QtCore.Slot(str)
    def _on_update_fetch_status(self, text: str):
        spinner = self._update_spinner
        if spinner:
            with contextlib.suppress(Exception):
                spinner.setLabelText(text)

    @QtCore.Slot()
    def _on_update_fetch_timeout(self):
        self._stop_update_timer()
        self._cleanup_update_thread(wait=True)
        if self._update_spinner:
            self._update_spinner.close()
        self._update_spinner = None
        QtWidgets.QMessageBox.critical(self, "Ошибка", "Таймаут: не удалось получить список релизов за 10 секунд")

    def _stop_update_timer(self):
        timer = self._update_fetch_timer
        if timer:
            with contextlib.suppress(Exception):
                timer.stop()
        self._update_fetch_timer = None

    def _cleanup_update_thread(self, wait: bool = False):
        thread = self._update_fetch_thread
        if thread:
            with contextlib.suppress(Exception):
                thread.requestInterruption()
                thread.quit()
                if wait:
                    thread.wait(3000)
        self._update_fetch_thread = None
        self._update_fetch_worker = None

    def closeEvent(self, event):
        with contextlib.suppress(Exception):
            self.save_settings()
        for window in QApplication.topLevelWidgets():
            window.close()
        super().closeEvent(event)

    def _show_download_link(self, url: str):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Загрузка обновления")
        dlg.resize(620, 180)
        layout = QtWidgets.QVBoxLayout(dlg)

        info = QtWidgets.QLabel("Ссылка на релиз. Нажмите, чтобы открыть в браузере, или скопируйте.")
        info.setWordWrap(True)
        layout.addWidget(info)

        link = QtWidgets.QLabel(f'<a href="{url}">{url}</a>')
        link.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        link.setOpenExternalLinks(True)
        layout.addWidget(link)

        btns = QtWidgets.QHBoxLayout()
        btn_copy = QtWidgets.QPushButton("Скопировать ссылку")
        btn_open = QtWidgets.QPushButton("Открыть в браузере")
        btn_close = QtWidgets.QPushButton("Закрыть")
        btns.addWidget(btn_copy)
        btns.addWidget(btn_open)
        btns.addStretch(1)
        btns.addWidget(btn_close)
        layout.addLayout(btns)

        def copy_link():
            cb = QApplication.clipboard()
            if cb:
                cb.setText(url)

        def open_link():
            QDesktopServices.openUrl(QtCore.QUrl(url))

        btn_copy.clicked.connect(copy_link)
        btn_open.clicked.connect(open_link)
        btn_close.clicked.connect(dlg.accept)

        dlg.exec()

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
        # Uncheck rows where Rn (Ω) is empty before calculations
        self.data_table.uncheck_rows_with_empty_rn()
        if not self.calc.calculate_results():
            return
        self.plot_current_data()
        # Update dirty flag for active cell if any
        with contextlib.suppress(Exception):
            if self.active_cell_index:
                self._update_dirty_flag_for_cell(self.active_cell_index)

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
            s_custom1=self.s_custom1.value(),
            s_custom2=self.s_custom2.value(),
            s_real_1=self.param_table.get_column_value(0, ParamTableColumns.S_REAL_1),
            s_real_custom1=self.param_table.get_column_value(0, ParamTableColumns.S_REAL_CUSTOM1),
            s_real_custom2=self.param_table.get_column_value(0, ParamTableColumns.S_REAL_CUSTOM2),
        )
        self.set_active_cell(cell)
        # After saving, clear dirty indicator for this cell
        with contextlib.suppress(Exception):
            self.cell_widgets[cell - 1].set_dirty(False)

    # ---- Helpers to detect unsaved changes for active cell ----
    def _param_table_snapshot(self) -> dict:
        values = {}
        for col in ParamTableColumns:
            with contextlib.suppress(Exception):
                values[col.slug] = self.param_table.get_column_value(0, col)
        return values

    @staticmethod
    def _float_equal(a, b, tol: float = 1e-6) -> bool:
        if a in (None, "") and b in (None, ""):
            return True
        try:
            fa = float(a)
            fb = float(b)
            return abs(fa - fb) <= tol
        except Exception:
            return a == b

    def _update_dirty_flag_for_cell(self, cell: int) -> None:
        item = self.repo.get(cell=cell)
        if not item:
            return
        snap = self._param_table_snapshot()
        compare_cols = [
            ParamTableColumns.SLOPE,
            ParamTableColumns.INTERCEPT,
            ParamTableColumns.DRIFT,
            ParamTableColumns.RNS,
            ParamTableColumns.DRIFT_ERROR,
            ParamTableColumns.RNS_ERROR,
            ParamTableColumns.RN_CONSISTENT,
            ParamTableColumns.ALLOWED_ERROR,
            ParamTableColumns.S_REAL_1,
            ParamTableColumns.S_REAL_CUSTOM1,
            ParamTableColumns.S_REAL_CUSTOM2,
        ]
        # Include custom area inputs if available
        with contextlib.suppress(Exception):
            compare_cols.extend([ParamTableColumns.S_CUSTOM1, ParamTableColumns.S_CUSTOM2])
        is_dirty = False
        for col in compare_cols:
            saved_val = getattr(item, col.slug, None)
            curr_val = snap.get(col.slug)
            if not self._float_equal(saved_val, curr_val):
                is_dirty = True
                break
        with contextlib.suppress(Exception):
            self.cell_widgets[cell - 1].set_dirty(is_dirty)

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
        start_dir = self._get_initial_directory()
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Cell Data",
            start_dir,
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )
        if not file_name:
            return
        self._remember_path(file_name)
        init_data = [(cell.name.text(), cell.drift.text(), cell.rns.text()) for cell in self.cell_widgets]
        self.excel_io.save(
            file_name=file_name,
            cell_grid_values=init_data,
            repo=self.repo,
        )

    def create_template(self):
        name, ok = QInputDialog.getText(self, "Создать шаблон", "Введите уникальное имя шаблона:")
        if not ok:
            return
        name = name.strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "Пустое имя", "Имя шаблона не может быть пустым.")
            return

        start_dir = self._get_initial_directory()
        default_path = os.path.join(start_dir, f"{name}.xlsx")
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Сохранить шаблон",
            default_path,
            "Excel Files (*.xlsx);;All Files (*)",
        )
        if not file_name:
            return
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        if os.path.exists(file_name):
            QtWidgets.QMessageBox.warning(
                self, "Имя уже занято", "Файл с таким именем уже существует, выберите другое."
            )
            return

        rows: list[dict] = []
        for row in range(self.data_table.rowCount()):
            name_item = self.data_table.item(row, DataTableColumns.NAME.index)
            name_val = name_item.text().strip() if name_item and name_item.text() else ""
            diam_item = self.data_table.item(row, DataTableColumns.DIAMETER.index)
            diam_val = None
            if diam_item and diam_item.text():
                with contextlib.suppress(Exception):
                    diam_val = float(str(diam_item.text()).replace(",", "."))
            cb = self.data_table.get_row_checkbox(row)
            selected = bool(cb.isChecked()) if cb else False
            if any([name_val, selected, diam_val not in (None, "")]):
                rows.append(
                    {
                        "number": row + 1,
                        "name": name_val,
                        "selected": selected,
                        "diameter": diam_val,
                    }
                )

        if not rows:
            QtWidgets.QMessageBox.warning(self, "Нет данных", "Не заполнено ни одной строки для шаблона.")
            return

        areas = {
            "s_custom1": float(self.s_custom1.value()) if self.s_custom1 is not None else None,
            "s_custom2": float(self.s_custom2.value()) if self.s_custom2 is not None else None,
        }

        try:
            saved_path = save_template(file_path=file_name, sheet_name=name, rows=rows, areas=areas)
            self._remember_path(saved_path)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка сохранения", f"Не удалось сохранить шаблон: {e}")
            logger.exception(e, exc_info=True)
            return

        QtWidgets.QMessageBox.information(self, "Готово", "Шаблон сохранён.")

    def open_template(self):
        options = QtWidgets.QFileDialog.Options()
        start_dir = self._get_initial_directory()
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Открыть шаблон",
            start_dir,
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )
        if not file_name:
            return
        self._remember_path(file_name)

        try:
            initial_data, areas, errors = load_template(file_name)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка чтения", f"Не удалось открыть шаблон: {e}")
            logger.exception(e, exc_info=True)
            return

        self.clean_all()
        self.data_table.load_data(initial_data)
        with contextlib.suppress(Exception):
            if areas.get("s_custom1") is not None:
                self.s_custom1.setValue(float(areas["s_custom1"]))
            if areas.get("s_custom2") is not None:
                self.s_custom2.setValue(float(areas["s_custom2"]))

        if errors:
            QtWidgets.QMessageBox.warning(self, "Шаблон загружен с предупреждениями", "\n".join(errors))
        else:
            QtWidgets.QMessageBox.information(self, "Готово", "Шаблон загружен.")

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
        # Load nominal custom areas into input fields if present
        with contextlib.suppress(Exception):
            if getattr(cell_data, "s_custom1", None) not in (None, ""):
                self.s_custom1.setValue(float(cell_data.s_custom1))
        with contextlib.suppress(Exception):
            if getattr(cell_data, "s_custom2", None) not in (None, ""):
                self.s_custom2.setValue(float(cell_data.s_custom2))
        # Clear dirty flag on load
        with contextlib.suppress(Exception):
            self.cell_widgets[cell - 1].set_dirty(False)

    def set_active_cell(self, cell: int):
        self.active_cell_index = cell if cell else None
        for cw in self.cell_widgets:
            is_active = cw.index == cell
            cw.set_active(is_active)
            # Star appears only on active cell; hide for others when switching
            with contextlib.suppress(Exception):
                if not is_active:
                    cw.set_dirty(False)

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
        start_dir = self._get_initial_directory()
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            None, "Выберите файл XLSX", start_dir, "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if not fileName:
            return
        self._remember_path(fileName)

        is_some_errors = False
        some_errors_text = ""
        try:
            items, errors = self.excel_io.load(fileName)
            first_cell = None
            for kw in items:
                cell_item = self.repo.update_or_create_item(**kw)
                if first_cell is None:
                    first_cell = cell_item.cell
                cell_widget = self.cell_widgets[cell_item.cell - 1]
                cell_widget.name.setText(cell_item.name)
                cell_widget.drift.setText(f"Уход: {round(cell_item.drift, 3)}")
                cell_widget.rns.setText(f"RnS: {round(cell_item.rns, 1)}")
                cell_widget.updateUI()
                self.calculate_means()

            # Load first imported cell into current tables (so ParamTable shows S_real_* etc.)
            if first_cell is not None:
                self.reload_tables_from_cell_data(first_cell)

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
