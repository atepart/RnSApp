import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pyqtgraph.exporters import ImageExporter
from PySide6 import QtCore, QtTest, QtWidgets

from domain.constants import DataTableColumns, ParamTableColumns
from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.xlsx_io import XlsxCellIO
from ui.app import RnSApp
from ui.widgets import TableWidgetItem


class RealUiWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = None
        artifact_dir = os.environ.get("RNS_UI_ARTIFACT_DIR")
        if artifact_dir:
            cls.artifact_dir = Path(artifact_dir)
            cls.artifact_dir.mkdir(parents=True, exist_ok=True)
        else:
            cls.temp_dir = tempfile.TemporaryDirectory()
            cls.artifact_dir = Path(cls.temp_dir.name)

        QtCore.QCoreApplication.setOrganizationName("RnSAppTests")
        QtCore.QCoreApplication.setApplicationName("RealUiWorkflow")
        QtCore.QSettings.setDefaultFormat(QtCore.QSettings.Format.IniFormat)
        QtCore.QSettings.setPath(
            QtCore.QSettings.Format.IniFormat,
            QtCore.QSettings.Scope.UserScope,
            str(cls.artifact_dir / "settings"),
        )
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    @classmethod
    def tearDownClass(cls):
        if cls.temp_dir is not None:
            cls.temp_dir.cleanup()

    def setUp(self):
        QtCore.QSettings().clear()
        self.repo = InMemoryCellRepository()
        self.xlsx_io = XlsxCellIO()
        self.window = RnSApp(repo=self.repo, excel_io=self.xlsx_io)
        self.window.resize(1920, 1039)
        self.window.show()
        self.app.processEvents()

    def tearDown(self):
        export_dialog = self.window.plot.scene().exportDialog
        if export_dialog is not None:
            export_dialog.close()
            export_dialog.deleteLater()
            self.window.plot.scene().exportDialog = None
        self.window.close()
        self.window.deleteLater()
        QtWidgets.QApplication.sendPostedEvents(None, QtCore.QEvent.Type.DeferredDelete.value)
        self.app.processEvents()

    def test_calculate_record_plot_export_and_save(self):
        expected_slope = 0.05
        expected_intercept = 0.1
        for row, diameter in enumerate((1.0, 2.0, 4.0, 8.0)):
            transformed_y = expected_slope * diameter + expected_intercept
            resistance = 1.0 / transformed_y**2
            self.window.data_table.setItem(row, DataTableColumns.NAME.index, TableWidgetItem(f"Sample {row + 1}"))
            self.window.data_table.setItem(row, DataTableColumns.DIAMETER.index, TableWidgetItem(str(diameter)))
            self.window.data_table.setItem(
                row,
                DataTableColumns.RESISTANCE.index,
                TableWidgetItem(f"{resistance:.12f}"),
            )

        QtTest.QTest.mouseClick(self.window.result_button, QtCore.Qt.MouseButton.LeftButton)
        self.app.processEvents()
        slope = self.window.param_table.get_column_value(0, ParamTableColumns.SLOPE)
        intercept = self.window.param_table.get_column_value(0, ParamTableColumns.INTERCEPT)
        self.assertAlmostEqual(slope, expected_slope, places=9)
        self.assertAlmostEqual(intercept, expected_intercept, places=9)
        assert (self.artifact_dir / "01-calculated-layout.png").parent.exists()
        assert self.window.grab().save(str(self.artifact_dir / "01-calculated-layout.png"))

        cell = self.window.cell_widgets[0]
        long_name = "Weighted integration cell with a long legend name"
        with mock.patch.object(QtWidgets.QInputDialog, "getText", return_value=(long_name, True)):
            QtTest.QTest.mouseClick(cell.writeButton, QtCore.Qt.MouseButton.LeftButton)
        self.app.processEvents()
        assert self.repo.get(cell=1) is not None
        assert cell.checkbox.isVisible()
        assert self.window.plot_data(1)
        cell.checkbox.click()
        self.app.processEvents()
        assert cell.checkbox.isChecked()
        assert self.window.grab().save(str(self.artifact_dir / "02-recorded-cell-and-legend.png"))

        plot_export = self.artifact_dir / "03-large-font-plot-export.png"
        exporter = ImageExporter(self.window.plot.getPlotItem())
        exporter.parameters()["width"] = 1800
        image = self.window.plot_service.export_visual(exporter, exporter.export, toBytes=True)
        assert image.save(str(plot_export))
        QtWidgets.QApplication.clipboard().setImage(image)
        assert not QtWidgets.QApplication.clipboard().image().isNull()
        QtWidgets.QApplication.clipboard().clear()
        assert plot_export.exists()
        assert plot_export.stat().st_size > 0

        xlsx_path = self.artifact_dir / "04-weighted-fit.xlsx"
        grid_values = [
            (widget.name.text(), widget.drift.text(), widget.rns.text(), widget.rns_error.text())
            for widget in self.window.cell_widgets
        ]
        self.xlsx_io.save(str(xlsx_path), grid_values, self.repo)
        assert xlsx_path.exists()
        assert xlsx_path.stat().st_size > 0


if __name__ == "__main__":
    unittest.main()
