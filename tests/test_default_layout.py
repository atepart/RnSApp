import os
import tempfile
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6 import QtCore, QtWidgets

from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.xlsx_io import XlsxCellIO
from ui.app import RnSApp


class DefaultDockLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.settings_dir = tempfile.TemporaryDirectory()
        QtCore.QCoreApplication.setOrganizationName("RnSAppTests")
        QtCore.QCoreApplication.setApplicationName("DockLayout")
        QtCore.QSettings.setDefaultFormat(QtCore.QSettings.Format.IniFormat)
        QtCore.QSettings.setPath(
            QtCore.QSettings.Format.IniFormat,
            QtCore.QSettings.Scope.UserScope,
            cls.settings_dir.name,
        )
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    @classmethod
    def tearDownClass(cls):
        cls.settings_dir.cleanup()

    def setUp(self):
        QtCore.QSettings().clear()
        self.window = RnSApp(repo=InMemoryCellRepository(), excel_io=XlsxCellIO())
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

    def test_reference_layout_has_expected_splitter_hierarchy(self):
        assert (
            self.window.data_dock.dockAreaWidget().parentSplitter()
            is self.window.cells_dock.dockAreaWidget().parentSplitter()
        )
        assert (
            self.window.calc_dock.dockAreaWidget().parentSplitter()
            is self.window.inputs_dock.dockAreaWidget().parentSplitter()
        )
        assert (
            self.window.plot_dock.dockAreaWidget().parentSplitter()
            is self.window.area_calc_dock.dockAreaWidget().parentSplitter()
        )
        assert self.window.area_calc_dock.dockAreaWidget() is not self.window.inputs_dock.dockAreaWidget()

        root_sizes = self.window.dock_manager.dockContainers()[0].rootSplitter().sizes()
        assert len(root_sizes) == 2
        assert root_sizes[1] > root_sizes[0]

    def test_saved_layout_is_restored_regardless_of_legacy_version(self):
        self.window.inputs_dock.toggleView(False)
        self.app.processEvents()
        saved_state = self.window.dock_manager.saveState()
        self.window.inputs_dock.toggleView(True)
        self.app.processEvents()
        assert not self.window.inputs_dock.isClosed()

        settings = QtCore.QSettings()
        settings.beginGroup("DockManager")
        settings.setValue("layout_version", 1)
        settings.setValue("state", saved_state)
        settings.endGroup()

        assert self.window.restore_settings()
        self.app.processEvents()
        assert self.window.inputs_dock.isClosed()

        self.window.save_settings()
        settings.beginGroup("DockManager")
        assert settings.value("state")
        assert settings.value("layout_version") is None
        settings.endGroup()

    def test_restore_default_layout_shows_all_panels(self):
        self.window.inputs_dock.toggleView(False)
        self.app.processEvents()
        assert self.window.inputs_dock.isClosed()
        self.window.restore_default_layout()
        self.app.processEvents()
        assert not self.window.inputs_dock.isClosed()


if __name__ == "__main__":
    unittest.main()
