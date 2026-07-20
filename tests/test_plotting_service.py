import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pyqtgraph as pg
from PySide6 import QtWidgets

from ui.plotting_service import (
    EXPORT_AXIS_LABEL_SIZE,
    EXPORT_LEGEND_TEXT_SIZE,
    EXPORT_TICK_FONT_SIZE,
    SCREEN_AXIS_LABEL_SIZE,
    SCREEN_LEGEND_TEXT_SIZE,
    SCREEN_TICK_FONT_SIZE,
    PlotService,
    StyledExportDialog,
)


class PlotServiceStyleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def setUp(self):
        self.plot = pg.PlotWidget()
        self.service = PlotService(self.plot, data_table=None, param_table=None)
        self.service.prepare_plot()
        self.plot.plot([1.0, 2.0], [2.0, 3.0], name="Long cell name")
        self.app.processEvents()

    def tearDown(self):
        export_dialog = self.plot.scene().exportDialog
        if export_dialog is not None:
            export_dialog.close()
            export_dialog.deleteLater()
            self.plot.scene().exportDialog = None
        self.plot.close()
        self.plot.deleteLater()
        QtWidgets.QApplication.sendPostedEvents(None, 0)
        self.app.processEvents()

    def _assert_profile(self, axis_size, tick_size, legend_size):
        plot_item = self.plot.getPlotItem()
        for axis_name in ("left", "bottom"):
            axis = plot_item.getAxis(axis_name)
            assert axis.labelStyle["font-size"] == axis_size
            assert axis.style["tickFont"].pointSize() == tick_size
        assert plot_item.legend.labelTextSize() == legend_size

    def test_prepare_plot_applies_screen_profile_and_custom_export_dialog(self):
        assert int(SCREEN_LEGEND_TEXT_SIZE.removesuffix("pt")) >= 16
        assert int(EXPORT_LEGEND_TEXT_SIZE.removesuffix("pt")) >= 26
        self._assert_profile(SCREEN_AXIS_LABEL_SIZE, SCREEN_TICK_FONT_SIZE, SCREEN_LEGEND_TEXT_SIZE)
        assert isinstance(self.plot.scene().exportDialog, StyledExportDialog)

    def test_export_profile_is_larger_and_always_restores_screen_profile(self):
        with self.service.export_font_profile():
            self._assert_profile(EXPORT_AXIS_LABEL_SIZE, EXPORT_TICK_FONT_SIZE, EXPORT_LEGEND_TEXT_SIZE)
        self._assert_profile(SCREEN_AXIS_LABEL_SIZE, SCREEN_TICK_FONT_SIZE, SCREEN_LEGEND_TEXT_SIZE)

        with self.assertRaisesRegex(RuntimeError, "render failed"):
            with self.service.export_font_profile():
                self._assert_profile(EXPORT_AXIS_LABEL_SIZE, EXPORT_TICK_FONT_SIZE, EXPORT_LEGEND_TEXT_SIZE)
                raise RuntimeError("render failed")
        self._assert_profile(SCREEN_AXIS_LABEL_SIZE, SCREEN_TICK_FONT_SIZE, SCREEN_LEGEND_TEXT_SIZE)


if __name__ == "__main__":
    unittest.main()
