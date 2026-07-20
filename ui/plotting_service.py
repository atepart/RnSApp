import contextlib
from contextlib import contextmanager

import numpy as np
import pyqtgraph as pg
from pyqtgraph.exporters import ImageExporter, SVGExporter
from pyqtgraph.GraphicsScene import exportDialog as pg_export_dialog
from PySide6 import QtCore, QtGui, QtWidgets

from domain.constants import PLOT_COLORS, DataTableColumns, ParamTableColumns
from domain.errors import ListsNotSameLength
from domain.utils import drop_nans, linear

SCREEN_AXIS_LABEL_SIZE = "20px"
SCREEN_TICK_FONT_SIZE = 12
SCREEN_LEGEND_TEXT_SIZE = "16pt"
EXPORT_AXIS_LABEL_SIZE = "30px"
EXPORT_TICK_FONT_SIZE = 18
EXPORT_LEGEND_TEXT_SIZE = "26pt"
EXPORT_SOURCE_BOTTOM_PADDING = 14.0


class StyledExportDialog(pg_export_dialog.ExportDialog):
    """Apply the larger plot profile only while a visual exporter renders."""

    def __init__(self, scene, plot_service) -> None:
        self.plot_service = plot_service
        super().__init__(scene)

    def exportFormatChanged(self, item, prev):  # noqa: N802 - Qt override
        super().exportFormatChanged(item, prev)
        exporter = self.currentExporter
        if exporter is None or not isinstance(exporter, (ImageExporter, SVGExporter)):
            return
        if getattr(exporter, "_rns_export_style_wrapped", False):
            return

        original_export = exporter.export

        def styled_export(*args, **kwargs):
            return self.plot_service.export_visual(exporter, original_export, *args, **kwargs)

        exporter.export = styled_export
        exporter._rns_export_style_wrapped = True


class PlotService:
    def __init__(self, plot_widget, data_table, param_table) -> None:
        self.plot = plot_widget
        self.data_table = data_table
        self.param_table = param_table

    @staticmethod
    def _cell_plot_names(item):
        return {item.name, f"{item.name} (fit)", f"{item.name} (data)"}

    @staticmethod
    def _mark_cell_plot_item(plot_item, cell: int):
        setattr(plot_item, "_rns_cell", cell)

    def prepare_plot(self):
        y_label = "1/√Rₙ"
        x_label = "Диаметр ACAD (μm)"
        self.plot.setBackground("w")
        self.plot.setLabel("left", y_label, color="#413C58")
        self.plot.setLabel("bottom", x_label, color="#413C58")
        plot_item = self.plot.getPlotItem()
        if plot_item.legend is None:
            self.plot.addLegend()
        self.apply_font_profile(
            axis_label_size=SCREEN_AXIS_LABEL_SIZE,
            tick_font_size=SCREEN_TICK_FONT_SIZE,
            legend_text_size=SCREEN_LEGEND_TEXT_SIZE,
        )
        self.configure_export_dialog()
        self.plot.showGrid(x=True, y=True)

    def configure_export_dialog(self) -> None:
        scene = self.plot.scene()
        if not isinstance(getattr(scene, "exportDialog", None), StyledExportDialog):
            scene.exportDialog = StyledExportDialog(scene, self)

    def apply_font_profile(self, axis_label_size: str, tick_font_size: int, legend_text_size: str) -> None:
        plot_item = self.plot.getPlotItem()
        for name in ("left", "bottom"):
            axis = plot_item.getAxis(name)
            label_style = dict(getattr(axis, "labelStyle", {}))
            label_style["font-size"] = axis_label_size
            axis.setLabel(
                axis.labelText,
                units=axis.labelUnits,
                unitPrefix=axis.labelUnitPrefix,
                **label_style,
            )
            tick_font = QtGui.QFont(axis.style.get("tickFont") or QtGui.QFont())
            tick_font.setPointSize(tick_font_size)
            axis.setTickFont(tick_font)

        if plot_item.legend is not None:
            plot_item.legend.setLabelTextSize(legend_text_size)
        self._refresh_plot_layout()

    def _refresh_plot_layout(self) -> None:
        """Recalculate axis geometry before rendering after a font change."""

        plot_item = self.plot.getPlotItem()
        for name in ("left", "bottom"):
            plot_item.getAxis(name).updateGeometry()
        plot_item.layout.invalidate()
        plot_item.layout.activate()
        plot_item.updateGeometry()
        self.plot.scene().update()
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.processEvents(QtCore.QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents)

    @contextmanager
    def export_font_profile(self):
        plot_item = self.plot.getPlotItem()
        axes_state = {}
        for name in ("left", "bottom"):
            axis = plot_item.getAxis(name)
            tick_font = axis.style.get("tickFont")
            axes_state[name] = {
                "label_style": dict(getattr(axis, "labelStyle", {})),
                "tick_font": QtGui.QFont(tick_font) if tick_font is not None else None,
                "size": axis.width() if name == "left" else axis.height(),
            }
        legend_size = plot_item.legend.labelTextSize() if plot_item.legend is not None else None

        self.apply_font_profile(
            axis_label_size=EXPORT_AXIS_LABEL_SIZE,
            tick_font_size=EXPORT_TICK_FONT_SIZE,
            legend_text_size=EXPORT_LEGEND_TEXT_SIZE,
        )
        # AxisItem reserves only 80% of the rich-text label bounding box by
        # default. That is acceptable for small labels but clips large export
        # labels, so reserve the missing space explicitly while rendering.
        bottom_axis = plot_item.getAxis("bottom")
        bottom_axis.setHeight(bottom_axis.height() + max(6.0, bottom_axis.label.boundingRect().height() * 0.25))
        self._refresh_plot_layout()
        try:
            yield
        finally:
            for name, state in axes_state.items():
                axis = plot_item.getAxis(name)
                axis.setLabel(
                    axis.labelText,
                    units=axis.labelUnits,
                    unitPrefix=axis.labelUnitPrefix,
                    **state["label_style"],
                )
                axis.setTickFont(state["tick_font"])
                if name == "left":
                    axis.setWidth(state["size"])
                else:
                    axis.setHeight(state["size"])
            if plot_item.legend is not None and legend_size is not None:
                plot_item.legend.setLabelTextSize(legend_size)
            self._refresh_plot_layout()

    def export_visual(self, exporter, export_callable, *args, **kwargs):
        """Render a visual exporter with large fonts and a safe bottom margin."""

        original_source_rect = exporter.getSourceRect

        def padded_source_rect():
            source_rect = QtCore.QRectF(original_source_rect())
            source_rect.setBottom(source_rect.bottom() + EXPORT_SOURCE_BOTTOM_PADDING)
            return source_rect

        exporter.getSourceRect = padded_source_rect
        try:
            with self.export_font_profile():
                return export_callable(*args, **kwargs)
        finally:
            exporter.getSourceRect = original_source_rect

    def apply_theme(self, dark: bool):
        bg = "#121212" if dark else "#FFFFFF"
        fg = "#E0E0E0" if dark else "#1F1F1F"
        self.plot.setBackground(bg)
        plot_item = self.plot.getPlotItem()
        for name in ("left", "bottom", "right", "top"):
            axis = plot_item.getAxis(name)
            if axis is not None:
                axis.setPen(fg)
                with contextlib.suppress(Exception):
                    axis.setTextPen(fg)

        # slightly adjust grid visibility
        with contextlib.suppress(Exception):
            plot_item.showGrid(x=True, y=True, alpha=0.3 if dark else 0.25)

    def plot_current_data(self):
        from domain.constants import DataTableColumns  # avoid cycles

        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN_SQRT)
        try:
            diameter_list, rn_sqrt_list = drop_nans(diameter_list, rn_sqrt_list)
            if diameter_list.size == 0:
                return False
            diameter_list, rn_sqrt_list = np.array(
                sorted(np.array([diameter_list, rn_sqrt_list]).T, key=lambda x: x[0]), dtype=float
            ).T
            diameter_list = diameter_list.tolist()
            rn_sqrt_list = rn_sqrt_list.tolist()
        except (ListsNotSameLength, ValueError):
            return False
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        slope = self.param_table.get_column_value(0, ParamTableColumns.SLOPE)
        intercept = self.param_table.get_column_value(0, ParamTableColumns.INTERCEPT)

        plotItem = self.plot.getPlotItem()
        items_data = [item for item in plotItem.items if item.name() == "Data"]
        items_fit = [item for item in plotItem.items if item.name() == "Fit"]

        if len(items_data):
            items_data[0].setData(diameter_list, rn_sqrt_list)
        else:
            self.plot.plot(diameter_list, rn_sqrt_list, name="Data", symbolSize=6, symbolBrush="#000000")

        if drift is None or slope is None or intercept is None:
            return True

        fit_x = list(diameter_list)
        if len(fit_x) == 0:
            return True
        if np.min(fit_x) > drift:
            fit_x.insert(0, drift)
        if np.max(fit_x) < drift:
            fit_x.append(drift)
        y_appr = np.vectorize(lambda x: linear(x, slope, intercept))(fit_x)

        if len(items_fit):
            items_fit[0].setData(fit_x, y_appr)
            return None
        else:
            pen2 = pg.mkPen(color="#000000", width=3)
            self.plot.plot(
                fit_x,
                y_appr,
                name="Fit",
                pen=pen2,
                symbolSize=0,
                symbolBrush=pen2.color(),
            )
            return None

    def plot_cell(self, cell: int, repo):
        item = repo.get(cell=cell)
        if not item:
            return False
        try:
            diameter, rn_sqrt = drop_nans(item.diameter_list, item.rn_sqrt_list)
        except Exception:
            diameter, rn_sqrt = np.array([], dtype=float), np.array([], dtype=float)
        # Fallback: rebuild series from initial_data selection if stored lists are empty
        if diameter.size == 0:
            try:
                selected = {
                    v.row for v in item.initial_data.filter(col=DataTableColumns.SELECT.index) if v.value == "True"
                }
                diam_by_row = {v.row: v.value for v in item.initial_data.filter(col=DataTableColumns.DIAMETER.index)}
                rn_sqrt_by_row = {v.row: v.value for v in item.initial_data.filter(col=DataTableColumns.RN_SQRT.index)}
                diam_list = [float(diam_by_row[r]) if diam_by_row.get(r) not in ("", None) else None for r in selected]
                rn_list = [
                    float(rn_sqrt_by_row[r]) if rn_sqrt_by_row.get(r) not in ("", None) else None for r in selected
                ]
                diameter, rn_sqrt = drop_nans(diam_list, rn_list)
            except Exception:
                return False
        if diameter.size == 0:
            return False

        # Sort by diameter for better visuals
        try:
            arr = np.array([diameter, rn_sqrt]).T
            arr = arr[arr[:, 0].argsort()]
            diameter_sorted = arr[:, 0].tolist()
            rn_sqrt_sorted = arr[:, 1].tolist()
        except Exception:
            diameter_sorted = diameter.tolist()
            rn_sqrt_sorted = rn_sqrt.tolist()

        # Prepare fit x range to include drift
        fit_x = list(diameter_sorted)
        can_plot_fit = False
        y_appr = []
        try:
            slope = float(item.slope)
            intercept = float(item.intercept)
            drift = float(item.drift)
            can_plot_fit = np.isfinite(slope) and np.isfinite(intercept) and np.isfinite(drift) and len(fit_x) > 0
        except Exception:
            can_plot_fit = False
        if can_plot_fit:
            with np.errstate(all="ignore"):
                if np.min(fit_x) > drift:
                    fit_x.insert(0, drift)
                if np.max(fit_x) < drift:
                    fit_x.append(drift)
                y_appr = np.vectorize(lambda x: linear(x, slope, intercept))(fit_x)

        # Color by cell number
        color = PLOT_COLORS[(cell - 1) % len(PLOT_COLORS)]
        pen = pg.mkPen(color=color, width=3)

        # Remove previous items for this cell (data and fit) if exist
        plotItem = self.plot.getPlotItem()
        to_remove = [
            it
            for it in plotItem.items
            if getattr(it, "_rns_cell", None) == cell or it.name() in self._cell_plot_names(item)
        ]
        for it in to_remove:
            plotItem.removeItem(it)

        # Plot scatter points for data (same color, no legend clutter)
        data_item = self.plot.plot(
            diameter_sorted,
            rn_sqrt_sorted,
            name=f"{item.name} (data)",
            pen=None,
            symbol="o",
            symbolSize=6,
            symbolBrush=color,
            symbolPen=pen,
        )
        self._mark_cell_plot_item(data_item, cell)

        # Plot fit line for this cell
        if can_plot_fit:
            fit_item = self.plot.plot(
                fit_x,
                y_appr,
                name=f"{item.name}",
                pen=pen,
                symbol=None,
            )
            self._mark_cell_plot_item(fit_item, cell)
        return True

    def remove_cell_plot(self, cell: int, store):
        cell_data = store.get(cell=cell)
        plotItem = self.plot.getPlotItem()
        target_names = set()
        if cell_data:
            target_names = self._cell_plot_names(cell_data)
        items_to_remove = [
            item for item in plotItem.items if getattr(item, "_rns_cell", None) == cell or item.name() in target_names
        ]
        for item in items_to_remove:
            plotItem.removeItem(item)
