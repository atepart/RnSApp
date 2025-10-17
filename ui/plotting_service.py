import contextlib

import numpy as np
import pyqtgraph as pg

from domain.constants import PLOT_COLORS, DataTableColumns, ParamTableColumns
from domain.errors import ListsNotSameLength
from domain.utils import drop_nans, linear


class PlotService:
    def __init__(self, plot_widget, data_table, param_table) -> None:
        self.plot = plot_widget
        self.data_table = data_table
        self.param_table = param_table

    def prepare_plot(self):
        y_label = "1/√Rₙ"
        x_label = "Диаметр ACAD (μm)"
        self.plot.setBackground("w")
        styles = {"color": "#413C58", "font-size": "15px"}
        self.plot.setLabel("left", y_label, **styles)
        self.plot.setLabel("bottom", x_label, **styles)
        self.plot.addLegend()
        self.plot.showGrid(x=True, y=True)

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
            return
        diameter, rn_sqrt = drop_nans(item.diameter_list, item.rn_sqrt_list)
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
                return
        if diameter.size == 0:
            return

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
        with np.errstate(all="ignore"):
            if len(fit_x):
                if np.min(fit_x) > item.drift:
                    fit_x.insert(0, item.drift)
                if np.max(fit_x) < item.drift:
                    fit_x.append(item.drift)
        y_appr = np.vectorize(lambda x: linear(x, item.slope, item.intercept))(fit_x)

        # Color by cell number
        color = PLOT_COLORS[cell - 1]
        pen = pg.mkPen(color=color, width=3)

        # Remove previous items for this cell (data and fit) if exist
        plotItem = self.plot.getPlotItem()
        to_remove = [
            it for it in plotItem.items if it.name() in {f"{item.name}", f"{item.name} (fit)", f"{item.name} (data)"}
        ]
        for it in to_remove:
            plotItem.removeItem(it)

        # Plot scatter points for data (same color, no legend clutter)
        self.plot.plot(
            diameter_sorted,
            rn_sqrt_sorted,
            name=f"{item.name} (data)",
            pen=None,
            symbol="o",
            symbolSize=6,
            symbolBrush=color,
            symbolPen=pen,
        )

        # Plot fit line for this cell
        self.plot.plot(
            fit_x,
            y_appr,
            name=f"{item.name}",
            pen=pen,
            symbol=None,
        )

    def remove_cell_plot(self, cell: int, store):
        cell_data = store.get(cell=cell)
        plotItem = self.plot.getPlotItem()
        target_names = set()
        if cell_data:
            target_names = {cell_data.name, f"{cell_data.name} (fit)", f"{cell_data.name} (data)"}
        items_to_remove = [item for item in plotItem.items if item.name() in target_names]
        for item in items_to_remove:
            plotItem.removeItem(item)
