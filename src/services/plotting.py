import numpy as np
import pyqtgraph as pg

from src.constants import PLOT_COLORS, ParamTableColumns
from src.errors import ListsNotSameLength
from src.utils import drop_nans, linear


class PlotService:
    def __init__(self, plot_widget, data_table, param_table) -> None:
        self.plot = plot_widget
        self.data_table = data_table
        self.param_table = param_table

    def prepare_plot(self):
        y_label = "Rn^-0.5"
        x_label = "Диаметр ACAD (μm)"
        self.plot.setBackground("w")
        styles = {"color": "#413C58", "font-size": "15px"}
        self.plot.setLabel("left", y_label, **styles)
        self.plot.setLabel("bottom", x_label, **styles)
        self.plot.addLegend()
        self.plot.showGrid(x=True, y=True)

    def plot_current_data(self):
        from src.constants import DataTableColumns  # local import to avoid cycles

        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN_SQRT)
        try:
            diameter_list, rn_sqrt_list = drop_nans(diameter_list, rn_sqrt_list)
            # If no valid points, stop updating curves
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

        # Skip fit if parameters are missing
        if drift is None or slope is None or intercept is None:
            return True

        if np.min(diameter_list) > drift:
            diameter_list.insert(0, drift)
        if np.max(diameter_list) < drift:
            diameter_list.append(drift)
        y_appr = np.vectorize(lambda x: linear(x, slope, intercept))(diameter_list)

        if items_fit:
            items_fit[0].setData(diameter_list, y_appr)
            return None
        else:
            pen2 = pg.mkPen(color="#000000", width=3)
            self.plot.plot(
                diameter_list,
                y_appr,
                name="Fit",
                pen=pen2,
                symbolSize=0,
                symbolBrush=pen2.color(),
            )
            return None

    def plot_cell(self, cell: int, store):
        item = store.data.get(cell=cell)
        if not item:
            return
        diameter, rn_sqrt = drop_nans(item.diameter_list, item.rn_sqrt_list)
        if diameter.size == 0:
            return
        diameter_list = diameter.tolist()
        if np.min(diameter_list) > item.drift:
            diameter_list.insert(0, item.drift)
        if np.max(diameter_list) < item.drift:
            diameter_list.append(item.drift)
        y_appr = np.vectorize(lambda x: linear(x, item.slope, item.intercept))(diameter_list)
        color = PLOT_COLORS[cell - 1]
        pen2 = pg.mkPen(color=color, width=2)
        self.plot.plot(
            diameter_list,
            y_appr,
            name=f"{item.name}",
            pen=pen2,
            symbolSize=0,
            symbolBrush=pen2.color(),
        )

    def remove_cell_plot(self, cell: int, store):
        cell_data = store.data.get(cell=cell)
        plotItem = self.plot.getPlotItem()
        items_to_remove = [item for item in plotItem.items if item.name() == (cell_data.name if cell_data else None)]
        for item in items_to_remove:
            plotItem.removeItem(item)
