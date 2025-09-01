import numpy as np
from PySide6 import QtWidgets

from src.constants import BLACK, RNS_ERROR_COLOR, WHITE, DataTableColumns, ParamTableColumns
from src.errors import ListsNotSameLength
from src.utils import (
    calculate_drift,
    calculate_drift_per_sample,
    calculate_rn_sqrt,
    calculate_rns,
    calculate_rns_error_diff,
    calculate_rns_error_per_sample,
    calculate_rns_per_sample,
    calculate_square,
    drop_nans,
    linear_fit,
)
from src.widgets import TableWidgetItem


class CalculationService:
    """Encapsulates calculation flows to keep the widget lean."""

    def __init__(self, data_table, param_table, rn_consistent_widget, allowed_error_widget) -> None:
        self.data_table = data_table
        self.param_table = param_table
        self.rn_consistent_widget = rn_consistent_widget
        self.allowed_error_widget = allowed_error_widget

    # High-level orchestration
    def calculate_results(self):
        self.data_table.clear_calculations()
        if not self.calculate_rn05():
            return False
        if not self.calculate_main_params():
            return False
        if not self.calculate_rns_drift_square_per_sample():
            return False
        if not self.calculate_error_params():
            return False
        return True

    # Individual steps
    def calculate_rn05(self):
        """Compute Rn^-0.5 for each selected row."""
        for row in range(self.data_table.rowCount()):
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            if not resistance or not diameter:
                continue
            rn_sqrt = calculate_rn_sqrt(resistance=resistance, rn_consistent=self.rn_consistent_widget.value())
            self.data_table.setItem(row, DataTableColumns.RN_SQRT.index, TableWidgetItem(str(rn_sqrt)))
        return True

    def calculate_main_params(self):
        """Compute slope/intercept, drift and RnS for current selection."""
        diameter_list = self.data_table.get_column_values(DataTableColumns.DIAMETER)
        rn_sqrt_list = self.data_table.get_column_values(DataTableColumns.RN_SQRT)
        try:
            diameter_list, rn_sqrt_list = drop_nans(diameter_list, rn_sqrt_list)
        except (ListsNotSameLength, ValueError):
            QtWidgets.QMessageBox.warning(
                None,
                "Не корректные данные!",
                "В таблице данных не корректные значения для 'Диаметр ACAD' и 'Rn^-0.5'",
            )
            return False

        slope, intercept = linear_fit(diameter_list, rn_sqrt_list)

        self.param_table.setItem(0, ParamTableColumns.SLOPE.index, TableWidgetItem(str(slope)))
        self.param_table.setItem(0, ParamTableColumns.INTERCEPT.index, TableWidgetItem(str(intercept)))

        drift = calculate_drift(slope=slope, intercept=intercept)
        self.param_table.setItem(0, ParamTableColumns.DRIFT.index, TableWidgetItem(str(drift)))

        rns = calculate_rns(slope)
        self.param_table.setItem(0, ParamTableColumns.RNS.index, TableWidgetItem(str(rns)))
        return True

    def calculate_error_params(self):
        """Compute RnS/Drift error metrics and color rows by allowed error."""
        rns = self.param_table.get_column_value(0, ParamTableColumns.RNS)
        rns_list = np.array([v for v in self.data_table.get_column_values(DataTableColumns.RNS) if v], dtype=float)

        try:
            rns_error = np.sqrt(np.sum((rns_list - rns) ** 2) / len(rns_list))
        except (ZeroDivisionError,):
            QtWidgets.QMessageBox.warning(
                None,
                "Ошибка в рассчете RNS_ERROR",
                "Ошибка деления на ноль при расчете RNS_ERROR!",
            )
            return False

        self.param_table.setItem(0, ParamTableColumns.RNS_ERROR.index, TableWidgetItem(str(rns_error)))

        for row in range(self.data_table.rowCount()):
            self.data_table.color_row(row=row, background_color=WHITE, text_color=BLACK)
            rns_value = self.data_table.get_column_value(row, DataTableColumns.RNS)
            if not rns_value:
                continue
            value = calculate_rns_error_per_sample(rns_i=rns_value, rns=rns)
            self.data_table.setItem(row, DataTableColumns.RNS_ERROR.index, TableWidgetItem(str(value)))
            error_diff = calculate_rns_error_diff(
                rns_error_per_sample=value,
                rns_error=rns_error,
                allowed_error=self.allowed_error_widget.value(),
            )
            if error_diff > 0:
                self.data_table.color_row(row=row, background_color=RNS_ERROR_COLOR, text_color=WHITE)
            else:
                self.data_table.color_row(row=row, background_color=WHITE, text_color=BLACK)

        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        drift_list = np.array([v for v in self.data_table.get_column_values(DataTableColumns.DRIFT) if v], dtype=float)
        drift_error = np.sqrt(np.sum((drift_list - drift) ** 2) / len(drift_list))
        self.param_table.setItem(0, ParamTableColumns.DRIFT_ERROR.index, TableWidgetItem(str(drift_error)))

        self.param_table.setItem(
            0, ParamTableColumns.RN_CONSISTENT.index, TableWidgetItem(str(self.rn_consistent_widget.value()))
        )
        self.param_table.setItem(
            0, ParamTableColumns.ALLOWED_ERROR.index, TableWidgetItem(str(self.allowed_error_widget.value()))
        )
        return True

    def calculate_rns_drift_square_per_sample(self):
        """Per-row RnS/Drift/Square using computed mean values."""
        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        if not drift:
            QtWidgets.QMessageBox.warning(None, "Не корректные данные!", "Не вычислен уход!")
            return False

        rns_mean = self.param_table.get_column_value(0, ParamTableColumns.RNS)
        if not rns_mean:
            QtWidgets.QMessageBox.warning(None, "Не корректные данные!", "Не вычислен RnS!")
            return False

        for row in range(self.data_table.rowCount()):
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            if not resistance or not diameter:
                continue

            square_value = calculate_square(diameter=diameter, drift=drift)
            self.data_table.setItem(row, DataTableColumns.SQUARE.index, TableWidgetItem(str(square_value)))

            rns_value = calculate_rns_per_sample(
                resistance=resistance,
                diameter=diameter,
                drift=drift,
                rn_persistent=self.rn_consistent_widget.value(),
            )
            self.data_table.setItem(row, DataTableColumns.RNS.index, TableWidgetItem(str(rns_value)))

            drift_value = calculate_drift_per_sample(
                diameter=diameter,
                resistance=resistance,
                rns=rns_mean,
                rn_persistent=self.rn_consistent_widget.value(),
            )
            self.data_table.setItem(row, DataTableColumns.DRIFT.index, TableWidgetItem(str(drift_value)))
        return True
