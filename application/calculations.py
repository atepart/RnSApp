import numpy as np
from PySide6 import QtWidgets

from domain.constants import BLACK, RNS_ERROR_COLOR, WHITE, DataTableColumns, ParamTableColumns
from domain.errors import ListsNotSameLength
from domain.utils import (
    calculate_drift,
    calculate_real_area,
    calculate_rn_sqrt,
    calculate_rns,
    calculate_rns_error_per_sample,
    calculate_rns_per_sample,
    calculate_square,
    drop_nans,
    linear_fit,
)
from ui.widgets import TableWidgetItem


class CalculationService:
    def __init__(
        self,
        data_table,
        param_table,
        rn_consistent_widget,
        allowed_error_widget,
        s_custom1_widget=None,
        s_custom2_widget=None,
        s_custom3_widget=None,
    ) -> None:
        self.data_table = data_table
        self.param_table = param_table
        self.rn_consistent_widget = rn_consistent_widget
        self.allowed_error_widget = allowed_error_widget
        self.s_custom1_widget = s_custom1_widget
        self.s_custom2_widget = s_custom2_widget
        self.s_custom3_widget = s_custom3_widget

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

    def calculate_rn05(self):
        for row in range(self.data_table.rowCount()):
            resistance = self.data_table.get_column_value(row, DataTableColumns.RESISTANCE)
            diameter = self.data_table.get_column_value(row, DataTableColumns.DIAMETER)
            if not resistance or not diameter:
                continue
            rn_sqrt = calculate_rn_sqrt(resistance=resistance, rn_consistent=self.rn_consistent_widget.value())
            self.data_table.setItem(row, DataTableColumns.RN_SQRT.index, TableWidgetItem(str(rn_sqrt)))
        return True

    def calculate_main_params(self):
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

        # Persist the user-entered nominal areas in hidden columns
        try:
            if self.s_custom1_widget is not None:
                s_nom1 = float(self.s_custom1_widget.value())
                self.param_table.setItem(0, ParamTableColumns.S_CUSTOM1.index, TableWidgetItem(str(s_nom1)))
        except Exception:
            pass
        try:
            if self.s_custom2_widget is not None:
                s_nom2 = float(self.s_custom2_widget.value())
                self.param_table.setItem(0, ParamTableColumns.S_CUSTOM2.index, TableWidgetItem(str(s_nom2)))
        except Exception:
            pass
        try:
            if self.s_custom3_widget is not None:
                s_nom3 = float(self.s_custom3_widget.value())
                self.param_table.setItem(0, ParamTableColumns.S_CUSTOM3.index, TableWidgetItem(str(s_nom3)))
        except Exception:
            pass

        # Real areas for nominal S1..S3 if provided
        try:
            if self.s_custom1_widget is not None:
                s_nom1 = float(self.s_custom1_widget.value())
                s_real_c1 = calculate_real_area(area_nominal=s_nom1, drift=drift)
                self.param_table.setItem(0, ParamTableColumns.S_REAL_CUSTOM1.index, TableWidgetItem(str(s_real_c1)))
        except Exception:
            pass
        try:
            if self.s_custom2_widget is not None:
                s_nom2 = float(self.s_custom2_widget.value())
                s_real_c2 = calculate_real_area(area_nominal=s_nom2, drift=drift)
                self.param_table.setItem(0, ParamTableColumns.S_REAL_CUSTOM2.index, TableWidgetItem(str(s_real_c2)))
        except Exception:
            pass
        try:
            if self.s_custom3_widget is not None:
                s_nom3 = float(self.s_custom3_widget.value())
                s_real_c3 = calculate_real_area(area_nominal=s_nom3, drift=drift)
                self.param_table.setItem(0, ParamTableColumns.S_REAL_CUSTOM3.index, TableWidgetItem(str(s_real_c3)))
        except Exception:
            pass
        return True

    def calculate_error_params(self):
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
            # Keep absolute deviation in the table column for RnS error
            value = calculate_rns_error_per_sample(rns_i=rns_value, rns=rns)
            self.data_table.setItem(row, DataTableColumns.RNS_ERROR.index, TableWidgetItem(str(value)))

            # Compare relative deviation in percent with user allowed deviation (%)
            try:
                rel_dev_percent = abs(rns_value - rns) / rns * 100
            except ZeroDivisionError:
                rel_dev_percent = float("inf")
            if rel_dev_percent > self.allowed_error_widget.value():
                self.data_table.color_row(row=row, background_color=RNS_ERROR_COLOR, text_color=WHITE)
            else:
                self.data_table.color_row(row=row, background_color=WHITE, text_color=BLACK)

        drift = self.param_table.get_column_value(0, ParamTableColumns.DRIFT)
        drift_list_raw = [v for v in self.data_table.get_column_values(DataTableColumns.DRIFT) if v not in (None, "")]
        try:
            drift_list = np.array(drift_list_raw, dtype=float)
        except Exception:
            drift_list = np.array([], dtype=float)
        if drift is None or len(drift_list) == 0:
            drift_error = 0.0
        else:
            with np.errstate(invalid="ignore"):
                drift_error = float(np.sqrt(np.sum((drift_list - drift) ** 2) / len(drift_list)))
                if np.isnan(drift_error):
                    drift_error = 0.0
        self.param_table.setItem(0, ParamTableColumns.DRIFT_ERROR.index, TableWidgetItem(str(drift_error)))

        self.param_table.setItem(
            0, ParamTableColumns.RN_CONSISTENT.index, TableWidgetItem(str(self.rn_consistent_widget.value()))
        )
        self.param_table.setItem(
            0, ParamTableColumns.ALLOWED_ERROR.index, TableWidgetItem(str(self.allowed_error_widget.value()))
        )
        return True

    def calculate_rns_drift_square_per_sample(self):
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

            # Per-sample drift is no longer calculated; drift is the same for all samples.
        return True
