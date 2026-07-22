import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell

from domain.constants import DataTableColumns, ParamTableColumns
from domain.models import InitialDataItem, InitialDataItemList
from domain.utils import calculate_rn_sqrt, inverse_diameter_linear_fit
from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.xlsx_io import (
    SAMPLE_SIZE_INPUT_MODE_HEADER,
    WEIGHT_HEADER,
    XLSX_CALCULATED_FILL_COLOR,
    XLSX_HEADER_CALCULATED_COLOR,
    XLSX_HEADER_INPUT_COLOR,
    XLSX_HEADER_METADATA_COLOR,
    XLSX_HEADER_RESULTS_COLOR,
    XLSX_HEADER_WEIGHT_COLOR,
    XLSX_IDENTIFIER_FILL_COLOR,
    XLSX_INPUT_FILL_COLOR,
    XLSX_RESULTS_FILL_COLOR,
    XLSX_WEIGHT_FILL_COLOR,
    XlsxCellIO,
)


class XlsxWeightedFitTests(unittest.TestCase):
    @staticmethod
    def _fill_rgb(cell: Cell) -> str:
        return cell.fill.fgColor.rgb[-6:]

    @staticmethod
    def _repo_with_cell() -> InMemoryCellRepository:
        repo = InMemoryCellRepository()
        rows = [
            (1.0, 8.0, True),
            (2.0, 4.0, True),
            (4.0, 2.0, True),
            (8.0, 1.0, True),
            (0.0, 10.0, True),
            (16.0, None, True),
        ]
        initial_data = InitialDataItemList()
        for row, (diameter, resistance, selected) in enumerate(rows):
            values = {
                DataTableColumns.NUMBER: row + 1,
                DataTableColumns.NAME: f"Sample {row + 1}",
                DataTableColumns.SELECT: selected,
                DataTableColumns.DIAMETER: diameter,
                DataTableColumns.RESISTANCE: resistance,
            }
            initial_data.extend(
                InitialDataItem(value=value, row=row, col=column.index) for column, value in values.items()
            )
        repo.update_or_create_item(
            cell=1,
            name="Weighted test",
            diameter_list=[row[0] for row in rows[:4]],
            rn_sqrt_list=[1 / (row[1] ** 0.5) for row in rows[:4]],
            slope=0.1,
            intercept=0.2,
            drift=-2.0,
            rns=10.0,
            drift_error=0.0,
            rns_error=0.0,
            initial_data=initial_data,
        )
        return repo

    def test_export_contains_weighted_formulas_and_fit_series(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            output = Path(tmp_dir) / "weighted.xlsx"
            XlsxCellIO().save(str(output), [("Cell №1", "Weighted test", "")], self._repo_with_cell())
            workbook = openpyxl.load_workbook(output, data_only=False)
            sheet = next(ws for ws in workbook.worksheets if ws.title != "Cells data")

            headers = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
            weight_col = headers[WEIGHT_HEADER]
            assert weight_col == headers[DataTableColumns.RN_SQRT.slug] + 1
            assert sheet.freeze_panes == "A2"
            weight_letter = sheet.cell(row=1, column=weight_col).column_letter
            assert sheet.auto_filter.ref == f"A1:{weight_letter}7"

            def fill_for(header: str, row: int = 1) -> str:
                cell = sheet.cell(row=row, column=headers[header])
                return self._fill_rgb(cell)

            input_header = fill_for(DataTableColumns.DIAMETER.slug)
            calculated_header = fill_for(DataTableColumns.RN_SQRT.slug)
            weight_header = self._fill_rgb(sheet.cell(row=1, column=weight_col))
            results_header = fill_for(ParamTableColumns.SLOPE.name)
            metadata_header = fill_for(SAMPLE_SIZE_INPUT_MODE_HEADER)
            identifier_body = fill_for(DataTableColumns.NUMBER.slug, row=2)
            input_body = fill_for(DataTableColumns.DIAMETER.slug, row=2)
            calculated_body = fill_for(DataTableColumns.RN_SQRT.slug, row=2)
            weight_body = self._fill_rgb(sheet.cell(row=2, column=weight_col))
            results_body = fill_for(ParamTableColumns.SLOPE.name, row=2)

            assert input_header == XLSX_HEADER_INPUT_COLOR
            assert calculated_header == XLSX_HEADER_CALCULATED_COLOR
            assert weight_header == XLSX_HEADER_WEIGHT_COLOR
            assert results_header == XLSX_HEADER_RESULTS_COLOR
            assert metadata_header == XLSX_HEADER_METADATA_COLOR
            assert identifier_body == XLSX_IDENTIFIER_FILL_COLOR
            assert input_body == XLSX_INPUT_FILL_COLOR
            assert calculated_body == XLSX_CALCULATED_FILL_COLOR
            assert weight_body == XLSX_WEIGHT_FILL_COLOR
            assert results_body == XLSX_RESULTS_FILL_COLOR

            cells_sheet = workbook["Cells data"]
            assert self._fill_rgb(cells_sheet["A1"]) == XLSX_HEADER_INPUT_COLOR
            assert self._fill_rgb(cells_sheet["A3"]) == XLSX_RESULTS_FILL_COLOR
            assert "1/D2" in sheet.cell(row=2, column=weight_col).value
            assert "ISNUMBER(J2)" in sheet.cell(row=2, column=weight_col).value
            assert "D6>0" in sheet.cell(row=6, column=weight_col).value
            assert "ISNUMBER(J7)" in sheet.cell(row=7, column=weight_col).value

            slope_formula = sheet.cell(row=2, column=headers[ParamTableColumns.SLOPE.name]).value
            intercept_formula = sheet.cell(row=2, column=headers[ParamTableColumns.INTERCEPT.name]).value
            wls_helper_cols = {
                sheet.cell(row=1, column=column).value: column
                for column in range(1, sheet.max_column + 1)
                if str(sheet.cell(row=1, column=column).value).startswith("_WLS ")
            }
            assert set(wls_helper_cols) == {
                "_WLS sum w",
                "_WLS sum wD",
                "_WLS sum wY",
                "_WLS sum wD2",
                "_WLS sum wDY",
            }
            diameter_letter = sheet.cell(row=1, column=headers[DataTableColumns.DIAMETER.slug]).column_letter
            rn_sqrt_letter = sheet.cell(row=1, column=headers[DataTableColumns.RN_SQRT.slug]).column_letter
            weight_range = f"{weight_letter}2:{weight_letter}7"
            diameter_range = f"{diameter_letter}2:{diameter_letter}7"
            rn_sqrt_range = f"{rn_sqrt_letter}2:{rn_sqrt_letter}7"
            expected_helper_formulas = {
                "_WLS sum w": f"=SUMPRODUCT({weight_range},{weight_range})",
                "_WLS sum wD": (f"=SUMPRODUCT({weight_range},{weight_range},{diameter_range})"),
                "_WLS sum wY": (f"=SUMPRODUCT({weight_range},{weight_range},{rn_sqrt_range})"),
                "_WLS sum wD2": (f"=SUMPRODUCT({weight_range},{weight_range},{diameter_range},{diameter_range})"),
                "_WLS sum wDY": (f"=SUMPRODUCT({weight_range},{weight_range},{diameter_range},{rn_sqrt_range})"),
            }
            for header, formula in expected_helper_formulas.items():
                assert sheet.cell(row=2, column=wls_helper_cols[header]).value == formula
                assert "IFERROR" not in formula
            assert sheet.cell(row=2, column=wls_helper_cols["_WLS sum w"]).coordinate in slope_formula
            assert sheet.cell(row=2, column=wls_helper_cols["_WLS sum wDY"]).coordinate in slope_formula
            assert "SLOPE(" not in slope_formula
            assert sheet.cell(row=2, column=wls_helper_cols["_WLS sum wY"]).coordinate in intercept_formula

            assert len(sheet._charts) == 1
            assert len(sheet._charts[0].series) == 2
            assert sheet._charts[0].visible_cells_only is False
            assert sheet._charts[0].scatterStyle == "lineMarker"
            assert sheet._charts[0].series[0].trendline is None
            fit_series = sheet._charts[0].series[1]
            cached_x = [point.v for point in fit_series.xVal.numRef.numCache.pt]
            cached_y = [point.v for point in fit_series.yVal.numRef.numCache.pt]
            assert cached_x == [-2.0, 8.0]
            assert cached_y == [0.0, 1.0]
            hidden_headers = {
                sheet.cell(row=1, column=column).value
                for column in range(1, sheet.max_column + 1)
                if sheet.column_dimensions[sheet.cell(row=1, column=column).column_letter].hidden
            }
            assert hidden_headers == {
                "_Weighted fit X",
                "_Weighted fit Y",
                "_WLS sum w",
                "_WLS sum wD",
                "_WLS sum wY",
                "_WLS sum wD2",
                "_WLS sum wDY",
            }

            loaded_items, load_errors = XlsxCellIO().load(str(output))
            assert load_errors == []
            expected_slope, expected_intercept = inverse_diameter_linear_fit(
                [1.0, 2.0, 4.0, 8.0],
                [calculate_rn_sqrt(resistance, 0.0) for resistance in (8.0, 4.0, 2.0, 1.0)],
            )
            self.assertAlmostEqual(loaded_items[0]["slope"], expected_slope, places=12)
            self.assertAlmostEqual(loaded_items[0]["intercept"], expected_intercept, places=12)


if __name__ == "__main__":
    unittest.main()
