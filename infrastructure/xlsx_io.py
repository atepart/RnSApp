import contextlib
import re
from typing import Any, Dict, List, Tuple

import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from domain.constants import BLUE, RED, DataTableColumns, ParamTableColumns
from domain.models import InitialDataItem, InitialDataItemList
from domain.ports import CellDataIO, CellRepository


def _export_cells_grid(ws_cells, cell_grid_values: List[Tuple[str, str, str]]):
    """Export the 4x4 grid of cell summary to the first sheet (visual aid)."""
    init_data = list(cell_grid_values)
    output = []
    blocks = [init_data[i : i + 4] for i in range(0, len(init_data), 4)]
    for block in blocks:
        block_transposed = list(map(list, zip(*block)))
        output.extend(block_transposed)

    def _maybe_numeric(val):
        if val in (None, ""):
            return None
        if isinstance(val, (int, float)):
            try:
                return int(val) if isinstance(val, int) else round(float(val), 2)
            except Exception:
                return val
        if isinstance(val, str):
            s = val.strip()
            # Values like "Уход: 0.123" or "RnS: 55.3" -> take right part
            if ":" in s:
                s = s.split(":", 1)[1].strip()
            s = s.replace(",", ".")
            # pure number?
            if re.fullmatch(r"-?\d+(?:\.\d+)?", s):
                try:
                    return round(float(s), 2)
                except Exception:
                    return val
        return val

    for row_ind, row in enumerate(output, 1):
        for col_ind, coll in enumerate(row, 1):
            # Header row every 3rd (names) stays text; others (drift, RnS) try numeric
            is_header = (row_ind - 1) % 3 == 0
            value = coll
            if not is_header:
                num = _maybe_numeric(coll)
                value = None if num in (None, "") else num
            cell = ws_cells.cell(row=row_ind, column=col_ind, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if (row_ind - 1) % 3 == 0:
                cell.border = Border(
                    right=Side(style="thick"), left=Side(style="thick"), top=Side("thick"), bottom=Side(style="thick")
                )
                cell.font = Font(bold=True)
            elif (row_ind - 3) % 3 == 0:
                cell.border = Border(right=Side(style="thick"), left=Side(style="thick"), bottom=Side(style="thick"))
            else:
                cell.border = Border(right=Side(style="thick"), left=Side(style="thick"))

    for col in ws_cells.columns:
        column = col[0].column_letter
        ws_cells.column_dimensions[column].width = 12
    for row in ws_cells.rows:
        ws_cells.row_dimensions[row[0].row].height = 21


_SANITIZE_MAP = {
    "/": "／",  # FULLWIDTH SOLIDUS
    "\\": "＼",  # FULLWIDTH REVERSE SOLIDUS
    ":": "：",  # FULLWIDTH COLON
    "*": "∗",  # ASTERISK OPERATOR
    "?": "？",  # FULLWIDTH QUESTION MARK
    "[": "［",  # FULLWIDTH LEFT SQUARE BRACKET
    "]": "］",  # FULLWIDTH RIGHT SQUARE BRACKET
}
_DESANITIZE_MAP = {v: k for k, v in _SANITIZE_MAP.items()}


def _sanitize_title_component(text: str) -> str:
    if not isinstance(text, str):
        return str(text)
    out = []
    for ch in text:
        out.append(_SANITIZE_MAP.get(ch, ch))
    return "".join(out)


def _desanitize_title_component(text: str) -> str:
    if not isinstance(text, str):
        return str(text)
    out = []
    for ch in text:
        out.append(_DESANITIZE_MAP.get(ch, ch))
    return "".join(out)


def _compose_cell_sheet_title(cell_index: int, name: str, existing: List[str]) -> str:
    prefix = f"Cell №{cell_index} "
    safe_name = _sanitize_title_component(name)
    # Excel sheet title max length = 31
    max_name_len = max(0, 31 - len(prefix))
    base = prefix + safe_name[:max_name_len]
    title = base
    suffix = 2
    while title in existing:
        # try appending _n within 31 chars
        suf = f"_{suffix}"
        title = (base[: max(0, 31 - len(suf))] + suf) if len(base) + len(suf) > 31 else base + suf
        suffix += 1
    return title


class XlsxCellIO(CellDataIO):
    """XLSX adapter implementing combined per-cell sheet with data+results and a chart."""

    def save(self, file_name: str, cell_grid_values: List[Tuple[str, str, str]], repo: CellRepository) -> None:
        wb = openpyxl.Workbook()
        ws_cells = wb.active
        ws_cells.title = "Cells data"
        _export_cells_grid(ws_cells, cell_grid_values)

        # Data columns to export: exclude DRIFT (per request)
        export_data_columns: List[DataTableColumns] = [c for c in DataTableColumns if c is not DataTableColumns.DRIFT]

        def _coerce_value(val, dtype):
            # Convert values to proper numeric types so Excel treats them as numbers
            if val in (None, ""):
                return None
            try:
                if dtype is float:
                    if isinstance(val, str):
                        # Normalize decimal separator: comma -> dot
                        val = val.replace(",", ".")
                    return round(float(val), 2)
                if dtype is int:
                    if isinstance(val, str):
                        val = val.replace(",", ".")
                        # Allow integers provided as "1.0"
                        return int(float(val))
                    return int(val)
                if dtype is bool:
                    if isinstance(val, str):
                        return val in ("True", "TRUE", "true", "1")
                    return bool(val)
            except Exception:
                return val
            return val

        # Create one sheet per recorded cell with data table on the left and results on the right
        for cell_data in repo:
            sheet_name = _compose_cell_sheet_title(cell_data.cell, cell_data.name, wb.sheetnames)
            ws = wb.create_sheet(sheet_name)

            # Write data header (row 1) with styling and width based on header text
            for col_idx, col_def in enumerate(export_data_columns, start=1):
                hcell = ws.cell(row=1, column=col_idx, value=col_def.slug)
                hcell.font = Font(bold=True)
                hcell.border = Border(bottom=Side(style="medium"))
                with contextlib.suppress(Exception):
                    ws.column_dimensions[get_column_letter(col_idx)].width = max(len(str(col_def.slug)) + 2, 10)

            # Write data values from InitialDataItemList
            # Determine how many rows are present in initial data
            max_row_index = 0
            for it in cell_data.initial_data:
                if isinstance(it, InitialDataItem):
                    max_row_index = max(max_row_index, it.row)
                else:
                    try:
                        max_row_index = max(max_row_index, it["row"])  # type: ignore[index]
                    except Exception:
                        pass

            # Build mapping from full column enum index to exported column position
            export_col_position = {col.index: pos for pos, col in enumerate(export_data_columns, start=1)}

            # Write rows 2..(max_row_index+2)
            for it in cell_data.initial_data:
                row = it.row if isinstance(it, InitialDataItem) else it["row"]
                col = it.col if isinstance(it, InitialDataItem) else it["col"]
                val = it.value if isinstance(it, InitialDataItem) else it["value"]
                pos = export_col_position.get(col)
                if pos is not None:
                    col_def = next(c for c in export_data_columns if c.index == col)
                    coerced = _coerce_value(val, col_def.dtype)
                    ws.cell(row=row + 2, column=pos, value=coerced)

            # Results header placed to the right with a gap column
            results_start_col = len(export_data_columns) + 2
            for i, param in enumerate(ParamTableColumns, start=0):
                hcell = ws.cell(row=1, column=results_start_col + i, value=param.name)
                hcell.font = Font(bold=True)
                hcell.border = Border(bottom=Side(style="medium"))
                with contextlib.suppress(Exception):
                    ws.column_dimensions[get_column_letter(results_start_col + i)].width = max(
                        len(str(param.name)) + 2, 10
                    )

                # Value row (2)
                raw_value = getattr(cell_data, param.slug, "")
                value = _coerce_value(raw_value, param.dtype)
                ws.cell(row=2, column=results_start_col + i, value=value)

            # Build chart: Rn^-0.5 vs Диаметр ACAD (μm)
            try:
                # Find exported columns for RN_SQRT and DIAMETER
                y_col_idx = next(i for i, c in enumerate(export_data_columns, start=1) if c is DataTableColumns.RN_SQRT)
                x_col_idx = next(
                    i for i, c in enumerate(export_data_columns, start=1) if c is DataTableColumns.DIAMETER
                )

                # Determine data max row based on last non-empty in x/y columns
                def last_non_empty_row_for_col(col_idx: int) -> int:
                    last = 1
                    for r in range(2, ws.max_row + 1):
                        v = ws.cell(row=r, column=col_idx).value
                        if v not in (None, ""):
                            last = r
                    return last

                max_row = max(last_non_empty_row_for_col(x_col_idx), last_non_empty_row_for_col(y_col_idx))
                if max_row < 2:
                    max_row = 2

                chart = ScatterChart()
                chart.title = "Rn^-0.5 vs Диаметр ACAD (μm)"
                # Show only markers (points) to avoid spurious lines/series
                chart.scatterStyle = "marker"
                chart.style = 2
                chart.x_axis.title = "Диаметр ACAD (μm)"
                chart.y_axis.title = "Rn^-0.5"
                # Title/legend and axes styling
                with contextlib.suppress(Exception):
                    chart.legend.position = "r"  # right, outside plot area
                    chart.legend.overlay = False
                    # put chart title outside overlay
                    if hasattr(chart, "title") and hasattr(chart.title, "overlay"):
                        chart.title.overlay = False
                    # Axis positions and numeric display
                    chart.x_axis.axPos = "b"
                    chart.y_axis.axPos = "l"
                    chart.x_axis.delete = False
                    chart.y_axis.delete = False
                    chart.x_axis.number_format = "General"
                    chart.y_axis.number_format = "General"
                    chart.x_axis.majorTickMark = "out"
                    chart.y_axis.majorTickMark = "out"
                    # Ensure tick labels appear next to axes and not skipped
                    chart.x_axis.tickLblPos = "nextTo"
                    chart.y_axis.tickLblPos = "nextTo"
                    with contextlib.suppress(Exception):
                        chart.x_axis.tickLblSkip = 1
                        chart.x_axis.tickMarkSkip = 1
                        chart.y_axis.tickLblSkip = 1
                        chart.y_axis.tickMarkSkip = 1
                    if getattr(chart.x_axis, "title", None):
                        chart.x_axis.title.overlay = False
                    if getattr(chart.y_axis, "title", None):
                        chart.y_axis.title.overlay = False
                # Gray grid and show tick labels next to axes
                with contextlib.suppress(Exception):
                    grid_color = "B3B3B3"  # light gray
                    for axis in (chart.x_axis, chart.y_axis):
                        axis.tickLblPos = "nextTo"
                        axis.majorGridlines = ChartLines()
                        axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=grid_color))

                xvalues = Reference(ws, min_col=x_col_idx, min_row=2, max_row=max_row)
                yvalues = Reference(ws, min_col=y_col_idx, min_row=2, max_row=max_row)
                series = Series(yvalues, xvalues, title="Data")
                # Uniform marker style & color; hide connecting line
                series.marker = Marker(symbol="circle")
                series.marker.size = 3
                with contextlib.suppress(Exception):
                    blue = BLUE.lstrip("#").upper()
                    series.marker.graphicalProperties.solidFill = blue
                    series.marker.graphicalProperties.line.solidFill = blue
                    if getattr(series.graphicalProperties, "line", None) is not None:
                        series.graphicalProperties.line.noFill = True
                # Trendline: show equation only (no R^2); paint it RED
                with contextlib.suppress(Exception):
                    series.trendline = Trendline(trendlineType="linear", dispEq=True, dispRSqr=False)
                    red = RED.lstrip("#").upper()
                    try:
                        series.trendline.graphicalProperties.line.solidFill = red
                    except Exception:
                        # Fallback property name in some versions
                        series.trendline.spPr = GraphicalProperties(ln=LineProperties(solidFill=red))

                chart.varyColors = False
                chart.series.append(series)

                # Ensure axis range includes x-intercept (drift); extend trendline via forecast to reach it
                with contextlib.suppress(Exception):
                    drift = float(getattr(cell_data, "drift", 0.0))
                    x_vals: List[float] = []
                    for r in range(2, max_row + 1):
                        xv = ws.cell(row=r, column=x_col_idx).value
                        if xv not in (None, ""):
                            x_vals.append(float(xv))
                    if x_vals:
                        chart.x_axis.scaling.min = min(min(x_vals), drift)
                        chart.x_axis.scaling.max = max(max(x_vals), drift)
                        # Forecast distances measured along X units
                        try:
                            backward = max(0.0, min(x_vals) - drift)
                            forward = max(0.0, drift - max(x_vals))
                            if hasattr(series.trendline, "backward"):
                                series.trendline.backward = round(backward, 2)
                            if hasattr(series.trendline, "forward"):
                                series.trendline.forward = round(forward, 2)
                        except Exception:
                            pass

                # Place chart under the results table
                chart_anchor_row = 4
                chart_anchor_col = results_start_col
                ws.add_chart(chart, ws.cell(row=chart_anchor_row, column=chart_anchor_col).coordinate)
            except Exception:
                # Chart creation should not break export
                pass

        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        wb.save(filename=file_name)

    def load(self, file_name: str) -> Tuple[List[Dict[str, Any]], List[str]]:
        wb = openpyxl.load_workbook(file_name)

        # Combined-sheet format: sheets starting with "Cell №"
        combined_sheet_names = [sh for sh in wb.sheetnames if sh.startswith("Cell №")]
        if combined_sheet_names:
            return self._load_combined(wb, combined_sheet_names)

        # Fallback to legacy format (separate Data/Results sheets)
        from .persistence_xlsx import load_cells_from_xlsx as _legacy_load  # lazy import to avoid cycles

        return _legacy_load(file_name)

    def _load_combined(self, wb, sheet_names: List[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
        items: List[Dict[str, Any]] = []
        errors: List[str] = []

        data_slugs = DataTableColumns.get_all_slugs()
        # We'll accept that DRIFT may be absent in combined export

        for sheet_name in sheet_names:
            try:
                i = int(re.findall(r"Cell №(\d+) ", sheet_name)[0])
                cell_name = re.findall(r"Cell №\d+ (.*)", sheet_name)[0]
                # de-sanitize sheet component back to original visible ASCII if needed
                cell_name = _desanitize_title_component(cell_name)
            except (IndexError, ValueError):
                errors.append(f"Неверное имя листа: {sheet_name}")
                continue

            ws = wb[sheet_name]

            # Map header titles to all column indexes (1-based)
            header_positions: Dict[str, List[int]] = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=c).value
                if isinstance(val, str) and val.strip():
                    key = val.strip()
                    header_positions.setdefault(key, []).append(c)

            # Build initial data for all DataTableColumns
            initial_data = InitialDataItemList()

            # Determine how many data rows based on presence of any data in known data columns
            def last_non_empty_row_for_col(col_idx: int) -> int:
                last = 1
                for r in range(2, ws.max_row + 1):
                    v = ws.cell(row=r, column=col_idx).value
                    if v not in (None, ""):
                        last = r
                return last

            def first_col_for(name: str) -> int | None:
                pos = header_positions.get(name)
                return pos[0] if pos else None

            def last_col_for(name: str) -> int | None:
                pos = header_positions.get(name)
                return pos[-1] if pos else None

            candidate_cols = [first_col_for(slug) for slug in data_slugs if first_col_for(slug)]
            max_data_row = 2
            for c in candidate_cols:
                max_data_row = max(max_data_row, last_non_empty_row_for_col(c))

            # Now iterate each DataTableColumns member and collect values if present
            for data_col in DataTableColumns:
                col_idx = first_col_for(data_col.slug)
                for row in range(2, max_data_row + 1):
                    try:
                        raw = ws.cell(row=row, column=col_idx).value if col_idx is not None else ""
                        value = "" if raw in (None, "None", "") else raw
                        # Normalize SELECT boolean as string 'True' or ''
                        if data_col is DataTableColumns.SELECT:
                            if raw in (True, "TRUE", "True", "true", 1, "1"):
                                value = "True"
                            else:
                                value = ""
                        initial_data.append(InitialDataItem(row=row - 2, col=data_col.index, value=value))
                    except Exception:
                        initial_data.append(InitialDataItem(row=row - 2, col=data_col.index, value=""))

            # Build diameter_list and rn_sqrt_list for convenience (floats or None)
            def list_from_initial(col: DataTableColumns):
                vals = [v.value for v in initial_data.filter(col=col.index)]
                out: List[Any] = []
                for v in vals:
                    try:
                        out.append(float(v))
                    except Exception:
                        out.append(None)
                return out

            diameter_list = list_from_initial(DataTableColumns.DIAMETER)
            rn_sqrt_list = list_from_initial(DataTableColumns.RN_SQRT)

            # Results: read from headers with ParamTableColumns names (one row of values)
            result_kwargs: Dict[str, Any] = {}
            optional_slugs = {
                ParamTableColumns.S_REAL_1.slug,
                ParamTableColumns.S_REAL_CUSTOM1.slug,
                ParamTableColumns.S_REAL_CUSTOM2.slug,
                ParamTableColumns.S_CUSTOM1.slug,
                ParamTableColumns.S_CUSTOM2.slug,
            }
            for param in ParamTableColumns:
                col = last_col_for(param.name)
                value = ws.cell(row=2, column=col).value if col is not None else None
                if value is None and param.slug in optional_slugs:
                    # Optional
                    continue
                if value is None:
                    value = 0
                result_kwargs[param.slug] = value

            items.append(
                dict(
                    cell=i,
                    name=cell_name,
                    diameter_list=diameter_list,
                    rn_sqrt_list=rn_sqrt_list,
                    initial_data=initial_data,
                    **result_kwargs,
                )
            )

        return items, errors
