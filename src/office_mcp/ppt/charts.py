"""ChartMixin — embedded Office chart lifecycle + data + title + type.

Mixed into PowerPointController. Relies on base for:
    self._validate_index(), self._require_shape(), self.pres

PowerPoint chart model (requires Excel installed — charts are backed by
an embedded Excel workbook):
  Shape (Type=msoChart)
    └─ .Chart
         ├─ .ChartType                    (XlChartType int)
         ├─ .HasTitle / .ChartTitle.Text
         ├─ .ChartData.Workbook           (xlwings-shaped Excel COM)
         │     └─ .Worksheets(1)
         │           └─ .Range("A1") ...   ← write cells here
         └─ .Refresh()                    (reload after edits)

Data layout written by chart_set_data / chart_add (categories+series):
   A      B       C       ...
1  ""     Plan    Actual
2  Q1     100     110
3  Q2     120     115
"""

from __future__ import annotations

import pywintypes

from .constants import (
    AXIS_TYPE_MAP,
    CHART_TYPE_MAP,
    DATA_LABEL_TYPE_MAP,
    LEGEND_POSITION_MAP,
)
from .errors import PowerPointError


def _col_letter(n: int) -> str:
    """1 → A, 2 → B, …, 27 → AA. Used to address Excel cells."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class ChartMixin:
    @staticmethod
    def _resolve_chart_type(chart_type: int | str) -> int:
        if isinstance(chart_type, int):
            return chart_type
        key = chart_type.lower().strip()
        if key not in CHART_TYPE_MAP:
            raise PowerPointError(
                f"Unknown chart_type '{chart_type}'. Choose from {sorted(CHART_TYPE_MAP)}"
            )
        return CHART_TYPE_MAP[key]

    def _require_chart(self, slide_index: int, shape_id: int):
        slide, shape = self._require_shape(slide_index, shape_id)
        try:
            chart = shape.Chart
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape id={shape_id} on slide {slide_index} is not a chart: {e}")
        return slide, shape, chart

    def _write_chart_data(
        self,
        chart,
        categories: list,
        series: dict | list,
    ) -> None:
        """Write categories (col A) + series (col B+) into the backing workbook.

        ``series`` accepts:
          - ``{"Plan": [100, 120], "Actual": [110, 115]}``  (dict of name→values)
          - ``[{"name": "Plan", "values": [100, 120]}, ...]`` (ordered list)
        """
        try:
            workbook = chart.ChartData.Workbook
        except pywintypes.com_error as e:
            raise PowerPointError(f"Cannot open chart workbook (is Excel installed?): {e}")
        try:
            ws = workbook.Worksheets(1)

            if isinstance(series, dict):
                pairs = list(series.items())
            else:
                pairs = [(s["name"], s["values"]) for s in series]

            n_cats = len(categories)
            n_series = len(pairs)
            if n_series == 0:
                raise PowerPointError("series must contain at least one entry")

            # Wipe a generous bounding box so leftover sample data goes away
            last_col = _col_letter(max(n_series + 1, 4))
            last_row = max(n_cats + 1, 5)
            ws.Range(f"A1:{last_col}{last_row}").ClearContents()

            # Headers: B1, C1, ...
            for i, (name, _) in enumerate(pairs, start=2):
                ws.Cells(1, i).Value = name
            # Categories: A2, A3, ...
            for i, cat in enumerate(categories, start=2):
                ws.Cells(i, 1).Value = cat
            # Values
            for c_idx, (_, vals) in enumerate(pairs, start=2):
                if len(vals) != n_cats:
                    raise PowerPointError(
                        f"Series '{pairs[c_idx - 2][0]}' has {len(vals)} values "
                        f"but {n_cats} categories"
                    )
                for r_idx, v in enumerate(vals, start=2):
                    ws.Cells(r_idx, c_idx).Value = v

            # Rebind source range so the chart picks up exact data extent
            try:
                chart.SetSourceData(
                    Source=ws.Range(f"A1:{_col_letter(n_series + 1)}{n_cats + 1}"),
                )
            except pywintypes.com_error:
                pass  # Best-effort; some chart types reject explicit source
            chart.Refresh()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Write chart data failed: {e}")
        finally:
            try:
                workbook.Application.Quit()
            except Exception:
                pass

    def chart_add(
        self,
        slide_index: int,
        chart_type: int | str,
        left: float,
        top: float,
        width: float,
        height: float,
        categories: list | None = None,
        series: dict | list | None = None,
        title: str | None = None,
    ) -> dict:
        """Add an embedded chart. Pass categories+series to overwrite the
        sample data PowerPoint inserts by default."""
        self._validate_index(slide_index)
        xl_type = self._resolve_chart_type(chart_type)
        slide = self.pres.Slides(slide_index)
        try:
            shape = slide.Shapes.AddChart2(
                -1,
                int(xl_type),
                float(left),
                float(top),
                float(width),
                float(height),
                True,  # NewLayout
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddChart2 failed: {e}")

        if categories is not None and series is not None:
            self._write_chart_data(shape.Chart, categories, series)

        if title is not None:
            try:
                shape.Chart.HasTitle = True
                shape.Chart.ChartTitle.Text = title
            except pywintypes.com_error as e:
                raise PowerPointError(f"Set chart title failed: {e}")

        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "chart_type": int(xl_type),
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
            "title": title,
        }

    def chart_set_data(
        self,
        slide_index: int,
        shape_id: int,
        categories: list,
        series: dict | list,
    ) -> dict:
        """Overwrite a chart's categories and series."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        self._write_chart_data(chart, categories, series)
        return {
            "success": True,
            "shape_id": int(shape_id),
            "categories": len(categories),
            "series": len(series) if hasattr(series, "__len__") else None,
        }

    def chart_set_title(
        self,
        slide_index: int,
        shape_id: int,
        title: str,
    ) -> dict:
        """Set or clear a chart's title. Pass empty string to hide."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        try:
            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            else:
                chart.HasTitle = False
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set chart title failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "title": title}

    def chart_set_type(
        self,
        slide_index: int,
        shape_id: int,
        chart_type: int | str,
    ) -> dict:
        """Switch a chart to a different XlChartType."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        xl_type = self._resolve_chart_type(chart_type)
        try:
            chart.ChartType = int(xl_type)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set chart type failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "chart_type": int(xl_type),
        }

    # ─────────────────────────────────────────────────────────
    # Polish — series colors, data labels, axes, legend, gridlines, style.
    # These are the difference between "default-styled chart" and
    # "board-ready visual." See Stage 5 plan.
    # ─────────────────────────────────────────────────────────

    @staticmethod
    def _resolve_series(chart, series: int | str):
        """Return SeriesCollection item by 1-based index or by series name."""
        if isinstance(series, int):
            return chart.SeriesCollection(int(series))
        # Name lookup — iterate until match (Series.Name returns the legend label)
        try:
            count = int(chart.SeriesCollection().Count)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Cannot enumerate series: {e}")
        for i in range(1, count + 1):
            try:
                s = chart.SeriesCollection(i)
                if str(s.Name) == series:
                    return s
            except pywintypes.com_error:
                continue
        raise PowerPointError(f"Series '{series}' not found")

    def chart_set_series_color(
        self,
        slide_index: int,
        shape_id: int,
        series: int | str,
        color,
    ) -> dict:
        """Set the fill/line color of one data series. ``series`` accepts a
        1-based index or the series name (legend label)."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        rgb = self._parse_color(color)
        ser = self._resolve_series(chart, series)
        try:
            # Fill applies to bar/column/area/pie; line applies to line/scatter.
            # Setting both is safe — PowerPoint ignores the one that doesn't apply.
            try:
                ser.Format.Fill.Solid()
                ser.Format.Fill.ForeColor.RGB = rgb
            except pywintypes.com_error:
                pass
            try:
                ser.Format.Line.ForeColor.RGB = rgb
            except pywintypes.com_error:
                pass
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set series color failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "series": series,
            "rgb": rgb,
        }

    def chart_set_data_labels(
        self,
        slide_index: int,
        shape_id: int,
        show: bool | str = True,
        number_format: str | None = None,
        series: int | str | None = None,
    ) -> dict:
        """Toggle data labels and optionally set their number format.

        ``show`` accepts True/False (binary on/off as value labels) or a
        string ∈ {none, value, percent, label, label_percent} for finer
        control. ``series`` scopes to one series; omit to apply chart-wide.
        ``number_format`` follows Excel conventions, e.g. ``"$#,##0"``.
        """
        _, _, chart = self._require_chart(slide_index, shape_id)
        if isinstance(show, bool):
            label_type = DATA_LABEL_TYPE_MAP["value" if show else "none"]
        else:
            key = show.lower().strip()
            if key not in DATA_LABEL_TYPE_MAP:
                raise PowerPointError(
                    f"Unknown data label type '{show}'. Choose from {sorted(DATA_LABEL_TYPE_MAP)}"
                )
            label_type = DATA_LABEL_TYPE_MAP[key]

        targets = [self._resolve_series(chart, series)] if series is not None else None
        try:
            if targets is None:
                chart.ApplyDataLabels(Type=label_type)
            else:
                for s in targets:
                    s.ApplyDataLabels(Type=label_type)
            if number_format is not None:
                if targets is None:
                    count = int(chart.SeriesCollection().Count)
                    targets = [chart.SeriesCollection(i) for i in range(1, count + 1)]
                for s in targets:
                    try:
                        # DataLabels is a callable accessor in pywin32's COM
                        # proxy — must invoke to reach the collection.
                        s.DataLabels().NumberFormat = str(number_format)
                    except (pywintypes.com_error, AttributeError):
                        continue
        except pywintypes.com_error as e:
            raise PowerPointError(f"Apply data labels failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "label_type": int(label_type),
            "series": series,
            "number_format": number_format,
        }

    @staticmethod
    def _resolve_axis_type(axis: int | str) -> int:
        if isinstance(axis, int):
            return axis
        key = axis.lower().strip()
        if key not in AXIS_TYPE_MAP:
            raise PowerPointError(f"Unknown axis '{axis}'. Choose from {sorted(AXIS_TYPE_MAP)}")
        return AXIS_TYPE_MAP[key]

    def chart_set_axis(
        self,
        slide_index: int,
        shape_id: int,
        axis: int | str = "y",
        minimum: float | None = None,
        maximum: float | None = None,
        number_format: str | None = None,
        title: str | None = None,
    ) -> dict:
        """Configure an axis. ``minimum`` / ``maximum`` force the scale range
        (typical: minimum=0 for revenue charts). ``number_format`` follows
        Excel conventions (``$#,##0``, ``0.0%``). ``title`` adds an axis label.
        """
        _, _, chart = self._require_chart(slide_index, shape_id)
        ax_type = self._resolve_axis_type(axis)
        applied: dict = {}
        try:
            ax = chart.Axes(int(ax_type))
            if minimum is not None:
                ax.MinimumScale = float(minimum)
                applied["minimum"] = float(minimum)
            if maximum is not None:
                ax.MaximumScale = float(maximum)
                applied["maximum"] = float(maximum)
            if number_format is not None:
                ax.TickLabels.NumberFormat = str(number_format)
                applied["number_format"] = number_format
            if title is not None:
                if title:
                    ax.HasTitle = True
                    ax.AxisTitle.Text = title
                else:
                    ax.HasTitle = False
                applied["title"] = title
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set axis failed: {e}")
        if not applied:
            raise PowerPointError(
                "chart_set_axis called with no attributes. Pass minimum, "
                "maximum, number_format, and/or title."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "axis": int(ax_type),
            "applied": applied,
        }

    def chart_set_legend(
        self,
        slide_index: int,
        shape_id: int,
        position: str | None = None,
        visible: bool | None = None,
    ) -> dict:
        """Toggle and position the legend. ``position`` ∈
        {top, bottom, left, right, corner}. Pass ``visible=False`` to hide."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        applied: dict = {}
        try:
            if visible is False:
                chart.HasLegend = False
                applied["visible"] = False
            elif position is not None or visible is True:
                chart.HasLegend = True
                applied["visible"] = True
                if position is not None:
                    key = position.lower().strip()
                    if key not in LEGEND_POSITION_MAP:
                        raise PowerPointError(
                            f"Unknown legend position '{position}'. "
                            f"Choose from {sorted(LEGEND_POSITION_MAP)}"
                        )
                    chart.Legend.Position = LEGEND_POSITION_MAP[key]
                    applied["position"] = key
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set legend failed: {e}")
        if not applied:
            raise PowerPointError(
                "chart_set_legend called with no attributes. Pass position and/or visible."
            )
        return {"success": True, "shape_id": int(shape_id), "applied": applied}

    def chart_set_gridlines(
        self,
        slide_index: int,
        shape_id: int,
        axis: int | str = "y",
        major: bool | None = None,
        minor: bool | None = None,
    ) -> dict:
        """Toggle major/minor gridlines on one axis. Business default is
        usually major=True on the value (Y) axis, minor=False everywhere."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        ax_type = self._resolve_axis_type(axis)
        applied: dict = {}
        try:
            ax = chart.Axes(int(ax_type))
            if major is not None:
                ax.HasMajorGridlines = bool(major)
                applied["major"] = bool(major)
            if minor is not None:
                ax.HasMinorGridlines = bool(minor)
                applied["minor"] = bool(minor)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set gridlines failed: {e}")
        if not applied:
            raise PowerPointError(
                "chart_set_gridlines called with no attributes. Pass major and/or minor."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "axis": int(ax_type),
            "applied": applied,
        }

    def chart_set_style(
        self,
        slide_index: int,
        shape_id: int,
        style: int,
    ) -> dict:
        """Apply a built-in chart style by number. PowerPoint exposes ~48
        basic styles plus theme-tinted variants (often 200+ values). Picking
        a polished preset is faster than configuring axes/labels/gridlines
        manually."""
        _, _, chart = self._require_chart(slide_index, shape_id)
        try:
            chart.ChartStyle = int(style)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set chart style failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "style": int(style),
        }
