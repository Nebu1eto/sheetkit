## 10. Charts

Charts render data from cell ranges and are anchored between two cells (top-left and bottom-right corners of the chart area).

### `add_chart` / `addChart`

Add a chart to a sheet.

**Rust:**

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType};

let config = ChartConfig {
    chart_type: ChartType::Col,
    title: Some("Quarterly Sales".to_string()),
    series: vec![ChartSeries {
        name: "Revenue".to_string(),
        categories: "Sheet1!$A$2:$A$5".to_string(),
        values: "Sheet1!$B$2:$B$5".to_string(),
        x_values: None,
        bubble_sizes: None,
    }],
    show_legend: true,
    view_3d: None,
};
wb.add_chart("Sheet1", "D1", "K15", &config)?;
```

**TypeScript:**

```typescript
wb.addChart("Sheet1", "D1", "K15", {
    chartType: "col",
    title: "Quarterly Sales",
    series: [{
        name: "Revenue",
        categories: "Sheet1!$A$2:$A$5",
        values: "Sheet1!$B$2:$B$5",
    }],
    showLegend: true,
});
```

> Note (Node.js): `chartType` must be one of the supported values. Unknown values now return an error instead of silently falling back.

### ChartConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `chart_type` | `ChartType` | `string` | Chart type (see table below) |
| `title` | `Option<String>` | `string?` | Chart title |
| `series` | `Vec<ChartSeries>` | `JsChartSeries[]` | Data series |
| `show_legend` | `bool` | `boolean?` | Show legend (default: true) |
| `view_3d` | `Option<View3DConfig>` | `JsView3DConfig?` | 3D rotation settings |

### ChartSeries

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Series name |
| `categories` | `String` | `string` | Category axis range (e.g., "Sheet1!$A$2:$A$5") |
| `values` | `String` | `string` | Value axis range (e.g., "Sheet1!$B$2:$B$5") |
| `x_values` | `Option<String>` | `string?` | X-axis values (scatter/bubble charts only) |
| `bubble_sizes` | `Option<String>` | `string?` | Bubble sizes (bubble charts only) |

### Supported Chart Types (57 types)

**Column charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Col` | `"col"` | Clustered column |
| `ChartType::ColStacked` | `"colStacked"` | Stacked column |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | 100% stacked column |
| `ChartType::Col3D` | `"col3D"` | 3D clustered column |
| `ChartType::Col3DStacked` | `"col3DStacked"` | 3D stacked column |
| `ChartType::Col3DPercentStacked` | `"col3DPercentStacked"` | 3D 100% stacked column |
| `ChartType::Col3DCone` | `"col3DCone"` | 3D cone column |
| `ChartType::Col3DConeStacked` | `"col3DConeStacked"` | 3D stacked cone column |
| `ChartType::Col3DConePercentStacked` | `"col3DConePercentStacked"` | 3D 100% stacked cone column |
| `ChartType::Col3DPyramid` | `"col3DPyramid"` | 3D pyramid column |
| `ChartType::Col3DPyramidStacked` | `"col3DPyramidStacked"` | 3D stacked pyramid column |
| `ChartType::Col3DPyramidPercentStacked` | `"col3DPyramidPercentStacked"` | 3D 100% stacked pyramid column |
| `ChartType::Col3DCylinder` | `"col3DCylinder"` | 3D cylinder column |
| `ChartType::Col3DCylinderStacked` | `"col3DCylinderStacked"` | 3D stacked cylinder column |
| `ChartType::Col3DCylinderPercentStacked` | `"col3DCylinderPercentStacked"` | 3D 100% stacked cylinder column |

**Bar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Bar` | `"bar"` | Clustered bar |
| `ChartType::BarStacked` | `"barStacked"` | Stacked bar |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | 100% stacked bar |
| `ChartType::Bar3D` | `"bar3D"` | 3D clustered bar |
| `ChartType::Bar3DStacked` | `"bar3DStacked"` | 3D stacked bar |
| `ChartType::Bar3DPercentStacked` | `"bar3DPercentStacked"` | 3D 100% stacked bar |

**Line charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Line` | `"line"` | Line |
| `ChartType::LineStacked` | `"lineStacked"` | Stacked line |
| `ChartType::LinePercentStacked` | `"linePercentStacked"` | 100% stacked line |
| `ChartType::Line3D` | `"line3D"` | 3D line |

**Pie charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Pie` | `"pie"` | Pie |
| `ChartType::Pie3D` | `"pie3D"` | 3D pie |
| `ChartType::PieOfPie` | `"pieOfPie"` | Pie of pie |
| `ChartType::BarOfPie` | `"barOfPie"` | Bar of pie |

**Area charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Area` | `"area"` | Area |
| `ChartType::AreaStacked` | `"areaStacked"` | Stacked area |
| `ChartType::AreaPercentStacked` | `"areaPercentStacked"` | 100% stacked area |
| `ChartType::Area3D` | `"area3D"` | 3D area |
| `ChartType::Area3DStacked` | `"area3DStacked"` | 3D stacked area |
| `ChartType::Area3DPercentStacked` | `"area3DPercentStacked"` | 3D 100% stacked area |

**Scatter charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Scatter` | `"scatter"` | Scatter (markers only) |
| `ChartType::ScatterSmooth` | `"scatterSmooth"` | Scatter with smooth lines |
| `ChartType::ScatterLine` | `"scatterLine"` | Scatter with straight lines |

**Radar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Radar` | `"radar"` | Radar |
| `ChartType::RadarFilled` | `"radarFilled"` | Filled radar |
| `ChartType::RadarMarker` | `"radarMarker"` | Radar with markers |

**Stock charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::StockHLC` | `"stockHLC"` | High-Low-Close |
| `ChartType::StockOHLC` | `"stockOHLC"` | Open-High-Low-Close |
| `ChartType::StockVHLC` | `"stockVHLC"` | Volume-High-Low-Close |
| `ChartType::StockVOHLC` | `"stockVOHLC"` | Volume-Open-High-Low-Close |

**Surface and contour charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Surface` | `"surface"` | 3D surface |
| `ChartType::Surface3D` | `"surface3D"` | 3D surface (top view) |
| `ChartType::SurfaceWireframe` | `"surfaceWireframe"` | Wireframe surface |
| `ChartType::SurfaceWireframe3D` | `"surfaceWireframe3D"` | Wireframe surface (top view) |
| `ChartType::Contour` | `"contour"` | Contour (2D surface projection) |
| `ChartType::WireframeContour` | `"wireframeContour"` | Wireframe contour |

**Other charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Doughnut` | `"doughnut"` | Doughnut |
| `ChartType::Bubble` | `"bubble"` | Bubble |
| `ChartType::Bubble3D` | `"bubble3D"` | 3D bubble |

**Combo charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::ColLine` | `"colLine"` | Column + line combo |
| `ChartType::ColLineStacked` | `"colLineStacked"` | Stacked column + line |
| `ChartType::ColLinePercentStacked` | `"colLinePercentStacked"` | 100% stacked column + line |

### `delete_chart` / `deleteChart`

Delete a chart anchored at the given cell. Removes the chart data, drawing anchor, relationship entry, and content type override associated with the chart.

Returns an error if no chart is found at the specified cell.

**Parameters:**

| Parameter | Rust Type | TS Type | Description |
|---|---|---|---|
| `sheet` | `&str` | `string` | Sheet name |
| `cell` | `&str` | `string` | Anchor cell of the chart (e.g., `"D1"`) |

**Rust:**

```rust
wb.delete_chart("Sheet1", "D1")?;
```

**TypeScript:**

```typescript
wb.deleteChart("Sheet1", "D1");
```

### View3DConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `rot_x` | `Option<i32>` | `number?` | X-axis rotation angle |
| `rot_y` | `Option<i32>` | `number?` | Y-axis rotation angle |
| `depth_percent` | `Option<u32>` | `number?` | Depth as percentage of chart width |
| `right_angle_axes` | `Option<bool>` | `boolean?` | Use right-angle axes |
| `perspective` | `Option<u32>` | `number?` | Perspective field of view |

---
