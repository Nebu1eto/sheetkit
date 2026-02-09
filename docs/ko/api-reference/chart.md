## 10. 차트

43가지 차트 유형을 지원한다. `add_chart`로 시트에 차트를 추가하며, 셀 범위로 위치와 크기를 지정한다.

### `add_chart(sheet, from_cell, to_cell, config)` / `addChart(sheet, fromCell, toCell, config)`

시트에 차트를 추가한다. `from_cell`은 차트의 왼쪽 상단, `to_cell`은 오른쪽 하단 위치를 나타낸다.

**Rust:**

```rust
use sheetkit::chart::*;

wb.add_chart("Sheet1", "E1", "L15", &ChartConfig {
    chart_type: ChartType::Col,
    title: Some("Sales Report".into()),
    series: vec![
        ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$2:$A$6".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        },
    ],
    show_legend: true,
    view_3d: None,
})?;
```

**TypeScript:**

```typescript
wb.addChart("Sheet1", "E1", "L15", {
    chartType: "col",
    title: "Sales Report",
    series: [
        {
            name: "Revenue",
            categories: "Sheet1!$A$2:$A$6",
            values: "Sheet1!$B$2:$B$6",
        },
    ],
    showLegend: true,
});
```

> Node.js에서 `chartType`은 지원되는 값만 허용된다. 지원되지 않는 값은 더 이상 기본값으로 대체되지 않고 오류를 반환한다.

### ChartSeries 구조

| 속성 | 타입 | 필수 | 설명 |
|------|------|------|------|
| `name` | `string` | O | 시리즈 이름 또는 셀 참조 |
| `categories` | `string` | O | 카테고리 데이터 범위 |
| `values` | `string` | O | 값 데이터 범위 |
| `x_values` / `xValues` | `string?` | X | Scatter/Bubble용 X 축 범위 |
| `bubble_sizes` / `bubbleSizes` | `string?` | X | Bubble 차트용 크기 범위 |

### View3DConfig 구조

3D 차트의 시점을 설정한다. 3D 차트 유형에서는 지정하지 않으면 자동으로 기본값이 적용된다.

| 속성 | 타입 | 설명 |
|------|------|------|
| `rot_x` / `rotX` | `i32?` / `number?` | X축 회전 각도 |
| `rot_y` / `rotY` | `i32?` / `number?` | Y축 회전 각도 |
| `depth_percent` / `depthPercent` | `u32?` / `number?` | 깊이 비율 (100 = 기본) |
| `right_angle_axes` / `rightAngleAxes` | `bool?` / `boolean?` | 직각 축 사용 여부 |
| `perspective` | `u32?` / `number?` | 원근 각도 |

### 차트 유형 전체 목록 (41종)

#### 세로 막대 (Column) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `col` | `ChartType::Col` | 세로 막대 |
| `colStacked` | `ChartType::ColStacked` | 누적 세로 막대 |
| `colPercentStacked` | `ChartType::ColPercentStacked` | 100% 누적 세로 막대 |
| `col3D` | `ChartType::Col3D` | 3D 세로 막대 |
| `col3DStacked` | `ChartType::Col3DStacked` | 3D 누적 세로 막대 |
| `col3DPercentStacked` | `ChartType::Col3DPercentStacked` | 3D 100% 누적 세로 막대 |

#### 가로 막대 (Bar) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `bar` | `ChartType::Bar` | 가로 막대 |
| `barStacked` | `ChartType::BarStacked` | 누적 가로 막대 |
| `barPercentStacked` | `ChartType::BarPercentStacked` | 100% 누적 가로 막대 |
| `bar3D` | `ChartType::Bar3D` | 3D 가로 막대 |
| `bar3DStacked` | `ChartType::Bar3DStacked` | 3D 누적 가로 막대 |
| `bar3DPercentStacked` | `ChartType::Bar3DPercentStacked` | 3D 100% 누적 가로 막대 |

#### 꺾은선 (Line) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `line` | `ChartType::Line` | 꺾은선 |
| `lineStacked` | `ChartType::LineStacked` | 누적 꺾은선 |
| `linePercentStacked` | `ChartType::LinePercentStacked` | 100% 누적 꺾은선 |
| `line3D` | `ChartType::Line3D` | 3D 꺾은선 |

#### 원형 (Pie) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `pie` | `ChartType::Pie` | 원형 |
| `pie3D` | `ChartType::Pie3D` | 3D 원형 |
| `doughnut` | `ChartType::Doughnut` | 도넛형 |

#### 영역 (Area) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `area` | `ChartType::Area` | 영역 |
| `areaStacked` | `ChartType::AreaStacked` | 누적 영역 |
| `areaPercentStacked` | `ChartType::AreaPercentStacked` | 100% 누적 영역 |
| `area3D` | `ChartType::Area3D` | 3D 영역 |
| `area3DStacked` | `ChartType::Area3DStacked` | 3D 누적 영역 |
| `area3DPercentStacked` | `ChartType::Area3DPercentStacked` | 3D 100% 누적 영역 |

#### 분산형 (Scatter) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `scatter` | `ChartType::Scatter` | 분산형 (표식만) |
| `scatterSmooth` | `ChartType::ScatterSmooth` | 부드러운 선 |
| `scatterStraight` | `ChartType::ScatterLine` | 직선 |

#### 방사형 (Radar) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `radar` | `ChartType::Radar` | 방사형 |
| `radarFilled` | `ChartType::RadarFilled` | 채워진 방사형 |
| `radarMarker` | `ChartType::RadarMarker` | 표식이 있는 방사형 |

#### 주식 (Stock) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `stockHLC` | `ChartType::StockHLC` | 고가-저가-종가 |
| `stockOHLC` | `ChartType::StockOHLC` | 시가-고가-저가-종가 |
| `stockVHLC` | `ChartType::StockVHLC` | 거래량-고가-저가-종가 |
| `stockVOHLC` | `ChartType::StockVOHLC` | 거래량-시가-고가-저가-종가 |

#### 기타 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `bubble` | `ChartType::Bubble` | 거품형 |
| `surface` | `ChartType::Surface` | 표면형 |
| `surfaceTop` | `ChartType::Surface3D` | 3D 표면형 |
| `surfaceWireframe` | `ChartType::SurfaceWireframe` | 와이어프레임 표면형 |
| `surfaceTopWireframe` | `ChartType::SurfaceWireframe3D` | 3D 와이어프레임 표면형 |

#### 콤보 (Combo) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `colLine` | `ChartType::ColLine` | 세로 막대 + 꺾은선 |
| `colLineStacked` | `ChartType::ColLineStacked` | 누적 세로 막대 + 꺾은선 |
| `colLinePercentStacked` | `ChartType::ColLinePercentStacked` | 100% 누적 세로 막대 + 꺾은선 |

---
