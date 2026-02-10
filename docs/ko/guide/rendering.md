# SVG Rendering

SheetKit은 워크시트를 SVG로 렌더링하여 시각적 미리보기, 썸네일, 또는 웹 애플리케이션 임베딩에 활용할 수 있습니다. 렌더러는 워크시트의 셀 값, 스타일, 레이아웃으로부터 자체 완결형 SVG 문자열을 생성합니다.

## 기본 사용법

```typescript
import { Workbook } from 'sheetkit';

const wb = Workbook.openSync('report.xlsx');
const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

// Write to file
import { writeFileSync } from 'node:fs';
writeFileSync('preview.svg', svg);
```

## Render Options

`renderToSvg` 메서드는 `JsRenderOptions` 객체를 인수로 받습니다:

| Option             | Type               | Default   | Description                                       |
|--------------------|--------------------|-----------|----------------------------------------------------|
| `sheetName`        | `string`           | 필수      | 렌더링할 시트 이름입니다.                           |
| `range`            | `string \| null`   | `null`    | 렌더링할 셀 범위입니다 (예: `"A1:F20"`). 생략하면 사용된 범위를 렌더링합니다. |
| `showGridlines`    | `boolean \| null`  | `true`    | 셀 사이에 격자선을 그릴지 여부입니다.               |
| `showHeaders`      | `boolean \| null`  | `true`    | 행/열 헤더(A, B, 1, 2)를 그릴지 여부입니다.        |
| `scale`            | `number \| null`   | `1.0`     | 출력 배율입니다 (2.0 = 2배 크기).                   |
| `defaultFontFamily`| `string \| null`   | `"Arial"` | 셀 텍스트의 기본 폰트 패밀리입니다.                 |
| `defaultFontSize`  | `number \| null`   | `11.0`    | 기본 폰트 크기(포인트)입니다.                       |

## 부분 범위 렌더링

시트의 특정 영역만 렌더링할 수 있습니다:

```typescript
const svg = wb.renderToSvg({
  sheetName: 'Sheet1',
  range: 'A1:D10',
});
```

## 시각적 출력 제어

```typescript
// Minimal rendering without headers or gridlines
const svg = wb.renderToSvg({
  sheetName: 'Sheet1',
  showGridlines: false,
  showHeaders: false,
});

// High-resolution rendering (2x scale)
const svg2x = wb.renderToSvg({
  sheetName: 'Sheet1',
  scale: 2,
});
```

## Rust API

Rust API는 `Workbook::render_to_svg` 메서드를 통해 사용할 수 있습니다:

```rust
use sheetkit::Workbook;
use sheetkit::RenderOptions;

let wb = Workbook::open("report.xlsx").unwrap();
let svg = wb.render_to_svg(&RenderOptions {
    sheet_name: "Sheet1".to_string(),
    ..RenderOptions::default()
}).unwrap();
```

하위 레벨의 `render::render_to_svg` 함수도 `WorksheetXml`, `SharedStringTable`, `StyleSheet` 참조를 직접 사용하여 호출할 수 있습니다.

## 지원 기능

SVG 렌더러는 다음과 같은 시각적 기능을 지원합니다:

- 셀 텍스트 값 (문자열, 숫자, 불리언, 날짜, 수식 캐시 결과)
- 열 너비 및 행 높이 (명시적 설정 및 기본값)
- 폰트 스타일: 굵게, 기울임, 밑줄, 취소선, 글꼴 색상, 글꼴 이름, 글꼴 크기
- 셀 채우기 색상 (단색 패턴 채우기)
- 셀 테두리 (좌, 우, 상, 하) - 선 스타일 및 색상 포함
- 텍스트 정렬 (수평: 왼쪽, 가운데, 오른쪽; 수직: 위, 가운데, 아래)
- 행/열 헤더 (배경 음영 포함)
- 격자선 (표시 여부 설정 가능)
- 출력 크기 배율 조정
- 부분 범위 렌더링

## 알려진 제한 사항

다음 기능은 아직 렌더러에서 지원되지 않습니다:

- 병합된 셀 (개별 셀로 렌더링됩니다)
- 조건부 서식 (SVG에 색상이 적용되지 않습니다)
- 이미지 및 차트
- Rich text (셀 내 개별 서식 적용)
- 그래디언트 채우기
- Theme 및 indexed 색상 해석 (검은색으로 대체됩니다)
- 숫자 서식 표시 (원시 값이 표시됩니다)
- 텍스트 줄바꿈 및 오버플로우
- 대각선 테두리
- 숨겨진 행 및 열
- 개요/그룹 접기
