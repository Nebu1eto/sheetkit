# 성능

SheetKit은 Rust와 TypeScript 애플리케이션 모두에 네이티브 Rust 성능을 제공합니다. 이 페이지에서는 SheetKit이 얼마나 빠른지, 그리고 이를 가능하게 하는 최적화 기법들을 설명합니다.

## SheetKit은 얼마나 빠른가요?

### ExcelJS, SheetJS와의 비교 (Node.js)

기존 Node.js 벤치마크 스위트(`benchmarks/node/RESULTS.md`) 기준으로, SheetKit은 대표적인 읽기/쓰기 시나리오에서 ExcelJS와 SheetJS보다 일관되게 빠르게 동작합니다.

| 시나리오 | SheetKit | ExcelJS | SheetJS |
|---------|----------|---------|---------|
| 대용량 읽기 (50k 행 x 20 열) | 541ms | 1.24s | 1.56s |
| 대용량 쓰기 (50k 행 x 20 열) | 469ms | 2.62s | 1.09s |
| 버퍼 왕복 (10k 행) | 123ms | 319ms | 163ms |
| 랜덤 접근 읽기 (50k 행 파일에서 1k 셀) | 453ms | 1.27s | 1.34s |

### Rust Excel 라이브러리와의 비교

Rust 라이브러리 중 SheetKit은 가장 빠른 writer입니다. 읽기의 경우 비교 가능한 전체 읽기 작업 기준에서 calamine(읽기 전용)이 더 빠르며, SheetKit은 단일 crate에서 읽기+수정+쓰기를 모두 지원하는 유일한 라이브러리입니다.

| 시나리오 | SheetKit | calamine | rust_xlsxwriter | edit-xlsx |
|---------|----------|----------|-----------------|-----------|
| 대용량 읽기 (50k 행) | 489ms | 323ms | N/A | 39ms* |
| 대용량 쓰기 (50k 행 x 20 열) | 497ms | N/A | 917ms | 946ms |
| 스트리밍 쓰기 (50k 행) | 200ms | N/A | 922ms | N/A |
| 수정 (50k 행 파일에서 1k 셀, lazy) | 688ms | N/A | N/A | N/A |

\* `edit-xlsx`에서 `*`가 표시된 값은 작업량 카운트 또는 값 프로브가 일치하지 않아 Winner 계산에서 제외된 결과입니다.

#### `edit-xlsx` 읽기 이상치 원인

SpreadsheetML에서 `workbook.xml`의 `fileVersion`, `workbookPr`, `bookViews`는 선택 요소(옵션)입니다.  
하지만 `edit-xlsx` 0.4.x는 일부 파일에서 이 요소들을 역직렬화 시 필수처럼 처리할 수 있습니다. 이때 파싱이 실패하면 기본 워크북/워크시트 구조로 fallback되어 런타임은 매우 짧게 측정되지만 `rows=0`, `cells=0`이 되는 경우가 발생합니다.

공정한 비교를 위해 Rust 비교 벤치마크는 다음 조건을 만족한 결과만 비교합니다.
- 라이브러리 간 행/셀 작업량 카운트 일치
- 동일 좌표 샘플에 대한 값 프로브 일치

이 조건을 만족하지 못한 결과는 non-comparable로 표시하고 Winner 계산에서 제외합니다.

### Rust vs Node.js 오버헤드

SheetKit의 Node.js 바인딩은 네이티브 Rust와 매우 가까운 성능을 유지합니다:

| 작업 | 오버헤드 |
|------|----------|
| **읽기 작업 (sync)** | 약 1.04배 (일반적으로 약 4% 느림) |
| **읽기 작업 (async)** | 약 1.04배 (일반적으로 약 4% 느림) |
| **쓰기 작업 (batch)** | 약 0.86배 (V8 문자열 처리로 더 빠름) |
| **스트리밍 쓰기** | 1.51배 (51% 느림) |
| **버퍼 왕복** | 약 1.0배 (거의 동일) |

대부분의 실제 워크로드에서 Node.js 성능은 네이티브 Rust와 매우 유사합니다.

### 읽기 성능 비교

| 시나리오 | Rust | Node.js | 오버헤드 |
|---------|------|---------|----------|
| 대용량 데이터 (50k 행 x 20 열) | 518ms | 541ms | +4% |
| 많은 스타일 (5k 행, 서식 적용) | 27ms | 27ms | 0% |
| 다중 시트 (10개 시트 x 5k 행) | 301ms | 290ms | -4% (더 빠름) |
| 수식 (10k 행) | 33ms | 34ms | +3% |
| 문자열 (20k 행 텍스트 중심) | 109ms | 117ms | +7% |

### 쓰기 성능 비교

| 시나리오 | Rust | Node.js | 오버헤드 |
|---------|------|---------|----------|
| 50k 행 x 20 열 | 544ms | 469ms | -14% (더 빠름) |
| 5k 스타일 적용 행 | 28ms | 32ms | +14% |
| 10k 행 (수식 포함) | 24ms | 27ms | +13% |
| 20k 텍스트 중심 행 | 108ms | 86ms | -20% (더 빠름) |

참고: 일부 쓰기 시나리오에서는 데이터 생성 시 V8의 효율적인 문자열 처리와 batch `setSheetData()` API 덕분에 Node.js가 Rust보다 약간 더 빠른 성능을 보입니다.

### 확장성 성능

읽기 성능은 다양한 파일 크기에서 일관성을 유지합니다:

| 행 수 | Rust | Node.js | 오버헤드 |
|------|------|---------|----------|
| 1k | 5ms | 5ms | 0% |
| 10k | 51ms | 50ms | -2% (더 빠름) |
| 100k | 539ms | 535ms | -1% (더 빠름) |

쓰기 성능은 선형적으로 확장됩니다:

| 행 수 | Rust | Node.js | 오버헤드 |
|------|------|---------|----------|
| 1k | 5ms | 4ms | -20% (더 빠름) |
| 10k | 48ms | 46ms | -4% (더 빠름) |
| 50k | 258ms | 231ms | -10% (더 빠름) |
| 100k | 531ms | 485ms | -9% (더 빠름) |

## Raw Buffer 전송과 메모리 동작

SheetKit은 셀별 JavaScript 객체 대신 Raw Buffer로 시트 데이터를 전달하여 Node.js-Rust 경계 비용을 줄입니다. 이 전송 모델은 FFI 경계를 큰 단위로 유지하여 객체 마샬링 오버헤드를 낮추고, 읽기 중심 경로에서 GC 부담을 줄입니다.

## 주요 최적화 기법

### 1. 버퍼 기반 FFI 전송

각 셀에 대해 개별 JavaScript 객체를 생성하는 대신, SheetKit은 전체 시트를 컴팩트한 바이너리 버퍼로 직렬화하여 단일 작업으로 FFI 경계를 넘습니다.

**최적화 전**: 셀 단위 객체를 FFI 경계로 개별 전송
**최적화 후**: 시트 단위 페이로드를 Raw Buffer로 단일 전송

이 최적화는:
- 읽기 경로의 FFI 오버헤드를 줄입니다
- 셀 단위 객체 생성으로 인한 할당 및 GC 부담을 줄입니다
- 완전한 타입 안전성을 유지합니다

### 2. 내부 데이터 구조 최적화

SheetKit의 내부 표현은 할당을 최소화합니다:

- **CompactCellRef**: 셀 참조를 heap `String` 대신 inline `[u8;10]` 배열로 저장합니다
- **CellTypeTag**: 셀 타입을 `Option<String>` 대신 1바이트 enum으로 저장합니다
- **Sparse-to-dense 변환**: 최적화된 행 반복으로 중간 할당을 방지합니다

이러한 최적화는 Rust와 Node.js 모두의 성능에 도움이 됩니다.

### 3. 밀도 기반 인코딩

버퍼 인코더는 셀 밀도에 따라 자동으로 dense와 sparse 레이아웃 중 하나를 선택합니다:
- 셀 점유율이 ≥30%인 파일은 dense 인코딩
- 셀 점유율이 <30%인 파일은 sparse 인코딩

이를 통해 모든 파일 타입에 대해 최적의 메모리 사용을 보장합니다.

## 벤치마크 환경

모든 벤치마크는 다음 환경에서 수행되었습니다:

| 구성 요소 | 버전 |
|---------|------|
| **CPU** | Apple M4 Pro |
| **RAM** | 24 GB |
| **OS** | macOS arm64 (Apple Silicon) |
| **Node.js** | v25.3.0 |
| **Rust** | rustc 1.93.0 |

결과는 시나리오당 1회 워밍업 실행 후 5회 실행의 중앙값입니다.

## 벤치마크 범위와 데이터

이 페이지의 수치는 리포지토리 내부의 SheetKit Rust/Node.js 벤치마크 스위트에서 측정한 결과입니다. 실제 결과는 데이터 형태, 사용 기능, 런타임 환경에 따라 달라집니다.

벤치마크 방법론과 원시 데이터는 리포지토리의 [`benchmarks/COMPARISON.md`](https://github.com/Nebu1eto/sheetkit/blob/main/benchmarks/COMPARISON.md)를 참조하세요.

## 성능 팁

### 읽기 모드 선택

SheetKit은 `open()` 시 파싱 범위를 제어하는 세 가지 읽기 모드를 지원합니다:

| 읽기 모드 | Open 비용 | Open 시 메모리 | 적합한 용도 |
|-----------|-----------|--------------|-----------|
| `lazy` (기본값) | 낮음 -- ZIP 인덱스 + 메타데이터만 | 최소 | 대부분의 워크로드. 시트는 첫 접근 시 파싱됩니다. |
| `eager` | 높음 -- 모든 시트 파싱 | 전체 워크북 | open 직후 모든 시트가 필요한 경우. |
| `stream` | 현재 lazy와 동일 | 현재 lazy와 동일 | 향후 호환성 모드이며 Rust 코어에서는 현재 lazy와 동일하게 동작합니다. |

```typescript
// Lazy open (기본값): 가장 빠른 open, 시트를 필요시 파싱
const wb = await Workbook.open("huge.xlsx");
const rows = wb.getRows("Sheet1"); // Sheet1이 여기서 파싱됨

// Eager open: open 시 모든 시트 파싱
const wb2 = await Workbook.open("huge.xlsx", { readMode: "eager" });

// Stream 모드: 메모리 제한 순방향 읽기
const wb3 = await Workbook.open("huge.xlsx", { readMode: "stream" });
const reader = await wb3.openSheetReader("Sheet1", { batchSize: 500 });
for await (const batch of reader) {
  for (const row of batch) {
    process(row);
  }
}
```

### 보조 파트 지연 로드

기본적으로 보조 파트(comments, charts, images, pivot table)는 open 시 파싱되지 않습니다. 해당 파트가 필요한 메서드를 처음 호출할 때 로드됩니다. 모든 파트를 즉시 사용해야 하는 경우 `auxParts: 'eager'`를 설정합니다:

```typescript
const wb = await Workbook.open("report.xlsx", { auxParts: "eager" });
```

### 읽기 위주 워크로드

`OpenOptions`를 사용하여 필요한 부분만 로드합니다:

```typescript
const wb = await Workbook.open("huge.xlsx", {
  readMode: "lazy",
  sheetRows: 1000,      // 시트당 첫 1000행만 읽기
  sheets: ["Sheet1"],   // Sheet1만 파싱
  maxUnzipSize: 100_000_000  // 압축 해제 크기 제한
});
```

### SheetStreamReader를 사용한 스트리밍 읽기

랜덤 셀 접근이 필요 없는 대용량 파일의 경우, `openSheetReader()`를 사용하여 메모리 제한 순방향 반복을 수행합니다:

```typescript
const wb = await Workbook.open("huge.xlsx", { readMode: "stream" });
const reader = await wb.openSheetReader("Sheet1", { batchSize: 1000 });

for await (const batch of reader) {
  for (const row of batch) {
    // 각 행 처리 -- 한 번에 하나의 배치만 메모리에 존재
  }
}
```

### Raw Buffer V2 전송

`getRowsBufferV2()`는 글로벌 문자열 테이블을 즉시 생성하지 않고 행별로 점진적으로 디코딩할 수 있는 inline 문자열 포함 v2 바이너리 buffer를 생성합니다:

```typescript
const bufV2 = wb.getRowsBufferV2("Sheet1");
```

### Copy-on-Write 저장

`lazy` 모드로 열린 워크북을 저장할 때, 변경되지 않은 시트는 파싱-직렬화 왕복 없이 원본 ZIP 엔트리에서 직접 기록됩니다. 이는 대용량 워크북에서 일부 시트만 수정하는 워크로드의 저장 지연 시간을 크게 줄여줍니다.

### 쓰기 위주 워크로드

순차적 행 쓰기에는 `StreamWriter`를 사용합니다. 각 `write_row()` 호출은 디스크의 임시 파일에 직접 기록하므로, 행 수에 관계없이 메모리 사용량이 일정하게 유지됩니다:

```typescript
const wb = new Workbook();
const sw = wb.newStreamWriter("LargeSheet");

for (let i = 1; i <= 100_000; i++) {
  sw.writeRow(i, [`Item_${i}`, i * 1.5]);
}

wb.applyStreamWriter(sw);
await wb.save("output.xlsx");
```

### 대용량 파일

Lazy open과 `StreamWriter`를 결합합니다:

```typescript
// Lazy open -- 메타데이터만 파싱됨
const wb = await Workbook.open("input.xlsx");

// 스트리밍으로 처리
const sw = wb.newStreamWriter("ProcessedData");
// ... 데이터 처리 ...
wb.applyStreamWriter(sw);
```

> **참고:** 스트리밍된 시트의 셀 값은 `applyStreamWriter` 이후 직접 읽을 수 없습니다. 데이터를 읽으려면 워크북을 저장한 후 다시 열어야 합니다.

## 다음 단계

- [시작 가이드](./getting-started.md) - 기본 사항 학습
- [아키텍처](./architecture.md) - 내부 설계 이해
- [API 레퍼런스](./api-reference/) - 사용 가능한 모든 메서드 탐색
