# 성능

SheetKit은 Rust와 TypeScript 애플리케이션 모두에 네이티브 Rust 성능을 제공합니다. 이 페이지에서는 SheetKit이 얼마나 빠른지, 그리고 이를 가능하게 하는 최적화 기법들을 설명합니다.

## SheetKit은 얼마나 빠른가요?

### ExcelJS, SheetJS와의 비교 (Node.js)

기존 Node.js 벤치마크 스위트(`benchmarks/node/RESULTS.md`) 기준으로, SheetKit은 대표적인 읽기/쓰기 시나리오에서 ExcelJS와 SheetJS보다 일관되게 빠르게 동작합니다.

| 시나리오 | SheetKit | ExcelJS | SheetJS |
|---------|----------|---------|---------|
| 대용량 읽기 (50k 행 × 20 열) | 680ms | 3.88s | 2.06s |
| 대용량 쓰기 (50k 행 × 20 열) | 657ms | 3.49s | 1.59s |
| 버퍼 왕복 (10k 행) | 167ms | 674ms | 211ms |
| 랜덤 접근 읽기 (50k 행 파일에서 1k 셀) | 550ms | 3.97s | 1.74s |

### Rust vs Node.js 오버헤드

SheetKit의 Node.js 바인딩은 네이티브 Rust와 매우 가까운 성능을 유지하며, 일부 쓰기 중심 경로에서는 더 빠르게 동작합니다:

| 작업 | 오버헤드 |
|------|----------|
| **읽기 작업 (sync)** | 약 1.10배 (일반적으로 약 10% 느림) |
| **읽기 작업 (async)** | 약 1.10배 (일반적으로 약 10% 느림) |
| **쓰기 작업 (batch)** | 약 0.90배 (일반적으로 약 10% 빠름) |
| **스트리밍 쓰기** | 1.21배 (21% 느림) |
| **버퍼 왕복** | 1.01배 (거의 동일) |

대부분의 실제 워크로드에서 Node.js 성능은 네이티브 Rust와 매우 유사합니다.

### 읽기 성능 비교

| 시나리오 | Rust | Node.js | 오버헤드 |
|---------|------|---------|----------|
| 대용량 데이터 (50k 행 × 20 열) | 616ms | 680ms | +10% |
| 많은 스타일 (5k 행, 서식 적용) | 33ms | 37ms | +12% |
| 다중 시트 (10개 시트 × 5k 행) | 360ms | 781ms | +117% |
| 수식 (10k 행) | 40ms | 52ms | +30% |
| 문자열 (20k 행 텍스트 중심) | 140ms | 126ms | -10% (더 빠름) |

### 쓰기 성능 비교

| 시나리오 | Rust | Node.js | 오버헤드 |
|---------|------|---------|----------|
| 50k 행 × 20 열 | 1.03s | 657ms | -36% (더 빠름) |
| 5k 스타일 적용 행 | 39ms | 48ms | +23% |
| 10k 행 (수식 포함) | 35ms | 39ms | +11% |
| 20k 텍스트 중심 행 | 145ms | 123ms | -15% (더 빠름) |

참고: 일부 쓰기 시나리오에서는 데이터 생성 시 V8의 효율적인 문자열 처리 덕분에 Node.js가 Rust보다 약간 더 빠른 성능을 보입니다.

### 확장성 성능

읽기 성능은 다양한 파일 크기에서 일관성을 유지합니다:

| 행 수 | Rust | Node.js | 오버헤드 |
|------|------|---------|----------|
| 1k | 6ms | 7ms | +17% |
| 10k | 62ms | 68ms | +10% |
| 100k | 659ms | 714ms | +8% |

쓰기 성능은 선형적으로 확장됩니다:

| 행 수 | Rust | Node.js | 오버헤드 |
|------|------|---------|----------|
| 1k | 7ms | 7ms | 0% |
| 10k | 68ms | 66ms | -3% (더 빠름) |
| 50k | 456ms | 332ms | -27% (더 빠름) |
| 100k | 735ms | 665ms | -10% (더 빠름) |

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

### 읽기 위주 워크로드

`OpenOptions`를 사용하여 필요한 부분만 로드합니다:

```typescript
const wb = await Workbook.open("huge.xlsx", {
  sheetRows: 1000,      // 시트당 첫 1000행만 읽기
  sheets: ["Sheet1"],   // Sheet1만 파싱
  maxUnzipSize: 100_000_000  // 압축 해제 크기 제한
});
```

### 쓰기 위주 워크로드

순차적 행 쓰기에는 `StreamWriter`를 사용합니다:

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

`OpenOptions`와 `StreamWriter`를 결합합니다:

```typescript
// 메타데이터만 읽기
const wb = await Workbook.open("input.xlsx", {
  sheetRows: 0  // 행을 파싱하지 않음
});

// 스트리밍으로 처리
const sw = wb.newStreamWriter("ProcessedData");
// ... 데이터 처리 ...
wb.applyStreamWriter(sw);
```

## 다음 단계

- [시작 가이드](./getting-started.md) - 기본 사항 학습
- [아키텍처](./architecture.md) - 내부 설계 이해
- [API 레퍼런스](./api-reference/) - 사용 가능한 모든 메서드 탐색
