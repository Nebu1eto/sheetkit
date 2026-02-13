# 마이그레이션 가이드: Async-First Lazy-Open

이 가이드에서는 async-first lazy-open 리팩터링에서 도입된 Breaking Change와 새로운 API를 다룹니다. 모든 변경 사항은 Node.js(`@sheetkit/node`) 패키지에 적용됩니다. Rust crate API는 영향을 받지 않습니다.

이 가이드는 `0.5.0` 이전 버전에서 `0.5.0` 이상 버전으로 마이그레이션할 때 적용됩니다.

## 변경 사항 요약

### Breaking Changes

1. **`ParseMode`가 `ReadMode`로 대체**: `full`/`readfast` 모드명이 `eager`/`lazy`/`stream`으로 변경되었습니다.
2. **`AuxParts` 정책**: 보조 파트(comments, charts, images)의 파싱 시점을 제어하는 새로운 옵션이 추가되었습니다.
3. **`OpenOptions` 재구성**: `readMode`, `auxParts`, `sheets`, `sheetRows`를 포함하는 새로운 타입 옵션 인터페이스가 적용되었습니다.
4. **`getRowsIterator`가 `openSheetReader`로 대체**: 비동기 iterator 프로토콜을 지원하는 스트림 기반 reader가 제공됩니다.
5. **기본 읽기 모드 변경**: `open()`의 기본값이 모든 시트를 즉시 파싱하는 대신 `readMode: 'lazy'`로 변경되었습니다.

### 새로운 API

1. `Workbook.openSheetReader(sheet, opts?)` -- `SheetStreamReader`를 반환합니다
2. `SheetStreamReader` 클래스 -- `for await...of` 지원
3. `getRowsBufferV2(sheet)` -- inline 문자열이 포함된 v2 raw buffer
4. `ReadMode` 타입: `'eager' | 'lazy' | 'stream'`
5. `AuxParts` 타입: `'eager' | 'deferred'`

## ReadMode 매핑 테이블

| 이전 (`parseMode`) | 이후 (`readMode`) | 동작 |
|-------------------|-----------------|------|
| `'full'` | `'eager'` | open 시 모든 시트와 보조 파트를 파싱합니다. 이전과 동일한 동작입니다. |
| `'readfast'` | `'lazy'` | ZIP 인덱스와 메타데이터만 파싱합니다. 시트 XML은 첫 접근 시(예: `getRows()`) 파싱됩니다. |
| (해당 없음) | `'stream'` | 최소 파싱 의도를 위한 모드입니다. 현재 Rust 코어 구현에서는 open 경로 동작이 `'lazy'`와 동일하며(시트 hydrate 지연, 보조 파트 eager 파싱 생략), `openSheetReader()`가 순방향 메모리 제한 반복을 제공합니다. |

### 이전 (v0.4)

```typescript
const wb = await Workbook.open('large.xlsx', {
  parseMode: 'readfast',
  sheetRows: 1000,
});
```

### 이후 (v0.5)

```typescript
const wb = await Workbook.open('large.xlsx', {
  readMode: 'lazy',
  sheetRows: 1000,
});
```

`parseMode: 'full'`을 사용하거나 `parseMode`를 지정하지 않았다면, `readMode: 'eager'`가 동일한 동작을 합니다. 그러나 open 지연 시간과 메모리 사용량 개선을 위해 `readMode: 'lazy'`로의 전환을 권장합니다. Lazy 모드는 첫 접근 시 자동으로 시트를 hydrate하므로, 대부분의 기존 코드가 추가 변경 없이 동작합니다.

## OpenOptions 변경 사항

새로운 `OpenOptions` 인터페이스:

```typescript
interface OpenOptions {
  readMode?: 'lazy' | 'stream' | 'eager';  // 기본값: 'lazy'
  auxParts?: 'deferred' | 'eager';          // 기본값: 'deferred'
  sheets?: string[];
  sheetRows?: number;
  maxUnzipSize?: number;
  maxZipEntries?: number;
}
```

이전 `JsOpenOptions`와의 주요 차이점:
- `parseMode`가 제거되었습니다. 대신 `readMode`를 사용합니다.
- `auxParts`가 새로 추가되었습니다. comments, charts, images 등 보조 파트를 open 시 파싱(`'eager'`)할지, 첫 접근 시 지연 로드(`'deferred'`)할지 제어합니다.
- 이전 `JsOpenOptions` 타입은 하위 호환성을 위해 여전히 허용되지만, `parseMode` 값은 새 코드 경로에서 인식되지 않습니다.

### AuxParts 정책

`auxParts`가 `'deferred'`(기본값)인 경우, comments, charts, images, pivot table 등 보조 파트는 `open()` 시 파싱되지 않습니다. 해당 파트가 필요한 메서드(예: `getComments()`, `getPivotTables()`)를 처음 호출할 때 로드됩니다.

이 방식은 셀 데이터만 필요한 워크로드에서 open 지연 시간과 메모리 사용량을 줄여줍니다.

```typescript
// Deferred (기본값): open 시 comments를 파싱하지 않음
const wb = await Workbook.open('report.xlsx');
// 첫 접근 시 comments가 로드됨:
const comments = wb.getComments('Sheet1');

// Eager: open 시 모든 파트를 파싱 (이전 동작)
const wb2 = await Workbook.open('report.xlsx', { auxParts: 'eager' });
```

## 스트리밍 Reader 마이그레이션

### 이전: `getRowsIterator()`

이전 `getRowsIterator()`는 전체 시트 buffer를 먼저 생성한 후 행을 하나씩 yield하는 동기 generator였습니다. 최대 JS 객체 수는 줄였지만, 최대 Rust 메모리는 줄이지 못했습니다.

```typescript
const wb = Workbook.openSync('large.xlsx');
for (const row of wb.getRowsIterator('Sheet1')) {
  process(row);
}
```

### 이후: `openSheetReader()`

새로운 `openSheetReader()`는 전체 시트를 메모리에 올리지 않고 ZIP 엔트리에서 직접 행을 배치 단위로 읽는 `SheetStreamReader`를 반환합니다. 메모리 사용량은 배치 크기로 제한됩니다.

```typescript
const wb = await Workbook.open('large.xlsx', { readMode: 'stream' });
const reader = await wb.openSheetReader('Sheet1', { batchSize: 500 });

for await (const batch of reader) {
  for (const row of batch) {
    process(row);
  }
}
```

주요 차이점:
- `openSheetReader()`는 비동기이며 `SheetStreamReader`를 반환합니다.
- Reader는 비동기 iterator 프로토콜(`for await...of`)을 구현합니다.
- 각 반복은 단일 행이 아닌 행 배열(배치)을 yield합니다.
- 메모리 사용량이 제한됩니다: 한 번에 하나의 배치만 메모리에 존재합니다.
- `reader.close()`를 호출하여 리소스를 일찍 해제하거나, `for await` 루프가 자동으로 처리하도록 할 수 있습니다.

### 수동 배치 제어

```typescript
const reader = await wb.openSheetReader('Sheet1');

while (true) {
  const batch = await reader.next(1000); // 호출당 커스텀 배치 크기
  if (batch === null) break;
  for (const row of batch) {
    process(row);
  }
}

await reader.close();
```

## Raw Buffer V2

`getRowsBufferV2()`는 별도의 글로벌 문자열 테이블 대신 문자열 데이터를 인라인으로 포함하는 v2 바이너리 buffer를 생성합니다. 이를 통해 모든 문자열을 즉시 생성하지 않고 행별로 점진적으로 디코딩할 수 있습니다.

v1 형식(`getRowsBuffer()`)은 여전히 사용 가능하며 완전히 지원됩니다.

```typescript
// V1 (변경 없음)
const bufV1 = wb.getRowsBuffer('Sheet1');

// V2 (새로운 형식): inline 문자열, 스트리밍 디코더에 적합
const bufV2 = wb.getRowsBufferV2('Sheet1');
```

커스텀 스트리밍 디코더를 구축하거나 행을 점진적으로 디코딩해야 할 때 v2를 사용합니다. 대부분의 사용 사례에서는 상위 수준 API(`getRows()`, `getRowsRaw()`, `SheetData`)가 형식 선택을 자동으로 처리합니다.

## Copy-on-Write 저장

`lazy` 모드로 열린 워크북은 copy-on-write 저장 파이프라인을 사용합니다. 시트를 수정하지 않고 lazy-open된 워크북을 저장하면, 변경되지 않은 시트 XML은 파싱 및 재직렬화 없이 원본 ZIP 엔트리에서 직접 기록됩니다. 이는 일부 시트만 수정하는 워크플로우에서 저장 지연 시간을 줄여줍니다.

```typescript
// 대용량 파일을 열고 하나의 시트만 수정 후 저장
const wb = await Workbook.open('100-sheets.xlsx');
wb.setCellValue('Sheet1', 'A1', 'Updated');
await wb.save('output.xlsx');
// Sheet1만 재직렬화됨. 나머지 99개 시트는 그대로 복사됨.
```

이 동작은 투명하게 처리되며 코드 변경이 필요하지 않습니다. 수정된 시트는 시트별 dirty flag를 통해 자동으로 추적됩니다.

## 전체 마이그레이션 체크리스트

1. `parseMode: 'full'`을 `readMode: 'eager'`로 변경합니다 (또는 기본 `'lazy'`를 사용하려면 제거).
2. `parseMode: 'readfast'`를 `readMode: 'lazy'`로 변경합니다.
3. `getRowsIterator()`를 `openSheetReader()`로 변경하고 비동기 반복으로 업데이트합니다.
4. open 직후 모든 보조 파트를 사용해야 한다면, open 옵션에 `auxParts: 'eager'`를 추가합니다.
5. 셀 접근 메서드가 여전히 동작하는지 테스트합니다 -- lazy 모드는 시트를 자동으로 hydrate합니다.
6. 랜덤 셀 접근이 필요 없는 대용량 파일의 경우 `readMode: 'stream'`과 `openSheetReader()` 사용을 고려합니다.

## 호환성 참고 사항

- `sheetRows`, `sheets`, `maxUnzipSize`, `maxZipEntries` 필드가 있는 이전 `JsOpenOptions` 타입은 모든 open 메서드에서 여전히 허용됩니다. `parseMode`만 인식되지 않습니다.
- `getRowsIterator()`는 여전히 존재하며 동작합니다. 제거되지 않았지만, 진정한 메모리 제한 스트리밍을 제공하는 `openSheetReader()`가 새 코드에 권장됩니다.
- `getRowsBuffer()` (v1)은 deprecated되지 않습니다. V1과 V2 형식이 공존합니다.
- 모든 동기 메서드(`openSync`, `saveSync`, `openBufferSync`, `writeBufferSync`)는 변경 없이 계속 동작합니다.
