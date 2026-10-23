import { isProxy } from 'node:util/types';
import type {
  JsAppProperties,
  JsCellEntry,
  JsChartConfig,
  JsColData,
  JsCommentConfig,
  JsConditionalFormatEntry,
  JsConditionalFormatRule,
  JsCryptLimits,
  JsDataValidationConfig,
  JsDefinedNameConfig,
  JsDefinedNameInfo,
  JsDocProperties,
  JsFormControlConfig,
  JsFormControlInfo,
  JsHeaderFooter,
  JsHyperlinkInfo,
  JsHyperlinkOptions,
  JsImageConfig,
  JsOpenOptions,
  JsPageMargins,
  JsPageSetup,
  JsPersonData,
  JsPersonInput,
  JsPictureInfo,
  JsPivotTableConfig,
  JsPivotTableInfo,
  JsPrintOptions,
  JsRenderOptions,
  JsRichTextRun,
  JsRowData,
  JsSaveLimits,
  JsShapeConfig,
  JsSheetProtectionConfig,
  JsSheetViewOptions,
  JsSlicerConfig,
  JsSlicerInfo,
  JsSparklineConfig,
  JsStreamLimits,
  JsStyle,
  JsTableConfig,
  JsTableInfo,
  JsThreadedCommentData,
  JsThreadedCommentInput,
  JsVbaLimits,
  JsVbaProject,
  JsWorkbookProtectionConfig,
} from './binding.js';
import {
  JsStreamWriter,
  type NativeSheetStreamReader,
  Workbook as NativeWorkbook,
} from './binding.js';
import type { RawRowsResult } from './buffer-codec.js';
import { decodeRowsBuffer, decodeRowsIterator, decodeRowsRawBuffer } from './buffer-codec.js';
import type { CellTypeName, CellValue as SheetDataCellValue } from './sheet-data.js';
import { SheetData } from './sheet-data.js';

export type { DateValue, OccupiedCell, OccupiedRow } from './sheet-data.js';

import type { DateValue } from './sheet-data.js';

export type {
  JsAlignmentStyle,
  JsAppProperties,
  JsBorderSideStyle,
  JsBorderStyle,
  JsCellEntry,
  JsChartConfig,
  JsChartSeries,
  JsColCell,
  JsColData,
  JsCommentConfig,
  JsConditionalFormatEntry,
  JsConditionalFormatRule,
  JsConditionalStyle,
  JsCryptLimits,
  JsDataValidationConfig,
  JsDefinedNameConfig,
  JsDefinedNameInfo,
  JsDocProperties,
  JsFillStyle,
  JsFontStyle,
  JsFormControlConfig,
  JsFormControlInfo,
  JsFormulaLimits,
  JsHeaderFooter,
  JsHyperlinkInfo,
  JsHyperlinkOptions,
  JsImageConfig,
  JsOpenOptions,
  JsPageMargins,
  JsPageSetup,
  JsPersonData,
  JsPersonInput,
  JsPictureInfo,
  JsPivotDataField,
  JsPivotField,
  JsPivotTableConfig,
  JsPivotTableInfo,
  JsPrintOptions,
  JsProtectionStyle,
  JsRenderOptions,
  JsRichTextRun,
  JsRowCell,
  JsRowData,
  JsSaveLimits,
  JsShapeConfig,
  JsSheetProtectionConfig,
  JsSheetViewOptions,
  JsSlicerConfig,
  JsSlicerInfo,
  JsSparklineConfig,
  JsStreamLimits,
  JsStyle,
  JsTableColumn,
  JsTableConfig,
  JsTableInfo,
  JsThreadedCommentData,
  JsThreadedCommentInput,
  JsVbaLimits,
  JsVbaModule,
  JsVbaProject,
  JsView3DConfig,
  JsWorkbookProtectionConfig,
} from './binding.js';

/** Read mode controls how much of the .xlsx file is parsed during open. */
export type ReadMode = 'lazy' | 'stream' | 'eager';

/**
 * Auxiliary parts loading policy.
 * - "deferred": auxiliary parts (comments, charts, images, etc.) are not parsed during open.
 * - "eager": all parts are parsed during open.
 */
export type AuxParts = 'deferred' | 'eager';

/**
 * Controls how date-formatted number cells are interpreted while reading.
 * - "numFmt" (default): promote `t="n"` (or untyped) cells whose style
 *   points at a built-in date format ID (14-22, 45-47) or a custom format
 *   code containing date/time tokens to a date cell. Matches how Microsoft
 *   Excel stores dates in practice.
 * - "cellType": follow the cell's `t` attribute strictly. A `t="n"` cell
 *   remains a number even when its style references a date number format;
 *   only `t="d"` cells are returned as dates. Opt in when you need
 *   spec-literal behavior.
 */
export type DateInterpretation = 'cellType' | 'numFmt';

/** Worksheet traversal layout. Dense traversal is available only for v1 buffers. */
export type TraversalMode = 'sparse' | 'dense';

/** Finite resource limits for one worksheet traversal. */
export interface TraversalLimits {
  /** Maximum occupied rows in sparse mode or spanned rows in dense mode. */
  maxRows?: number;
  /** Maximum occupied or spanned columns. */
  maxColumns?: number;
  /** Maximum occupied cells in sparse mode or spanned cells in dense mode. */
  maxCells?: number;
  /** Maximum binary output size in bytes. */
  maxOutputBytes?: number;
}

/** Options shared by worksheet row traversal and export methods. */
export interface TraversalOptions {
  /** Sparse by default. Dense is an explicit, finite v1 compatibility mode. */
  mode?: TraversalMode;
  /** Per-operation finite limits. Omitted fields use secure finite defaults. */
  limits?: TraversalLimits;
}

/** Finite limits for encrypted Office containers and password derivation. */
export interface CryptLimits {
  outerCfbBytes?: number;
  namedStreamCount?: number;
  namedStreamDeclaredBytes?: number;
  namedStreamReadBytes?: number;
  encryptedPackageBytes?: number;
  kdfIterations?: number;
  kdfWork?: number;
}

/** Finite limits for one workbook save operation. */
export interface SaveLimits {
  outputBytes?: number;
  xmlPartBytes?: number;
  rowCount?: number;
  cellCount?: number;
}

/** Finite limits for one streaming worksheet writer. */
export interface StreamLimits {
  rowCount?: number;
  cellCount?: number;
  rowXmlBytes?: number;
  outputBytes?: number;
  configCount?: number;
  configBytes?: number;
}

/** Finite limits for VBA project loading and module extraction. */
export interface VbaLimits {
  inputBytes?: number;
  cfbStreamCount?: number;
  cfbStreamDeclaredBytes?: number;
  cfbAggregateDeclaredBytes?: number;
  cfbStreamReadBytes?: number;
  cfbAggregateReadBytes?: number;
  compressedBytes?: number;
  moduleCount?: number;
  chunkCount?: number;
  moduleDecodedBytes?: number;
  aggregateDecodedBytes?: number;
  retainedBytes?: number;
  outputBytes?: number;
}

/** Finite limits for one formula parse, mutation, or calculation operation. */
export interface FormulaLimits {
  perFormulaInputBytes?: number;
  perFormulaNestingDepth?: number;
  perFormulaBinaryChainDepth?: number;
  perFormulaTokens?: number;
  perFormulaAstNodes?: number;
  operationInputBytes?: number;
  operationTokens?: number;
  operationAstNodes?: number;
  operationRangeCells?: number;
  operationDependencyEdges?: number;
  operationReadyQueueOps?: number;
  operationEvaluationSteps?: number;
  operationMatchWork?: number;
  operationOutputBytes?: number;
}

/** Options for controlling how a workbook is opened and parsed. */
export interface OpenOptions {
  /** Maximum bytes accepted from the outer ZIP or encrypted container. */
  maxOuterInputBytes?: number;
  /** Maximum bytes read from one XML part, including lazy hydration. */
  maxXmlPartBytes?: number;
  /**
   * Decompressed entry size at or below which the compression ratio is not
   * enforced. Default: 8 MiB. `0` enforces the ratio for every entry.
   */
  compressionRatioFloorBytes?: number;
  /** Read mode: "lazy" (recommended for read-only), "eager", or "stream". Default: "lazy". */
  readMode?: ReadMode;
  /** Only parse sheets whose names are in this list. Omit to parse all sheets. */
  sheets?: string[];
  /** Maximum rows read per sheet. Omit to use the worksheet format maximum. */
  sheetRows?: number;
  /** Maximum total decompressed size of all ZIP entries in bytes. */
  maxUnzipSize?: number;
  /** Maximum number of ZIP entries allowed. */
  maxZipEntries?: number;
  /** Finite encrypted-container and password-derivation limits. */
  cryptLimits?: CryptLimits | JsCryptLimits;
  /** Finite VBA project loading limits. */
  vbaLimits?: VbaLimits | JsVbaLimits;
  /** Auxiliary parts loading policy. Default: "deferred". */
  auxParts?: AuxParts;
  /**
   * How to interpret date-formatted number cells while reading. Default:
   * "numFmt" (promote `t="n"` cells with a date number format into date
   * cells). Set to "cellType" for spec-literal behavior where only
   * `t="d"` cells become dates.
   */
  dateInterpretation?: DateInterpretation;
}

const RESOURCE_OPERATIONS = [
  'workbook_open',
  'workbook_save',
  'formula_parse',
  'formula_mutate',
  'formula_calculate',
  'data_traverse',
  'render',
  'export',
  'encrypt',
  'decrypt',
  'macro_read',
  'stream_read',
  'stream_write',
] as const;

const RESOURCE_CATEGORIES = [
  'archive',
  'container',
  'xml',
  'formula',
  'traversal',
  'output',
  'encryption',
  'macro',
  'streaming',
  'structure',
] as const;

const RESOURCE_DIMENSIONS = [
  'input_bytes',
  'output_bytes',
  'retained_bytes',
  'entry_count',
  'entry_declared_bytes',
  'entry_actual_bytes',
  'aggregate_declared_bytes',
  'aggregate_actual_bytes',
  'compression_ratio',
  'compression_ratio_floor_bytes',
  'xml_bytes',
  'nesting_depth',
  'token_count',
  'node_count',
  'range_cell_count',
  'dependency_edge_count',
  'operation_count',
  'match_work',
  'row_count',
  'column_count',
  'cell_count',
  'stream_count',
  'module_count',
  'chunk_count',
  'iteration_count',
  'collection_count',
] as const;

const RESOURCE_STAGES = [
  'inspect',
  'read',
  'parse',
  'decode',
  'validate',
  'calculate',
  'render',
  'serialize',
  'write',
  'replace',
] as const;

/** Public operation that exhausted a finite resource budget. */
export type ResourceOperation = (typeof RESOURCE_OPERATIONS)[number];

/** Subsystem whose finite resource budget was exhausted. */
export type ResourceCategory = (typeof RESOURCE_CATEGORIES)[number];

/** Independently bounded resource dimension that was exhausted. */
export type ResourceDimension = (typeof RESOURCE_DIMENSIONS)[number];

/** Processing stage at which resource exhaustion was detected. */
export type ResourceStage = (typeof RESOURCE_STAGES)[number];

/** Stable machine-readable contract for bounded resource failures. */
export interface ResourceLimitError extends Error {
  readonly code: 'resource_limit_exceeded';
  readonly operation: ResourceOperation;
  readonly category: ResourceCategory;
  readonly dimension: ResourceDimension;
  readonly limit: bigint;
  readonly observed: bigint;
  readonly stage: ResourceStage | null;
}

function createResourceLimitError(
  message: string,
  metadata: Omit<ResourceLimitError, keyof Error>,
): ResourceLimitError {
  return Object.assign(new Error(message), metadata);
}

function isClosedCode<const Values extends readonly string[]>(
  values: Values,
  value: unknown,
): value is Values[number] {
  return typeof value === 'string' && (values as readonly string[]).includes(value);
}

/** Returns whether a thrown value satisfies the stable resource-error contract. */
export function isResourceLimitError(error: unknown): error is ResourceLimitError {
  if (typeof error !== 'object' || error === null) return false;

  try {
    const candidate = error as Record<string, unknown>;
    const { limit, observed, stage } = candidate;
    const maxU64 = 18_446_744_073_709_551_615n;
    return (
      candidate.code === 'resource_limit_exceeded' &&
      isClosedCode(RESOURCE_OPERATIONS, candidate.operation) &&
      isClosedCode(RESOURCE_CATEGORIES, candidate.category) &&
      isClosedCode(RESOURCE_DIMENSIONS, candidate.dimension) &&
      typeof limit === 'bigint' &&
      limit >= 0n &&
      limit <= maxU64 &&
      typeof observed === 'bigint' &&
      observed > limit &&
      observed <= maxU64 &&
      (stage === null || isClosedCode(RESOURCE_STAGES, stage))
    );
  } catch {
    return false;
  }
}

/** Cached output associated with a formula expression. */
export interface FormulaResultValue {
  valueType: 'empty' | 'string' | 'number' | 'boolean' | 'date' | 'error' | 'formula';
  value?: string;
  numberValue?: number;
  boolValue?: boolean;
  date?: DateValue;
}

/** A formula expression with the cached result read from the workbook. */
export interface FormulaValue {
  type: 'formula';
  formula: string;
  result?: FormulaResultValue;
}

/** A cell's public value representation. */
export type CellValue = string | number | boolean | DateValue | FormulaValue | null;

type CellValueInput = string | number | boolean | DateValue | FormulaValue | null;
type NativeCellValueInput = Parameters<NativeWorkbook['setCellValue']>[2];

function toNativeCellValue(value: CellValueInput): NativeCellValueInput {
  if (typeof value !== 'object' || value === null) {
    return value as NativeCellValueInput;
  }
  if ('type' in value && value.type === 'date') {
    return { ...value, iso: value.iso ?? undefined } as NativeCellValueInput;
  }
  if ('type' in value && value.type === 'formula' && value.result?.date) {
    return {
      ...value,
      result: {
        ...value.result,
        date: { ...value.result.date, iso: value.result.date.iso ?? undefined },
      },
    } as NativeCellValueInput;
  }
  return value as NativeCellValueInput;
}

export interface ToJsonOptions extends TraversalOptions {
  /** Use first row as keys (true), or provide custom key names. */
  header?: boolean | string[];
}

/** Policy for text that spreadsheet applications may interpret as a formula. */
export type CsvFormulaPolicy = 'spreadsheet-safe' | 'raw-exchange';

/** Maximum UTF-8 bytes returned by one CSV serialization. */
export const MAX_CSV_OUTPUT_BYTES = 64 * 1024 * 1024;

const CSV_TRANSFER_BASE_BYTES = 64;
const CSV_TRANSFER_OVERHEAD_FACTOR = 21;
const DEFAULT_CSV_TRANSFER_BYTES = 512 * 1024 * 1024;

export interface ToCsvOptions extends TraversalOptions {
  /** Traversal limits; maxOutputBytes bounds the returned UTF-8 CSV. */
  limits?: TraversalLimits;
  /** Field delimiter. Default: "," */
  delimiter?: string;
  /** Quote character. Default: '"' */
  quote?: string;
  /** Line ending. Default: "\n" */
  lineEnding?: string;
  /** Formula-capable text policy. Default: "spreadsheet-safe". */
  formulaPolicy?: CsvFormulaPolicy;
}

/** Default traversal output ceiling, matching the native traversal default. */
const DEFAULT_TRAVERSAL_OUTPUT_BYTES = 512 * 1024 * 1024;

export interface ToHtmlOptions extends TraversalOptions {
  /**
   * Traversal limits; maxOutputBytes bounds both the native transfer buffer
   * and the returned UTF-8 HTML. Default: 512 MiB.
   */
  limits?: TraversalLimits;
  /** CSS class name for the table element. */
  className?: string;
}

export interface FromJsonOptions {
  /** Write column headers as the first row (true), or provide custom headers. */
  header?: boolean | string[];
  /** Top-left cell to start writing. Default: "A1" */
  startCell?: string;
  /** Per-operation finite import limits. Omitted fields use secure defaults. */
  limits?: FromJsonLimits;
}

/** Finite limits for one `fromJSON` import operation. */
export interface FromJsonLimits {
  /** Maximum number of data and header rows written. */
  maxRows?: number;
  /** Maximum number of distinct imported columns. */
  maxColumns?: number;
  /** Maximum number of coordinate-bearing cells staged for one import. */
  maxCells?: number;
  /** Maximum logical bytes retained by the import staging representation. */
  maxRetainedBytes?: number;
}

const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLUMNS = 16_384;
const DEFAULT_FROM_JSON_MAX_CELLS = 10_000_000;
const HARD_FROM_JSON_MAX_CELLS = 100_000_000;
const DEFAULT_FROM_JSON_MAX_RETAINED_BYTES = 512 * 1_024 * 1_024;
const HARD_FROM_JSON_MAX_RETAINED_BYTES = 2 * 1_024 * 1_024 * 1_024;
const FROM_JSON_KEY_OVERHEAD_BYTES = 16;
const FROM_JSON_ENTRY_OVERHEAD_BYTES = 64;
const MAX_FROM_JSON_CELL_STRING_BYTES = 32_767;
const MAX_FROM_JSON_DATE_ISO_BYTES = 64;

interface ResolvedFromJsonLimits {
  maxRows: number;
  maxColumns: number;
  maxCells: number;
  maxRetainedBytes: number;
}

function fromJsonLimitError(
  dimension: 'row_count' | 'column_count' | 'cell_count' | 'retained_bytes',
  limit: number,
  observed: number,
): ResourceLimitError {
  return createResourceLimitError('JSON import resource limit exceeded', {
    code: 'resource_limit_exceeded',
    operation: 'data_traverse',
    category: 'structure',
    dimension,
    limit: BigInt(limit),
    observed: BigInt(Math.max(observed, limit + 1)),
    stage: 'validate',
  });
}

function resolveFromJsonLimit(
  value: number | undefined,
  fallback: number,
  hardLimit: number,
  dimension: 'row_count' | 'column_count' | 'cell_count' | 'retained_bytes',
): number {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new TypeError('JSON import limits must be non-negative safe integers');
  }
  if (value === 0) throw fromJsonLimitError(dimension, 0, 1);
  if (value > hardLimit) throw fromJsonLimitError(dimension, hardLimit, value);
  return value;
}

function resolveFromJsonLimits(limits: FromJsonLimits | undefined): ResolvedFromJsonLimits {
  if (
    limits !== undefined &&
    (typeof limits !== 'object' || limits === null || Array.isArray(limits))
  ) {
    throw new TypeError('JSON import limits must be an object');
  }
  return {
    maxRows: resolveFromJsonLimit(limits?.maxRows, EXCEL_MAX_ROWS, EXCEL_MAX_ROWS, 'row_count'),
    maxColumns: resolveFromJsonLimit(
      limits?.maxColumns,
      EXCEL_MAX_COLUMNS,
      EXCEL_MAX_COLUMNS,
      'column_count',
    ),
    maxCells: resolveFromJsonLimit(
      limits?.maxCells,
      DEFAULT_FROM_JSON_MAX_CELLS,
      HARD_FROM_JSON_MAX_CELLS,
      'cell_count',
    ),
    maxRetainedBytes: resolveFromJsonLimit(
      limits?.maxRetainedBytes,
      DEFAULT_FROM_JSON_MAX_RETAINED_BYTES,
      HARD_FROM_JSON_MAX_RETAINED_BYTES,
      'retained_bytes',
    ),
  };
}

function checkFromJsonLimit(
  dimension: 'row_count' | 'column_count' | 'cell_count' | 'retained_bytes',
  limit: number,
  observed: number,
): void {
  if (observed > limit) throw fromJsonLimitError(dimension, limit, observed);
}

function checkedFromJsonAdd(
  current: number,
  additional: number,
  limit: number,
  dimension: 'cell_count' | 'retained_bytes',
): number {
  const observed = current + additional;
  if (!Number.isSafeInteger(observed)) {
    throw fromJsonLimitError(dimension, limit, limit + 1);
  }
  checkFromJsonLimit(dimension, limit, observed);
  return observed;
}

interface OwnDataProperty {
  present: boolean;
  value: unknown;
}

interface JsonCellSnapshot {
  value: JsCellEntry['value'];
  retainedBytes: number;
}

function ownDataProperty(value: Record<string, unknown>, key: string): OwnDataProperty {
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  if (descriptor === undefined) return { present: false, value: undefined };
  if (!('value' in descriptor)) throw new TypeError('JSON values must use data properties');
  return { present: true, value: descriptor.value };
}

function snapshotDateValue(value: unknown, knownKind?: 'date'): JsonCellSnapshot {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('Invalid JSON cell value');
  }
  const candidate = value as Record<string, unknown>;
  const kind = knownKind ?? ownDataProperty(candidate, 'type').value;
  const serial = ownDataProperty(candidate, 'serial');
  const iso = ownDataProperty(candidate, 'iso');
  if (
    kind !== 'date' ||
    typeof serial.value !== 'number' ||
    !Number.isFinite(serial.value) ||
    (iso.value !== null && iso.value !== undefined && typeof iso.value !== 'string')
  ) {
    throw new TypeError('Invalid JSON cell value');
  }
  const isoBytes = typeof iso.value === 'string' ? Buffer.byteLength(iso.value) : 0;
  if (isoBytes > MAX_FROM_JSON_DATE_ISO_BYTES) {
    throw new TypeError('Invalid JSON cell value');
  }
  return {
    value: Object.freeze({
      type: 'date',
      serial: serial.value,
      iso: typeof iso.value === 'string' ? iso.value : undefined,
    }) as JsCellEntry['value'],
    retainedBytes: 16 + isoBytes,
  };
}

function snapshotFormulaResult(value: unknown): {
  value: Record<string, unknown>;
  retainedBytes: number;
} {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('Invalid JSON formula result');
  }
  const candidate = value as Record<string, unknown>;
  const valueType = ownDataProperty(candidate, 'valueType');
  const stringValue = ownDataProperty(candidate, 'value');
  const numberValue = ownDataProperty(candidate, 'numberValue');
  const boolValue = ownDataProperty(candidate, 'boolValue');
  const date = ownDataProperty(candidate, 'date');
  const hasValue = stringValue.present && stringValue.value !== undefined;
  const hasNumber = numberValue.present && numberValue.value !== undefined;
  const hasBoolean = boolValue.present && boolValue.value !== undefined;
  const hasDate = date.present && date.value !== undefined;

  switch (valueType.value) {
    case 'empty':
      if (hasValue || hasNumber || hasBoolean || hasDate) {
        throw new TypeError('Invalid JSON formula result');
      }
      return { value: Object.freeze({ valueType: 'empty' }), retainedBytes: 0 };
    case 'boolean':
      if (typeof boolValue.value !== 'boolean' || hasValue || hasNumber || hasDate) {
        throw new TypeError('Invalid JSON formula result');
      }
      return {
        value: Object.freeze({ valueType: 'boolean', boolValue: boolValue.value }),
        retainedBytes: 1,
      };
    case 'number':
      if (
        typeof numberValue.value !== 'number' ||
        !Number.isFinite(numberValue.value) ||
        hasValue ||
        hasBoolean ||
        hasDate
      ) {
        throw new TypeError('Invalid JSON formula result');
      }
      return {
        value: Object.freeze({ valueType: 'number', numberValue: numberValue.value }),
        retainedBytes: 8,
      };
    case 'date': {
      if (hasValue || hasNumber || hasBoolean || !hasDate) {
        throw new TypeError('Invalid JSON formula result');
      }
      const snapshot = snapshotDateValue(date.value);
      return {
        value: Object.freeze({ valueType: 'date', date: snapshot.value }),
        retainedBytes: snapshot.retainedBytes,
      };
    }
    case 'error':
    case 'string': {
      if (typeof stringValue.value !== 'string' || hasNumber || hasBoolean || hasDate) {
        throw new TypeError('Invalid JSON formula result');
      }
      const bytes = Buffer.byteLength(stringValue.value);
      if (bytes > MAX_FROM_JSON_CELL_STRING_BYTES) {
        throw new TypeError('JSON cell string exceeds the Excel limit');
      }
      return {
        value: Object.freeze({ valueType: valueType.value, value: stringValue.value }),
        retainedBytes: bytes,
      };
    }
    default:
      throw new TypeError('Invalid JSON formula result');
  }
}

function snapshotJsonCellValue(value: unknown): JsonCellSnapshot {
  if (value === undefined || value === null) return { value: null, retainedBytes: 0 };
  if (typeof value === 'string') {
    const bytes = Buffer.byteLength(value);
    if (bytes > MAX_FROM_JSON_CELL_STRING_BYTES) {
      throw new TypeError('JSON cell string exceeds the Excel limit');
    }
    return { value, retainedBytes: bytes };
  }
  if (typeof value === 'boolean') return { value, retainedBytes: 1 };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError('Invalid JSON cell value');
    return { value, retainedBytes: 8 };
  }
  if (typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Invalid JSON cell value');
  }

  const candidate = value as Record<string, unknown>;
  const kind = ownDataProperty(candidate, 'type');
  if (kind.value === 'date') return snapshotDateValue(candidate, 'date');
  if (kind.value !== 'formula') throw new TypeError('Invalid JSON cell value');

  const formula = ownDataProperty(candidate, 'formula');
  const result = ownDataProperty(candidate, 'result');
  if (typeof formula.value !== 'string') throw new TypeError('Invalid JSON cell value');
  const resultSnapshot = result.value === undefined ? null : snapshotFormulaResult(result.value);
  return {
    value: Object.freeze({
      type: 'formula',
      formula: formula.value,
      result: resultSnapshot?.value,
    }) as JsCellEntry['value'],
    retainedBytes: Buffer.byteLength(formula.value) + (resultSnapshot?.retainedBytes ?? 0),
  };
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function* jsonRecordKeys(record: Record<string, unknown>): Generator<string> {
  if (isProxy(record)) {
    throw new TypeError('JSON import rows must be plain objects');
  }
  const prototype = Object.getPrototypeOf(record);
  if (prototype !== null && prototype !== Object.prototype) {
    throw new TypeError('JSON import rows must be plain objects');
  }
  if (prototype === Object.prototype) {
    for (const _key in Object.prototype) {
      throw new TypeError('JSON import rows require an unmodified object prototype');
    }
  }
  for (const key in record) {
    yield key;
  }
}

function jsonRecordValue(record: Record<string, unknown>, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(record, key);
  if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
    throw new TypeError('JSON import rows must use enumerable data properties');
  }
  return descriptor.value;
}

function jsonArrayLength(value: unknown[]): number {
  const descriptor = Object.getOwnPropertyDescriptor(value, 'length');
  if (
    descriptor === undefined ||
    !('value' in descriptor) ||
    !Number.isSafeInteger(descriptor.value) ||
    descriptor.value < 0
  ) {
    throw new TypeError('Invalid JSON array');
  }
  return descriptor.value;
}

function jsonArrayElement(value: unknown[], index: number): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
  if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
    throw new TypeError('JSON arrays must use dense data properties');
  }
  return descriptor.value;
}

function cellValueToString(v: SheetDataCellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.type === 'date') return v.iso ?? String(v.serial);
  return String(v);
}

function isFormulaLeadingCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0);
  return (
    character === '\uFEFF' ||
    /\s/u.test(character) ||
    (codePoint !== undefined && (codePoint <= 0x1f || (codePoint >= 0x7f && codePoint <= 0x9f)))
  );
}

function isFormulaCapableText(value: string): boolean {
  for (const character of value) {
    if (isFormulaLeadingCharacter(character)) continue;
    return '=+-@'.includes(character);
  }
  return false;
}

function csvFormulaPolicy(value: CsvFormulaPolicy | undefined): CsvFormulaPolicy {
  const policy = value ?? 'spreadsheet-safe';
  if (policy !== 'spreadsheet-safe' && policy !== 'raw-exchange') {
    throw new TypeError('formulaPolicy must be "spreadsheet-safe" or "raw-exchange"');
  }
  return policy;
}

function validateCsvSyntax(
  delimiter: string,
  quote: string,
  lineEnding: string,
  formulaPolicy: CsvFormulaPolicy,
): void {
  if ([...delimiter].length !== 1 || [...quote].length !== 1) {
    throw new TypeError('delimiter and quote must each be one character');
  }
  const isControl = (value: string): boolean => {
    const codePoint = value.codePointAt(0);
    return (
      codePoint === undefined ||
      codePoint <= 0x1f ||
      codePoint === 0x7f ||
      (codePoint >= 0x80 && codePoint <= 0x9f)
    );
  };
  if (delimiter === quote || isControl(delimiter) || isControl(quote)) {
    throw new TypeError('delimiter and quote must be distinct non-control characters');
  }
  if (lineEnding !== '\n' && lineEnding !== '\r\n') {
    throw new TypeError('lineEnding must be "\\n" or "\\r\\n"');
  }
  if (formulaPolicy === 'spreadsheet-safe' && [delimiter, quote].some(isFormulaCapableText)) {
    throw new TypeError('CSV syntax must not introduce a spreadsheet formula');
  }
}

function csvOutputLimit(options: ToCsvOptions | undefined): number {
  const requested = options?.limits?.maxOutputBytes;
  if (requested === undefined) return MAX_CSV_OUTPUT_BYTES;
  if (!Number.isSafeInteger(requested) || requested <= 0) {
    throw new TypeError('maxOutputBytes must be a positive safe integer');
  }
  return Math.min(requested, MAX_CSV_OUTPUT_BYTES);
}

function csvTraversalOptions(
  options: ToCsvOptions | undefined,
  outputLimit: number,
): TraversalOptions {
  const limits = options?.limits;
  // With R rows, C cells, and F field bytes, sparse v2 is at most
  // 16 + 10R + 10C + F. CSV has R <= C and at least F + C - 1
  // bytes, so 64 + 21 * CSV is a conservative sparse admission bound.
  // Dense v1 may also carry the workbook SST and fails closed at this
  // derived limit. Retain the library's finite default ceiling overall.
  const derivedTransferCandidate =
    CSV_TRANSFER_BASE_BYTES + outputLimit * CSV_TRANSFER_OVERHEAD_FACTOR;
  if (!Number.isSafeInteger(derivedTransferCandidate)) {
    throw new TypeError('derived CSV transfer limit must be a safe integer');
  }
  const derivedTransferLimit = Math.min(DEFAULT_CSV_TRANSFER_BYTES, derivedTransferCandidate);
  return {
    mode: options?.mode,
    limits: {
      maxRows: limits?.maxRows,
      maxColumns: limits?.maxColumns,
      maxCells: limits?.maxCells,
      maxOutputBytes: derivedTransferLimit,
    },
  };
}

function csvOutputLimitError(limit: number, observed: number): ResourceLimitError {
  return createResourceLimitError('resource limit exceeded during CSV serialization', {
    code: 'resource_limit_exceeded' as const,
    operation: 'export' as const,
    category: 'output' as const,
    dimension: 'output_bytes' as const,
    limit: BigInt(limit),
    observed: BigInt(observed),
    stage: 'serialize' as const,
  });
}

const HTML_TABLE_CLOSE = '</table>';

/** Add `bytes` to `observed`; throw the HTML output limit error when the total exceeds `limit`. */
function chargeHtmlOutput(observed: number, bytes: number, limit: number): number {
  const next = observed + bytes;
  if (next > limit || !Number.isSafeInteger(next)) {
    throw htmlOutputLimitError(limit, Number.isSafeInteger(next) ? next : limit + 1);
  }
  return next;
}

function htmlOutputLimitError(limit: number, observed: number): ResourceLimitError {
  return createResourceLimitError('resource limit exceeded during HTML serialization', {
    code: 'resource_limit_exceeded' as const,
    operation: 'export' as const,
    category: 'output' as const,
    dimension: 'output_bytes' as const,
    limit: BigInt(limit),
    observed: BigInt(observed),
    stage: 'serialize' as const,
  });
}

interface CsvCell {
  value: SheetDataCellValue;
  formula?: boolean;
  cellType?: CellTypeName;
}

function countOccurrences(value: string, needle: string): number {
  let count = 0;
  let offset = 0;
  while (true) {
    const index = value.indexOf(needle, offset);
    if (index === -1) return count;
    count++;
    offset = index + needle.length;
  }
}

function csvField(
  row: CsvCell[],
  column: number,
  leadingEmptyFields: number,
  formulaPolicy: CsvFormulaPolicy,
): string {
  const rowColumn = column - leadingEmptyFields;
  const cell = rowColumn >= 0 && rowColumn < row.length ? row[rowColumn] : undefined;
  const value = cell?.value ?? null;
  const isFormula = cell?.formula === true || cell?.cellType === 'formula';
  const field = isFormula ? `=${cellValueToString(value)}` : cellValueToString(value);
  return formulaPolicy === 'spreadsheet-safe' &&
    (isFormula || typeof value === 'string') &&
    isFormulaCapableText(field)
    ? `'${field}`
    : field;
}

function encodeCsvOutput(
  rows: CsvCell[][],
  leadingEmptyFields: number,
  maxColumns: number | null,
  delimiter: string,
  quote: string,
  lineEnding: string,
  formulaPolicy: CsvFormulaPolicy,
  limit: number,
): string {
  const delimiterBytes = Buffer.byteLength(delimiter);
  const quoteBytes = Buffer.byteLength(quote);
  const lineEndingBytes = Buffer.byteLength(lineEnding);
  const quotePattern = new RegExp(escapeRegExp(quote), 'g');
  const lines: string[] = [];
  let observed = 0;
  const charge = (bytes: number): void => {
    const next = observed + bytes;
    if (!Number.isSafeInteger(next) || next > limit) {
      throw csvOutputLimitError(limit, Number.isSafeInteger(next) ? next : limit + 1);
    }
    observed = next;
  };

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    if (rowIndex > 0) charge(lineEndingBytes);
    const row = rows[rowIndex];
    const fieldCount = maxColumns === null ? row.length : leadingEmptyFields + maxColumns;
    const escapedFields: string[] = [];
    for (let column = 0; column < fieldCount; column++) {
      if (column > 0) charge(delimiterBytes);
      const field = csvField(row, column, leadingEmptyFields, formulaPolicy);

      const needsQuotes =
        field.includes(delimiter) ||
        field.includes(quote) ||
        field.includes('\n') ||
        field.includes('\r');
      const escapedBytes =
        Buffer.byteLength(field) +
        (needsQuotes ? (countOccurrences(field, quote) + 2) * quoteBytes : 0);
      charge(escapedBytes);
      const escapedField = needsQuotes
        ? quote + field.replace(quotePattern, quote + quote) + quote
        : field;
      escapedFields.push(escapedField);
    }
    lines.push(escapedFields.join(delimiter));
  }
  return lines.join(lineEnding);
}

function createJsonOutputObject(): Record<string, CellValue> {
  return Object.create(null) as Record<string, CellValue>;
}

function finishJsonOutputObject(object: Record<string, CellValue>): Record<string, CellValue> {
  Object.setPrototypeOf(object, Object.prototype);
  return object;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseCellRef(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    throw new TypeError('Invalid cell reference');
  }
  const letters = match[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
    if (col > EXCEL_MAX_COLUMNS) {
      throw fromJsonLimitError('column_count', EXCEL_MAX_COLUMNS, col);
    }
  }

  let row = 0;
  for (const digit of match[2]) {
    row = row * 10 + (digit.charCodeAt(0) - 48);
    if (row > EXCEL_MAX_ROWS) {
      throw fromJsonLimitError('row_count', EXCEL_MAX_ROWS, row);
    }
  }
  if (row === 0) throw new TypeError('Invalid cell reference');
  return { col, row };
}

function columnNumberToLetter(n: number): string {
  let name = '';
  let num = n;
  while (num > 0) {
    num--;
    name = String.fromCharCode(65 + (num % 26)) + name;
    num = Math.floor(num / 26);
  }
  return name;
}

/**
 * Convert user-facing `OpenOptions` to the internal napi `JsOpenOptions`.
 * Also accepts raw `JsOpenOptions` for backward compatibility.
 */
function toNativeOpenOptions(
  options?: OpenOptions | JsOpenOptions | null,
): JsOpenOptions | undefined {
  // Preserve raw napi parseMode behavior when callers pass legacy JsOpenOptions.
  if (
    options &&
    'parseMode' in options &&
    typeof (options as JsOpenOptions).parseMode === 'string' &&
    (options as JsOpenOptions).readMode == null
  ) {
    return options as JsOpenOptions;
  }
  return {
    ...(options ?? {}),
    readMode: (options as OpenOptions | JsOpenOptions | undefined)?.readMode ?? 'lazy',
    auxParts: (options as OpenOptions | JsOpenOptions | undefined)?.auxParts ?? 'deferred',
  } satisfies JsOpenOptions;
}

/** Row data from streaming reader, matching JsRowData shape. */
type RowData = JsRowData;

/** Tag strings shared with `encode_stream_batch` in `src/conversions.rs`. */
const STREAM_TAGS = ['empty', 'string', 'number', 'boolean', 'error', 'formula', 'date'];
const STREAM_TAG_LITERAL = 0xff;
const STREAM_HAS_VALUE = 0x01;
const STREAM_HAS_NUMBER = 0x02;
const STREAM_HAS_BOOL = 0x04;
const STREAM_HAS_NESTED = 0x08;

const streamColumnNames: string[] = [''];

function streamColumnName(col: number): string {
  for (let next = streamColumnNames.length; next <= col; next++) {
    let num = next;
    let name = '';
    while (num > 0) {
      num--;
      name = String.fromCharCode(65 + (num % 26)) + name;
      num = Math.floor(num / 26);
    }
    streamColumnNames.push(name);
  }
  return streamColumnNames[col];
}

/**
 * Decode one buffer produced by `NativeSheetStreamReader.nextBatchBuffer`
 * into the same row objects `nextBatch` returns.
 */
class StreamBatchDecoder {
  #buf: Buffer;
  #view: DataView;
  #pos = 0;

  constructor(buf: Buffer) {
    // Zero-copy JS view with a stable hidden class; see buffer-codec.ts.
    this.#buf = Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength);
    this.#view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  }

  decode(): RowData[] {
    const rowCount = this.#u32();
    const rows: RowData[] = new Array(rowCount);
    for (let r = 0; r < rowCount; r++) {
      const row = this.#u32();
      const cellCount = this.#u32();
      const cells: JsRowData['cells'] = new Array(cellCount);
      for (let c = 0; c < cellCount; c++) {
        cells[c] = this.#cell();
      }
      rows[r] = { row, cells };
    }
    if (this.#pos !== this.#buf.length) {
      throw new Error('Invalid stream batch: trailing bytes');
    }
    return rows;
  }

  #cell(): JsRowData['cells'][number] {
    const col = this.#view.getUint16(this.#pos, true);
    this.#pos += 2;
    const cell: JsRowData['cells'][number] = {
      column: streamColumnName(col),
      valueType: this.#tag(),
    };
    const flags = this.#u8();
    if (flags & STREAM_HAS_VALUE) cell.value = this.#str();
    if (flags & STREAM_HAS_NUMBER) cell.numberValue = this.#f64();
    if (flags & STREAM_HAS_BOOL) cell.boolValue = this.#u8() !== 0;
    if (flags & STREAM_HAS_NESTED) {
      const formula: NonNullable<JsRowData['cells'][number]['formula']> = {
        type: this.#tag(),
        formula: this.#str(),
      };
      if (this.#u8() !== 0) {
        const result: NonNullable<typeof formula.result> = { valueType: this.#tag() };
        const resultFlags = this.#u8();
        if (resultFlags & STREAM_HAS_VALUE) result.value = this.#str();
        if (resultFlags & STREAM_HAS_NUMBER) result.numberValue = this.#f64();
        if (resultFlags & STREAM_HAS_BOOL) result.boolValue = this.#u8() !== 0;
        if (resultFlags & STREAM_HAS_NESTED) {
          const date: NonNullable<typeof result.date> = {
            type: this.#tag(),
            serial: this.#f64(),
          };
          if (this.#u8() !== 0) date.iso = this.#str();
          result.date = date;
        }
        formula.result = result;
      }
      cell.formula = formula;
    }
    return cell;
  }

  #u8(): number {
    const value = this.#view.getUint8(this.#pos);
    this.#pos += 1;
    return value;
  }

  #u32(): number {
    const value = this.#view.getUint32(this.#pos, true);
    this.#pos += 4;
    return value;
  }

  #f64(): number {
    const value = this.#view.getFloat64(this.#pos, true);
    this.#pos += 8;
    return value;
  }

  #str(): string {
    const length = this.#u32();
    const start = this.#pos;
    const end = start + length;
    if (end > this.#buf.length) throw new Error('Invalid stream batch: truncated string');
    this.#pos = end;
    return this.#buf.toString('utf8', start, end);
  }

  #tag(): string {
    const code = this.#u8();
    if (code === STREAM_TAG_LITERAL) return this.#str();
    const tag = STREAM_TAGS[code];
    if (tag === undefined) throw new Error('Invalid stream batch: unknown tag');
    return tag;
  }
}

/** Forward-only streaming reader for worksheet data. Reads rows in batches. */
class SheetStreamReader {
  #native: NativeSheetStreamReader;
  #defaultBatchSize: number;

  /** @internal Use Workbook.openSheetReader() instead. */
  constructor(native: NativeSheetStreamReader, defaultBatchSize: number) {
    this.#native = native;
    this.#defaultBatchSize = defaultBatchSize;
  }

  /**
   * Read the next batch of rows. Returns an array of row data objects, or
   * null when there are no more rows.
   */
  async next(batchSize?: number): Promise<RowData[] | null> {
    const size = batchSize ?? this.#defaultBatchSize;
    const buf = this.#native.nextBatchBuffer(size);
    return buf === null ? null : new StreamBatchDecoder(buf).decode();
  }

  /** Close the reader and release resources. */
  async close(): Promise<void> {
    this.#native.close();
  }

  /** Async iterator that yields batches of rows. */
  async *[Symbol.asyncIterator](): AsyncGenerator<RowData[], void, undefined> {
    try {
      while (true) {
        const batch = await this.next();
        if (batch === null) break;
        yield batch;
      }
    } finally {
      await this.close();
    }
  }
}

/** Excel workbook for reading and writing .xlsx files. */
class Workbook {
  #native: NativeWorkbook;

  constructor() {
    this.#native = new NativeWorkbook();
  }

  static #wrap(native: NativeWorkbook): Workbook {
    const wb = new Workbook();
    wb.#native = native;
    return wb;
  }

  /** Open an existing .xlsx file from disk. */
  static openSync(
    path: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Workbook {
    return Workbook.#wrap(NativeWorkbook.openSync(path, toNativeOpenOptions(options)));
  }

  /** Open an existing .xlsx file from disk asynchronously. */
  static async open(
    path: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Promise<Workbook> {
    return Workbook.#wrap(await NativeWorkbook.open(path, toNativeOpenOptions(options)));
  }

  /** Open a workbook from an in-memory Buffer. */
  static openBufferSync(
    data: Buffer,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Workbook {
    return Workbook.#wrap(NativeWorkbook.openBufferSync(data, toNativeOpenOptions(options)));
  }

  /** Open a workbook from an in-memory Buffer asynchronously. */
  static async openBuffer(
    data: Buffer,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Promise<Workbook> {
    return Workbook.#wrap(await NativeWorkbook.openBuffer(data, toNativeOpenOptions(options)));
  }

  /** Open an encrypted .xlsx file using a password. */
  static openWithPasswordSync(
    path: string,
    password: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Workbook {
    return Workbook.#wrap(
      NativeWorkbook.openWithPasswordSync(path, password, toNativeOpenOptions(options)),
    );
  }

  /** Open an encrypted .xlsx file using a password asynchronously. */
  static async openWithPassword(
    path: string,
    password: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Promise<Workbook> {
    return Workbook.#wrap(
      await NativeWorkbook.openWithPassword(path, password, toNativeOpenOptions(options)),
    );
  }

  /** Open an encrypted workbook from an in-memory Buffer. */
  static openBufferWithPasswordSync(
    data: Buffer,
    password: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Workbook {
    return Workbook.#wrap(
      NativeWorkbook.openBufferWithPasswordSync(data, password, toNativeOpenOptions(options)),
    );
  }

  /** Open an encrypted in-memory Buffer asynchronously. */
  static async openBufferWithPassword(
    data: Buffer,
    password: string,
    options?: OpenOptions | JsOpenOptions | undefined | null,
  ): Promise<Workbook> {
    return Workbook.#wrap(
      await NativeWorkbook.openBufferWithPassword(data, password, toNativeOpenOptions(options)),
    );
  }

  /** Get the names of all sheets in workbook order. */
  get sheetNames(): string[] {
    return this.#native.sheetNames;
  }

  /** Save the workbook to a .xlsx file. */
  saveSync(path: string, limits?: SaveLimits | JsSaveLimits): void {
    this.#native.saveSync(path, limits);
  }

  /** Save the workbook using destination-preserving atomic replacement. */
  async save(path: string, limits?: SaveLimits | JsSaveLimits): Promise<void> {
    await this.#native.save(path, limits);
  }

  /** Serialize the workbook to an in-memory Buffer. */
  writeBufferSync(limits?: SaveLimits | JsSaveLimits): Buffer {
    return this.#native.writeBufferSync(limits);
  }

  /** Serialize the workbook to an in-memory Buffer asynchronously. */
  async writeBuffer(limits?: SaveLimits | JsSaveLimits): Promise<Buffer> {
    return this.#native.writeBuffer(limits);
  }

  /** Save the workbook as an encrypted .xlsx file. */
  saveWithPasswordSync(path: string, password: string, limits?: CryptLimits | JsCryptLimits): void {
    this.#native.saveWithPasswordSync(path, password, limits);
  }

  /** Save the workbook as an encrypted .xlsx file asynchronously. Both sync and async encrypted saves buffer in memory before encrypting, so peak memory is similar. */
  async saveWithPassword(
    path: string,
    password: string,
    limits?: CryptLimits | JsCryptLimits,
  ): Promise<void> {
    await this.#native.saveWithPassword(path, password, limits);
  }

  /** Get the value of a cell. */
  getCellValue(sheet: string, cell: string): CellValue {
    return this.#native.getCellValue(sheet, cell) as CellValue;
  }

  /** Get the formatted display text for a cell, applying its number format. */
  getCellFormattedValue(sheet: string, cell: string): string {
    return this.#native.getCellFormattedValue(sheet, cell);
  }

  /** Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear. */
  setCellValue(sheet: string, cell: string, value: CellValueInput): void {
    this.#native.setCellValue(sheet, cell, toNativeCellValue(value));
  }

  /** Set multiple cell values at once. */
  setCellValues(sheet: string, cells: Array<{ cell: string; value: CellValueInput }>): void {
    this.#native.setCellValues(
      sheet,
      cells.map(({ cell, value }) => ({ cell, value: toNativeCellValue(value) })),
    );
  }

  /** Set values in a single row starting from the given column. */
  setRowValues(sheet: string, row: number, startCol: string, values: CellValueInput[]): void {
    this.#native.setRowValues(sheet, row, startCol, values.map(toNativeCellValue));
  }

  /** Set a block of cell values from a 2D array. */
  setSheetData(
    sheet: string,
    data: CellValueInput[][],
    startCell?: string | undefined | null,
  ): void {
    this.#native.setSheetData(
      sheet,
      data.map((row) => row.map(toNativeCellValue)),
      startCell,
    );
  }

  /** Create a new empty sheet. Returns the 0-based sheet index. */
  newSheet(name: string): number {
    return this.#native.newSheet(name);
  }

  /** Delete a sheet by name. */
  deleteSheet(name: string): void {
    this.#native.deleteSheet(name);
  }

  /** Rename a sheet. */
  setSheetName(oldName: string, newName: string): void {
    this.#native.setSheetName(oldName, newName);
  }

  /** Copy a sheet. Returns the new sheet's 0-based index. */
  copySheet(source: string, target: string): number {
    return this.#native.copySheet(source, target);
  }

  /** Get the 0-based index of a sheet, or null if not found. */
  getSheetIndex(name: string): number | null {
    return this.#native.getSheetIndex(name);
  }

  /** Get the name of the active sheet. */
  getActiveSheet(): string {
    return this.#native.getActiveSheet();
  }

  /** Set the active sheet by name. */
  setActiveSheet(name: string): void {
    this.#native.setActiveSheet(name);
  }

  /** Insert empty rows starting at the given 1-based row number. */
  insertRows(sheet: string, startRow: number, count: number): void {
    this.#native.insertRows(sheet, startRow, count);
  }

  /** Remove a row (1-based). */
  removeRow(sheet: string, row: number): void {
    this.#native.removeRow(sheet, row);
  }

  /** Duplicate a row (1-based). */
  duplicateRow(sheet: string, row: number): void {
    this.#native.duplicateRow(sheet, row);
  }

  /** Set the height of a row (1-based). */
  setRowHeight(sheet: string, row: number, height: number): void {
    this.#native.setRowHeight(sheet, row, height);
  }

  /** Get the height of a row, or null if not explicitly set. */
  getRowHeight(sheet: string, row: number): number | null {
    return this.#native.getRowHeight(sheet, row);
  }

  /** Set whether a row is visible. */
  setRowVisible(sheet: string, row: number, visible: boolean): void {
    this.#native.setRowVisible(sheet, row, visible);
  }

  /** Get whether a row is visible. */
  getRowVisible(sheet: string, row: number): boolean {
    return this.#native.getRowVisible(sheet, row);
  }

  /** Set the outline level of a row (0-7). */
  setRowOutlineLevel(sheet: string, row: number, level: number): void {
    this.#native.setRowOutlineLevel(sheet, row, level);
  }

  /** Get the outline level of a row. Returns 0 if not set. */
  getRowOutlineLevel(sheet: string, row: number): number {
    return this.#native.getRowOutlineLevel(sheet, row);
  }

  /** Set the width of a column (e.g., "A", "B", "AA"). */
  setColWidth(sheet: string, col: string, width: number): void {
    this.#native.setColWidth(sheet, col, width);
  }

  /** Get the width of a column, or null if not explicitly set. */
  getColWidth(sheet: string, col: string): number | null {
    return this.#native.getColWidth(sheet, col);
  }

  /** Set whether a column is visible. */
  setColVisible(sheet: string, col: string, visible: boolean): void {
    this.#native.setColVisible(sheet, col, visible);
  }

  /** Get whether a column is visible. */
  getColVisible(sheet: string, col: string): boolean {
    return this.#native.getColVisible(sheet, col);
  }

  /** Set the outline level of a column (0-7). */
  setColOutlineLevel(sheet: string, col: string, level: number): void {
    this.#native.setColOutlineLevel(sheet, col, level);
  }

  /** Get the outline level of a column. Returns 0 if not set. */
  getColOutlineLevel(sheet: string, col: string): number {
    return this.#native.getColOutlineLevel(sheet, col);
  }

  /** Insert empty columns starting at the given column letter. */
  insertCols(sheet: string, col: string, count: number): void {
    this.#native.insertCols(sheet, col, count);
  }

  /** Remove a column by letter. */
  removeCol(sheet: string, col: string): void {
    this.#native.removeCol(sheet, col);
  }

  /** Add a style definition. Returns the style ID for use with setCellStyle. */
  addStyle(style: JsStyle): number {
    return this.#native.addStyle(style);
  }

  /** Get the style ID applied to a cell, or null if default. */
  getCellStyle(sheet: string, cell: string): number | null {
    return this.#native.getCellStyle(sheet, cell);
  }

  /** Apply a style ID to a cell. */
  setCellStyle(sheet: string, cell: string, styleId: number): void {
    this.#native.setCellStyle(sheet, cell, styleId);
  }

  /** Apply a style ID to an entire row. */
  setRowStyle(sheet: string, row: number, styleId: number): void {
    this.#native.setRowStyle(sheet, row, styleId);
  }

  /** Get the style ID for a row. Returns 0 if not set. */
  getRowStyle(sheet: string, row: number): number {
    return this.#native.getRowStyle(sheet, row);
  }

  /** Apply a style ID to an entire column. */
  setColStyle(sheet: string, col: string, styleId: number): void {
    this.#native.setColStyle(sheet, col, styleId);
  }

  /** Get the style ID for a column. Returns 0 if not set. */
  getColStyle(sheet: string, col: string): number {
    return this.#native.getColStyle(sheet, col);
  }

  /** Add a chart to a sheet. */
  addChart(sheet: string, fromCell: string, toCell: string, config: JsChartConfig): void {
    this.#native.addChart(sheet, fromCell, toCell, config);
  }

  /** Add an image to a sheet. */
  addImage(sheet: string, config: JsImageConfig): void {
    this.#native.addImage(sheet, config);
  }

  /** Add a shape to a sheet, anchored between two cells. */
  addShape(sheet: string, config: JsShapeConfig): void {
    this.#native.addShape(sheet, config);
  }

  /** Delete a chart anchored at the given cell. */
  deleteChart(sheet: string, cell: string): void {
    this.#native.deleteChart(sheet, cell);
  }

  /** Delete a picture anchored at the given cell. */
  deletePicture(sheet: string, cell: string): void {
    this.#native.deletePicture(sheet, cell);
  }

  /** Get all pictures anchored at the given cell. */
  getPictures(sheet: string, cell: string): JsPictureInfo[] {
    return this.#native.getPictures(sheet, cell);
  }

  /** Get all cells that have pictures anchored to them on the given sheet. */
  getPictureCells(sheet: string): string[] {
    return this.#native.getPictureCells(sheet);
  }

  /** Add a form control to a sheet. */
  addFormControl(sheet: string, config: JsFormControlConfig): void {
    this.#native.addFormControl(sheet, config);
  }

  /** Get all form controls on a sheet. */
  getFormControls(sheet: string): JsFormControlInfo[] {
    return this.#native.getFormControls(sheet);
  }

  /** Delete a form control by index. */
  deleteFormControl(sheet: string, index: number): void {
    this.#native.deleteFormControl(sheet, index);
  }

  /** Merge a range of cells on a sheet. */
  mergeCells(sheet: string, topLeft: string, bottomRight: string): void {
    this.#native.mergeCells(sheet, topLeft, bottomRight);
  }

  /** Remove a merged cell range from a sheet. */
  unmergeCell(sheet: string, reference: string): void {
    this.#native.unmergeCell(sheet, reference);
  }

  /** Get all merged cell ranges on a sheet. */
  getMergeCells(sheet: string): string[] {
    return this.#native.getMergeCells(sheet);
  }

  /** Add a data validation rule to a sheet. */
  addDataValidation(sheet: string, config: JsDataValidationConfig): void {
    this.#native.addDataValidation(sheet, config);
  }

  /** Get all data validations on a sheet. */
  getDataValidations(sheet: string): JsDataValidationConfig[] {
    return this.#native.getDataValidations(sheet);
  }

  /** Remove a data validation by sqref. */
  removeDataValidation(sheet: string, sqref: string): void {
    this.#native.removeDataValidation(sheet, sqref);
  }

  /** Set conditional formatting rules on a cell range. */
  setConditionalFormat(sheet: string, sqref: string, rules: JsConditionalFormatRule[]): void {
    this.#native.setConditionalFormat(sheet, sqref, rules);
  }

  /** Get all conditional formatting rules for a sheet. */
  getConditionalFormats(sheet: string): JsConditionalFormatEntry[] {
    return this.#native.getConditionalFormats(sheet);
  }

  /** Delete conditional formatting for a specific cell range. */
  deleteConditionalFormat(sheet: string, sqref: string): void {
    this.#native.deleteConditionalFormat(sheet, sqref);
  }

  /** Add a comment to a cell. */
  addComment(sheet: string, config: JsCommentConfig): void {
    this.#native.addComment(sheet, config);
  }

  /** Get all comments on a sheet. */
  getComments(sheet: string): JsCommentConfig[] {
    return this.#native.getComments(sheet);
  }

  /** Remove a comment from a cell. */
  removeComment(sheet: string, cell: string): void {
    this.#native.removeComment(sheet, cell);
  }

  /** Add a threaded comment to a cell. Returns the comment ID. */
  addThreadedComment(sheet: string, cell: string, input: JsThreadedCommentInput): string {
    return this.#native.addThreadedComment(sheet, cell, input);
  }

  /** Get all threaded comments on a sheet. */
  getThreadedComments(sheet: string): JsThreadedCommentData[] {
    return this.#native.getThreadedComments(sheet);
  }

  /** Get threaded comments for a specific cell on a sheet. */
  getThreadedCommentsByCell(sheet: string, cell: string): JsThreadedCommentData[] {
    return this.#native.getThreadedCommentsByCell(sheet, cell);
  }

  /** Delete a threaded comment by ID. */
  deleteThreadedComment(sheet: string, commentId: string): void {
    this.#native.deleteThreadedComment(sheet, commentId);
  }

  /** Set the resolved (done) state of a threaded comment. */
  resolveThreadedComment(sheet: string, commentId: string, done: boolean): void {
    this.#native.resolveThreadedComment(sheet, commentId, done);
  }

  /** Add a person to the person list. Returns the person ID. */
  addPerson(input: JsPersonInput): string {
    return this.#native.addPerson(input);
  }

  /** Get all persons in the person list. */
  getPersons(): JsPersonData[] {
    return this.#native.getPersons();
  }

  /** Set an auto-filter on a sheet. */
  setAutoFilter(sheet: string, range: string): void {
    this.#native.setAutoFilter(sheet, range);
  }

  /** Remove the auto-filter from a sheet. */
  removeAutoFilter(sheet: string): void {
    this.#native.removeAutoFilter(sheet);
  }

  /** Create a new stream writer for a new sheet. */
  newStreamWriter(sheetName: string, limits?: StreamLimits | JsStreamLimits): JsStreamWriter {
    return this.#native.newStreamWriter(sheetName, limits);
  }

  /** Apply a stream writer's output to the workbook. Returns the sheet index. */
  applyStreamWriter(writer: JsStreamWriter): number {
    return this.#native.applyStreamWriter(writer);
  }

  /**
   * Open a forward-only streaming reader for the named sheet.
   * Reads rows from a bounded incremental source without loading the entire
   * worksheet into memory.
   */
  async openSheetReader(sheet: string, opts?: { batchSize?: number }): Promise<SheetStreamReader> {
    const batchSize = opts?.batchSize ?? 1000;
    const native = this.#native.openSheetReader(sheet, batchSize);
    return new SheetStreamReader(native, batchSize);
  }

  /** Set core document properties (title, creator, etc.). */
  setDocProps(props: JsDocProperties): void {
    this.#native.setDocProps(props);
  }

  /** Get core document properties. */
  getDocProps(): JsDocProperties {
    return this.#native.getDocProps();
  }

  /** Set application properties (company, app version, etc.). */
  setAppProps(props: JsAppProperties): void {
    this.#native.setAppProps(props);
  }

  /** Get application properties. */
  getAppProps(): JsAppProperties {
    return this.#native.getAppProps();
  }

  /** Set a custom property. Value can be string, number, or boolean. */
  setCustomProperty(name: string, value: string | number | boolean): void {
    this.#native.setCustomProperty(name, value);
  }

  /** Get a custom property value, or null if not found. */
  getCustomProperty(name: string): string | number | boolean | null {
    return this.#native.getCustomProperty(name);
  }

  /** Delete a custom property. Returns true if it existed. */
  deleteCustomProperty(name: string): boolean {
    return this.#native.deleteCustomProperty(name);
  }

  /** Protect the workbook structure/windows with optional password. */
  protectWorkbook(config: JsWorkbookProtectionConfig): void {
    this.#native.protectWorkbook(config);
  }

  /** Remove workbook protection. */
  unprotectWorkbook(): void {
    this.#native.unprotectWorkbook();
  }

  /** Check if the workbook is protected. */
  isWorkbookProtected(): boolean {
    return this.#native.isWorkbookProtected();
  }

  /** Set freeze panes on a sheet. */
  setPanes(sheet: string, cell: string): void {
    this.#native.setPanes(sheet, cell);
  }

  /** Remove any freeze or split panes from a sheet. */
  unsetPanes(sheet: string): void {
    this.#native.unsetPanes(sheet);
  }

  /** Get the current freeze pane cell reference for a sheet, or null if none. */
  getPanes(sheet: string): string | null {
    return this.#native.getPanes(sheet);
  }

  /** Set page margins on a sheet (values in inches). */
  setPageMargins(sheet: string, margins: JsPageMargins): void {
    this.#native.setPageMargins(sheet, margins);
  }

  /** Get page margins for a sheet. Returns defaults if not explicitly set. */
  getPageMargins(sheet: string): JsPageMargins {
    return this.#native.getPageMargins(sheet);
  }

  /** Set page setup options (paper size, orientation, scale, fit-to-page). */
  setPageSetup(sheet: string, setup: JsPageSetup): void {
    this.#native.setPageSetup(sheet, setup);
  }

  /** Get the page setup for a sheet. */
  getPageSetup(sheet: string): JsPageSetup {
    return this.#native.getPageSetup(sheet);
  }

  /** Set header and footer text for printing. */
  setHeaderFooter(
    sheet: string,
    header?: string | undefined | null,
    footer?: string | undefined | null,
  ): void {
    this.#native.setHeaderFooter(sheet, header, footer);
  }

  /** Get the header and footer text for a sheet. */
  getHeaderFooter(sheet: string): JsHeaderFooter {
    return this.#native.getHeaderFooter(sheet);
  }

  /** Set print options on a sheet. */
  setPrintOptions(sheet: string, opts: JsPrintOptions): void {
    this.#native.setPrintOptions(sheet, opts);
  }

  /** Get print options for a sheet. */
  getPrintOptions(sheet: string): JsPrintOptions {
    return this.#native.getPrintOptions(sheet);
  }

  /** Insert a horizontal page break before the given 1-based row. */
  insertPageBreak(sheet: string, row: number): void {
    this.#native.insertPageBreak(sheet, row);
  }

  /** Remove a horizontal page break at the given 1-based row. */
  removePageBreak(sheet: string, row: number): void {
    this.#native.removePageBreak(sheet, row);
  }

  /** Get all row page break positions (1-based row numbers). */
  getPageBreaks(sheet: string): number[] {
    return this.#native.getPageBreaks(sheet);
  }

  /** Set a hyperlink on a cell. */
  setCellHyperlink(sheet: string, cell: string, opts: JsHyperlinkOptions): void {
    this.#native.setCellHyperlink(sheet, cell, opts);
  }

  /** Get hyperlink information for a cell, or null if no hyperlink exists. */
  getCellHyperlink(sheet: string, cell: string): JsHyperlinkInfo | null {
    return this.#native.getCellHyperlink(sheet, cell);
  }

  /** Delete a hyperlink from a cell. */
  deleteCellHyperlink(sheet: string, cell: string): void {
    this.#native.deleteCellHyperlink(sheet, cell);
  }

  #rowsBuffer(sheet: string, options?: TraversalOptions): Buffer {
    const nativeOptions =
      options?.mode === undefined && options?.limits === undefined
        ? undefined
        : { mode: options.mode, limits: options.limits };
    return options?.mode === 'dense'
      ? this.#native.getRowsBuffer(sheet, nativeOptions)
      : this.#native.getRowsBufferV2(sheet, nativeOptions);
  }

  /** Get all rows with their data from a sheet using buffer-based transfer. */
  getRows(sheet: string, options?: TraversalOptions): JsRowData[] {
    const buf = this.#rowsBuffer(sheet, options);
    return decodeRowsBuffer(buf);
  }

  /** Serialize a sheet's cell data into a compact binary buffer (v1 format). */
  getRowsBuffer(sheet: string, options?: TraversalOptions): Buffer {
    return this.#native.getRowsBuffer(sheet, options);
  }

  /**
   * Serialize a sheet's cell data into a v2 binary buffer with inline strings.
   * The v2 format eliminates the global string table, enabling incremental
   * row-by-row decoding without eagerly materializing all strings.
   */
  getRowsBufferV2(sheet: string, options?: TraversalOptions): Buffer {
    return this.#native.getRowsBufferV2(sheet, options);
  }

  /**
   * Get all rows as typed arrays with minimal JS object creation.
   * Returns column numbers (not names), cell types, and values in parallel
   * typed arrays. Use this when processing large sheets where object creation
   * overhead matters.
   */
  getRowsRaw(sheet: string, options?: TraversalOptions): RawRowsResult {
    const buf = this.#rowsBuffer(sheet, options);
    return decodeRowsRawBuffer(buf);
  }

  /**
   * Get a generator that yields one row at a time from the sheet, avoiding
   * materializing the entire result array in memory at once.
   */
  *getRowsIterator(sheet: string, options?: TraversalOptions): Generator<JsRowData> {
    const buf = this.#rowsBuffer(sheet, options);
    yield* decodeRowsIterator(buf);
  }

  /** Apply cell data from a binary buffer to a sheet. */
  setSheetDataBuffer(sheet: string, buf: Buffer, startCell?: string | undefined | null): void {
    this.#native.setSheetDataBuffer(sheet, buf, startCell);
  }

  /** Get all occupied columns using finite sparse traversal limits. */
  getCols(sheet: string, options?: TraversalOptions): JsColData[] {
    return this.#native.getCols(sheet, options);
  }

  /** Set a formula on a cell using finite mutation limits. */
  setCellFormula(sheet: string, cell: string, formula: string, limits?: FormulaLimits): void {
    this.#native.setCellFormula(sheet, cell, formula, limits);
  }

  /** Fill a range with a formula using finite mutation limits. */
  fillFormula(sheet: string, range: string, formula: string, limits?: FormulaLimits): void {
    this.#native.fillFormula(sheet, range, formula, limits);
  }

  /** Evaluate a formula string against the current workbook data. */
  evaluateFormula(sheet: string, formula: string, limits?: FormulaLimits): CellValue {
    return this.#native.evaluateFormula(sheet, formula, limits) as CellValue;
  }

  /** Recalculate all formula cells in the workbook. */
  calculateAll(limits?: FormulaLimits): void {
    this.#native.calculateAll(limits);
  }

  /** Add a pivot table using finite dense source traversal limits. */
  addPivotTable(config: JsPivotTableConfig, options?: TraversalOptions): void {
    this.#native.addPivotTable(config, options);
  }

  /** Get all pivot tables in the workbook. */
  getPivotTables(): JsPivotTableInfo[] {
    return this.#native.getPivotTables();
  }

  /** Delete a pivot table by name. */
  deletePivotTable(name: string): void {
    this.#native.deletePivotTable(name);
  }

  /** Add a sparkline to a worksheet. */
  addSparkline(sheet: string, config: JsSparklineConfig): void {
    this.#native.addSparkline(sheet, config);
  }

  /** Get all sparklines for a worksheet. */
  getSparklines(sheet: string): JsSparklineConfig[] {
    return this.#native.getSparklines(sheet);
  }

  /** Remove a sparkline by its location cell reference. */
  removeSparkline(sheet: string, location: string): void {
    this.#native.removeSparkline(sheet, location);
  }

  /** Register a table in the workbook. */
  addTable(sheet: string, config: JsTableConfig): void {
    this.#native.addTable(sheet, config);
  }

  /** Get all tables on a sheet. */
  getTables(sheet: string): JsTableInfo[] {
    return this.#native.getTables(sheet);
  }

  /** Delete a table from a sheet by name. */
  deleteTable(sheet: string, name: string): void {
    this.#native.deleteTable(sheet, name);
  }

  /** Add a slicer to a sheet targeting a table column. */
  addSlicer(sheet: string, config: JsSlicerConfig): void {
    this.#native.addSlicer(sheet, config);
  }

  /** Get all slicers on a sheet. */
  getSlicers(sheet: string): JsSlicerInfo[] {
    return this.#native.getSlicers(sheet);
  }

  /** Delete a slicer by name from a sheet. */
  deleteSlicer(sheet: string, name: string): void {
    this.#native.deleteSlicer(sheet, name);
  }

  /** Set a cell to a rich text value with multiple formatted runs. */
  setCellRichText(sheet: string, cell: string, runs: JsRichTextRun[]): void {
    this.#native.setCellRichText(sheet, cell, runs);
  }

  /** Get rich text runs for a cell, or null if not rich text. */
  getCellRichText(sheet: string, cell: string): JsRichTextRun[] | null {
    return this.#native.getCellRichText(sheet, cell);
  }

  /** Resolve a theme color by index (0-11) with optional tint. */
  getThemeColor(index: number, tint?: number | undefined | null): string | null {
    return this.#native.getThemeColor(index, tint);
  }

  /** Add or update a defined name. */
  setDefinedName(config: JsDefinedNameConfig): void {
    this.#native.setDefinedName(config);
  }

  /** Get a defined name by name and optional scope. */
  getDefinedName(name: string, scope?: string | undefined | null): JsDefinedNameInfo | null {
    return this.#native.getDefinedName(name, scope);
  }

  /** Get all defined names in the workbook. */
  getDefinedNames(): JsDefinedNameInfo[] {
    return this.#native.getDefinedNames();
  }

  /** Delete a defined name by name and optional scope. */
  deleteDefinedName(name: string, scope?: string | undefined | null): void {
    this.#native.deleteDefinedName(name, scope);
  }

  /** Protect a sheet with optional password and permission settings. */
  protectSheet(sheet: string, config?: JsSheetProtectionConfig | undefined | null): void {
    this.#native.protectSheet(sheet, config);
  }

  /** Remove sheet protection. */
  unprotectSheet(sheet: string): void {
    this.#native.unprotectSheet(sheet);
  }

  /** Check if a sheet is protected. */
  isSheetProtected(sheet: string): boolean {
    return this.#native.isSheetProtected(sheet);
  }

  /** Convert sheet data to an array of JSON objects. */
  toJSON(sheet: string, options?: ToJsonOptions): Record<string, CellValue>[] {
    const buf = this.#rowsBuffer(sheet, options);
    const sd = new SheetData(buf);
    if (options?.mode !== 'dense') {
      const rows = [...sd.occupiedRows()];
      if (rows.length === 0) return [];

      const headerOpt = options?.header;
      let keys: Array<string | undefined> | null = null;
      let dataStartIndex = 0;
      if (Array.isArray(headerOpt)) {
        keys = headerOpt;
      } else if (headerOpt === true || headerOpt === undefined) {
        keys = [];
        for (const { column, value } of rows[0].cells) {
          keys[column - sd.minCol] = cellValueToString(value);
        }
        dataStartIndex = 1;
      }

      const result: Record<string, CellValue>[] = [];
      for (let rowIndex = dataStartIndex; rowIndex < rows.length; rowIndex++) {
        const obj = createJsonOutputObject();
        for (const { column, value } of rows[rowIndex].cells) {
          const key = keys === null ? columnNumberToLetter(column) : keys[column - sd.minCol];
          if (key === undefined) continue;
          obj[key] = value;
        }
        result.push(finishJsonOutputObject(obj));
      }
      return result;
    }

    const rows = sd.toArray();
    if (rows.length === 0) return [];

    const headerOpt = options?.header;
    let keys: string[];
    let dataStartIndex: number;

    if (Array.isArray(headerOpt)) {
      keys = headerOpt;
      dataStartIndex = 0;
    } else if (headerOpt === true || headerOpt === undefined) {
      keys = rows[0].map((v) => cellValueToString(v));
      dataStartIndex = 1;
    } else {
      const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
      keys = Array.from({ length: colCount }, (_, i) => columnNumberToLetter(sd.minCol + i));
      dataStartIndex = 0;
    }

    const result: Record<string, CellValue>[] = [];
    for (let r = dataStartIndex; r < rows.length; r++) {
      const row = rows[r];
      const obj = createJsonOutputObject();
      for (let c = 0; c < keys.length; c++) {
        obj[keys[c]] = c < row.length ? row[c] : null;
      }
      result.push(finishJsonOutputObject(obj));
    }
    return result;
  }

  /**
   * Convert sheet data to CSV. Formula cells export expressions regardless of
   * cached results. Records use LF by default without a trailing terminator.
   */
  toCSV(sheet: string, options?: ToCsvOptions): string {
    const delimiter = options?.delimiter ?? ',';
    const quote = options?.quote ?? '"';
    const lineEnding = options?.lineEnding ?? '\n';
    const formulaPolicy = csvFormulaPolicy(options?.formulaPolicy);
    validateCsvSyntax(delimiter, quote, lineEnding, formulaPolicy);
    const outputLimit = csvOutputLimit(options);

    const buf = this.#rowsBuffer(sheet, csvTraversalOptions(options, outputLimit));
    const sd = new SheetData(buf);
    const dense = options?.mode === 'dense';
    const rows: CsvCell[][] = dense
      ? sd.toArray().map((row, rowIndex) =>
          row.map((value, columnIndex) => ({
            value,
            formula: sd.getCellType(sd.minRow + rowIndex, sd.minCol + columnIndex) === 'formula',
          })),
        )
      : [...sd.typedOccupiedRows()].map(({ cells }) => cells);
    if (rows.length === 0) return '';

    const maxCols = dense ? rows.reduce((max, row) => Math.max(max, row.length), 0) : null;
    const leadingEmptyFields = dense ? sd.minCol - 1 : 0;
    return encodeCsvOutput(
      rows,
      leadingEmptyFields,
      maxCols,
      delimiter,
      quote,
      lineEnding,
      formulaPolicy,
      outputLimit,
    );
  }

  /** Convert sheet data to an HTML table string. */
  toHTML(sheet: string, options?: ToHtmlOptions): string {
    const className = options?.className;
    const buf = this.#rowsBuffer(sheet, options);
    const sd = new SheetData(buf);
    const rows =
      options?.mode === 'dense'
        ? sd.toArray()
        : [...sd.occupiedRows()].map(({ cells }) => cells.map(({ value }) => value));

    const limit = options?.limits?.maxOutputBytes ?? DEFAULT_TRAVERSAL_OUTPUT_BYTES;
    const classAttr = className ? ` class="${escapeHtml(className)}"` : '';
    const open = `<table${classAttr}>`;
    let observed = chargeHtmlOutput(0, Buffer.byteLength(open), limit);
    const parts: string[] = [open];
    // Build and charge one string per row; per-cell byte counting dominated runtime.
    for (const row of rows) {
      let line = '<tr>';
      for (const cell of row) {
        line += `<td>${escapeHtml(cellValueToString(cell))}</td>`;
      }
      line += '</tr>';
      observed = chargeHtmlOutput(observed, Buffer.byteLength(line), limit);
      parts.push(line);
    }
    chargeHtmlOutput(observed, HTML_TABLE_CLOSE.length, limit);
    parts.push(HTML_TABLE_CLOSE);
    return parts.join('');
  }

  /** Write an array of JSON objects to a sheet. */
  fromJSON(sheet: string, data: Record<string, CellValueInput>[], options?: FromJsonOptions): void {
    if (!Array.isArray(data)) throw new TypeError('JSON import data must be an array');
    if (
      options !== undefined &&
      (typeof options !== 'object' || options === null || Array.isArray(options))
    ) {
      throw new TypeError('JSON import options must be an object');
    }

    const limits = resolveFromJsonLimits(options?.limits);
    const startCell = options?.startCell ?? 'A1';
    if (typeof startCell !== 'string') throw new TypeError('Invalid cell reference');
    const { col: startCol, row: startRow } = parseCellRef(startCell);

    const headerOpt = options?.header;
    if (headerOpt !== undefined && typeof headerOpt !== 'boolean' && !Array.isArray(headerOpt)) {
      throw new TypeError('header must be a boolean or an array of strings');
    }
    const writeHeader = headerOpt !== false;
    const dataLength = jsonArrayLength(data as unknown[]);
    const rowCount = dataLength === 0 ? 0 : dataLength + (writeHeader ? 1 : 0);
    if (!Number.isSafeInteger(rowCount)) {
      throw fromJsonLimitError('row_count', limits.maxRows, limits.maxRows + 1);
    }
    checkFromJsonLimit('row_count', limits.maxRows, rowCount);
    if (rowCount > 0) {
      const lastRow = startRow + rowCount - 1;
      checkFromJsonLimit('row_count', EXCEL_MAX_ROWS, lastRow);
    }

    const keys: string[] = [];
    const keyToOffset = new Map<string, number>();
    let retainedBytes = 0;
    const addKey = (key: string): number => {
      const existing = keyToOffset.get(key);
      if (existing !== undefined) return existing;
      const columnCount = keys.length + 1;
      checkFromJsonLimit('column_count', limits.maxColumns, columnCount);
      checkFromJsonLimit('column_count', EXCEL_MAX_COLUMNS, columnCount);
      checkFromJsonLimit('column_count', EXCEL_MAX_COLUMNS, startCol + columnCount - 1);
      retainedBytes = checkedFromJsonAdd(
        retainedBytes,
        FROM_JSON_KEY_OVERHEAD_BYTES + Buffer.byteLength(key),
        limits.maxRetainedBytes,
        'retained_bytes',
      );
      const offset = keys.length;
      keys.push(key);
      keyToOffset.set(key, offset);
      return offset;
    };

    if (Array.isArray(headerOpt)) {
      const headerLength = jsonArrayLength(headerOpt as unknown[]);
      checkFromJsonLimit('column_count', limits.maxColumns, headerLength);
      checkFromJsonLimit('column_count', EXCEL_MAX_COLUMNS, headerLength);
      for (let index = 0; index < headerLength; index++) {
        const key = jsonArrayElement(headerOpt as unknown[], index);
        if (typeof key !== 'string') throw new TypeError('header names must be strings');
        if (keyToOffset.has(key)) throw new TypeError('header names must be unique');
        addKey(key);
      }
    }
    if (dataLength === 0) return;

    const entries: JsCellEntry[] = [];
    let cellCount = 0;

    for (let rowOffset = 0; rowOffset < dataLength; rowOffset++) {
      const record = jsonArrayElement(data as unknown[], rowOffset);
      if (!isJsonRecord(record)) throw new TypeError('JSON import rows must be objects');
      for (const key of jsonRecordKeys(record)) {
        let columnOffset = keyToOffset.get(key);
        if (columnOffset === undefined) {
          if (Array.isArray(headerOpt)) {
            throw new TypeError('JSON import row keys must match custom headers');
          }
          columnOffset = addKey(key);
        }
        cellCount = checkedFromJsonAdd(cellCount, 1, limits.maxCells, 'cell_count');
        const cell = `${columnNumberToLetter(startCol + columnOffset)}${
          startRow + rowOffset + (writeHeader ? 1 : 0)
        }`;
        retainedBytes = checkedFromJsonAdd(
          retainedBytes,
          FROM_JSON_ENTRY_OVERHEAD_BYTES + Buffer.byteLength(cell),
          limits.maxRetainedBytes,
          'retained_bytes',
        );
        const snapshot = snapshotJsonCellValue(jsonRecordValue(record, key));
        retainedBytes = checkedFromJsonAdd(
          retainedBytes,
          snapshot.retainedBytes,
          limits.maxRetainedBytes,
          'retained_bytes',
        );
        entries.push({
          cell,
          value: snapshot.value,
        });
      }
    }

    if (writeHeader) {
      for (let offset = 0; offset < keys.length; offset++) {
        const key = keys[offset];
        const keyBytes = Buffer.byteLength(key);
        if (keyBytes > MAX_FROM_JSON_CELL_STRING_BYTES) {
          throw new TypeError('JSON cell string exceeds the Excel limit');
        }
        cellCount = checkedFromJsonAdd(cellCount, 1, limits.maxCells, 'cell_count');
        const cell = `${columnNumberToLetter(startCol + offset)}${startRow}`;
        retainedBytes = checkedFromJsonAdd(
          retainedBytes,
          FROM_JSON_ENTRY_OVERHEAD_BYTES + Buffer.byteLength(cell) + keyBytes,
          limits.maxRetainedBytes,
          'retained_bytes',
        );
        entries.push({ cell, value: key });
      }
    }

    if (entries.length === 0) return;

    this.#native.setCellValues(sheet, entries);
  }

  /** Get the workbook format ("xlsx", "xlsm", "xltx", "xltm", "xlam"). */
  getFormat(): string {
    return this.#native.getFormat();
  }

  /** Set sheet view options (gridlines, zoom, view mode, etc.). */
  setSheetViewOptions(sheet: string, opts: JsSheetViewOptions): void {
    this.#native.setSheetViewOptions(sheet, opts);
  }

  /** Get sheet view options. */
  getSheetViewOptions(sheet: string): JsSheetViewOptions {
    return this.#native.getSheetViewOptions(sheet);
  }

  /** Set sheet visibility ("visible", "hidden", or "veryHidden"). */
  setSheetVisibility(sheet: string, visibility: string): void {
    this.#native.setSheetVisibility(sheet, visibility);
  }

  /** Get sheet visibility. Returns "visible", "hidden", or "veryHidden". */
  getSheetVisibility(sheet: string): string {
    return this.#native.getSheetVisibility(sheet);
  }

  /** Render a worksheet to SVG using finite dense traversal limits. */
  renderToSvg(options: JsRenderOptions, traversal?: TraversalOptions): string {
    return this.#native.renderToSvg(options, traversal);
  }

  /** Get the raw VBA project binary (xl/vbaProject.bin), or null if not present. */
  getVbaProject(): Buffer | null {
    return this.#native.getVbaProject();
  }

  /** Extract VBA module source code from the workbook's VBA project. Returns null if no VBA project. */
  getVbaModules(limits?: VbaLimits | JsVbaLimits): JsVbaProject | null {
    return this.#native.getVbaModules(limits);
  }
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Format a numeric value using an Excel format code string. */
function formatNumber(value: number, formatCode: string): string {
  return NativeWorkbook.formatNumber(value, formatCode);
}

/** Get the format code string for a built-in number format ID (0-49). Returns null if unknown. */
function builtinFormatCode(id: number): string | null {
  return NativeWorkbook.builtinFormatCode(id);
}

export { builtinFormatCode, formatNumber, JsStreamWriter, SheetData, SheetStreamReader, Workbook };
export type { CellTypeName, RawRowsResult, RowData };
