import {
  type FormulaLimits,
  isResourceLimitError,
  type OpenOptions,
  type ResourceLimitError,
  type SaveOptions,
  type TraversalOptions,
  Workbook,
} from '@sheetkit/node';
import {
  type JsFormulaLimits,
  type JsOpenOptions,
  type JsSaveOptions,
  type JsTraversalOptions,
  Workbook as NativeWorkbook,
} from '@sheetkit/node/binding.js';

const openOptions: OpenOptions = {
  readMode: 'lazy',
  auxParts: 'deferred',
  maxOuterInputBytes: 1024,
  maxUnzipSize: 1024,
  maxZipEntries: 16,
  cryptLimits: { kdfIterations: 100_000 },
  vbaLimits: { moduleCount: 16 },
};

const legacyOpenOptions: JsOpenOptions = { parseMode: 'readfast' };

const traversal: TraversalOptions = {
  mode: 'sparse',
  limits: { maxRows: 10, maxColumns: 10, maxCells: 100, maxOutputBytes: 4096 },
};

const nativeTraversal: JsTraversalOptions = traversal;

const formulaLimits: FormulaLimits = {
  perFormulaInputBytes: 1024,
  operationInputBytes: 4096,
  operationEvaluationSteps: 4096,
};

const nativeFormulaLimits: JsFormulaLimits = formulaLimits;

const workbook = new Workbook();

workbook.getRows('Sheet1', traversal);

workbook.getCols('Sheet1', traversal);

workbook.evaluateFormula('Sheet1', '1+1', formulaLimits);

workbook.calculateAll(formulaLimits);

Workbook.openSync('workbook.xlsx', openOptions);

Workbook.open('workbook.xlsx', legacyOpenOptions);

const saveOptions: SaveOptions = { durable: true };
const nativeSaveOptions: JsSaveOptions = saveOptions;

workbook.saveSync('workbook.xlsx', undefined, saveOptions);

workbook.save('workbook.xlsx', { outputBytes: 1024 }, nativeSaveOptions);

workbook.saveWithPasswordSync('workbook.xlsx', 'password', undefined, saveOptions);

const nativeWorkbook = new NativeWorkbook();

nativeWorkbook.getRows('Sheet1', nativeTraversal);

nativeWorkbook.getCols('Sheet1', nativeTraversal);

nativeWorkbook.evaluateFormula('Sheet1', '1+1', nativeFormulaLimits);

nativeWorkbook.calculateAll(nativeFormulaLimits);

declare const unknownError: unknown;

if (isResourceLimitError(unknownError)) {
  const error: ResourceLimitError = unknownError;
  const observed: bigint = error.observed;
  void observed;
}
