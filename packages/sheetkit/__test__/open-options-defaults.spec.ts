import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { JsOpenOptions } from '../binding.js';

const nativeCalls = vi.hoisted(() => ({
  openSync: vi.fn<(path: string, options?: unknown) => unknown>(() => ({})),
  open: vi.fn<(path: string, options?: unknown) => Promise<unknown>>(async () => ({})),
  openBufferSync: vi.fn<(data: Buffer, options?: unknown) => unknown>(() => ({})),
  openBuffer: vi.fn<(data: Buffer, options?: unknown) => Promise<unknown>>(async () => ({})),
}));

vi.mock('../binding.js', () => {
  class MockNativeWorkbook {
    constructor() {}

    static openSync(path: string, options?: unknown): unknown {
      return nativeCalls.openSync(path, options);
    }

    static async open(path: string, options?: unknown): Promise<unknown> {
      return nativeCalls.open(path, options);
    }

    static openBufferSync(data: Buffer, options?: unknown): unknown {
      return nativeCalls.openBufferSync(data, options);
    }

    static async openBuffer(data: Buffer, options?: unknown): Promise<unknown> {
      return nativeCalls.openBuffer(data, options);
    }

    static formatNumber(value: number, formatCode: string): string {
      return `${value}:${formatCode}`;
    }

    static builtinFormatCode(_id: number): string | null {
      return null;
    }
  }

  class MockNativeSheetStreamReader {
    nextBatch(_size: number): null {
      return null;
    }

    close(): void {}
  }

  class MockJsStreamWriter {}

  return {
    JsStreamWriter: MockJsStreamWriter,
    NativeSheetStreamReader: MockNativeSheetStreamReader,
    Workbook: MockNativeWorkbook,
  };
});

import { Workbook } from '../index.js';

describe('OpenOptions wrapper defaults', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('uses readMode=lazy and auxParts=deferred when options are omitted', async () => {
    const path = '/tmp/test.xlsx';
    const buffer = Buffer.from('xlsx');

    Workbook.openSync(path);
    await Workbook.open(path);
    Workbook.openBufferSync(buffer);
    await Workbook.openBuffer(buffer);

    const expectedDefaults = { readMode: 'lazy', auxParts: 'deferred' } as const;
    expect(nativeCalls.openSync).toHaveBeenCalledWith(path, expectedDefaults);
    expect(nativeCalls.open).toHaveBeenCalledWith(path, expectedDefaults);
    expect(nativeCalls.openBufferSync).toHaveBeenCalledWith(buffer, expectedDefaults);
    expect(nativeCalls.openBuffer).toHaveBeenCalledWith(buffer, expectedDefaults);
  });

  it('applies defaults when options are partial', async () => {
    const path = '/tmp/test.xlsx';
    const buffer = Buffer.from('xlsx');

    Workbook.openSync(path, { sheetRows: 5 });
    await Workbook.open(path, { auxParts: 'eager' });
    Workbook.openBufferSync(buffer, { readMode: 'eager' });
    await Workbook.openBuffer(buffer, { sheets: ['Sheet1'] });

    expect(nativeCalls.openSync).toHaveBeenCalledWith(path, {
      sheetRows: 5,
      readMode: 'lazy',
      auxParts: 'deferred',
    });
    expect(nativeCalls.open).toHaveBeenCalledWith(path, {
      auxParts: 'eager',
      readMode: 'lazy',
    });
    expect(nativeCalls.openBufferSync).toHaveBeenCalledWith(buffer, {
      readMode: 'eager',
      auxParts: 'deferred',
    });
    expect(nativeCalls.openBuffer).toHaveBeenCalledWith(buffer, {
      sheets: ['Sheet1'],
      readMode: 'lazy',
      auxParts: 'deferred',
    });
  });

  it('preserves legacy parseMode-only options', () => {
    const path = '/tmp/test.xlsx';
    Workbook.openSync(path, { parseMode: 'full' } as JsOpenOptions);
    expect(nativeCalls.openSync).toHaveBeenCalledWith(path, { parseMode: 'full' });
  });
});
