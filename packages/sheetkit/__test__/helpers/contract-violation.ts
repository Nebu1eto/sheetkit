import type { Workbook } from '../../index.js';

/** One record accepted by `Workbook.fromJSON`. */
export type JsonImportRecord = Parameters<Workbook['fromJSON']>[1][number];

/** A primitive that untyped JavaScript callers can place in any option or value field. */
export type ForeignPrimitive = string | number | boolean | null;

/**
 * Overwrites `key` on `target` with an enumerable, writable, configurable data property
 * holding `value`, which may lie outside the field's declared type.
 *
 * Tests use this to reproduce what untyped JavaScript callers can pass at runtime
 * (for example `{ durable: 'yes' }`) without weakening the static type of `target`.
 * Returns `target` itself.
 */
export function withForeignField<T extends object>(
  target: T,
  key: string,
  value: ForeignPrimitive,
): T {
  Object.defineProperty(target, key, {
    value,
    enumerable: true,
    writable: true,
    configurable: true,
  });

  return target;
}
