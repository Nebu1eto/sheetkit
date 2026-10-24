import type { JsOpenOptions } from './binding.js';
import type { OpenOptions } from './index.js';

/** True for legacy napi options that select `parseMode` without a `readMode`. */
function isLegacyParseModeOptions(options: OpenOptions | JsOpenOptions): options is JsOpenOptions {
  return (
    'parseMode' in options && typeof options.parseMode === 'string' && options.readMode == null
  );
}

/**
 * Convert user-facing `OpenOptions` to the internal napi `JsOpenOptions`.
 * Also accepts raw `JsOpenOptions` for backward compatibility.
 */
export function toNativeOpenOptions(
  options?: OpenOptions | JsOpenOptions | null,
): JsOpenOptions | undefined {
  // Preserve raw napi parseMode behavior when callers pass legacy JsOpenOptions.
  if (options && isLegacyParseModeOptions(options)) return options;

  return {
    ...(options ?? {}),
    readMode: options?.readMode ?? 'lazy',
    auxParts: options?.auxParts ?? 'deferred',
  } satisfies JsOpenOptions;
}
