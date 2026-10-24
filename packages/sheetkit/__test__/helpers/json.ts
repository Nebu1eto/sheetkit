/** Any value `JSON.parse` can produce. */
export type JsonValue = string | number | boolean | null | JsonValue[] | JsonObject;

/** A JSON object as produced by `JSON.parse`. */
export interface JsonObject {
  [key: string]: JsonValue;
}

/** A JSON object whose own property values are all strings. */
export interface JsonStringRecord {
  [key: string]: string;
}

function isJsonObject(value: JsonValue): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isString(value: JsonValue): value is string {
  return typeof value === 'string';
}

function isJsonStringRecord(value: JsonObject): value is JsonStringRecord {
  return Object.values(value).every(isString);
}

/**
 * Parses `text` as JSON and returns the top-level object.
 * Throws `SyntaxError` for malformed JSON and `TypeError` when the top-level value
 * is not an object, so a test fails at the parse boundary.
 */
export function parseJsonObject(text: string): JsonObject {
  const parsed: JsonValue = JSON.parse(text);

  if (!isJsonObject(parsed)) throw new TypeError('Expected a top-level JSON object');

  return parsed;
}

/**
 * Parses `text` as a JSON object whose own property values are all strings.
 * Throws for malformed JSON or any other JSON document.
 */
export function parseJsonStringRecord(text: string): JsonStringRecord {
  const parsed = parseJsonObject(text);

  if (!isJsonStringRecord(parsed)) throw new TypeError('Expected only string JSON values');

  return parsed;
}
