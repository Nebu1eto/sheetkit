import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const binding = require('./binding.cjs');
export const { Workbook, JsStreamWriter } = binding;
