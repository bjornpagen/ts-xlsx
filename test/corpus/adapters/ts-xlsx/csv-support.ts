// CSV spec ↔ model translation, including the oracle's ExcelJS-shaped option names.

import type {Untyped} from '../../untyped.ts';

// A JSON-serializable view of a read-back CSV cell value, mirroring the oracle's normalizeCsvValue.
export const normalizeCsvValue = (v: Untyped) => {
  if (v instanceof Date) return {date: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  if (v && typeof v === 'object' && 'error' in v) return {error: v.error};
  return v ?? null;
};

// A declarative CSV write-spec cell → a live model value: { date } → Date, { formula, result } →
// formula value, { error } → error value, primitive passes through.
export const specCsvValue = (c: Untyped) => {
  if (c && typeof c === 'object') {
    if (c.date) return new Date(c.date);
    if ('formula' in c) return {formula: c.formula, result: c.result};
    if ('error' in c) return {error: c.error};
  }
  return c;
};

// The oracle's ExcelJS-shaped read options → the rewrite's CsvReadOptions.
export const translateCsvReadOptions = (options: Untyped = {}) => {
  const parser = options.parserOptions || {};
  const translated: Record<string, Untyped> = {};
  if (parser.delimiter !== undefined) translated.delimiter = parser.delimiter;
  if (parser.headers) translated.headers = true;
  if (typeof options.map === 'function') translated.map = options.map;
  return translated;
};

// The oracle's ExcelJS-shaped write options → the rewrite's CsvWriteOptions.
export const translateCsvWriteOptions = (options: Untyped = {}) => {
  const formatter = options.formatterOptions || {};
  const translated: Record<string, Untyped> = {};
  if (formatter.delimiter !== undefined) translated.delimiter = formatter.delimiter;
  if (formatter.rowDelimiter !== undefined) translated.rowDelimiter = formatter.rowDelimiter;
  if (options.dateFormat !== undefined) translated.dateFormat = options.dateFormat;
  if (options.dateUTC !== undefined) translated.dateUTC = options.dateUTC;
  return translated;
};
