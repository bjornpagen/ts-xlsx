// The implementation under test, loaded once for the whole adapter.
//
// Every other module here imports its `src/` bindings from this one so the CORPUS_TARGET
// switch, stripped .ts sources by default, emitted dist/ JS on demand, is decided in a
// single place and cannot drift between concerns.

import fs from 'node:fs';
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

// Retarget the implementation under test. Default: the src/ .ts sources, run
// directly via Node's type-stripping (the zero-build dev/test loop). Set
// CORPUS_TARGET=dist to run the *emitted* artifact instead, the exact ESM `tsc`
// produces for consumers, putting the full behavioral corpus behind the same
// gate. test:src and this adapter's default only ever see *stripped* source;
// dist runs catch strip-vs-emit divergence (import-specifier rewrite, a
// runtime reference type-stripping tolerated) across every case, not just the
// smoke round-trip. dist mirrors src's tree (rootDir=src), so the only change is
// the base dir and the .ts→.js extension.
export const target =
  process.env.CORPUS_TARGET === 'dist' ? {dir: 'dist', ext: 'js'} : {dir: 'src', ext: 'ts'};
export const loadModule = <T>(rel: string): Promise<T> =>
  import(
    new URL(`../../../../${target.dir}/${rel}.${target.ext}`, import.meta.url).href
  ) as Promise<T>;

export const {decodeAddress, decodeRange, encodeAddress, MAX_COLUMN, MAX_ROW} =
  await loadModule<typeof import('../../../../src/core/address.ts')>('core/address');
export const {XlsxError} = await loadModule<typeof import('../../../../src/errors.ts')>('errors');
// The codec's back channel, for the two facts a cell carries that have no public accessor. A case
// never sees these; the adapter reaches them so it can report what survived a copy.
export const {NAMED_STYLE_ID} =
  await loadModule<typeof import('../../../../src/core/internal.ts')>('core/internal');
export const {detectValueType, isRichTextValue} =
  await loadModule<typeof import('../../../../src/core/value.ts')>('core/value');
export const {Workbook} =
  await loadModule<typeof import('../../../../src/core/workbook.ts')>('core/workbook');
export const {readCsv} =
  await loadModule<typeof import('../../../../src/io/csv/read.ts')>('io/csv/read');
export const {writeCsv, writeCsvText} =
  await loadModule<typeof import('../../../../src/io/csv/write.ts')>('io/csv/write');
export const {readXlsx} =
  await loadModule<typeof import('../../../../src/io/xlsx/read.ts')>('io/xlsx/read');
export const {readXlsb} =
  await loadModule<typeof import('../../../../src/io/xlsb/read.ts')>('io/xlsb/read');
export const {readWorkbookStream} =
  await loadModule<typeof import('../../../../src/io/xlsx/read-rows.ts')>('io/xlsx/read-rows');
export const {writeXlsx} =
  await loadModule<typeof import('../../../../src/io/xlsx/write.ts')>('io/xlsx/write');
export const {WorkbookStreamWriter} =
  await loadModule<typeof import('../../../../src/io/xlsx/write-stream.ts')>(
    'io/xlsx/write-stream',
  );
export const {CompoundFile} =
  await loadModule<typeof import('../../../../src/vba/cfb.ts')>('vba/cfb');
export const {writeCompoundFile} =
  await loadModule<typeof import('../../../../src/vba/cfb-writer.ts')>('vba/cfb-writer');
export const {compressContainer, decompressContainer} =
  await loadModule<typeof import('../../../../src/vba/ms-ovba.ts')>('vba/ms-ovba');
export const {parseVbaProject} =
  await loadModule<typeof import('../../../../src/vba/project.ts')>('vba/project');
export const {addVbaReference, removeVbaModule} =
  await loadModule<typeof import('../../../../src/vba/project-editor.ts')>('vba/project-editor');

// The model's instance types, so the adapter can annotate what it is holding instead of reaching for
// `Untyped`. Derived from the bindings above rather than imported separately: `runtime.ts` stays the
// only module that knows where `src/` is, and a type here cannot drift from the implementation it
// describes because it is read off that implementation. Nothing is emitted: these are erased before
// anything runs, so the CORPUS_TARGET switch is still the only thing deciding what executes.
export type WorkbookInstance = InstanceType<typeof Workbook>;
export type WorksheetInstance = ReturnType<WorkbookInstance['addWorksheet']>;
export type CellInstance = ReturnType<WorksheetInstance['getCell']>;
export type RowInput = Parameters<WorksheetInstance['addRow']>[0];

// JSZip is an independent zip implementation used only to VERIFY the streaming writer's output (CRC
// integrity), a hostile-input posture toward our own archive, never in the production src path.
const require = createRequire(import.meta.url);
export const JSZip = require('jszip');

// Durable sample inputs live under test/corpus/fixtures/<case-slug>/: the SAME tree the
// oracle adapter reads, so a fixture-backed case measures both implementations against one
// real-world file. The rewrite reads them straight through readXlsx (a fixture is just a
// foreign `.xlsx` buffer).
export const FIXTURES_ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  'fixtures',
);
export const fixtureBytes = (rel: string) => fs.readFileSync(path.join(FIXTURES_ROOT, rel));
export const readFixture = (rel: string) => readXlsx(fixtureBytes(rel));
