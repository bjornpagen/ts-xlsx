#!/usr/bin/env node

// The OOXML gate: emit the writers' real output and hold it against a frozen baseline,
// using `ooxml-validate`, the shared oracle this project and `ts-pptx` both validate
// against. Everything this harness used to own below the assertions (the .NET build, the
// process spawn, the conformance pin, the report types) belongs to that package now; the
// two repos had each grown their own validator on a different Open XML SDK version, so
// they enforced different rule sets while appearing to enforce the same one.

import assert from 'node:assert/strict';
import {readFile, rm, writeFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

import type JSZipType from 'jszip';
import {
  type ValidationDiagnostic,
  type ValidationResult,
  FILE_FORMAT,
  validate,
  validatorAvailable,
} from 'ooxml-validate';

import {scratchDir} from '../../scripts/repo.ts';
import {Workbook} from '../../src/core/workbook.ts';
import {readXlsx} from '../../src/io/xlsx/read.ts';
import {WorkbookStreamWriter} from '../../src/io/xlsx/write-stream.ts';
import {writeXlsx} from '../../src/io/xlsx/write.ts';

/** The stable subset of a diagnostic used to detect baseline drift. */
type ValidationFingerprint = Pick<ValidationDiagnostic, 'id' | 'type' | 'partUri' | 'xpath'>;

/** Baselined-until-fixed diagnostics, keyed by workbook basename. Empty while the writer is clean:
 * an entry is a *known-open* writer bug we've chosen to track, never a mute button for a new one. */
type Baseline = Readonly<Record<string, readonly ValidationFingerprint[]>>;

const require = createRequire(import.meta.url);
const JSZip = require('jszip') as typeof JSZipType;

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BASELINE = JSON.parse(
  await readFile(path.join(HERE, 'allowed-errors.json'), 'utf8'),
) as Baseline;

// The buffered and both streaming outputs, and a Strict workbook written back, are the packages under
// test: every one must validate against the frozen baseline (empty today, so: clean). A new diagnostic on any of them fails the gate.
const WRITER_FILES = [
  'buffered.xlsx',
  'streaming-inline.xlsx',
  'streaming-shared.xlsx',
  'strict-rewritten.xlsx',
] as const;

// A Strict workbook read and written back: every write is Transitional, so a part carried through as
// bytes has to come out Transitional too. Its theme, spelled in Strict, is one the oracle once could not
// load at all.
async function writeStrictRewrite(file: string): Promise<void> {
  const strict = path.join(HERE, '../corpus/fixtures/strict-mode-iso8601-date/sample.xlsx');
  await writeFile(file, writeXlsx(readXlsx(await readFile(strict))));
}

// Exercise a representative slice of the buffered writer: styled font, data validation, a formula, and
// a table over its own cells with a totals row carrying a custom <totalsRowFormula>, so the oracle sees
// more than a bare grid and validates the totals-row markup against the schema.
async function writeBufferedWorkbook(file: string): Promise<void> {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  Object.assign(sheet.getColumn(1), {key: 'name', width: 20});
  Object.assign(sheet.getColumn(2), {key: 'value', width: 12});
  sheet.addRow(['Name', 'Value']);
  sheet.addRow(['alpha', 42]);
  sheet.getCell('A2').font = {bold: true, color: {argb: 'FF336699'}};
  sheet.addDataValidation('B2:B20', {type: 'whole', operator: 'between', formulae: [0, 100]});
  sheet.getCell('B3').value = {formula: 'SUM(B2:B2)', result: 42};
  sheet.getCell('D1').value = 'Label';
  sheet.getCell('E1').value = 'Amount';
  sheet.getCell('D2').value = 'alpha';
  sheet.getCell('E2').value = 42;
  // The totals row's cells (label + custom SUBTOTAL-less formula) are materialised by addTable, so the
  // oracle checks both the <totalsRowFormula> child and the grid cells it writes.
  sheet.addTable({
    name: 'DataTable',
    ref: 'D1',
    headerRow: true,
    totalsRow: true,
    columns: [
      {name: 'Label', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(DataTable[Amount])*1.1'},
    ],
    rowCount: 1,
  });
  // Every attribute the writer checks against a closed OOXML enumeration, so the oracle rules on the
  // tokens rather than the guard merely agreeing with itself: page orientation and order, three
  // cfRule types with their operator/timePeriod/iconSet spellings, a hyperlink's free-string
  // attributes, a hidden tab, and a non-default window visibility.
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.pageOrder = 'overThenDown';
  sheet.addConditionalFormatting({
    ref: 'B2:B20',
    rules: [
      {type: 'cellIs', operator: 'greaterThan', formulae: [10], priority: 1},
      {type: 'timePeriod', timePeriod: 'lastWeek', priority: 2},
      {
        type: 'iconSet',
        iconSet: '3TrafficLights1',
        cfvo: [
          {type: 'percent', value: 0},
          {type: 'percent', value: 33},
          {type: 'percent', value: 67},
        ],
        priority: 3,
      },
    ],
  });
  sheet.getCell('A5').value = {
    text: 'link',
    hyperlink: 'https://example.invalid/',
    tooltip: 'a tooltip & an ampersand',
  };
  workbook.addWorksheet('Hidden', {state: 'veryHidden'});
  workbook.view.visibility = 'hidden';
  await writeFile(file, writeXlsx(workbook));
}

// The streaming writer must be clean in both string-storage modes: inline (eager per-row flush) and
// shared-strings (whole-workbook pool). They travel different serialisation paths, so both are exercised.
async function writeStreamingWorkbook(file: string, useSharedStrings: boolean): Promise<void> {
  const writer = new WorkbookStreamWriter({useSharedStrings});
  const sheet = writer.addWorksheet('Stream');
  sheet.addRow(['Name', 'Value']).commit();
  sheet.addRow(['alpha', 42]).commit();
  sheet.addRow(['beta', 7]).commit();
  sheet.commit();
  // No sink was supplied, so the writer retains the package and `commit` resolves with it.
  const bytes = await writer.commit();
  if (bytes === undefined) throw new Error('the streaming writer retained no package');
  await writeFile(file, bytes);
}

async function rewritePackage(
  source: string,
  destination: string,
  transform: (zip: JSZipType) => Promise<void>,
): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(source));
  await transform(zip);
  await writeFile(destination, await zip.generateAsync({type: 'nodebuffer'}));
}

// A negative control: inject an element the worksheet schema forbids, so a passing run proves the oracle
// still discriminates rather than rubber-stamping. Derived from a known-clean package.
async function makeSchemaInvalidControl(source: string, destination: string): Promise<void> {
  await rewritePackage(source, destination, async (zip) => {
    const sheetPart = zip.file('xl/worksheets/sheet1.xml');
    assert.ok(sheetPart, 'generated workbook must contain xl/worksheets/sheet1.xml');
    const xml = await sheetPart.async('string');
    const invalid = xml.replace('</worksheet>', '<unexpectedValidatorProbe/></worksheet>');
    assert.notStrictEqual(invalid, xml, 'worksheet close tag must be present');
    zip.file('xl/worksheets/sheet1.xml', invalid);
  });
}

function fingerprint(error: ValidationDiagnostic): ValidationFingerprint {
  return {
    id: error.id,
    type: error.type,
    partUri: error.partUri,
    xpath: error.xpath,
  };
}

async function main(): Promise<void> {
  // The oracle is obtained, not assumed: the package downloads and verifies its binary on
  // first use. Under CI this throws rather than returning false, which is the property
  // that stops a missing oracle from turning the gate into a no-op, and here, where
  // running this command IS asking for the oracle, "unavailable" is a failure either way.
  if (!(await validatorAvailable())) {
    throw new Error('the OOXML oracle is unavailable, so this gate proves nothing');
  }

  const temp = scratchDir('ooxml');
  try {
    const at = (name: string) => path.join(temp, name);
    const invalid = at('invalid.xlsx');
    const truncated = at('truncated.xlsx');
    const unsupported = at('unsupported.txt');

    await writeBufferedWorkbook(at('buffered.xlsx'));
    await writeStreamingWorkbook(at('streaming-inline.xlsx'), false);
    await writeStreamingWorkbook(at('streaming-shared.xlsx'), true);
    await writeStrictRewrite(at('strict-rewritten.xlsx'));
    await makeSchemaInvalidControl(at('buffered.xlsx'), invalid);
    await writeFile(truncated, (await readFile(at('buffered.xlsx'))).subarray(0, 128));
    await writeFile(unsupported, 'not an xlsx');

    const inputs = [...WRITER_FILES.map(at), invalid, truncated];
    const report = await validate(inputs);
    assert.strictEqual(report.format, FILE_FORMAT);
    // Every input appears in the report with an explicit `valid` flag, including the clean
    // ones. Absence is never cleanliness, so a short report is a broken contract, not a pass.
    assert.strictEqual(report.results.length, inputs.length);
    const byName = new Map<string, ValidationResult>(
      report.results.map((result) => [path.basename(result.file), result]),
    );

    for (const name of WRITER_FILES) {
      const result = byName.get(name);
      assert.ok(result, `missing validator result for ${name}`);
      const expected = BASELINE[name] ?? [];
      assert.deepStrictEqual(
        result.errors.map(fingerprint),
        expected,
        `${name} diverged from its baseline: fix the writer, do not baseline a new error`,
      );
      assert.strictEqual(
        result.valid,
        expected.length === 0,
        `${name} validity must match its (empty) baseline`,
      );
    }

    const invalidResult = byName.get('invalid.xlsx');
    assert.strictEqual(invalidResult?.valid, false);
    assert.ok(
      invalidResult?.errors.some(
        (error) => error.type === 'Schema' && error.partUri === '/xl/worksheets/sheet1.xml',
      ),
      'invalid worksheet must produce a structured schema diagnostic',
    );

    // A package that cannot be opened is a *finding* about that file, not a tool failure: it
    // comes back as a result with `valid: false`, which is what keeps a corrupt workbook from
    // reading as a clean one.
    const truncatedResult = byName.get('truncated.xlsx');
    assert.strictEqual(truncatedResult?.valid, false);
    assert.deepStrictEqual(
      truncatedResult?.errors.map((error) => error.id),
      ['PackageOpenError'],
    );
    assert.deepStrictEqual(
      truncatedResult?.errors.map((error) => error.type),
      ['Package'],
    );

    // The other half of that distinction: an input the oracle cannot even dispatch on is a
    // tool failure (its exit 2), and the package surfaces it by rejecting rather than by
    // handing back a report with a made-up verdict in it.
    await assert.rejects(validate([unsupported]), /Unsupported file extension/);

    console.log(
      `ooxml validation (${report.format}, SDK ${report.sdkVersion}): ` +
        'buffered, streaming and Strict-rewrite outputs clean; error controls detected',
    );
  } finally {
    await rm(temp, {recursive: true, force: true});
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? (error.stack ?? error.message) : error);
  process.exitCode = 1;
});
