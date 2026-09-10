import assert from 'node:assert/strict';
import {readFileSync, rmSync} from 'node:fs';
import {join} from 'node:path';
import {Duplex, PassThrough, type Readable} from 'node:stream';
import {test} from 'node:test';

import {scratchDir} from '../../../scripts/repo.ts';
import {isOneCellAnchor} from '../../core/image.ts';
import {partsOf, partText} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';

// A 1×1 transparent PNG: enough bytes to prove the streamed media round-trips verbatim.
const ONE_PX_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

// Drain a Node readable into one buffer, resolving once it ends.
function drain(stream: NodeJS.ReadableStream): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    stream.on('data', (chunk: Buffer) => chunks.push(chunk));
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
}

// The committed package, for a writer that keeps one. `commit` resolves with `undefined` when the
// caller supplied a sink and never touched `.stream`; the two tests that assert that contract call
// `commit` directly, and everywhere below the bytes are what the test is about.
async function committed(writer: WorkbookStreamWriter): Promise<Uint8Array> {
  const bytes = await writer.commit();
  assert.ok(bytes !== undefined, 'commit resolved with the package');
  return bytes;
}

test('a streamed workbook reloads to the same sheet names and cell values as a whole-file write', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  for (let i = 1; i <= 20; i++) sheet.addRow([`r${i}`, i]).commit();
  sheet.commit();
  const bytes = await committed(writer);

  const workbook = readXlsx(bytes);
  assert.deepEqual(
    workbook.worksheets.map((s) => s.name),
    ['S'],
  );
  const reread = workbook.getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'r1');
  assert.equal(reread.getCell('B1').value, 1);
  assert.equal(reread.getCell('A20').value, 'r20');
});

test('addRows appends a batch identically to adding rows one at a time', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRows([
    ['a', 1],
    ['b', 2],
  ]);
  sheet.commit();
  const workbook = readXlsx(await committed(writer));
  const reread = workbook.getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.rowCount, 2);
  assert.equal(reread.getCell('A1').value, 'a');
  assert.equal(reread.getCell('B1').value, 1);
  assert.equal(reread.getCell('A2').value, 'b');
  assert.equal(reread.getCell('B2').value, 2);
});

test('a row added to a committed sheet is rejected with a legible "already committed" error', () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['a']).commit();
  sheet.commit();
  assert.throws(() => sheet.addRow(['b']), /already committed/i);
  assert.throws(() => sheet.getCell('A2'), /already committed/i);
});

test('getCell into a committed row is refused rather than emitting a second row of that number', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['a']).commit();
  sheet.addRow(['b']).commit();
  // An eager writer has already rendered row 1 and released its cells, so this cell could only reach
  // the package as a duplicate <row r="1">, which is what the guard exists to prevent.
  assert.throws(() => sheet.getCell('B1'), /row 1, which is already committed/);
  sheet.getCell('B3').value = 'live';
  sheet.commit();

  const sheetXml = partText(await committed(writer), 'xl/worksheets/sheet1.xml');
  assert.deepEqual(
    [...sheetXml.matchAll(/<row r="(\d+)"/g)].map((m) => m[1]),
    ['1', '2', '3'],
    'each row number appears exactly once, and an uncommitted row is still writable',
  );
});

test('writer.stream.pipe(dest) returns dest and delivers the whole package', async () => {
  const writer = new WorkbookStreamWriter();
  const sink = new PassThrough();
  const drained = drain(sink);
  const returned = writer.stream.pipe(sink);
  assert.equal(returned, sink, 'pipe(dest) must return dest so .pipe(out).on(...) composes');

  const sheet = writer.addWorksheet('S');
  sheet.addRow(['a', 'b']).commit();
  sheet.commit();
  const bytes = await committed(writer);

  const piped = await drained;
  assert.deepEqual(Uint8Array.from(piped), bytes, 'the piped bytes match the committed package');
  const reread = readXlsx(piped).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

test('fullCalcOnLoad set through calcProperties is emitted; unset it is absent', async () => {
  const withFlag = new WorkbookStreamWriter();
  withFlag.calcProperties.fullCalcOnLoad = true;
  withFlag.addWorksheet('S').getCell('A1').value = 1;
  const flagged = await committed(withFlag);

  const without = new WorkbookStreamWriter();
  without.addWorksheet('S').getCell('A1').value = 1;
  const plain = await committed(without);

  assert.match(partText(flagged, 'xl/workbook.xml'), /fullCalcOnLoad="1"/);
  assert.doesNotMatch(partText(plain, 'xl/workbook.xml'), /fullCalcOnLoad/);
});

test('useSharedStrings pools streamed string cells into a shared table that reads back intact', async () => {
  const writer = new WorkbookStreamWriter({useSharedStrings: true});
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['dup']).commit();
  sheet.addRow(['dup']).commit();
  sheet.addRow(['other']).commit();
  sheet.commit();
  const bytes = await committed(writer);

  const sst = partText(bytes, 'xl/sharedStrings.xml');
  assert.match(sst, /uniqueCount="2"/);
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /t="s"><v>0<\/v>/);

  const reread = readXlsx(bytes).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'dup');
  assert.equal(reread.getCell('A3').value, 'other');
});

test('a streamed workbook without the option keeps strings inline and writes no shared table', async () => {
  const writer = new WorkbookStreamWriter();
  writer.addWorksheet('S').addRow(['inline']).commit();
  const bytes = await committed(writer);

  assert.throws(() => partText(bytes, 'xl/sharedStrings.xml'), /expected part/);
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /t="inlineStr"/);
});

test('commit over a caller-supplied PassThrough sink resolves and delivers a valid package', async () => {
  const sink = new PassThrough();
  const chunks: Buffer[] = [];
  sink.on('data', (chunk: Buffer) => chunks.push(chunk));

  const writer = new WorkbookStreamWriter({stream: sink});
  writer.addWorksheet('S').addRow(['a', 'b']).commit();
  const bytes = await writer.commit();

  const delivered = Buffer.concat(chunks);
  assert.ok(delivered.length > 0, 'the sink received bytes');
  // Nobody asked for the archive: the sink is where it went, and `.stream` was never touched. Holding
  // every chunk to concatenate them at the end would cost a second copy of the whole package, which
  // is the memory a caller passes a sink to avoid.
  assert.equal(bytes, undefined, 'the package was streamed, not retained');
  const reread = readXlsx(delivered).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

test('a sink plus a touched stream still resolves with the package', async () => {
  // `stream` is a `PassThrough` this writer owns, so a caller who reached for it may be draining it
  // and may not; the bytes are retained either way rather than deciding for them.
  const sink = new PassThrough();
  sink.resume();
  const writer = new WorkbookStreamWriter({stream: sink});
  const drained = drain(writer.stream);
  writer.addWorksheet('S').addRow(['a']).commit();
  const bytes = await writer.commit();
  assert.ok(bytes !== undefined, 'the archive comes back');
  assert.deepEqual(
    Uint8Array.from(await drained),
    bytes,
    'and it is the same bytes the stream carried',
  );
});

test('commit over a Duplex sink resolves: completion does not depend on the writer owning the stream', async () => {
  const chunks: Buffer[] = [];
  const sink = new Duplex({
    read() {},
    write(chunk: Buffer, _enc, cb) {
      chunks.push(chunk);
      cb();
    },
  });

  const writer = new WorkbookStreamWriter({stream: sink});
  writer.addWorksheet('S').addRow(['a']).commit();
  await writer.commit();

  const reread = readXlsx(Buffer.concat(chunks)).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

// The stream's terminal event: 'error' with its message, or 'end'. A stream that settles neither way
// reports 'open' after a short grace period, so the assertion below names what went wrong instead of
// hanging the suite on a promise that will never resolve. The grace period is only ever paid by a
// regression: a stream that does settle settles in the same tick as the commit it belongs to.
function terminalEvent(stream: Readable): Promise<{event: string; message: string}> {
  return new Promise((resolve) => {
    stream.on('error', (error: Error) => resolve({event: 'error', message: error.message}));
    stream.on('end', () => resolve({event: 'end', message: ''}));
    setTimeout(() => resolve({event: 'open', message: ''}), 250).unref();
    stream.resume();
  });
}

// A title carrying a character XML 1.0 cannot spell. Core properties have no `_xHHHH_` convention to
// escape it into, so this is refused when the package is serialised rather than when it is assigned:
// the way to make a commit fail after the caller already holds the output stream.
const UNSPELLABLE_TITLE = `bad${String.fromCharCode(1)}title`;
const UNSPELLABLE = /XML 1\.0 has no representation/;

test('a commit that fails destroys the writer-owned stream instead of leaving it open forever', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = 'ok';
  writer.properties.title = UNSPELLABLE_TITLE;
  sheet.commit();

  const terminal = terminalEvent(writer.stream);
  await assert.rejects(writer.commit(), UNSPELLABLE);

  const outcome = await terminal;
  // Not `end`: a consumer told the stream ended cleanly would treat a truncated package as whole.
  assert.equal(outcome.event, 'error', 'the failure reaches the stream as a failure');
  assert.match(outcome.message, UNSPELLABLE);
});

test('a commit that fails destroys a caller-supplied sink, so an outer await on it settles', async () => {
  const sink = new PassThrough();
  const writer = new WorkbookStreamWriter({stream: sink});
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = 'ok';
  writer.properties.title = UNSPELLABLE_TITLE;
  sheet.commit();

  const terminal = terminalEvent(sink);
  await assert.rejects(writer.commit(), UNSPELLABLE);

  const outcome = await terminal;
  assert.equal(outcome.event, 'error', 'the sink the caller is awaiting learns the commit failed');
  assert.match(outcome.message, UNSPELLABLE);
  assert.equal(sink.destroyed, true);
});

test('a successful commit still ends the stream cleanly, with no error event', async () => {
  const writer = new WorkbookStreamWriter();
  writer.addWorksheet('S').addRow(['a']).commit();

  const terminal = terminalEvent(writer.stream);
  const bytes = await committed(writer);

  assert.deepEqual(await terminal, {event: 'end', message: ''});
  assert.equal(readXlsx(bytes).getWorksheet('S')?.getCell('A1').value, 'a');
});

test('commit to a valid filename writes a re-openable package to disk', async () => {
  // Under the repo's own `.tmp/`, and a fresh directory rather than a name built from the process
  // id. A pid collides between two concurrent checkouts, and a throw before the `finally` leaves the
  // file in the system temp directory for good; `.tmp/` is git-ignored, inspectable, and swept.
  const dir = scratchDir('stream-commit');
  try {
    const target = join(dir, 'out.xlsx');
    const writer = new WorkbookStreamWriter({filename: target});
    writer.addWorksheet('S').addRow(['a']).commit();
    await writer.commit();

    const reread = readXlsx(readFileSync(target)).getWorksheet('S');
    assert.ok(reread);
    assert.equal(reread.getCell('A1').value, 'a');
  } finally {
    rmSync(dir, {recursive: true, force: true});
  }
});

test('commit to an unopenable filename rejects with the underlying I/O error rather than hanging', async () => {
  // A path whose parent directory does not exist cannot be opened for writing; the write stream errors
  // on a later tick and commit must surface it.
  const dir = scratchDir('stream-bad');
  try {
    const badPath = join(dir, 'no-such-dir', `${'x'.repeat(300)}.xlsx`);
    const writer = new WorkbookStreamWriter({filename: badPath});
    writer.addWorksheet('S').addRow(['a']).commit();

    await assert.rejects(writer.commit(), (err: NodeJS.ErrnoException) => {
      assert.match(String(err.code ?? err.message), /ENOENT|ENAMETOOLONG/);
      return true;
    });
  } finally {
    rmSync(dir, {recursive: true, force: true});
  }
});

test('supplying both a stream and a filename is rejected at construction', () => {
  // Mutual exclusivity is a compile-time error for a typed caller (see WorkbookStreamWriterOptions'
  // discriminated union); this exercises the runtime backstop that still catches an untyped JS caller.
  const options = {stream: new PassThrough(), filename: 'out.xlsx'} as ConstructorParameters<
    typeof WorkbookStreamWriter
  >[0];
  assert.throws(() => new WorkbookStreamWriter(options), /either a stream or a filename/i);
});

test('shared-formula slave cells authored on the stream reload populated, not empty', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('yua');
  for (let i = 1; i <= 10; i++) sheet.getCell(`A${i}`).value = i * 10;
  sheet.getCell('B1').value = {formula: 'A1*2', result: 20};
  for (let j = 2; j <= 10; j++) sheet.getCell(`B${j}`).value = {sharedFormula: 'B1'};
  sheet.commit();

  const reread = readXlsx(await committed(writer)).getWorksheet('yua');
  assert.ok(reread);
  const master = reread.getCell('B1').value;
  assert.ok(
    master && typeof master === 'object' && 'formula' in master,
    'master is a formula cell',
  );
  const slave = reread.getCell('B3').value;
  assert.ok(slave && typeof slave === 'object', 'slave reloads as a value object, not empty');
  assert.ok('sharedFormula' in slave || 'formula' in slave, 'slave carries a (shared) formula');
});

test('a streamed sheet emits <conditionalFormatting> before <hyperlinks>, per the CT_Worksheet sequence', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addConditionalFormatting({
    ref: 'A1:A10',
    rules: [
      {
        type: 'expression',
        formulae: ['MOD(ROW(),2)=0'],
        style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFEEEEEE'}}},
      },
    ],
  });
  sheet.addRow(['x']).commit();
  sheet.commit();

  const xml = partText(await committed(writer), 'xl/worksheets/sheet1.xml');
  const posCf = xml.indexOf('<conditionalFormatting');
  const posHl = xml.indexOf('<hyperlinks');
  assert.ok(posCf >= 0 && posHl >= 0, 'both blocks are present');
  assert.ok(posCf < posHl, 'conditionalFormatting precedes hyperlinks');
});

test('a streamed sheet emits <dataValidations> before <hyperlinks>, per the CT_Worksheet sequence', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
  sheet.addRow(['r']).commit();
  sheet.commit();

  const xml = partText(await committed(writer), 'xl/worksheets/sheet1.xml');
  const posDv = xml.indexOf('<dataValidations');
  const posHl = xml.indexOf('<hyperlinks');
  assert.ok(posDv >= 0 && posHl >= 0, 'both blocks are present');
  assert.ok(posDv < posHl, 'dataValidations precedes hyperlinks');
});

test('streamed conditional formatting and data validations reload through the tolerant reader', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
  sheet.addConditionalFormatting({
    ref: 'A1:A10',
    rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [3], priority: 1}],
  });
  sheet.addRow(['r']).commit();
  sheet.commit();

  const reread = readXlsx(await committed(writer)).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.dataValidations.length, 1, 'the data validation survives');
  assert.equal(reread.conditionalFormattings.length, 1, 'the conditional formatting survives');
  assert.equal(reread.conditionalFormattings[0]?.rules[0]?.type, 'cellIs');
});

test('authoring conditional formatting or a data validation on a committed streamed sheet is rejected legibly', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.commit();
  assert.throws(() => sheet.addConditionalFormatting({ref: 'A1', rules: []}), /already committed/);
  assert.throws(
    () => sheet.addDataValidation('A1', {type: 'list', formulae: ['"a"']}),
    /already committed/,
  );
  await committed(writer);
});

test('an image anchored on the stream reloads with its anchor and bytes intact', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  sheet.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
  sheet.commit();

  const bytes = await committed(writer);
  const reloaded = readXlsx(bytes);
  const [image] = reloaded.getWorksheet('S')?.images ?? [];
  const anchor = image?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 0, row: 5, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.to, {col: 2, row: 8, colOff: 0, rowOff: 0});
  const media = reloaded.getImage(image?.imageId ?? -1);
  assert.strictEqual(media?.extension, 'png');
  assert.deepStrictEqual(media?.data, ONE_PX_PNG, 'the streamed media bytes survive verbatim');
});

test('a streamed image emits the drawing, media, and <drawing> reference like a buffered write', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  sheet.addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  sheet.commit();

  const bytes = await committed(writer);
  const names = Object.keys(partsOf(bytes));
  assert.ok(names.includes('xl/drawings/drawing1.xml'), 'a drawing part is streamed');
  assert.ok(names.includes('xl/media/image1.png'), 'the media bytes are streamed');
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /<drawing r:id="[^"]+"\/>/);
  const contentTypes = partText(bytes, '[Content_Types].xml');
  assert.match(contentTypes, /Extension="png" ContentType="image\/png"/);
  assert.match(contentTypes, /PartName="\/xl\/drawings\/drawing1\.xml"/);
});

test('one streamed image anchored on two sheets is stored as a single media part', async () => {
  const writer = new WorkbookStreamWriter();
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  writer.addWorksheet('A').addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  writer.addWorksheet('B').addImage(id, {tl: {col: 3, row: 3}, br: {col: 4, row: 4}});

  const files = partsOf(await committed(writer));
  const mediaParts = Object.keys(files).filter((n) => n.startsWith('xl/media/'));
  assert.strictEqual(mediaParts.length, 1, 'the shared image is streamed once');
});

test('registering an image on a committed streamed workbook is rejected legibly', async () => {
  const writer = new WorkbookStreamWriter();
  writer.addWorksheet('S').commit();
  await committed(writer);
  assert.throws(() => writer.addImage({buffer: ONE_PX_PNG, extension: 'png'}), /already committed/);
});

test('a streamed sheet carrying both protection and an autofilter emits them in CT_Worksheet order', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['H1', 'H2']).commit();
  sheet.addRow(['a', 'b']).commit();
  sheet.autoFilter = 'A1:B1';
  sheet.protect('pw', {});
  sheet.commit();

  const bytes = await committed(writer);
  const xml = partText(bytes, 'xl/worksheets/sheet1.xml');
  const posProtection = xml.indexOf('<sheetProtection');
  const posAutoFilter = xml.indexOf('<autoFilter');
  assert.ok(posProtection >= 0, 'the streamed sheet carries sheetProtection');
  assert.ok(posAutoFilter >= 0, 'the streamed sheet carries autoFilter');
  assert.ok(
    posProtection < posAutoFilter,
    '<sheetProtection> precedes <autoFilter> per the schema',
  );
  // The whole package still reloads, and the autofilter survives.
  assert.deepEqual(readXlsx(bytes).getWorksheet('S')?.autoFilter, {ref: 'A1:B1', columns: []});
});

test('the streamed sheet autoFilter getter reflects what was set', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  assert.strictEqual(sheet.autoFilter, undefined);
  sheet.autoFilter = 'A1:C3';
  assert.deepEqual(sheet.autoFilter, {ref: 'A1:C3', columns: []});
});

test('setting protection or an autofilter on a committed streamed sheet is rejected legibly', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  sheet.commit();
  assert.throws(() => sheet.protect('pw', {}), /already committed/);
  assert.throws(() => {
    sheet.autoFilter = 'A1:B2';
  }, /already committed/);
});

test('committing a streamed row evicts its cells from the model, bounding peak memory', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['a', 'b']);
  assert.ok(sheet.model.hasCell(1, 1), 'the cell is live while the row is being styled');
  row.commit();
  assert.equal(sheet.model.hasCell(1, 1), false, 'commit released the cell graph');
  assert.equal(sheet.model.hasCell(1, 2), false);
  // rowCount survives the eviction, so the next append lands below the evicted row rather than reusing
  // its number: the correctness hazard that makes owning the row counter necessary.
  assert.equal(sheet.rowCount, 1);
  sheet.addRow(['c']).commit();
  const reread = readXlsx(await committed(writer)).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
  assert.equal(
    reread.getCell('A2').value,
    'c',
    'the second append did not collide with the evicted first',
  );
});

test('a fully committed streamed sheet holds no live cells at commit: every row was evicted', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  for (let i = 1; i <= 50; i++) sheet.addRow([i, `v${i}`]).commit();
  for (let i = 1; i <= 50; i++)
    assert.equal(sheet.model.hasCell(i, 1), false, `row ${i} was freed`);
  assert.equal(sheet.rowCount, 50, 'rowCount still reflects every appended row');
  const reread = readXlsx(await committed(writer)).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A50').value, 50);
  assert.equal(reread.getCell('B1').value, 'v1');
});

test('a streamed sheet reports the outline depth of rows already evicted', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['summary']).commit();
  const detail = sheet.addRow(['detail']);
  sheet.model.getRow(2).outlineLevel = 2;
  detail.commit();
  // <sheetFormatPr> is written at commit, long after both rows were flushed and freed, so the depth has
  // to have been carried across the eviction rather than re-read off the model.
  assert.match(
    partText(await committed(writer), 'xl/worksheets/sheet1.xml'),
    /<sheetFormatPr [^>]*\boutlineLevelRow="2"/,
  );
});

test('with useSharedStrings a streamed row stays live until commit: the shared pool defeats bounding', async () => {
  const writer = new WorkbookStreamWriter({useSharedStrings: true});
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['x']).commit();
  assert.ok(
    sheet.model.hasCell(1, 1),
    'shared-strings mode keeps the row live rather than evicting it',
  );
  const reread = readXlsx(await committed(writer)).getWorksheet('S');
  assert.equal(reread?.getCell('A1').value, 'x');
});

test('committing a row that carries a shared-formula cell is rejected legibly', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  const row = sheet.addRow([{sharedFormula: 'A1'}]);
  assert.throws(() => row.commit(), /shared-formula/);
});

test('committing a streamed row twice does not duplicate it in the sheet', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['only']);
  row.commit();
  row.commit();
  const xml = partText(await committed(writer), 'xl/worksheets/sheet1.xml');
  assert.equal(
    (xml.match(/<row /g) ?? []).length,
    1,
    'the row appears exactly once despite the double commit',
  );
});

test('a styled row committed eagerly round-trips its fill through the shared style registry', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['tinted']);
  const cell = row.cells[0];
  assert.ok(cell);
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  row.commit();

  const fill = readXlsx(await committed(writer))
    .getWorksheet('S')
    ?.getCell('A1').fill;
  assert.ok(
    fill && fill.type === 'pattern' && fill.pattern === 'solid',
    'the eagerly-flushed row kept its fill',
  );
});

test('a live getCell row and a committed appended row serialise in ascending order and both reload', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = 'live'; // row 1, never committed → stays in the model
  sheet.addRow(['flushed']).commit(); // row 2, serialised and evicted
  const bytes = await committed(writer);

  const xml = partText(bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('r="1"') < xml.indexOf('r="2"'),
    'the live row 1 precedes the flushed row 2',
  );
  const reread = readXlsx(bytes).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'live');
  assert.equal(reread.getCell('A2').value, 'flushed');
});

// Committing a row evicts it, which keeps the model small and the append point cheap to find. With
// `useSharedStrings` on there is no eviction (the strings stay interned until commit), so every
// appended row is still live when the next one asks where to land, and appending was quadratic while
// the used range was derived by scanning: 8k rows took a second where 2k took fifty milliseconds.
//
// That property used to be guarded here by a wall-clock ratio between two sizes, which failed a
// `verify` run on nothing but a neighbouring gate holding the CPU. It is now counted instead of
// timed, one layer down, where the seam to count at exists: `core/used-extent.test.ts` hands
// `UsedExtent` a row map that reports every row it looks at, and pins the answer at one lookup
// however many rows the sheet already holds. This writer numbers its appended rows through exactly
// that reading (`Math.max(#lastRow, sheet.rowCount) + 1`), so what is left to assert here is the
// numbering itself, which the eviction cases above and below already do.
