import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsb} from '../xlsb/read.ts';
import {readSheetRows} from '../xlsx/read-rows.ts';
import {readXlsx} from '../xlsx/read.ts';
import {writeXlsx} from '../xlsx/write.ts';
import {capturePartClosure, packageAccessors, readPartRelationships} from './read-opc.ts';

// An externalLink part points at its source workbook through a `TargetMode="External"` relationship.
// The closure must keep that wiring verbatim, since dropping it (as it once did) orphans the link and
// dangles every `[n]` external reference a formula resolves through, while never trying to walk into
// the out-of-package target.
test('capturePartClosure retains an external relationship verbatim without walking it', () => {
  const rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" ' +
    'Target="C:\\Sources\\Linked.xlsm" TargetMode="External"/></Relationships>';
  const files = {
    'xl/externalLinks/externalLink1.xml': strToU8('<externalLink/>'),
    'xl/externalLinks/_rels/externalLink1.xml.rels': strToU8(rels),
  };
  const {partText, partBytes} = packageAccessors(files);
  const closure = capturePartClosure(
    'xl/externalLinks/externalLink1.xml',
    partText,
    partBytes,
    () => 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
  );

  assert.ok(closure !== undefined, 'the entry part is present');
  assert.strictEqual(closure.length, 1, 'the external target is not visited as a package part');
  const [entry] = closure;
  assert.deepStrictEqual(
    entry?.rels,
    [
      {
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
        targetPath: 'C:\\Sources\\Linked.xlsm',
        external: true,
      },
    ],
    'the external relationship is kept with its raw target and the external flag',
  );
});

test('packageAccessors answers a differently cased path with the package entry it found', () => {
  // OPC part names compare case-insensitively. `partKey` is what an editor uses to replace or delete
  // the part a read found, so it has to name the package's own entry, not the spelling it was asked for.
  const files = {
    'xl/VbaProject.bin': Uint8Array.of(1),
    'xl/workbook.xml': strToU8('<workbook/>'),
  };
  const {partBytes, partKey} = packageAccessors(files);
  assert.deepStrictEqual(partBytes('xl/vbaProject.bin'), Uint8Array.of(1));
  assert.strictEqual(partKey('xl/vbaProject.bin'), 'xl/VbaProject.bin');
  assert.strictEqual(partKey('XL/WORKBOOK.XML'), 'xl/workbook.xml');
  assert.strictEqual(
    partKey('xl/workbook.xml'),
    'xl/workbook.xml',
    'an exact spelling is its own key',
  );
  assert.strictEqual(partKey('xl/styles.xml'), undefined);
});

// `readPartRelationships` is the single parse of a part's rels every sheet-part lookup goes through, so
// the four queries it answers are pinned here rather than through eight readers that each used to
// re-parse the XML for themselves.

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function sheetWithRels(relationships: string): (path: string) => string | undefined {
  const files = {
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${RELS_NS}">${relationships}</Relationships>`,
  };
  return (path) => (files as Record<string, string>)[path];
}

test('readPartRelationships resolves each target against the owning part, not the package root', () => {
  const rels = readPartRelationships(
    'xl/worksheets/sheet1.xml',
    sheetWithRels(
      `<Relationship Id="rId1" Type="${TYPE}/table" Target="../tables/table1.xml"/>` +
        `<Relationship Id="rId2" Type="${TYPE}/table" Target="/xl/tables/table2.xml"/>`,
    ),
  );
  assert.deepStrictEqual(rels.targetPaths('table'), [
    'xl/tables/table1.xml',
    'xl/tables/table2.xml',
  ]);
  assert.strictEqual(rels.targetPath('table'), 'xl/tables/table1.xml');
});

test('readPartRelationships matches a type on its trailing segment, not the whole URI', () => {
  // The suffix match is what lets a foreign package name the type in a namespace of its own; a
  // whole-URI comparison would drop the relationship and silently lose the part it reaches.
  const rels = readPartRelationships(
    'xl/worksheets/sheet1.xml',
    sheetWithRels(
      `<Relationship Id="rId1" Type="http://example.invalid/rel/printerSettings" Target="../printerSettings/printerSettings1.bin"/>`,
    ),
  );
  assert.strictEqual(rels.targetPath('printerSettings'), 'xl/printerSettings/printerSettings1.bin');
});

test('readPartRelationships answers an id lookup with the raw target, unresolved', () => {
  // A hyperlink's target is a URL, so resolving it against the sheet's directory would corrupt it.
  const rels = readPartRelationships(
    'xl/worksheets/sheet1.xml',
    sheetWithRels(
      `<Relationship Id="rId9" Type="${TYPE}/hyperlink" Target="https://example.invalid/a" TargetMode="External"/>`,
    ),
  );
  assert.strictEqual(rels.byId('rId9')?.target, 'https://example.invalid/a');
  assert.strictEqual(rels.byId('rId9')?.external, true);
  assert.strictEqual(rels.byId('missing'), undefined);
});

test('a part with no rels part reads as an empty relationship set, not a failure', () => {
  const rels = readPartRelationships('xl/worksheets/sheet9.xml', () => undefined);
  assert.deepStrictEqual(rels.records, []);
  assert.deepStrictEqual(rels.targetPaths('table'), []);
  assert.strictEqual(rels.targetPath('drawing'), undefined);
});

test('readPartRelationships reads the rels part once, however many queries follow', () => {
  // The point of the type: the sheet loop asks it eight questions per sheet, and a re-read per
  // question is exactly what this replaced.
  let reads = 0;
  const partText = sheetWithRels(
    `<Relationship Id="rId1" Type="${TYPE}/drawing" Target="../drawings/drawing1.xml"/>`,
  );
  const rels = readPartRelationships('xl/worksheets/sheet1.xml', (path) => {
    reads += 1;
    return partText(path);
  });
  rels.targetPath('drawing');
  rels.targetPath('comments');
  rels.targetPaths('table');
  rels.byId('rId1');
  assert.strictEqual(reads, 1);
});

// `openSpreadsheetPackage` is the one place the inflate bound is defaulted and applied, so the
// property worth pinning is that every entry point built on it enforces the same bound at the same
// byte. Before it, each of the three carried its own `?? DEFAULT_MAX_UNCOMPRESSED`, and a reader
// that dropped it or reached for a different constant would have been caught by nothing.

/** The bomb-guard message if `run` raised it, else null. A reader may fail for its own reasons at a
 * bound it accepts (an `.xlsb` reader handed an `.xlsx` package, say); only the guard is under test. */
function zipBombMessage(run: () => void): string | null {
  try {
    run();
    return null;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return /possible zip bomb/.test(message) ? message : null;
  }
}

test('every reader entry point enforces the inflate bound at the same byte', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addRow(['x']);
  const bytes = writeXlsx(workbook);
  const inflated = Object.values(unzipSync(bytes)).reduce((total, part) => total + part.length, 0);

  const entryPoints: readonly [string, (cap: number) => void][] = [
    ['readXlsx', (cap) => void readXlsx(bytes, {maxUncompressedBytes: cap})],
    ['readXlsb', (cap) => void readXlsb(bytes, {maxUncompressedBytes: cap})],
    ['readSheetRows', (cap) => void [...readSheetRows(bytes, {maxUncompressedBytes: cap})]],
  ];

  for (const [name, read] of entryPoints) {
    assert.notStrictEqual(
      zipBombMessage(() => read(inflated - 1)),
      null,
      `${name} refuses a package one byte over its bound`,
    );
    assert.strictEqual(
      zipBombMessage(() => read(inflated)),
      null,
      `${name} accepts a package exactly at its bound`,
    );
  }
});
