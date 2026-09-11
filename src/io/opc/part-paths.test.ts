import assert from 'node:assert/strict';
import {test} from 'node:test';

import {extensionOf, relativePartPath, resolveRelativePart} from './part-paths.ts';

// A part path comes verbatim from an untrusted package (both a rels Target and a raw media path),
// so this must only recognise a `.` after the last `/`: a dot in a directory segment is not a file
// extension, even when the filename itself has none.
test('extensionOf only recognises a `.` after the last `/`, not one in an earlier segment', () => {
  assert.strictEqual(extensionOf('xl/media/image1.png'), 'png');
  assert.strictEqual(
    extensionOf('xl/media.v2/image1'),
    '',
    'a dot in a directory segment is not a file extension, even if the filename has none',
  );
});

// Path resolution is a hostile-input parser path: a relationship Target comes verbatim from an
// untrusted package. These pin the OPC-legal shapes a well-formed writer never emits (absolute
// package-root targets and `.`/`..`/empty segments) so a real or malicious file that uses them
// still resolves to a bounded part path.

test('resolveRelativePart treats a leading slash as package-root-absolute', () => {
  assert.strictEqual(
    resolveRelativePart('xl/worksheets/sheet1.xml', '/xl/media/image1.png'),
    'xl/media/image1.png',
    'the base directory is ignored and the leading slash is stripped',
  );
});

test('resolveRelativePart collapses `..` and doubled slashes in an absolute target too', () => {
  // An absolute target used to return with only its leading slash stripped, so both of these came back
  // as keys no part answers to, and the sheet or stylesheet they name read as missing.
  assert.strictEqual(
    resolveRelativePart('xl/workbook.xml', '/xl/../xl/worksheets/sheet1.xml'),
    'xl/worksheets/sheet1.xml',
  );
  assert.strictEqual(resolveRelativePart('xl/workbook.xml', '//xl/styles.xml'), 'xl/styles.xml');
  assert.strictEqual(
    resolveRelativePart('xl/workbook.xml', '/../../xl/styles.xml'),
    'xl/styles.xml',
    'a `..` at the root stays at the root',
  );
});

test('resolveRelativePart collapses `.` and `..` segments against the base directory', () => {
  assert.strictEqual(
    resolveRelativePart('xl/worksheets/sheet1.xml', '../media/./image1.png'),
    'xl/media/image1.png',
    '`..` pops the parent and `.` is dropped',
  );
});

test('resolveRelativePart drops empty segments from a doubled slash', () => {
  assert.strictEqual(
    resolveRelativePart('xl/drawings/drawing1.xml', 'sub//child.xml'),
    'xl/drawings/sub/child.xml',
    'the empty segment between the slashes is skipped',
  );
});

test('a workbook target escaping `xl/` resolves to a part path outside it', () => {
  // The workbook's own relationships resolve against the workbook part like any other part's do, so a
  // target that climbs out of `xl/` lands where OPC says it does. A resolver that only prefixed `xl/`
  // would yield a path with the `..` still in it, which no package part can ever answer to.
  assert.strictEqual(
    resolveRelativePart('xl/workbook.xml', '../docProps/custom.xml'),
    'docProps/custom.xml',
  );
});

// The two functions are inverses, and that is the property the writer's output depends on: every
// Target the writer names through `relativePartPath` is read back by the reader (and by Excel)
// through the resolution `resolveRelativePart` implements. Either one drifting alone re-points a
// relationship at a part that is not there, which no schema check catches because the markup stays
// well-formed. The pairs below cover the shapes a package actually produces: same directory,
// down into a child, up and across a sibling, up to the root, and down from the root.
const REFERENCE_PAIRS: readonly (readonly [from: string, to: string])[] = [
  ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'],
  ['xl/workbook.xml', 'xl/worksheets/sheet1.xml'],
  ['xl/drawings/drawing1.xml', 'xl/media/image1.png'],
  ['xl/workbook.xml', 'docProps/custom.xml'],
  ['_rels/.rels', 'xl/workbook.xml'],
  ['xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png'],
  ['docProps/app.xml', 'docProps/core.xml'],
];

test('a target named relative to its referencing part resolves back to the part it named', () => {
  for (const [from, to] of REFERENCE_PAIRS) {
    const target = relativePartPath(from, to);
    assert.strictEqual(
      resolveRelativePart(from, target),
      to,
      `${from} -> ${to} was written as "${target}", which did not resolve back`,
    );
  }
});
