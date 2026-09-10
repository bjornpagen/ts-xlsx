import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import {test} from 'node:test';
import {fileURLToPath} from 'node:url';

import {partBytes} from '../xlsx/package.test-support.ts';
import {parseStyleTable} from './read-styles.ts';

// The BIFF12 half of the label check in `../xlsx/read-styles.test.ts`: a `BrtStyle`'s name and
// builtinId are invisible to a round trip, so nothing but a direct assertion on the parse can tell a
// reader that keeps them from one that drops them on the floor.
//
// It runs against the real binary package rather than a hand-framed record because the label's
// encoding (a flag word deciding whether the gallery index means anything, and a wide string after
// it) is exactly the part a hand-written fixture would get wrong in the same direction the reader
// does. The workbook model cannot be asked instead: a package whose only named style is Normal never
// reaches `restoreNamedStyles`, which is why this gap went unwatched.
const FIXTURE = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../../../test/corpus/fixtures/xlsb-binary-workbook-reads-like-its-xlsx-twin/source.xlsb',
);

test("a BrtStyle's label reaches the style table, name and builtinId both", () => {
  const styles = partBytes(readFileSync(FIXTURE), 'xl/styles.bin');
  const {namedStyles} = parseStyleTable(styles);
  assert.deepEqual(
    namedStyles.map((style) => [style.name, style.builtinId]),
    [['Normal', 0]],
  );
});
