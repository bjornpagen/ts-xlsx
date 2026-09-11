# Reading a workbook

`readXlsx` takes bytes and gives you a `Workbook`. It is synchronous, and it accepts both
OOXML serialisations: an XML `.xlsx` and a binary `.xlsb`. The two are told apart from the
package itself rather than from a file extension, so you never branch on which one you are
holding, and the model you get is the same either way.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const original = new Workbook();
const built = original.addWorksheet('People');
built.addRow(['Name', 'Joined']);
built.addRow(['Ada', new Date('2026-01-01')]);

const workbook = readXlsx(writeXlsx(original));
const sheet = workbook.requireWorksheet('People');
console.log(sheet.getCell('A2').value); // 'Ada'
console.log(sheet.usedRange?.address); // 'A1:B2'
```

`getWorksheet` returns `undefined` for a name that is not there. `requireWorksheet` throws
instead, which is the one you want when the sheet's absence is a bug in your input rather
than a case to handle.

## Walking a sheet

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Data');
built.addRow(['a', 1]);
built.addRow(['b', 2]);
built.addRow(['c', 3]);

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Data');

for (const row of sheet.rows()) {
  console.log(row.number, row.values);
}
console.log(sheet.actualRowCount); // 3
```

`rows()` yields only the rows the sheet declares, so a sparse sheet costs what it holds
rather than what its bounds imply. `usedRange` is the rectangle those cells occupy, and it
is `undefined` for a sheet with nothing in it.

To read a rectangle rather than a sheet, ask for a `Range`:

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Grid');
built.addRow([1, 2, 3]);
built.addRow([4, 5, 6]);

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Grid');
const block = sheet.getRange('A1:C2');
console.log(block.cellCount); // 6
console.log(block.cells.map((cell) => cell.value)); // [1, 2, 3, 4, 5, 6]
```

## What a cell reads back as

The same union you wrote. A cell you never touched is `null`, not `undefined` and not an
empty string, and `cellValueToText` is the total function over the union when you want text
and would rather not switch on nine kinds yourself.

```ts
import {cellValueToText, readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Kinds');
built.getCell('A1').value = 1234.5;
built.getCell('A2').value = {formula: 'A1*2', result: 2469};
built.getCell('A3').value = {error: '#DIV/0!'};

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Kinds');
console.log(sheet.getCell('A1').type); // 'number'
console.log(cellValueToText(sheet.getCell('A2').value)); // '2469'
console.log(sheet.getCell('B9').value); // null
```

Note what `cellValueToText` gives you for a formula: the cached _result_, not the formula
source, because the source is not text the sheet ever displayed. And note what it does not
do. It is the value's text, with no number format applied, because the format lives on the
style and this function is handed only the value. A currency cell has no currency sign here.
This library models number formats; it does not implement a formatting engine, and
[styles and number formats](./styles-and-formats.md) says what that means for you.

## What the model does not interpret

A real workbook contains parts this library does not model: a pivot cache, a slicer, a
chart, a linked-workbook reference. It reads those bytes, keeps them, and writes them back
untouched, so a load and save does not quietly delete them. They surface as _preserved_
references, and each one carries the package parts it reaches.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
source.addWorksheet('Plain').getCell('A1').value = 'nothing exotic';

const workbook = readXlsx(writeXlsx(source));
// A workbook this library wrote has nothing it had to preserve.
console.log(workbook.preservedReferences.length); // 0
```

Seeing a preserved part is the system working rather than a gap. A library that modelled
everything would have to be finished before it could be safe; preserving what it does not
model is what lets it be safe first. What is preserved rather than modelled today is
recorded in [the preserved-content reference](../api/preserved.md), and the decision behind
it in
[ADR-0014](../decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md).

## Reading untrusted input

Assume every file you did not write is hostile, because the format makes that cheap for an
attacker. Two defences are on by default and are not optional.

Inflation is bounded by counting the bytes actually produced, never by trusting the
archive's declared sizes, so a zip bomb that lies about its uncompressed size is refused all
the same. The ceiling defaults to 512 MiB and you can lower it:

```ts
import {PackageReadError, readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Big');
for (let row = 1; row <= 200; row++) built.addRow([row, 'x'.repeat(200)]);
const bytes = writeXlsx(source);

try {
  readXlsx(bytes, {maxUncompressedBytes: 1024});
  console.log('no bound was enforced');
} catch (error) {
  console.log(error instanceof PackageReadError); // true
}
```

XML entities are decoded but never expanded, so the billion-laughs family of attacks has
nothing to expand into. Neither defence needs configuring, and neither can be switched off.

## Why there is no `readXlsxAsync`

Writing has an async form because DEFLATE dominates it and workers can take that off your
thread. Reading is dominated by XML parsing and model building, which no worker can move, so
an async read would advertise a non-blocking call and then block for most of its duration.
The zip-bomb ceiling is also enforced by counting output between synchronous input slices,
and that guarantee weakens the moment inflation moves to a worker.

If you need a read that does not block, run the whole read in a worker.
[ADR-0024](../decisions/0024-async-is-one-writer-not-a-mirrored-pair.md) has the argument in
full.

For a workbook too large to hold in memory at all, see [streaming](./streaming.md).

Next: [styles and number formats](./styles-and-formats.md), or [errors](./errors.md) for
what the failures above mean.
