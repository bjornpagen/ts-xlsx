# Writing a workbook

Build a `Workbook`, then hand it to `writeXlsx`. The call is synchronous and returns the
package as a `Uint8Array`; nothing is written to disk unless you write it.

```ts
import {Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Orders');

sheet.addRow(['Order', 'Placed', 'Total', 'Paid']);
sheet.addRow(['A-1001', new Date('2026-03-04'), 1240.5, true]);
sheet.addRow(['A-1002', new Date('2026-03-05'), 980, false]);

const bytes = writeXlsx(workbook);
console.log(bytes.length > 0); // true
```

## Putting values in cells

There are three ways in, and they are the same underneath.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Sheet1');

// By address.
sheet.getCell('A1').value = 'Name';

// By row, appended after the last one that exists.
sheet.addRow(['Ada', 36]);

// By row and column, when you are iterating.
sheet.getRow(3).getCell('A').value = 'Grace';

console.log(sheet.getCell('A2').value); // 'Ada'
```

`addRow` also takes an object, which is how you write rows without counting columns. Give
the columns keys first:

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('People');
sheet.getColumn(1).key = 'name';
sheet.getColumn(2).key = 'joined';

sheet.addRow({name: 'Ada', joined: new Date('2026-01-01')});
sheet.addRow({name: 'Grace', joined: new Date('2026-02-14')});

console.log(sheet.getCell('A2').value); // 'Grace'
```

## What a cell can hold

`cell.value` is one union, and it is the whole story. There is no separate "type" to set:
the value you assign decides what the cell is.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Kinds');

sheet.getCell('A1').value = 42;
sheet.getCell('A2').value = 'text';
sheet.getCell('A3').value = true;
sheet.getCell('A4').value = new Date('2026-03-04');
sheet.getCell('A5').value = null; // an empty cell
sheet.getCell('A6').value = {formula: 'SUM(A1:A1)', result: 42};
sheet.getCell('A7').value = {error: '#N/A'};
sheet.getCell('A8').value = {
  richText: [{text: 'half ', font: {bold: true}}, {text: 'and half'}],
};

console.log(sheet.getCell('A4').type); // 'date'
```

A hyperlink is not one of those kinds. It sits on the sheet, over one cell or a range, and the cells
beneath it keep whatever values they hold, so a total or a date can carry a link as readily as a label:

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Links');

sheet.getCell('A1').value = 'example';
sheet.getCell('A2').value = 1250;
sheet.addHyperlink({ref: 'A1', target: 'https://example.com'});
sheet.addHyperlink({ref: 'A2', target: '#Links!A1', tooltip: 'Back to the top'});

console.log(sheet.hyperlinkAt('A2')?.target); // '#Links!A1'
```

Two things are worth knowing about formulas. A formula carries an optional cached `result`,
which is the value a consumer sees before anything recalculates; this library computes
nothing, so if you leave the result out, the cell reads as blank until the file is opened by
something that calculates. If you would rather ask every consumer to recalculate, say so on
the workbook:

```ts
import {Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Totals');
sheet.getCell('A1').value = 2;
sheet.getCell('A2').value = 3;
sheet.getCell('A3').value = {formula: 'SUM(A1:A2)'};
workbook.fullCalcOnLoad = true;

console.log(writeXlsx(workbook).length > 0); // true
```

## Several sheets, and document metadata

```ts
import {Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
workbook.properties.title = 'Q1 report';
workbook.properties.creator = 'reporting service';

workbook.addWorksheet('Summary').getCell('A1').value = 'Q1';
workbook.addWorksheet('Detail').getCell('A1').value = 'line items';

console.log(workbook.worksheets.map((sheet) => sheet.name)); // ['Summary', 'Detail']
console.log(writeXlsx(workbook).length > 0); // true
```

A name can be attached to a range, which is what Excel's Name Manager shows and what a
formula elsewhere can refer to:

```ts
import {Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Rates');
sheet.getCell('B1').value = 0.2;
workbook.defineName({name: 'TaxRate', refersTo: 'Rates!$B$1'});

sheet.getCell('B2').value = {formula: '100*TaxRate', result: 20};
console.log(writeXlsx(workbook).length > 0); // true
```

## The bytes are a pure function of the model

An unchanged workbook written twice produces two identical archives: entry timestamps are
pinned to a fixed date rather than taken from the clock. A committed `.xlsx` therefore
changes only when something about it changed, which is what makes one reviewable in a diff
and cacheable by content.

```ts
import {Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const build = () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 'same';
  return workbook;
};

const first = writeXlsx(build());
const second = writeXlsx(build());
console.log(first.length === second.length && first.every((b, i) => b === second[i])); // true
```

## Not blocking the event loop

`writeXlsx` spends the whole cost of DEFLATE on the calling thread, which for a large
workbook is seconds during which a server answers nothing. `writeXlsxAsync` produces the
same package with the compression handed to worker threads.

```ts
import {Workbook, writeXlsxAsync} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Big');
for (let row = 1; row <= 500; row++) sheet.addRow([row, row * row]);

const bytes = await writeXlsxAsync(workbook);
console.log(bytes.length > 0); // true
```

Same package, byte for byte, entry timestamps included. Expect responsiveness always, and
speed only when there is more than one substantial part to deflate in parallel.

There is deliberately no `readXlsxAsync`; [reading](./reading.md) explains why.

## What you cannot author yet

Charts, vector shapes, slicers and legacy form controls have no authoring API. A workbook
that already contains them keeps them through a load and save, byte for byte, but there is
no call that creates one. That is a decision rather than an oversight, and
[ADR-0014](../decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md) records
why. [Reading a workbook](./reading.md) covers what preservation means in practice.

Next: [reading a workbook](./reading.md), or [styles and number
formats](./styles-and-formats.md) to make this one look like something.
