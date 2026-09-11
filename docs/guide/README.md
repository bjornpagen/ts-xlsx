# Getting started

`ts-xlsx` reads and writes spreadsheet documents. The buffered path is synchronous and
speaks `Uint8Array`, so writing a workbook is a function call that hands you bytes, and
reading one is a function call that hands you a model.

```shell
npm install @shbernal/ts-xlsx
```

Node 24 or later, or any modern browser bundler. ESM only.

## Five lines

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('People');
sheet.addRow(['Name', 'Joined']);
sheet.addRow(['Ada', new Date('2026-01-01')]);

const bytes: Uint8Array = writeXlsx(workbook);
console.log(readXlsx(bytes).requireWorksheet('People').getCell('A2').value); // 'Ada'
```

`writeXlsx` returns the file. Putting it somewhere is your platform's job, not the
library's:

```ts
import {mkdtempSync, readFileSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
workbook.addWorksheet('Sheet1').getCell('A1').value = 'hello';

const path = join(mkdtempSync(join(tmpdir(), 'ts-xlsx-')), 'book.xlsx');
writeFileSync(path, writeXlsx(workbook));

// A Node Buffer is a Uint8Array, so it goes straight back in.
const reopened = readXlsx(readFileSync(path));
console.log(reopened.requireWorksheet('Sheet1').getCell('A1').value); // 'hello'
```

## The four things you will touch

| Type               | What it is                                                               |
| ------------------ | ------------------------------------------------------------------------ |
| `Workbook`         | The document. Sheets, defined names, images, theme, document properties. |
| `Worksheet`        | One sheet. Cells, rows, columns, merges, panes, tables, page setup.      |
| `Row` and `Column` | Handles onto one line of the grid. Reading one creates nothing.          |
| `Cell`             | One cell. `cell.value` is the whole story, and it is precisely typed.    |

`Row`, `Column` and `Cell` are handles rather than snapshots. They read and write straight
through to the sheet, so two handles on the same cell always agree, and asking about row 500
of an empty sheet costs nothing and does not extend the used range.

## Where to go next

- [Writing a workbook](./writing.md), if you are producing a file.
- [Reading a workbook](./reading.md), if you are consuming one, including what the model
  carries without interpreting.
- [Styles and number formats](./styles-and-formats.md) for fonts, fills, borders and the
  format codes.
- [Structure](./tables-and-structure.md) for merges, widths, frozen panes and tables.
- [Streaming](./streaming.md) when the workbook is larger than the memory you want to spend.
- [Errors](./errors.md) for what can fail and what each kind means for your code.
- [In the browser](./browser.md) for what runs in a tab and what does not.
- [Migrating from ExcelJS](../migrating-from-exceljs.md) if you are coming across.

The [API reference](../api/README.md) is generated from the public types, so it cannot
describe a shape the compiler would reject. This guide is the part that explains why you
would reach for one.
