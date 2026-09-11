# Streaming

The buffered path builds the whole workbook in memory. That is the right default, and it
stops being the right default at the point where the workbook is larger than the memory you
are willing to spend on it. Streaming trades the model for bounded memory: you see rows go
past, once, in order, and nothing accumulates.

Reach for it when the file is large and your job is a pass over the data. Stay with
`readXlsx` and `writeXlsx` when you need to look at a cell twice, when you need styles, or
when the file is small enough that none of this matters.

## Reading rows without building a workbook

```ts
import {readSheetRows, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('People');
built.addRow(['Name', 'Age']);
built.addRow(['Ada', 36]);
built.addRow(['Grace', 45]);
const bytes = writeXlsx(source);

let total = 0;
for (const row of readSheetRows(bytes, {sheet: 'People'})) {
  const age = row.cells[1]?.value;
  if (typeof age === 'number') total += age;
}
console.log(total); // 81
```

`readSheetRows` is a generator, so nothing is read until you ask for the next row and
nothing is kept once you have moved past it. Two things about what it yields are worth
knowing, because both differ from the buffered model:

- only rows the sheet actually declares appear, so a gap in the numbering is a gap in the
  file;
- within a row, only the cells that carry something appear, because a blank or style-only
  cell is not data.

`sheet` takes a name or a 1-based position, and defaults to the first sheet.

## Reading every sheet

```ts
import {readWorkbookStream, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
source.addWorksheet('First').addRow(['a']);
source.addWorksheet('Second').addRow(['b']);
const bytes = writeXlsx(source);

const names: string[] = [];
for (const sheet of readWorkbookStream(bytes)) {
  names.push(sheet.name);
}
console.log(names); // ['First', 'Second']
```

This is the streaming analogue of walking `readXlsx(bytes).worksheets`. Each yielded sheet
lets you stream its rows and read its merge and hidden-column summaries, without the model
behind it ever existing.

## Writing a workbook incrementally

The streaming writer is the mirror: add a sheet, append rows, and commit each one to
serialise it and let it go. It comes from `@shbernal/ts-xlsx/node` rather than from the root
specifier, because it is the one part of the library that opens files and pipes Node streams;
[in the browser](./browser.md) says what that boundary buys everything else.

```ts
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

const writer = new WorkbookStreamWriter();
const sheet = writer.addWorksheet('Big');
for (let n = 1; n <= 1000; n++) sheet.addRow([n, n * n]).commit();
sheet.commit();

const bytes = await writer.commit();
console.log(bytes !== undefined && bytes.length > 0); // true
```

`commit()` on a row serialises it and frees it, which is what keeps peak memory flat; a row
you never commit is a row still in memory. `commit()` on the sheet freezes it, and any
further mutation after that is rejected with a legible error rather than silently accepted.
`commit()` on the writer assembles the package.

`commit()` is therefore also a deadline. A row stays reachable through `getCell` for as long
as it is uncommitted, which is when to style it or fill a cell you skipped; once committed,
its `<row>` is rendered and its cells are released, so the same call is refused rather than
emitting a second row carrying that number.

```ts
import {AuthoringError} from '@shbernal/ts-xlsx';
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

const writer = new WorkbookStreamWriter();
const sheet = writer.addWorksheet('Report');
const header = sheet.addRow(['Region', 'Revenue']);
sheet.getCell('B1').value = 'Revenue (EUR)'; // still live, so still editable
header.commit();

let refused = false;
try {
  sheet.getCell('B1').value = 'too late';
} catch (error) {
  refused = error instanceof AuthoringError;
}
console.log(refused); // true

sheet.commit();
const written = await writer.commit();
console.log(written !== undefined && written.length > 0); // true
```

This half of the API is asynchronous where the buffered path is synchronous, and the bytes
arrive two ways: as the resolved value above, and through `writer.stream`, a Node `Readable`.
The second is the one that matters on a server, because it means a workbook can go straight
to an HTTP response without ever being one object in memory:

<!-- sample: illustrative. It needs a live HTTP response object, which a guide has none of. -->

```ts
// Inside a request handler:
const writer = new WorkbookStreamWriter();
writer.stream.pipe(response);
const sheet = writer.addWorksheet('Export');
for (const record of records) sheet.addRow([record.id, record.total]).commit();
sheet.commit();
await writer.commit();
```

`commit()` resolves with the package _only when something is going to want it_: when you
supplied no sink, or when you touched `writer.stream`. Hand the writer a `stream` or a
`filename` and never reach for `writer.stream`, as the handler above does not, and it resolves
with `undefined` and the archive is never assembled as a whole object. That is the point of
passing a sink, and the return type says so rather than handing back a copy of everything you
just streamed away.

## What streaming costs you

It is a genuinely narrower surface, and pretending otherwise would waste your afternoon.

- The streaming **writer is Node-only**, which is why it is published from `/node` and not
  from the root specifier: it writes to a Node stream. A browser build that imports it links
  a module that throws by name rather than one a bundler cannot resolve.
  [In the browser](./browser.md) says what runs in a tab.
- The streaming **reader does not read `.xlsb`**. A binary package cannot be row-streamed,
  and asking raises `UnsupportedFormatError` with `format` saying which kind it was, rather
  than a vague refusal.
- You do not get a `Worksheet`. `getCell` is there, but only ahead of the commit line, and
  the whole-sheet reads (`usedRange`, `getRange`, `hasCell`, `actualRowCount`) and the
  structural edits (`spliceRows`, `insertRow`, the column verbs) are absent altogether: a
  streamed sheet could only answer them over whichever rows happened to still be in memory,
  which is a plausible answer rather than a true one.

If you are streaming because writing is slow rather than because it is large, look at
`writeXlsxAsync` in [writing a workbook](./writing.md) first: it keeps the whole buffered
API and just moves compression off your thread.

Next: [errors](./errors.md), which is where `UnsupportedFormatError` above is explained.
