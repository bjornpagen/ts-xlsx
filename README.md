# ts-xlsx

[![npm](https://img.shields.io/npm/v/%40shbernal%2Fts-xlsx)](https://www.npmjs.com/package/@shbernal/ts-xlsx)
[![weekly downloads](https://img.shields.io/npm/dw/%40shbernal%2Fts-xlsx.svg?label=npm%20downloads&logo=npm)](https://www.npmjs.com/package/@shbernal/ts-xlsx)
[![total downloads](https://img.shields.io/npm/dt/%40shbernal%2Fts-xlsx.svg?label=npm%20total%20downloads&logo=npm)](https://www.npmjs.com/package/@shbernal/ts-xlsx)
[![CI](https://github.com/shbernal/ts-xlsx/actions/workflows/build.yml/badge.svg)](https://github.com/shbernal/ts-xlsx/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A TypeScript-first library for reading and writing spreadsheet documents
(`.xlsx` / OOXML, and CSV) — synchronous, `Uint8Array`-native, and dependency-lean.

**[Documentation, guide and a browser playground](https://shbernal.github.io/ts-xlsx/)**

> **Status — 1.0.0, the first release.**
> `ts-xlsx` began as a hard fork of [ExcelJS](https://github.com/exceljs/exceljs)
> and has been rebuilt from the ground up into a modern, strict-TypeScript library.
> **It carries no backwards-compatibility guarantee with ExcelJS** — the API below is
> its own, not a drop-in. From 1.0.0 onward it follows
> [SemVer](https://semver.org/) against _itself_; see the [changelog](CHANGELOG.md).
> See [`CLAUDE.md`](CLAUDE.md) for the goals and [`docs/architecture.md`](docs/architecture.md)
> for the design, and [migrating from ExcelJS](docs/migrating-from-exceljs.md) if you are coming across.

## Why it exists

Upstream ExcelJS is effectively unmaintained — no release since 2023, a backlog of
hundreds of issues and PRs — while still serving tens of millions of downloads a month.
`ts-xlsx` extracts the accumulated value from that backlog (as a permanent regression
corpus) and discards the debt: no untyped surfaces, no callback APIs, a single small
runtime dependency, and every behavior pinned by a test. This is not a compatibility
shim. It is a different, better library that happens to share ancestry.

What that buys you today:

- **Strict, precise types are the contract.** `strict` + `noUncheckedIndexedAccess` +
  `exactOptionalPropertyTypes`. The published `.d.ts` _is_ the documentation, and the
  [API reference](docs/api/README.md) is generated straight from it.
- **Synchronous, buffer-native I/O.** `readXlsx`/`writeXlsx` take and return a
  `Uint8Array` — no streams to await for the buffered path, no Node `Buffer` assumption,
  so the same code runs in Node and the browser.
- **Streaming when you need it.** Bounded-memory row streaming both ways — read a large
  workbook without materializing it, or write one incrementally to a Node stream so peak
  memory stays flat.
- **One runtime dependency** ([`fflate`](https://github.com/101arrowz/fflate) for zip),
  a hand-written SAX reader with bounded allocation on every parser path, and a
  build-free, strict-typed source tree.

## Install

```shell
npm install @shbernal/ts-xlsx
```

Requires Node ≥ 24 (or any modern browser bundler). ESM only.

## Quick start

The buffered path is synchronous. `writeXlsx` returns the file bytes; `readXlsx` takes them back.

```ts
import {Workbook, writeXlsx, readXlsx} from '@shbernal/ts-xlsx';

// --- write ---
const wb = new Workbook();
const sheet = wb.addWorksheet('People');

sheet.getCell('A1').value = 'Name';
sheet.getCell('B1').value = 'Joined';
sheet.addRow(['Ada', new Date('2026-01-01')]);
sheet.addRow(['Grace', new Date('2026-02-14')]);
sheet.getCell('C1').value = {formula: 'COUNTA(A:A)', result: 3};

const bytes: Uint8Array = writeXlsx(wb);

// --- read it back ---
const reopened = readXlsx(bytes);
const people = reopened.getWorksheet('People');
console.log(people?.getCell('A2').value); // 'Ada'
```

Persisting to disk is your platform's job, not the library's — `writeXlsx` hands you the
bytes:

```ts
import {writeFileSync, readFileSync} from 'node:fs';

writeFileSync('people.xlsx', writeXlsx(wb));
const wb2 = readXlsx(readFileSync('people.xlsx')); // a Buffer is a Uint8Array
```

## Core model

- **`Workbook`** — the document. `addWorksheet(name)`, `getWorksheet(nameOrId)`,
  `worksheets`, defined names, images, and workbook-level properties.
- **`Worksheet`** — a sheet. `getCell('B3')`, `addRow(values)`, `addTable(...)`,
  `mergeCells('A1:B2')`, `getRow(n)` / `getColumn(n)`, `rows()` / `columns()`, page setup,
  and print options.
- **`Row`** / **`Column`** — one line of the grid. Formatting is flat (`row.height = 20`,
  `column.width = 12`, `column.key = 'name'`), and cells are reachable from it:
  `row.getCell('B')`, `row.cells`, `row.values`.
- **`Cell`** — one cell. `cell.value` is the whole story: a `number`, `string`,
  `boolean`, `Date`, `null` (empty), a formula (`{formula, result}`), rich text, or an
  error — all precisely typed as [`CellValue`](docs/api/cell-values.md). A hyperlink is not
  a value: it sits on the sheet, over any cell, through
  `sheet.addHyperlink({ref: 'B2', target: 'https://example.com'})`.

```ts
sheet.getRow(1).height = 20;
sheet.getRow(1).values = ['Name', 'Joined'];
sheet.getColumn(1).width = 24;
sheet.getRow(2).getCell('B').value = new Date('2026-01-01');

for (const row of sheet.rows()) {
  console.log(row.number, row.values);
}
```

`Row` and `Column` are _handles_, not snapshots: they read and write straight through to the
sheet, so two handles on the same line always agree, and reading one creates nothing — asking
about row 500 costs nothing and does not extend the used range. Position is fixed, exactly as a
`Cell`'s is: after a splice, `getRow(3)` still means row 3, now holding whatever moved there.

Addresses are honest: an axis a reference doesn't mention is `undefined`, never a
sentinel — see [`decodeAddress`](docs/api/addresses-ranges.md).

Charts, vector shapes, slicers, and legacy form controls are **round-trip-only**: a
workbook that has them keeps them byte-faithfully through a load/edit/save, but there is
no API to author a new one — see [`docs/api/preserved.md`](docs/api/preserved.md) and
[ADR-0014](docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md).

## Reading, writing, streaming, CSV

```ts
import {
  readXlsx,
  writeXlsx, // buffered .xlsx  (Uint8Array ⇄ Workbook)
  readSheetRows, // stream one sheet's rows, bounded memory
  readWorkbookStream, // stream every sheet, rows one at a time
  readCsv,
  writeCsv, // CSV as Uint8Array
  writeCsvText, // CSV as a string
} from '@shbernal/ts-xlsx';
// The streaming writer opens files and pipes Node streams, so it is the one Node-only face
// and lives behind its own subpath — which is what keeps everything above browser-safe:
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

// Bounded-memory extraction — the whole workbook is never materialized:
for (const row of readSheetRows(bytes, {sheet: 'People'})) {
  console.log(
    row.number,
    row.cells.map((c) => c.value),
  );
}

// Bounded-memory generation — commit each row to serialize and free it as you go:
const writer = new WorkbookStreamWriter();
const out = writer.addWorksheet('Big');
for (let i = 1; i <= 1_000_000; i++) out.addRow([i, i * i]).commit();
out.commit();
const packaged = await writer.commit(); // Uint8Array here; also delivered via writer.stream
```

The streaming writer is asynchronous where the buffered path is synchronous: `commit()`
resolves to the package bytes and simultaneously pipes them through `writer.stream` (a Node
`Readable`), so `writer.stream.pipe(res)` streams a workbook straight to an HTTP response.
Hand the writer its own sink instead (`{stream}` or `{filename}`) and never touch
`writer.stream`, and `commit()` resolves with `undefined`: the archive goes to the sink and is
never assembled as one object, which is what passing a sink was for.

## Writing without blocking the event loop

`writeXlsx` spends the whole cost of DEFLATE on the calling thread — seconds, for a large
workbook, during which a server answers nothing. `writeXlsxAsync` produces the same package
with the compression handed to worker threads:

```ts
import {writeXlsxAsync} from '@shbernal/ts-xlsx';

const bytes = await writeXlsxAsync(wb); // same package, event loop stays live
```

Measured on a ~42 MB part map: one large sheet takes the same wall-clock either way, but the
longest event-loop stall drops from the entire write to ~17 ms; a twenty-sheet workbook also
finishes about 2.4× sooner, because its parts deflate in parallel. Expect responsiveness
always and speed only when there is more than one substantial part.

There is deliberately **no `readXlsxAsync`**. Reading is dominated by XML parsing and model
building, which no worker can take off the calling thread, so it would advertise a
non-blocking read and then block for most of its duration — and the reader's zip-bomb ceiling
is enforced by counting output between synchronous input slices, a guarantee that weakens the
moment inflation moves to a worker. If you need a non-blocking read, run the whole read in a
worker. See [ADR-0024](docs/decisions/0024-async-is-one-writer-not-a-mirrored-pair.md).

The reader decodes untrusted input defensively — entities are decoded but never
expanded, and inflation is bounded by a running output counter rather than any declared
size, so a malformed or hostile package can't exhaust memory.

## Entry points

The bare package name gives you everything, and with a bundler that is the right default:
`"sideEffects": false` is declared, so anything you don't reference is dropped. The subpaths
are for when you'd rather the module graph itself said which half of the library you depend
on — a Lambda with no bundler, a service that only classifies failures:

| Import from                  | You get                                                                        | It loads |
| ---------------------------- | ------------------------------------------------------------------------------ | -------- |
| `@shbernal/ts-xlsx`          | everything except the streaming writer                                         | 501 KB   |
| `@shbernal/ts-xlsx/core`     | `Workbook`, `Worksheet`, `Cell`, styles, values, addresses                     | 185 KB   |
| `@shbernal/ts-xlsx/xlsx`     | `readXlsx`, `writeXlsx`/`writeXlsxAsync`, the streaming reader, VBA part edits | 490 KB   |
| `@shbernal/ts-xlsx/xlsb`     | `readXlsb`                                                                     | 256 KB   |
| `@shbernal/ts-xlsx/csv`      | `readCsv`, `writeCsv`, `writeCsvText`                                          | 191 KB   |
| `@shbernal/ts-xlsx/node`     | `WorkbookStreamWriter` and the rest of the streaming writer                    | 358 KB   |
| `@shbernal/ts-xlsx/vba`      | `parseVbaProject`, `addVbaReference`, `removeVbaModule`                        | 46 KB    |
| `@shbernal/ts-xlsx/customui` | `parseCustomUi` and the ribbon types                                           | 15 KB    |
| `@shbernal/ts-xlsx/errors`   | every error class the library throws                                           | 3 KB     |

`/node` is the only one the root specifier does not carry, and the reason is its imports
rather than its size: it reaches `node:fs` and `node:stream`, and every other entry reaches
no Node built-in at all, which is what makes them bundle for a browser with nothing to
configure ([ADR-0040](docs/decisions/0040-the-browser-boundary-is-an-entry-point.md)).

Every error class lives in `/errors` and nowhere else, because a container-level failure
belongs to no single codec — `readXlsx` and `readXlsb` both raise `UnsupportedFormatError`.
Catching and classifying therefore costs 3 KB, not a parser.

`/xlsx` is barely cheaper than the whole package, and that is honest rather than a defect:
`readXlsx` sniffs the bytes and hands a binary package to the BIFF12 reader, so the `.xlsb`
codec is not optional on that path. The numbers above are the static-import closure of each
entry, measured by `pnpm run size` — a bundler can only shrink them further.

## API reference

The full reference is generated from the public types — it cannot drift from what the
compiler accepts — and lives in **[`docs/api/`](docs/api/README.md)**. Regenerate it with:

```shell
pnpm run docs
```

## Coming from ExcelJS?

Read **[migrating from ExcelJS](docs/migrating-from-exceljs.md)**. The short version: the
shapes are deliberately different (synchronous `Uint8Array` I/O instead of async
`Buffer`/stream methods, `readXlsx`/`writeXlsx` free functions instead of
`workbook.xlsx.*`), because the goal was the _right_ API, not the familiar one.

## Hit a bug? There is a skill for that

Most code that uses this library is written by an agent, and an agent that hits a library
defect will usually route around it silently — so the defect is never reported and never
fixed. `ts-xlsx-upstream` is a skill that turns that moment into a filed issue with a
minimal reproduction, which is what the maintainers convert into a permanent regression
test. It ships inside the package, so it is already on disk:

```shell
npx skills add ./node_modules/@shbernal/ts-xlsx     # offline, matches your installed version
npx skills add shbernal/ts-xlsx                     # or straight from the repo
```

That writes the skill under `.agents/skills/`, links it into whichever agent directories it
finds, and records what it did in `skills-lock.json`. Track the lock file and ignore the
copies: they are derived, `npx skills experimental_install` rebuilds them from the lock, and
`npx skills update` overwrites them in place.

It covers triage (is this bug ours or your file's?), reducing a failure to a script that
builds its own input, and — because spreadsheets carry real data — never uploading a
workbook to a public tracker. Once the reproduction stands on its own, it files without
interrupting you, and tells you the issue number afterwards.

You do not need it to report something: <https://github.com/shbernal/ts-xlsx/issues> is
open, and errors the library knows are its own fault print that link themselves.

## Design principles

This project is optimized to be built and maintained largely by autonomous agents, with
a machine-checkable safety net as the primary guarantor of correctness. The rules that
govern every change are in [`CLAUDE.md`](CLAUDE.md); the design and working agreements are
in [`docs/architecture.md`](docs/architecture.md); notable decisions are recorded under
[`docs/decisions/`](docs/decisions/).

## License

[MIT](LICENSE) — inherited from ExcelJS and retained.
