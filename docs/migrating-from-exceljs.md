# Migrating from ExcelJS

`ts-xlsx` is a hard fork of ExcelJS, but **it is not a drop-in replacement** and does
not try to be. The API is deliberately different because the goal of the fork was the
_right_ shape, not the familiar one (see [`CLAUDE.md`](../CLAUDE.md) §1). This page maps
the common ExcelJS patterns to their `ts-xlsx` equivalents so a port is mechanical, and
is honest about what has not been rebuilt yet.

Treat this as a translation guide, not a compatibility promise. Pin a version; the API
is still moving toward a `0.x` release.

## The three shifts that cover most code

### 1. I/O is synchronous and byte-native: free functions, not `workbook.xlsx.*`

ExcelJS routed I/O through async methods on the workbook that assumed Node `Buffer`s and
streams. `ts-xlsx` reads and writes plain `Uint8Array` synchronously, so the same call
works in Node and the browser and there is nothing to `await` for the buffered path.

```ts
// ExcelJS
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('in.xlsx');
await wb.xlsx.writeFile('out.xlsx');
const buf = await wb.xlsx.writeBuffer();

// ts-xlsx: I/O is separate from the model, synchronous, and Uint8Array in/out
import {Workbook, readXlsx, writeXlsx} from '@shbernal/ts-xlsx';
import {readFileSync, writeFileSync} from 'node:fs';

const wb = readXlsx(readFileSync('in.xlsx'));
writeFileSync('out.xlsx', writeXlsx(wb));
const bytes = writeXlsx(wb); // Uint8Array
```

Reading and writing files from disk is the caller's job. `ts-xlsx` never touches the
filesystem, which is what keeps it browser-safe.

### 2. Cell values are one precisely-typed union

ExcelJS overloaded `cell.value` loosely and expressed types like formulas, hyperlinks,
and rich text as ad-hoc object shapes. In `ts-xlsx`, `cell.value` is the single
[`CellValue`](api/cell-values.md) union: `number | string | boolean | Date | null`
plus typed formula, rich-text, hyperlink, and error shapes. `null` is the empty cell.

```ts
sheet.getCell('A1').value = 42;
sheet.getCell('A2').value = new Date();
sheet.getCell('A3').value = {formula: 'SUM(A1:A2)', result: 42};
sheet.getCell('A4').value = null; // empty
```

### 3. Absent axes are `undefined`, never sentinels

A whole-row reference (`$1`) has no column; a whole-column reference (`$A:$A`) has no
row. ExcelJS let those decay into `NaN`/`"undefined"` and leak into serialized addresses.
`ts-xlsx` models an omitted axis as `undefined` on
[`CellAddress`](api/addresses-ranges.md) and never fabricates a sentinel.

## Quick reference

| ExcelJS                         | ts-xlsx                                                       |
| ------------------------------- | ------------------------------------------------------------- |
| `await wb.xlsx.readFile(path)`  | `readXlsx(readFileSync(path))`                                |
| `await wb.xlsx.load(buffer)`    | `readXlsx(bytes)`                                             |
| `await wb.xlsx.writeFile(path)` | `writeFileSync(path, writeXlsx(wb))`                          |
| `await wb.xlsx.writeBuffer()`   | `writeXlsx(wb)` → `Uint8Array`                                |
| `await wb.csv.readFile(path)`   | `readCsv(readFileSync(path, 'utf8'))`                         |
| `await wb.csv.writeBuffer()`    | `writeCsv(wb)` / `writeCsvText(wb)`                           |
| `wb.addWorksheet('S')`          | `wb.addWorksheet('S')` _(unchanged)_                          |
| `wb.getWorksheet('S')`          | `wb.getWorksheet('S')` → `Worksheet \| undefined`             |
| `sheet.getCell('A1').value = …` | `sheet.getCell('A1').value = …` _(unchanged)_                 |
| `sheet.addRow([…])`             | `sheet.addRow([…])` _(unchanged)_                             |
| streaming `WorkbookReader`      | `readSheetRows(bytes, {sheet})` / `readWorkbookStream(bytes)` |

Where a method name is unchanged, its types are still stricter. `getWorksheet` returns
`Worksheet | undefined`, so handle the miss, and the arguments are precisely typed.

## What is not here (yet)

The rewrite is corpus-driven: an API lands only once it is strict-typed and pinned by
tests. Some ExcelJS features are still on the way, and the buffered writer refuses a
value it cannot represent faithfully rather than emitting a lossy package. If you depend
on a feature not yet in the [API reference](api/README.md), then, per the project's
working agreement (see [`architecture.md`](architecture.md)), a missing behavior is best
reported as a corpus case so it is fixed once and never regresses.

**Charts, vector shapes, slicers, and legacy form controls are round-trip-only, by
decision, not by omission.** ExcelJS let you create these from scratch, with varying
fidelity. `ts-xlsx` preserves them byte-faithfully through a load/edit/save but has no
authoring API for any of the four. See [`docs/api/preserved.md`](api/preserved.md) and
[ADR-0014](decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md) for the
reasoning and what it would take to pick one up.

## Why break compatibility at all?

Because keeping the old surface would make the library worse but easier, and replacing it
makes the library better but harder, and the fork exists to do the harder, better thing
([`CLAUDE.md`](../CLAUDE.md) §5). What you get for the port: `.d.ts` types precise enough
that the API reference is generated from them, an I/O path with no Node built-ins in it
so the same code runs in a browser, and a behavior set that is pinned by the regression
corpus rather than assumed.
