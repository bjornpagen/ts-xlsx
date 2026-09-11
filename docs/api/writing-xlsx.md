# Writing .xlsx

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `WriteOptions`

<sub>interface</sub>

Options controlling how [`writeXlsx`](./writing-xlsx.md#writexlsx) serialises a workbook.

```ts
interface WriteOptions {
  /**
   * Pool string and rich-text cell values into a shared-strings table (`xl/sharedStrings.xml`) that
   * cells reference by index, rather than storing each inline in its cell. Deduplicates repeated text
   * and matches Excel's own storage; a rich value is pooled as a `<si>` of its runs, so its formatting
   * survives. Off by default, which keeps strings inline and omits the part.
   */
  readonly useSharedStrings?: boolean;
}
```

---

### `writeXlsx`

<sub>function</sub>

Serialise a workbook into an `.xlsx` package.

The bytes are a pure function of the workbook: an unchanged model written twice produces two
identical archives, because entry timestamps are pinned to a fixed date rather than taken from the
clock. A committed `.xlsx` therefore only changes when something about it changed.

```ts
function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array;
```

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the workbook has no worksheets (a zero-sheet package is corrupt),
or holds a value the writer cannot yet represent.

---

### `writeXlsxAsync`

<sub>function</sub>

Serialise a workbook into an `.xlsx` package, deflating off the calling thread.

Produces the same package [`writeXlsx`](./writing-xlsx.md#writexlsx) does, byte for byte and entry timestamps included, and
exists for one reason: DEFLATE dominates the cost of writing a large workbook, and [`writeXlsx`](./writing-xlsx.md#writexlsx)
spends all of it on the caller's thread. Here `fflate` deflates each part in a worker, so the event
loop keeps turning (stalls drop from the whole write to tens of milliseconds) and parts compress in
parallel, which on a multi-sheet workbook also finishes sooner. On a single-sheet workbook there is
only one part to deflate, so expect responsiveness rather than speed.

Building the parts still happens on the calling thread; only compression moves. That is why there
is no `readXlsxAsync` mirroring this: reading is dominated by XML parsing and model building, which
no worker can take, and the reader's zip-bomb ceiling is enforced by counting output between
synchronous input slices. See ADR-0024.

```ts
async function writeXlsxAsync(
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array>;
```

**Throws:** [`AuthoringError`](./errors.md#authoringerror) , as a rejection, under the same conditions as [`writeXlsx`](./writing-xlsx.md#writexlsx);
the part-building it shares happens before any worker is involved. A failure raised by the zip
layer itself (including an environment that cannot spawn a worker) propagates unwrapped, exactly
as it does from [`writeXlsx`](./writing-xlsx.md#writexlsx).
