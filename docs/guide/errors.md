# Errors

A spreadsheet library spends its life reading files someone else wrote, so failing well is
part of the job rather than an afterthought. Every error this library raises deliberately
descends from `XlsxError`, and every one of those carries a `code` saying which kind of
failure it was.

```ts
import {readXlsx, XlsxError} from '@shbernal/ts-xlsx';

try {
  readXlsx(new Uint8Array([1, 2, 3, 4]));
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof XlsxError); // true
  console.log(error instanceof XlsxError ? error.code : 'not ours'); // 'unsupported-format'
}
```

`error instanceof XlsxError` is the one-line answer to "was that us?". Everything else is a
bug in your code or in your runtime, and this library will not have wrapped it.

## The four codes, and what you would do about each

The taxonomy is deliberately coarse, because four is the number of genuinely different next
actions.

| `code`                 | What happened                                      | What you do                                    |
| ---------------------- | -------------------------------------------------- | ---------------------------------------------- |
| `'unsupported-format'` | Not a container this library reads at all.         | Reject the file, or hand it to something else. |
| `'malformed-input'`    | A part we do read is corrupt or off-specification. | Reject the file. It is broken, or hostile.     |
| `'authoring'`          | Your code described a document that cannot exist.  | Fix the calling code.                          |
| `'internal'`           | An invariant the library maintains did not hold.   | Report it. The bug is ours.                    |

Every subclass fixes `code` to a literal, so the hierarchy is a discriminated union and
narrowing on `error.code` narrows the type.

```ts
import {readXlsx, XlsxError} from '@shbernal/ts-xlsx';

function classify(bytes: Uint8Array): string {
  try {
    readXlsx(bytes);
    return 'read';
  } catch (error) {
    if (!(error instanceof XlsxError)) throw error;
    switch (error.code) {
      case 'unsupported-format':
        return 'not a spreadsheet we read';
      case 'malformed-input':
        return 'broken or hostile';
      case 'authoring':
        return 'our bug';
      case 'internal':
        return 'their bug, please report it';
    }
  }
}

console.log(classify(new Uint8Array([0, 0, 0, 0]))); // 'not a spreadsheet we read'
```

## Telling apart the two ways a file can be wrong

`UnsupportedFormatError` says the input is a different _kind_ of thing. `PackageReadError`
says it is the right kind and cannot be unpacked. Keeping them apart is what lets you answer
"should I try another reader, or reject this?", and the first carries a `format` field so you
never have to match on a message.

```ts
import {readXlsx, UnsupportedFormatError} from '@shbernal/ts-xlsx';

// The magic bytes of a legacy OLE2 .xls.
const legacy = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);

try {
  readXlsx(legacy);
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof UnsupportedFormatError); // true
  console.log(error instanceof UnsupportedFormatError ? error.format : ''); // 'xls'
}
```

`format` is `'xls'`, `'xlsb'` or `'unknown'`. A message never carries a filesystem path or
the zip library's internals, and that is on purpose: the classification is what the caller
sees, so no lower-layer string can leak into a log or a user-facing error.

## Errors that are your fault

`AuthoringError` is raised when the _document you described_ cannot exist: a workbook with no
worksheets, a table whose columns do not span its range, a merge overlapping another.

```ts
import {AuthoringError, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

try {
  writeXlsx(new Workbook()); // a package with no worksheets is corrupt
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof AuthoringError); // true
  console.log(error instanceof AuthoringError ? error.code : ''); // 'authoring'
}
```

Note where the line falls. A single scalar that is out of range or the wrong type stays a
native `RangeError` or `TypeError`, because those types exist for exactly that and wrapping
them would make this taxonomy a re-implementation of the language's. `getColumn(0)` is a
`RangeError`; a table that names a column twice is an `AuthoringError`.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('S');
try {
  sheet.getColumn(0);
  console.log('that should not have worked');
} catch (error) {
  console.log(error instanceof RangeError); // true
}
```

## `InternalError` is not a failure mode to handle

It means an invariant this library maintains did not hold. No caller can provoke one, so
there is nothing to catch it for. It is a distinct type so that "unreachable" is stated
rather than implied by a bare `Error`, and it is the one class that appends a "please report
this" notice to its own message. Every other class leaves the message exactly as given,
because `'malformed-input'` is a routine outcome for a library that reads untrusted files,
and a report-this banner on every corrupt input would train you to ignore the one banner that
always means something.

## Where each class is documented

The reference pages are generated from the types, so they cannot drift:

- [`errors`](../api/errors.md) for `XlsxError`, `AuthoringError` and `InternalError`;
- [`opc-errors`](../api/opc-errors.md) for `UnsupportedFormatError` and `PackageReadError`;
- [`xlsx-errors`](../api/xlsx-errors.md), [`xlsb-errors`](../api/xlsb-errors.md) and
  [`xml-errors`](../api/xml-errors.md) for the per-codec parse failures.

Every one of them lives in the `@shbernal/ts-xlsx/errors` subpath and nowhere else, because a
container-level failure belongs to no single codec. A service that only classifies failures
can import that subpath alone and pay 12 KB rather than a parser.

Next: [in the browser](./browser.md).
