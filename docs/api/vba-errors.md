# Vba Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `VbaAuthorError`

<sub>class</sub>

Thrown when a structural edit of a VBA project (removing a module, adding a library reference) is
asked to produce something that cannot be encoded to a well-formed container: a stream name longer
than the [MS-CFB] 31-character limit, a duplicate stream name, or a project so large it would exceed
the writer's single-header DIFAT bound. This is a caller-side contract violation, distinct from
[`VbaParseError`](./vba-errors.md#vbaparseerror) (which reports a malformed blob *read* from an untrusted file).

It shares `code: 'authoring'` with `AuthoringError` rather than being one, and it is the only
subsystem in the tree that got its own authoring class. The reason is symmetry with
[`VbaParseError`](./vba-errors.md#vbaparseerror) and not the VBA subsystem's importance: a caller who catches the parse error
to skip a macro project it cannot read is the same caller who catches the author error to skip one
it cannot write, and the two halves of that pair have to be nameable the same way. A CSV or table
authoring failure has no read-side twin to pair with, which is why those stay `AuthoringError`.

```ts
class VbaAuthorError extends XlsxError {
  override readonly name = 'VbaAuthorError';
  override readonly code = 'authoring';
}
```

---

### `VbaParseError`

<sub>class</sub>

Thrown when a VBA project (`vbaProject.bin`) is present but cannot be parsed: a malformed
compound file, a corrupt compressed stream, or a `dir` record that does not conform to [MS-OVBA].
A workbook with no macros never produces this: [`Workbook.vbaProject`](./workbook.md#workbookvbaproject) is `undefined` instead.

The parser treats the blob as hostile input (a spreadsheet library parses untrusted files), so every
malformed structure fails closed with this error rather than crashing, hanging, or over-allocating.

```ts
class VbaParseError extends XlsxError {
  override readonly name = 'VbaParseError';
  override readonly code = 'malformed-input';
}
```
