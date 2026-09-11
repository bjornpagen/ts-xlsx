# Reject non-.xlsx input with a clear, typed error, not raw zip internals

Cluster: xlsx-io

## Scenario

A user hands the reader a legacy binary spreadsheet, the old BIFF `.xls` format, or any other file
that is not an OOXML `.xlsx` package. Because `.xlsx` is a ZIP container and `.xls` is not, the zip
layer fails deep in its own internals with an opaque message like _"Can't find end of central
directory : is this a zip file?"_. The user cannot tell from that message that the real problem is a
wrong or unsupported format, and the raw error may even leak absolute local filesystem paths from the
zip layer.

> Spec note, not a corpus case: the desired behavior is an error _contract_ covering message clarity,
> error type and no path leakage, rather than a serialization property the corpus asserts on a
> package. The durable value is the classification of input-format failures and their typed errors.

## Desired behavior

- Given input that is not a valid OOXML `.xlsx` ZIP package, the reader rejects with a **clear,
  format-aware, typed error** rather than leaking a raw zip-parsing failure.
- If the input is a **legacy BIFF `.xls`** file, detectable by the OLE Compound File magic bytes
  `D0 CF 11 E0 A1 B1 1A E1`, the error says the `.xls` binary format is not supported and only
  `.xlsx` and OOXML are handled.
- If the input is **not a ZIP at all**, with no `PK\x03\x04` and no end-of-central-directory record,
  the error says the file is not a valid `.xlsx` package rather than surfacing zip-internals text.
- The error is a **distinct, catchable type or category** (`UnsupportedFormatError` against
  `CorruptPackageError`) so callers can branch programmatically, and it **does not expose absolute
  local filesystem paths** from the zip layer.
- A **corrupt or truncated ZIP** is a different answer from an unrecognised one. The container is the
  right kind of thing; it simply cannot be unpacked. That is `PackageReadError` and
  `'malformed-input'`, alongside the zip-bomb refusal, not `'unsupported-format'`.

## Every message names the check that actually ran

The classification is only useful if the message points at the layer that refused. Three sites, three
answers (`io/opc/sniff-format.ts`, `io/xlsx/read.ts`):

| Input                                         | Error                                            | What the message may say                                                                 |
| --------------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------------------------------------- |
| Not a ZIP at all (CSV, `.xls`)                | `UnsupportedFormatError` (`'unknown'` / `'xls'`) | the input is not a ZIP, _not_ that a workbook part is missing, since none was looked for |
| `PK`-headed, the zip layer rejects it         | `PackageReadError`                               | the container is corrupt or truncated                                                    |
| Inflated fine, no `xl/workbook.xml` or `.bin` | `UnsupportedFormatError('unknown')`              | no OOXML workbook part was found, the one place that sentence is true                    |

Note the middle row is narrower than "any corrupt-looking `PK` blob": a streaming unzip silently
_skips_ a stub it cannot make an entry out of rather than failing, so a few hand-made `PK` bytes
inflate to zero parts and land in the third row. The fixture that exercises the real path is a
genuine package cut off mid-stream, which is also what a corrupt file looks like in the wild.

The underlying zip text is discarded, never wrapped and never attached as `cause`: it can name
internals, or an absolute filesystem path, and `cause` is printed by default for an uncaught error.
The typed classification exists precisely so no lower-layer string has to travel.

## Open questions

- How much format sniffing is in scope beyond `.xls`? Detect `.xlsb` (a ZIP but binary-parts),
  `.ods`, and CSV-handed-to-the-xlsx-reader, and give each a tailored message?
- Where does detection live? A small magic-byte probe in front of the zip layer, so a non-ZIP fails
  fast before the zip library runs?
- Error taxonomy: one `UnsupportedFormatError` with a `format` field, or distinct subclasses?

Related: `path-reader-is-node-only-clear-error`, `load-accepts-arraybuffer-and-typed-arrays`,
`tolerant-parse-unclosed-vml-tags`.
