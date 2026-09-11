# OOXML reading must be namespace-agnostic, BOM-tolerant, and order-independent

## The scenario

Spreadsheet files are produced by far more than desktop Excel: server-side
generators, LibreOffice and Calc, reporting stacks, and libraries in other languages.
These files are valid OOXML but exercise corners Excel itself never emits. A reader
built by matching literal tag names against Excel's exact output, rather than by
(namespace, local-name), chokes on them, even though every mainstream consumer opens
them fine and Excel repairs them silently on open.

Three concrete shapes recur in real-world reports:

1. **Namespace-prefixed roots.** A part's root element and its children carry an
   explicit prefix instead of the default-namespace form. `xl/workbook.xml` may be
   `<x:workbook …><x:sheets><x:sheet …/></x:sheets></x:workbook>`; `docProps/app.xml`
   may be `<properties:Properties xmlns:properties="…extended-properties">…`. A
   tag-literal reader never enters `<sheets>` or `<Properties>`, so the model is built
   empty and the next access throws (`Cannot read properties of undefined (reading
'sheets')`, or an undefined extended-property dereference).

2. **A leading byte-order mark or preamble.** A part begins with a UTF-8 BOM, or other
   insignificant bytes, before `<?xml`. A strict parser fed the raw bytes reports
   _"Non-whitespace before first tag"_. Files that traveled over email or chat, or were
   saved by tools that prepend a BOM, hit this, and re-saving in Excel, which normalizes
   the preamble, is the only current workaround.

3. **Unusual zip entry ordering.** `xl/workbook.xml` appears in the archive after a
   worksheet part. A reader that assumes a fixed part order parses a sheet before the
   workbook model exists. This particular shape is already handled correctly today
   and is locked by the corpus; it is recorded here so the rebuild does not regress
   the guarantee.

In every case the data is valid, since Excel repairs and re-saves the file cleanly, so the
defect is purely on the read side.

## Desired behaviour

- Every OOXML part is matched by **(namespace URI, local-name)**, never by a literal
  prefixed string. `x:workbook`, `workbook`, and any other prefix bound to the
  spreadsheetml namespace are the same element to the reader.
- The reader **skips a leading BOM and insignificant whitespace** before the XML
  declaration in every part, decoding as if the preamble were absent.
- Reading does **not depend on the order** of entries within the package: the workbook
  model is available, or lazily resolvable, before any worksheet is materialized,
  regardless of zip order.
- Non-ASCII sheet names, defined names, and text survive read unchanged.

## Root cause (legacy)

The SAX-style parsers key handlers off literal local tag names as emitted by Excel and
assume Excel's canonical part ordering and preamble. There is no namespace resolution
layer and no BOM or preamble skip, so any generator that differs on prefix, ordering, or
preamble produces an empty model and a downstream crash.

## Prior art

- The failure family is well attested across foreign generators (SpreadsheetLight,
  LibreOffice and Calc, assorted server libraries), and the common thread in every report is
  "opens everywhere else, Excel repairs it, this reader crashes."
- The fix that repeatedly works in the community is to normalize on namespace and to
  strip the BOM before parsing, that is, to parse against the schema rather than against
  Excel's byte-for-byte output.

## Open questions for the rebuild

- Which BOMs to honor (UTF-8 against UTF-16LE and BE) and whether other non-whitespace bytes
  before the first tag should hard-fail or be tolerated leniently.
- Whether to carry a full namespace-prefix map per part or to canonicalize by
  stripping prefixes bound to known OOXML namespaces at the tokenizer boundary.
- How strict-mode parsing interacts with this leniency, since a strict or diagnostic mode may
  want to _report_ that a file used a non-canonical shape even while reading it.
- Interaction with the write side: files we read leniently must still be written back
  in the canonical unprefixed form (see
  [[excel-repair-on-open-structural-constraints]]).
