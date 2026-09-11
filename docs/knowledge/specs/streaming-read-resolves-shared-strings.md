# Streaming read must resolve shared strings regardless of part order

## The scenario

A large `.xlsx` is read in streaming mode to process the first rows without loading the
whole file into memory. String cell values must arrive as their actual text, not as
unresolved shared-string index placeholders. In real files the `sharedStrings.xml` part
and the worksheet parts can appear in either order within the zip, and the streaming reader
consumes entries in the order they appear. When a worksheet entry is reached _before_
the shared-string table has been read, string cells surface as `{ sharedString: N }`
objects, a raw index, instead of the resolved string.

Because the trigger is the zip entry ordering of whichever tool produced the file, the
same consumer code fails on some files and succeeds on others, which makes the defect
look intermittent and hard to pin.

## Desired behaviour

- A streamed string cell is always delivered as its resolved value, a `string` or a
  `{ richText: [...] }` for a rich string, never as a `{ sharedString: N }` index,
  regardless of whether `sharedStrings.xml` precedes or follows the worksheet in the
  archive.
- Rich strings resolve to structured rich text, runs with text and font, preserved
  end-to-end through the streaming path.
- The streaming reader does not require the whole shared-string table to be buffered
  before any row is emitted if that would defeat the memory goal; the design must
  resolve strings correctly _and_ stay bounded (see
  [[bounded-memory-large-workbook-read]]).

## Root cause (legacy)

The streaming reader resolves a cell's shared-string reference only if
`sharedStrings.xml` has already been parsed and cached when the worksheet is processed.
Entries are consumed in zip order with no step that guarantees the shared-string table
is available first, so a worksheet-before-table ordering leaks the raw index to the
consumer. There is no deferred-resolution or two-pass strategy.

## Open questions for the rebuild

- Resolution strategy for the worksheet-before-table case: read the central directory
  and parse `sharedStrings.xml` first, a targeted and cheap seek; emit rows with
  deferred string handles resolved when the table arrives; or fall back to a bounded
  buffer. The central-directory-first approach is simplest and keeps memory bounded.
- Whether the public streaming row event should ever expose an unresolved handle, for
  extreme memory cases, behind an explicit opt-in, or never.
- A regression corpus case needs a fixture whose worksheet entry is deliberately
  ordered before `sharedStrings.xml`. Such a file must be hand-assembled, since no mainstream
  writer emits that order on demand, so this is recorded as design intent until a
  durable fixture exists.
