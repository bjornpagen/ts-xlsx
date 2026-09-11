# Streaming read must emit every worksheet, including large sheet counts

Cluster: streaming

## Scenario

A workbook with a large number of worksheets, well over 100 and one report cites about 188, is read
with the streaming reader. The streaming reader should emit every worksheet the workbook declares, in
order, and complete without stalling. In practice the many-sheet path is unreliable: worksheets
are dropped, or enumeration crashes.

> A synthesized 150-sheet workbook fed through the streaming reader crashes in the worksheet
> parser, where `this.model.sheets` is undefined for a worksheet, indicating the workbook model is not
> reliably available before worksheets are parsed when there are many sheets. The originally
> attached repro file is a broken HTML download rather than a zip, so the exact reported symptom
> ("drops worksheets") could not be reproduced from it, and this is captured here as desired behavior
> rather than a corpus case with a fragile assertion.

## Desired behavior

- Streaming-reading a workbook with any number of worksheets emits a worksheet for **every** sheet
  part the workbook declares, so the count surfaced equals the count of declared sheets.
- Worksheets are emitted in workbook order, each discoverable by its declared name.
- Enumeration completes, and the stream never stalls, regardless of sheet count.
- The workbook-level model (sheet list, relationships, shared strings) needed to resolve a
  worksheet must be available _before_ that worksheet is emitted. Parsing order must guarantee it,
  so a worksheet handler never dereferences an unpopulated model.
- **Foreign-generated workbooks keep their sheet names.** A file emitted by a non-Excel producer,
  openpyxl in the report, was streamed with its worksheet _names_ lost: the sheets
  emitted, but each `name` came back empty or wrong because the streaming reader bound names by an
  Excel-specific assumption the foreign `workbook.xml` did not satisfy. Names must be resolved from
  the workbook part's declared sheet list, via the relationship graph, so a streamed worksheet
  carries the same name the buffered reader would surface. This is the streaming analogue of
  `worksheet-enumeration-tolerates-foreign-generators`.
- **Stream end is deterministic and fires only when the input is truly exhausted.** The row and
  worksheet iteration signalled completion early under a race: the underlying stream reported "ended"
  while more data was still in flight, so the last rows, or a whole trailing worksheet, were silently
  dropped, non-deterministically. The reader must not treat a transient drain or pause as
  end-of-input; a worksheet is complete only when its part is fully consumed, and the workbook stream
  is complete only when every declared part has been. Because the failure is a timing-dependent race,
  its regression guard belongs in a stress or soak harness, with repeated runs under artificial chunk
  fragmentation, not a single deterministic corpus fixture.

## Open questions

- Ordering guarantee: does the streaming reader parse `workbook.xml` and `_rels` fully before
  emitting any worksheet, or lazily? The crash above suggests the model must be resolved first.
- Memory: many sheets must not force the whole workbook into memory, only the workbook-level
  index, which is small, plus the current worksheet's bounded window. This ties into
  `bounded-memory-large-workbook-read`.
- Error semantics if a declared sheet part is missing or unreadable: skip-with-signal, or throw?
- A regression check for this belongs in a harness that can synthesize an N-sheet workbook and
  assert `emitted === declared`, since a durable fixture would be large.

Related: `bounded-memory-large-workbook-read`, `streaming-read-resolves-shared-strings`,
`streaming-delete-worksheet-before-write`, `worksheet-enumeration-tolerates-foreign-generators`,
`streaming-reader-releases-temp-resources-on-early-abort`.
