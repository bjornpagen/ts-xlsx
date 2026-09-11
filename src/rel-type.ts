// What a relationship Type says, asked once.
//
// A root leaf importing nothing, because everything asks it. The reader resolves a part through its
// `.rels`; the writer decides which of a preserved sheet's relationships it is re-emitting; the model
// finds the macro project among a workbook's preserved references; the ribbon reader recognises a
// `customUI` part. Put beside the reader, this one-line test dragged the whole read path -- the
// bounded inflater included -- into the writer's module closure, which the per-entry size budget is
// there to notice.
//
// It lived in `io/opc/` for exactly as long as only the codecs asked it. `core/` and `customui/` may
// not import `src/io`, so they each spelled `relType.endsWith('/vbaProject')` and
// `type.endsWith('/ui/extensibility')` by hand, three sites carrying the separator gotcha below with
// nowhere to say why it was there. The question that decides where this lives is whether a
// relationship Type is a concept the model may own, and it is: a `Workbook` holds preserved
// references *keyed by* their Type, so the model already carries the string. What it lacked was the
// one place that reads it. Reading a URI's last segment is no more a serialisation concern than
// rendering a number as hex, which is why this sits beside `hex.ts` rather than being duplicated
// across three layers.

/**
 * The last segment of a relationship Type (`vbaProjectSignatureAgile` in
 * `http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile`), or the whole Type
 * when it has no `/`: for a caller that looks the segment up rather than testing for one class.
 *
 * {@link isRelType} is not built on it, because a class name may span segments: the ribbon's is
 * `ui/extensibility`.
 */
export function relTypeSegment(type: string): string {
  return type.slice(type.lastIndexOf('/') + 1);
}

/**
 * Does a relationship Type name this class of part?
 *
 * A Type is a URI whose last segment names the part class, and the namespace in front of it varies by
 * relationship family and by the schema version a producer wrote against. So the match is on that
 * last segment.
 *
 * **The separator is the load-bearing character.** `/slicer` and `slicer` are not the same test:
 * without the slash, `slicer` matches `slicerCache`, `table` matches `pivotTable`, and a lookup
 * quietly gathers a part class nobody asked for. Every caller used to spell the slash itself, ten
 * times over, with nowhere to say why it was there.
 */
export function isRelType(type: string, name: string): boolean {
  return type.endsWith(`/${name}`);
}

/** {@link isRelType} against several part classes at once, for the predicates that partition
 * relationships into "modeled", "preserved verbatim", and "dropped". */
export function isAnyRelType(type: string, ...names: readonly string[]): boolean {
  return names.some((name) => isRelType(type, name));
}
