// The codec's back channel into the model.
//
// A reader has to push state into a `Workbook` that no authoring path can produce: a `<dxfs>` table
// whose indices existing rules already point at, a theme part carried as opaque bytes, a protection
// credential in finished agile form with no recoverable password. The streaming writer likewise has
// to drop a row's cells the moment they are serialised. All of that used to be public methods on the
// model classes, so they shipped in the `.d.ts`, appeared in the generated API reference, and a
// caller who tried one put the workbook in a state nothing else could reach or repair. There was no
// authoring/codec boundary at all: the model class *was* the codec's mutation interface.
//
// These symbols are that boundary. The operations still live on the classes and still mutate exactly
// the state they always did (this moves no data) but reaching them requires a symbol that is not on
// the public barrel and, because `package.json` exposes only the root entry, is not importable from
// outside the package at all. Inside the tree, the import is the audit trail: `grep` for this module
// and you have the complete list of code that can restore preserved state.
//
// Prefer an ordinary public method whenever a caller could reasonably want the operation. This is for
// operations that are *only* meaningful to the library's own machinery: mid-deserialisation, or
// mid-serialisation, which is how the streaming writer came to use the same key for its own plumbing.

/**
 * Keys the operations the library's own machinery may perform on a published class and a caller may
 * not (see `WorkbookInternals` / `WorksheetInternals`, declared beside their classes).
 *
 * Not only the model: `WorksheetStreamWriter` hangs its construction and its row-flush plumbing off
 * the same key, on the static side and the instance side respectively, because that class has the
 * identical problem one layer up. Its constructor took the writer's style registry and its
 * `flushedSheet()` returned the writer's flushed-row record, so seven internal types were named by a
 * published signature and none of them was a type a consumer could write down. `StreamedRow` is
 * constructed through the same static key, because a row built outside its writer could commit, and
 * so evict, a row number that writer never handed out.
 */
export const INTERNAL: unique symbol = Symbol('ts-xlsx codec channel');

/**
 * Keys a `Cell`'s link to a named cell style, its OOXML `xfId`. A single hidden property rather than
 * a channel object, so it stays a prototype accessor: cells are the one model class allocated in the
 * millions, and a per-instance internals object would be a real cost for state most cells never carry.
 */
export const NAMED_STYLE_ID: unique symbol = Symbol('ts-xlsx named style link');

/**
 * The exhaustiveness proof the mirror types are built on: instantiate it with a `keyof` difference
 * that should be empty. An unmirrored field leaves that difference non-`never`, which does not
 * satisfy the constraint, so the build fails naming the field that was missed: a type error at the
 * declaration rather than a silent hole discovered by a round trip. `Row`, `Column` and
 * `WORKSHEET_MODEL_FACETS` each carry one of these proofs; this is the mechanism all three use.
 */
export type AssertNever<T extends never> = T;
