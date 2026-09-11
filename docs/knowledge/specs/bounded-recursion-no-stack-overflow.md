# Serialization must be iterative, not deeply recursive, so large data cannot overflow the stack

Cluster: security/deps

## Scenario

A user builds a worksheet from a large JSON dataset and calls `writeBuffer()`. On a development
machine the export succeeds; on another environment (a production bundle, a smaller default stack, a
different engine) the same export dies with `RangeError: Maximum call stack size exceeded`. The
failure is data-size-dependent and environment-dependent: the write walks the model (or an XML/render
tree) with a recursion depth that scales with the number of rows, cells or nodes, so past some
threshold the call stack is exhausted. The document is not malformed and the data is not adversarial.
It is simply big enough, on a stack small enough, to overflow.

> Spec note, not a corpus case: reproducing this reliably requires driving the model to whatever depth
> overflows the _current_ stack, which is engine- and build-dependent and would either not trigger in
> CI or risk crashing the runner. The durable requirement is the invariant, verified by design review
> and a bounded large-input smoke test, not by a data file that must overflow to assert.

## Desired behavior

- **Serialization depth is bounded by document structure, not by data volume.** Emitting N rows or N
  cells must not consume O(N) stack frames. Row, cell and token emission is iterative (a loop or an
  explicit work-stack), so a million-row sheet writes at constant stack depth. Recursion, where used,
  is bounded by the genuine nesting depth of the format, a handful of levels, never by row count.
- **The same guarantee holds on the read path.** Parsing a large or deeply-repetitive part must not
  recurse per element, and a hostile file cannot force unbounded recursion. This is the
  stack-exhaustion sibling of the memory-bound and zip-bomb guards: a crafted file must fail cleanly,
  not blow the stack.
- **Large exports are streaming-friendly.** The streaming writer already emits row-by-row; the
  buffered writer must share the same non-recursive emission core so both paths inherit the bound.
- **Failure, if a real limit is hit, is a typed, catchable error** with context (which part, how far),
  never an opaque `RangeError` from the engine's call stack.

## Open questions

- Where legacy actually recurses per-element (the XML render/stringify tree is the prime suspect) and
  whether the modern rewrite's serializer is iterative by construction or needs an explicit guard.
- What a meaningful large-input smoke test looks like, with a row count high enough to have overflowed
  legacy on a typical stack, without making CI slow or flaky.
- Whether to expose a configurable soft cap on document size that fails fast with guidance ("use the
  streaming writer") before the engine's hard limit is reached.

Related: `bounded-memory-on-load`, `zip-bomb-resistant-unzip`, `streaming-write-memory-and-shared-strings-tradeoff`.
