# ADR 0008: Upgrade to TypeScript 6; hold at 6 (not 7) until the printer API ports

**Status:** Superseded in part (2026-08-05) by [ADR 0028](./0028-typescript-7-adoption.md) · Accepted (2026-07-19) · Phase 4 · advances the toolchain from [ADR 0029](./0029-toolchain-standup.md), constrained by the docs generator of [ADR 0006](./0006-docs-from-types.md)

> The "hold at 6" half of this decision is spent: the tree is on `typescript@^7.0.2`. The
> blocker diagnosed below was real and was removed exactly as the _Negative / deferred_ note
> predicted, by rendering signatures from source-text spans instead of `transform`+`printNode`.
> The `@types/node` reasoning still stands unchanged. Kept for the diagnosis, which is the part
> worth reading.

## Context

The tree was on `typescript@^5.9.3` and `@types/node@^22`. Two newer TypeScript
lines are now published: **6.0.x**, the transitional release and the last of the
classic C#/JS-hosted compiler, which turns on deprecation warnings for what the
native compiler removes; and **7.0.x**, npm's new `latest`, which is the native
Go port ("tsgo"). CLAUDE.md §3 authorizes tooling modernization with bias to
action, so the question was simply which line the toolchain can actually move to
while every gate stays green.

## Decision

Move to **`typescript@^6.0.3`** and **`@types/node@^24.13.3`**. Do **not** move to
TS 7 yet.

### TypeScript 6.0.3

All seven gates pass unchanged: `typecheck` (strict config), `build` (tsc ESM plus
`.d.ts` with `rewriteRelativeImportExtensions`), `smoke:dist`, `lint` (Biome, which
is TS-version-independent), `test:src` (647 unit tests), `corpus` (671 behaviors),
and `docs:check`. No source or config change was required. 6.0 is a clean lift,
and being the deprecation-surfacing bridge release it is the right place to sit
before any eventual 7 move.

### `@types/node` matched to the runtime, not to `latest`

`@types/node` `latest` is `26.x`, but the rewrite runs and tests on Node 24 (the
no-build `.ts` type-stripping workflow of [ADR 0001](./0001-rewrite-runtime-and-toolchain.md)).
We pin the types to the **Node 24 line** (`^24.13.3`) so the ambient types describe
exactly the APIs present at runtime. Taking the v26 types would let source
reference Node-26-only APIs that typecheck but throw on the Node 24 we actually
run, which is a silent gap the types exist to prevent.

### Why not TypeScript 7 (yet)

Under TS 7, `typecheck`, `build`, and `smoke:dist` all pass, but **`docs:check`
fails**:

```
TypeError: ts.createPrinter is not a function
    at scripts/gen-docs.mjs:112
```

TS 7's `typescript` package ships the native Go compiler plus a JavaScript API
shim that currently exposes only a _subset_ of the classic compiler API. It splits
in two:

- **Analysis calls, present in 7.** `createProgram`, the type checker,
  diagnostics, `getDocumentationComment`, `displayPartsToString`, `SyntaxKind`,
  `ScriptTarget`. This is why typecheck and build work, since they emit through the
  compiler's own internal Go pipeline via the CLI.
- **Printer and transform calls, absent in 7.** The in-process AST-to-source
  machinery: `ts.createPrinter()` and `ts.transform()`.

Our docs generator ([ADR 0006](./0006-docs-from-types.md)) lives entirely in that
second half. Its "body-stripped signature" feature runs a `transform` to strip each
function body, then a `printer.printNode(...)` to render the trimmed AST node back
into a signature string. TS 7 hands us the AST and the checker but not the printer
to turn a node back into text, so the generator dies. The blocker is our own
programmatic use of a narrow corner of the compiler API, not the language, and not
the codebase, both of which TS 7 accepts.

## Consequences

- **Positive:** the toolchain is on the current classic line with zero source
  changes and every gate green; deprecation warnings for the 7 transition now
  surface early; the ambient Node types match the runtime we test on.
- **Negative / deferred:** TS 7 is blocked on the docs step alone. When we want it,
  the fix is to rework `scripts/gen-docs.mjs` to reconstruct signatures from source
  text spans (`node.getText(sourceFile)` plus a manual body-strip via the body
  node's range) instead of `transform`+`printNode`. That needs no printer and works
  on both 6 and 7. Splitting the toolchain to keep a TS-6 binary only for docs was
  considered and rejected, because it fractures the one-toolchain principle of ADR 0029.
- **Revisit when:** tsgo ports the printer/transform API (then re-test 7 directly),
  or we rewrite the docs generator off that API for another reason, whichever comes
  first.
