# ADR 0028: Move to TypeScript 7; the compiler API scripts move to `unstable/*`

**Status:** Accepted (2026-08-05) · supersedes the "hold at 6" half of [ADR 0008](./0008-typescript-6-upgrade.md) · completes the generator rework of [ADR 0006](./0006-docs-from-types.md)

## Context

[ADR 0008](./0008-typescript-6-upgrade.md) moved the toolchain to `typescript@^6.0.3` and
explicitly declined TypeScript 7, the native Go compiler ("tsgo"), for exactly one reason:
`scripts/gen-docs.ts` rendered signatures with `ts.createPrinter()` over a `ts.transform()`
body-stripping pass, and 7 ships neither. Every other gate already passed under 7.

That blocker no longer exists. The generator now slices a declaration's own source text from
`getStart` to the start of its body, which needs no emit machinery at all, and it landed on 6
ahead of and independent of this move. What remained was mechanical: the two scripts that
consume the compiler programmatically had to move to the API 7 actually publishes.

## Decision

Move to **`typescript@^7.0.2`**, and port `scripts/gen-docs.ts` and `scripts/check-entries.ts`
from the `typescript` default export to `typescript/unstable/ast` and `typescript/unstable/sync`.

### The compiler is a separate process, and a project is the only way in

TypeScript 7's main export is only a version string; the JavaScript API lives under
`typescript/unstable/*` and is a _client_: an `API` that spawns a tsgo server, an
`updateSnapshot()` that returns a disposable `Snapshot`, and a `Project` carrying the `program`
and `checker`. Both scripts therefore acquire and release two resources under `finally`, or a
throw would strand a tsgo process and leave the script hanging on its pipe. `check-entries.ts`
reports failure through `process.exitCode` rather than `process.exit` for the same reason: the
close has to run before the exit is taken.

There is no inline-`compilerOptions` door in this API, since a project comes from a tsconfig. That
removed a real hazard rather than imposing one: `gen-docs.ts` had been passing a hand-written
option set of its own that no longer matched `tsconfig.json`, so the reference was being
rendered under a checker configured differently from the gate. It now reads the gate's own
config.

### A summary is read from the JSDoc block, not from the checker

`Symbol.getDocumentationComment()` under 7 returns a string in which `{@link Target}` has
already been flattened to a bare `Target`. Under 6 the markup survived into the string, and the
generator turned it into a code span. Taking 7's rendering as-is silently stripped the
monospace off every cross-reference in the reference, 155 lines across 33 pages.

The summary now comes from the declaration's JSDoc block via `getTextOfJSDocComment`, which
still carries the tag intact. This is the same source-of-truth shift the signature rework made:
the reference shows what the author wrote. One attachment quirk is handled explicitly, since a
`const` carries its block on the enclosing statement rather than on the declarator.

### `check-entries.ts` gives up its standalone parse

That script deliberately used `ts.createSourceFile` because the entry barrels are pure
re-export lists and parsing answered every question in milliseconds. TypeScript 7 publishes no
standalone parser (`factory.createSourceFile` builds a synthetic node from statements, not a
tree from text) so a `SourceFile` is reachable only through a project. It now builds one over
`tsconfig.json`, which already includes `src/**/*.ts`.

Measured, this costs nothing: 445 ms before, 600 ms after, because the old figure was dominated
by loading the TypeScript 6 module in the first place. No diagnostics are requested, so a tree
that does not typecheck still gets a verdict, which is the property that made the parse-only
choice worth stating.

### Two rendering divergences are accepted, not worked around

Regenerating the reference under 7 leaves it byte-identical across all 205 symbols except three
lines, both differences being the compiler's own rendering rather than anything the generator
controls:

- **Quote style.** `typeToString` preserves the source's single quotes where 6 normalised to
  double (`ValueType`). 7 is the more faithful of the two.
- **Union ordering.** 7 sorts union constituents alphabetically where 6 kept declaration order
  (`STRIPE_ELEMENT_TYPES`, `DEFAULT_THEME_COLOR_SCHEME`). This loses a little incidental
  meaning, since the declared order tracked OOXML's, but it is deterministic, so the
  `docs:check` drift gate stays exact.

The same quote-style change appears in emitted JS, which made the comment at
`scripts/size-budget.ts` false; its matcher already accepted both forms, so only the comment
needed correcting.

### What it bought

`typecheck`, the gate that runs most often, goes from 4.1 s to 0.9 s, and a full `verify --full`
from 31.4 s to 25.5 s. `gen-docs` alone runs in ~590 ms. Against that, `invariants` grows from
0.8 s to 1.1 s for the project `check-entries` now builds.

## Consequences

- **Positive:** one toolchain, on npm's `latest` line, with every gate green; the inner-loop
  type gate is ~4.5× faster; the docs generator now runs under the same config the gate
  enforces; the reference gained back nothing and lost nothing but three lines of compiler
  formatting.
- **Negative / deferred:** the API is named `unstable` and is documented as such. It can break
  in a minor release, and two scripts now depend on it. The exposure is bounded and gated:
  `typecheck` catches a shape change, and `docs:check` catches a silent behavioural one by
  diffing generated output byte for byte. Renovate ([ADR 0027](./0027-dependencies-are-updated-by-a-bot-and-ci-is-the-reviewer.md))
  raises the bump and CI is the reviewer, which is the intended shape for exactly this risk.
- **Negative / deferred:** there is still no printer, no `transform`, and no `EmitHint`. Nothing
  in the tree needs them, and the signature renderer is deliberately built so nothing will.
- **Revisit when:** the `unstable/*` API stabilises under a supported name (drop the caveat
  above), or a gate needs a compiler capability the API does not expose.
