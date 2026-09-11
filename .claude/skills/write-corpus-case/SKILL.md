---
name: write-corpus-case
description: Author a regression corpus case under test/corpus/cases. Use when distilling a bug/reproduction into a corpus case, adding a permanent regression lock, writing a *.case.ts file, or extending the adapter's capability contract.
metadata:
  # For working *on* ts-xlsx, not *with* it: it writes into this repo's test tree. `npx skills add
  # shbernal/ts-xlsx` walks .claude/skills/ too, and this flag is what keeps it out of the menu a
  # consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install it here.
  internal: true
---

# Writing a corpus case

The corpus is **the product's spine** (`docs/architecture.md`). A case encodes "correct
behavior" as assertions written against the _behavior_ rather than the code — which is
how the corpus outlived the rewrite it was built to survive. Full reference:
`test/corpus/README.md`.

## The three rules that make a case durable

1. **Reach the library only through a capability.** A case NEVER imports `src/`. It calls
   the adapter's vocabulary (`api.decodeAddress(...)` etc.); the adapter is the only file
   allowed to know how the library is shaped. That decoupling is why the corpus survived a
   full rewrite. It does _not_ mean the capability list is hidden — `CorpusApi` is typed,
   so a name or argument you get wrong is a compile error.
2. **Describe the real-world scenario, not the source thread.** The durable knowledge
   is "real `.xlsx` files declare whole-column defined names like `$A:$A`", not "issue
   140". Write the scenario into `description` and each behavior `name`. Do **not**
   put upstream issue/PR numbers in that durable text (they die with the fork). A
   number may sit in the optional `provenance` block as a disposable trace only.
3. **A behavior asserts what must be true, and nothing weaker.** There is no field for
   "what the code does today". If the behavior does not hold, the case reddens the build —
   that is the whole point of putting it here. Never soften an assertion to land it green;
   either the library is wrong and this is the failing reproduction that proves it, or the
   assertion is wrong and does not belong in the corpus.

## Shape

`test/corpus/cases/<descriptive-slug>.case.ts` (a slug like
`whole-column-defined-names` — no number prefix) imports the shared `Case` type and
default-exports an object pinned to it with `satisfies Case`:

```ts
import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'whole-column-defined-names',
  cluster: 'address-decoding', // one of the known clusters
  description:
    'Defined names referencing whole rows/columns must decode without ' +
    'crashing or leaking undefined/NaN into serialized addresses.',
  provenance: {source: 'upstream-issue'}, // OPTIONAL, disposable — never the identity
  behavior: [
    {
      name: 'decodeRange("$1:$1") — a full-row range — resolves its known row bounds',
      expect(api: CorpusApi, assert: Assert) {
        const range = api.decodeRange('$1:$1');
        assert.strictEqual(range.top, 1);
        assert.strictEqual(range.bottom, 1);
      },
    },
  ],
} satisfies Case;
```

- **`CorpusApi`** is the adapter's capability surface, derived from the adapter object so it
  cannot drift from it. Annotate `api` with it; let values you pull _out_ of a capability be
  inferred from its return type rather than annotating them. If you genuinely must opt out
  of typing, import `Untyped` from `../untyped.ts` — it is a named `any` that says so, and
  it is countable. `Assert` is Node's strict `assert`. Both `expect` params must be
  explicitly annotated — `assert`'s assertion signatures need it (TS2775). The harness is
  type-checked (`pnpm run typecheck:test`), so a new case must be green there too, not just
  under `pnpm run corpus`.
- **`behavior[]`** — each entry is one assertion about _observable_ behavior.
  `expect` gets the adapter and Node's strict `assert`; throw to fail, return to pass.
  Keep each behavior single-purpose so a failure names one thing.
- **`cluster`** — group by theme (`address-decoding`, tables, styles, streaming,
  pivot, images, conditional-formatting, dates, formulas, csv, types, security/deps).

## Landing it (run it, do not guess)

1. Write the behavior asserting what _correct_ looks like.
2. Run `pnpm run corpus --case <your-slug>` and read the result.
3. Green → it is a regression lock, and you are done. Red → you have a failing
   reproduction, which is the _good_ outcome for a bug you are about to fix: land the fix in
   the same change (CLAUDE.md §2, "bugs are fixed test-first"). Do not weaken the assertion
   to get green, and do not land a red case on its own.

Also run `pnpm run typecheck:test` — a mistyped capability name or argument is a compile
error there, before the corpus ever runs.

## Fixtures

If the behavior needs a sample spreadsheet, commit it to the durable corpus and
reference it by path:

```
test/corpus/fixtures/<case-slug>/<file>.xlsx
```

Load it inside `expect` via a capability the adapter provides (add one if needed —
see below), never by reaching into an implementation's reader directly.

## Extending the adapter contract

If a case needs a capability the contract doesn't have yet:

1. Add the capability to the module that owns its concern under
   `test/corpus/adapters/ts-xlsx/` (or `adapters/ooxml-facts.ts` for package-level OOXML
   facts). Each returns plain JSON-serializable data — corners, dimensions, cell values —
   never a live model object.
2. Give it a return shape a case can _use_: build a keyed map with `Object.fromEntries` over
   a mapped list rather than filling a pre-declared `Record<string, unknown>`, so the shape
   is inferred and reaches the case. A declared-then-filled record has to name its value
   type before the literal exists, and the only name available is one every case then has to
   defeat.
3. Document it in the contract table in `test/corpus/README.md`.
4. Keep it minimal and behavior-shaped; the vocabulary grows only as cases demand.

## Validate

```
pnpm run corpus
pnpm run typecheck:test
```

Every behavior `✓`, zero failures. There is no healthy non-green state.
