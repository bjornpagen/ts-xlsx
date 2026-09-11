---
name: validate-ooxml
description: Validate generated .xlsx output against Microsoft's OOXML schema/semantic oracle (OpenXmlValidator). Use after changing any writer path, when a workbook "opens broken in Excel", when checking that emitted XML conforms to ECMA-376, or when you need an authoritative pass/fail on a generated package. Pairs with write-corpus-case; see docs/agent-correctness-playbook.md.
metadata:
  # For working *on* ts-xlsx, not *with* it: needs the repo's dev dependencies and its docs.
  # `npx skills add shbernal/ts-xlsx` walks .claude/skills/ too, and this flag is what keeps it
  # out of the menu a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install it here.
  internal: true
---

# Validating OOXML output

The corpus (`pnpm run corpus`) proves well-formedness, package structure, and no
regression — but it is _our_ assertions. The **authoritative** schema + semantic check
is Microsoft's own `OpenXmlValidator`, reached through `ooxml-validate` (ADR-0002) — the
shared oracle this repo and `ts-pptx` both validate against, so the two enforce one rule
set. It is the independent oracle: if it says a package is clean, Excel's own conformance
layer agrees. Reach for it after any change to a writer path.

## Requirement

A dev install (`pnpm install`) and, on the very first call, network access: the package
fetches a prebuilt oracle binary from its GitHub release, verifies its checksum and build
provenance, and caches it under `~/.cache/ooxml-validate`. No .NET, no `dotnet` on PATH.
The published npm package is unaffected either way — this is development-only tooling.

If the binary cannot be fetched, do **not** substitute an XSD/`xmllint` validator
(deliberately not wired — ADR-0002, ADR-0034). Fall back to `pnpm run corpus` locally and
let CI's `ooxml-validation` workflow run the oracle on your PR. See
`docs/agent-correctness-playbook.md`.

The `ooxml-lookup` skill is the other half of this loop, not a replacement for it: it does
not validate anything, but given a diagnostic's `id`, `description` and `xpath` its
`explain` subcommand says what _would_ have been legal at that position.

## Validate a file you already have

```bash
pnpm run validate:ooxml path/to/workbook.xlsx another.xlsx
```

Pass every file you want checked in one call — process startup dominates (~0.4 s, then
~9 ms per additional package). Write the filenames straight after the script name: a
`--` separator reaches the oracle as an argument and it will reject it.

Deterministic JSON to stdout. Exit codes: **0** every input clean · **1** validation or
package-open errors found · **2** the tool could not run (bad args / internal failure).
Each error carries `id`, `type` (`Schema` | `Semantic` | `MarkupCompatibility` |
`Package`), `partUri`, and `xpath` — enough to locate the offending element without
opening the zip. Every input file appears in the report with an explicit `valid` flag;
a file is never absent because it was clean.

## Emit a representative workbook, then validate

Node 24 runs the `.ts` sources directly, so a repro imports straight from `src`. Write
one that exercises the feature you changed, then point the validator at it:

```js
// .tmp/repro.mjs
import {writeFileSync} from 'node:fs';
import {Workbook} from '../src/core/workbook.ts';
import {writeXlsx} from '../src/io/xlsx/write.ts';

const wb = new Workbook();
const ws = wb.addWorksheet('Data');
ws.addRow(['Name', 'Value']);
ws.addRow(['alpha', 42]);
// …exercise the exact path you touched: styles, tables, formulas, images, CF…

writeFileSync('.tmp/repro.xlsx', writeXlsx(wb)); // writeXlsx → Uint8Array; cwd is the repo root
```

```bash
node .tmp/repro.mjs && pnpm run validate:ooxml .tmp/repro.xlsx
```

For the streaming writer, use `WorkbookStreamWriter` (`src/io/xlsx/write-stream.ts`) and
collect its output the same way, then validate — streaming and buffered output must both
be clean.

## The baseline (do not paper over new errors)

`test/ooxml-validation/run.ts` (`pnpm run test:ooxml`) generates buffered + streaming
workbooks and control cases, then compares each package's errors to the frozen set in
`test/ooxml-validation/allowed-errors.json`.

- A **new** diagnostic that isn't baselined → the gate fails. Fix the writer; do not add
  it to the baseline to silence it. The baseline is a record of _known-open_ writer bugs
  we've chosen to track, not a mute button.
- A **stale** baseline (an error you _fixed_) also fails the gate — remove that entry in
  the same change so the file stays honest.
- Match errors by the exact `{id, type, partUri, xpath}` fingerprint.

## After validating

- Clean (exit 0)? Good — but the oracle is not a substitute for the corpus. If the change
  is new behavior, also land an implementation-blind corpus case (`write-corpus-case`
  skill) so the guarantee is permanent and survives without the oracle.
- Errors (exit 1)? The `xpath` + `partUri` point at the element. Fix the serializer, not
  the assertion. Re-run until clean.
