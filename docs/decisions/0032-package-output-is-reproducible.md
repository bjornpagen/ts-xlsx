# ADR 0032: Package output is reproducible, with entry timestamps pinned rather than clocked

**Status:** Accepted (2026-08-11) · answers the question [ADR 0024](./0024-async-is-one-writer-not-a-mirrored-pair.md) deferred when it declined to pin `mtime`

## Context

A zip entry carries a modification time in its local header and again in the central
directory. `fflate` defaults it to `Date.now()`, so until now every writer here stamped
the moment it ran. Writing an unchanged workbook twice produced two packages that differed
in a few bytes per entry and in nothing else.

[ADR 0024](./0024-async-is-one-writer-not-a-mirrored-pair.md) met this while comparing
`writeXlsx` and `writeXlsxAsync`, listed pinning `mtime` under rejected alternatives, and
was right to: the two-writer comparison did not need it, and pinning would have changed
`writeXlsx`'s output for a reason internal to a test. It named the real question and left it
open, _should `.xlsx` output be reproducible at all?_ This record answers it.

What settled it was a consumer. The library's first authoring consumer commits the workbooks
it generates, so a regenerated deliverable arrived as a diff of the whole file with no
content change behind it. Three costs, none visible from inside the library:

- **The diff lies.** A file that changes on every regeneration cannot be reviewed, so it
  stops being read, and a real change to a committed deliverable hides in the noise.
- **Nothing downstream can cache on output bytes.** A build step that skips work when its
  output is unchanged never skips.
- **A byte-comparison gate cannot fail for the right reason.** It cannot separate "the
  writer changed" from "the clock moved", which is the same failure as a gate that is always
  red: it gets ignored, and then it is not a gate.

## Decision

1. **Every zip entry this library writes is stamped with a fixed timestamp**, so package
   bytes are a function of the workbook alone. `src/io/opc/zip-mtime.ts` owns the constant,
   at the container layer, because it is a property of the OPC package rather than of the
   `.xlsx` codec that happens to be the first to fill one.

2. **All four writing paths use it**: `writeXlsx` (`zipSync`), `writeXlsxAsync` (`zip`),
   `WorkbookStreamWriter` (the streamed `Zip`/`ZipDeflate` container), and the package-level
   VBA edits in `edit-vba.ts`, which re-zip after splicing. Three different `fflate` calls
   with three different ways of accepting the stamp, since the streamed container takes it as
   a field on the entry rather than as an option, and any one left out would have been silent.

3. **The value is built from local components, not a UTC instant.** `fflate` encodes the DOS
   date through local-time getters, so a fixed instant would still stamp differently in
   different timezones: reproducible on one machine and not across two, which is the half
   of the property that a CI comparison needs. `new Date(2001, 0, 1, 12).getTime()` reads
   back as 2001-01-01 12:00 everywhere. Midday, so no DST transition can shift the date
   under it; 2001 because zip stores DOS dates, which start at 1980, so the epoch is not
   expressible.

4. **There is no option to restore clock stamps.** A caller who wants wall-clock times can
   re-stamp an archive; a caller who wants reproducibility could not recover it. The
   asymmetry decides the default, and a flag on a property this quiet would mostly serve to
   let it be turned off by accident.

## Consequences

- **Positive:** a committed `.xlsx` changes only when its content changes; `writeXlsx` and
  `writeXlsxAsync` are now byte-identical, which pins them to the same compression settings
  in a way comparing inflated parts never could; and a byte-comparison gate over generated
  workbooks becomes worth having.
- **The stamp is visible to users.** Explorer and Finder show 2001-01-01 as each entry's
  date inside the zip. This is the same trade every reproducible-build toolchain makes, it
  does not touch the file's own filesystem timestamp, and Excel reads none of it.
- **Output bytes changed once, at this release.** Any stored hash of a previously written
  package no longer matches. Nothing in the model, the content, or how the file reads
  changed with it.
- **Reproducibility is now a property, so it needs a test that can fail.** Byte-equality of
  two writes is not enough, because two calls in the same second stamp the same DOS bucket and
  pass regardless. `write-determinism.test.ts` therefore decodes the DOS date out of the central
  directory and asserts the literal 2001-01-01 12:00, spelled out rather than imported from
  the constant so the test can disagree with it.
- **`edit-vba.ts` re-stamps rather than preserves.** `unzipSync` returns bytes and drops the
  original entry times, so a package-level edit has nothing to carry through. The choice
  there was a pinned stamp or the clock, and it is the same pinned stamp.
