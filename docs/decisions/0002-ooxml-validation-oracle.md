# ADR 0002: Microsoft OpenXmlValidator as an external conformance oracle

**Status:** Accepted (2026-07-11) · Phase 3 · **mechanism superseded** (2026-08-15) by
[ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md): the oracle is no longer a
repository-owned .NET tool but the shared `ooxml-validate` package, which carries the pin
and the binary. Everything below about _why_ this oracle, at `Microsoft365`, as
development-only tooling, still holds. Read 0033 for what runs it.

## Context

The regression corpus proves observable spreadsheet behavior, but many checks reload a workbook
with the same implementation that wrote it. A writer and reader can therefore agree on malformed
OOXML. Hand-written ZIP, XML, and element-order assertions cover known failures but are not a
complete independent conformance check.

The legacy writer also currently emits its default font children in an order that Microsoft's
validator rejects. Phase 2 deliberately froze that tree instead of fixing code that the TypeScript
rewrite will replace, so a new validation gate must distinguish that one known error from new ones.

## Decision

- Use a repository-owned .NET console tool around Microsoft's `OpenXmlValidator`, pinned to
  `DocumentFormat.OpenXml` 3.5.1 with a committed NuGet lockfile.
- Target `FileFormatVersions.Microsoft365`. The validator is development and CI tooling only; the
  published JavaScript package gains no .NET dependency.
- Validate representative buffered and streaming output in a separate required Linux workflow.
- Baseline the legacy default-font error by exact ID, type, part, and XPath. Additional errors and a
  stale baseline both fail the check.
- Keep intentionally malformed and tolerant-reader corpus fixtures out of the blanket validation
  set. Add generated cases deliberately as writer capabilities move into `src/`.

## Consequences

The project gains an independent OOXML schema and semantic oracle with structured diagnostics and
reproducible dependencies. Contributors running this check need the .NET 10 SDK. Passing it still
does not prove that Excel will render or calculate a workbook correctly, so corpus behavior and
interoperability testing remain separate gates.
