# Authoring workbook-level structure protection

Cluster: security

## Scenario

A user wants to lock a workbook's _structure_ so that recipients cannot add, delete, rename, move,
copy, hide, or unhide worksheets, the protection Excel exposes as "Protect Workbook → Structure."
The library today has no authoring surface for this: a caller can protect individual worksheets
(locked cells, `sheetProtection`) but cannot set the workbook-level `<workbookProtection
lockStructure="1">` element that guards the sheet collection itself. The only workbook-wide integrity
signal available is per-sheet, which does not stop a recipient from deleting or reordering sheets.

A separate, already-locked defect is that structure protection present in a _loaded_ file is dropped
on a passthrough save (corpus case `workbook-structure-protection-survives-roundtrip`). This note is
about the complementary gap: there is no way to _set_ it in the first place.

> Spec note, not a corpus case: round-trip preservation of an existing lock is already a corpus case,
> and what is missing is an authoring API and its shape, a design decision rather than a
> malformed-output bug. It becomes a corpus case once the setter exists and its emitted
> `workbookProtection` is asserted.

## Desired behavior

- **A workbook exposes a structure-protection setting** that, when enabled, emits
  `<workbookProtection lockStructure="1">`, and, if a windows lock is also chosen,
  `lockWindows="1"`, so a produced file opens with its structure locked.
- **Optional password hardening**: the API accepts a password and emits the hashed
  `workbookPassword`, `algorithmName`, `hashValue`, `saltValue` and `spinCount` attributes using the
  same password-hashing path as sheet protection (`sheet-protection-password-hash-compatibility`),
  rather than a legacy weak hash. Structure protection without a password is also valid, a soft lock.
- **Honest about strength**: workbook structure protection is an integrity _signal_, not encryption.
  It is trivially removable and must not be confused with `workbook-password-encryption`. The API and
  its docs make that distinction explicit so callers do not mistake it for confidentiality.
- **Round-trips**: a workbook the caller protects writes the element, and reloading then re-saving
  preserves it, converging with the existing round-trip lock.

## Open questions

- Surface shape: a `workbook.protect(password?, {lockStructure, lockWindows})` method mirroring the
  worksheet API, or a declarative `workbook.protection = {...}` property? The method form matches
  the sheet-protection precedent.
- Default `spinCount` and whether it is bounded on read to avoid a hostile file forcing an extreme
  iteration count, which ties to `public-type-surface-matches-runtime`'s `spinCount` note and the
  encryption bounding.
- Whether enabling structure protection without a password should warn that it is trivially bypassed.

Related: `sheet-protection-password-hash-compatibility`, `workbook-password-encryption`,
`sheet-protection-permits-requested-operations`.
