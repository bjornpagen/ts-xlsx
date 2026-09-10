# Sheet-protection password hashing and cross-application compatibility

Cluster: security

## Scenario

A user protects a worksheet with a password. The library writes a `sheetProtection` element with
a modern hash: `algorithmName="SHA-512"`, a `hashValue`, a `saltValue`, and a `spinCount`. This
is valid and honored by current Excel. But some older spreadsheet applications, and some
consumers, only understand the legacy 16-bit `password="XXXX"` hash attribute and silently ignore
the modern algorithm form, so on those consumers the sheet appears unprotected. The user expects
the protection to be enforced everywhere they open the file.

## The design question

This is not a case of wrong output, since the SHA-512 `sheetProtection` the library emits is
spec-valid and enforced by modern Excel. The question is a **compatibility policy**: whether, and
how, to also satisfy older consumers that only read the legacy hash.

## Desired behavior (to decide)

- Protecting a worksheet with a non-empty password writes a `sheetProtection` that records enough
  to enforce it. Decide whether that means:
  1. **Modern only** (current): `algorithmName`, `hashValue`, `saltValue` and `spinCount`. Correct
     and strong for modern Excel; ignored by legacy-only consumers.
  2. **Legacy only**: the 16-bit `password` attribute. Broad compatibility, cryptographically
     weak, deprecated.
  3. **Both**: emit the modern hash and the legacy attribute so every consumer enforces
     _something_. Maximizes compatibility at the cost of shipping the weak hash too.
- Each protection permission passed to `protect()` (objects, scenarios, `selectLockedCells`,
  `formatCells`, `insertRows`, …) is serialized as its corresponding `sheetProtection` flag.
- Protecting with **no** password still writes the permission flags but fabricates **no** hash.
- The protection state round-trips: reading the produced file reports the sheet as protected with
  the same permission flags.

## Open questions

- Which policy (modern, legacy, or both) is the fork's default, and is it configurable per call?
- If "both", document explicitly that the legacy attribute is weak and present only for
  compatibility, and do not let its presence imply strong protection.
- Worksheet protection is not encryption, since the sheet data is still readable in the zip, so the
  API docs must not imply confidentiality. This belongs alongside any future workbook-encryption work.

## What a file already carries

A separate matter from what to emit for a password a caller supplies. A file written by pre-2010 Excel,
XlsxWriter, openpyxl or LibreOffice often protects a sheet with the legacy hash alone. The reader keeps
that hash verbatim as `SheetProtection.legacyPasswordHash` and the writer re-emits it before `sheet="1"`,
so a save never turns a password-guarded sheet into one anyone can unprotect. The hash is held to its
schema type (`ST_UnsignedShortHex`, four hexadecimal digits) on both sides: a malformed one is dropped on
read and refused on write. It is never derived from a password, since `protect` still writes only the
agile credential, so the policy question above is untouched. An element without `sheet="1"` records an
unprotected sheet, because `sheet` defaults to false.

## Prior art

The modern `sheetProtection` hash (SHA-512 plus salt plus spin count) is what current Excel writes and
reads; the legacy 16-bit hash predates it and is what older applications key on. The gap is purely
which forms are emitted.
