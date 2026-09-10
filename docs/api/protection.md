# Protection

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `SheetProtection`

<sub>interface</sub>

A sheet's protection: which operations stay allowed, and the optional password guard.

```ts
interface SheetProtection {
  readonly flags: SheetProtectionFlags;
  readonly credential?: SheetProtectionCredential;
  /**
   * The legacy 16-bit password hash (`password="CC3D"`) a file protected the sheet with, kept verbatim
   * so the sheet stays guarded after a save. Pre-2010 Excel, XlsxWriter, openpyxl and LibreOffice write
   * it instead of the agile credential. It is never derived from a password: `protect` writes only the
   * agile form, and this hash is weak enough that nobody should want it authored.
   */
  readonly legacyPasswordHash?: string;
}
```

---

### `SheetProtectionCredential`

<sub>interface</sub>

A password-derived credential, in OOXML's agile form: the hash algorithm, the salted
iterated hash of the password, the salt, and the iteration count: everything a consumer
needs to verify a supplied password without the password ever being stored.

```ts
interface SheetProtectionCredential {
  readonly algorithmName: string;
  readonly hashValue: string;
  readonly saltValue: string;
  readonly spinCount: number;
}
```

---

### `SheetProtectionFlags`

<sub>interface</sub>

Whether each protected-sheet operation stays available to a user. Every flag is an
*allow* flag: `true` keeps the operation permitted, `false` forbids it, and an absent
flag falls to Excel's default for that operation (most editing operations default to
forbidden once a sheet is protected; selecting cells defaults to permitted).

```ts
interface SheetProtectionFlags {
  /** Select locked cells (Excel permits this by default). */
  readonly selectLockedCells?: boolean;
  /** Select unlocked cells (permitted by default). */
  readonly selectUnlockedCells?: boolean;
  readonly formatCells?: boolean;
  readonly formatColumns?: boolean;
  readonly formatRows?: boolean;
  readonly insertColumns?: boolean;
  readonly insertRows?: boolean;
  readonly insertHyperlinks?: boolean;
  readonly deleteColumns?: boolean;
  readonly deleteRows?: boolean;
  readonly sort?: boolean;
  readonly autoFilter?: boolean;
  readonly pivotTables?: boolean;
  readonly objects?: boolean;
  readonly scenarios?: boolean;
}
```

---

### `SheetProtectionOptions`

<sub>interface</sub>

[`SheetProtectionFlags`](./protection.md#sheetprotectionflags) plus the password-hardening knob accepted by `protect`.

```ts
interface SheetProtectionOptions extends SheetProtectionFlags {
  /**
   * Iteration count for the password hash. Higher is slower to brute-force; Excel writes
   * 100000 by default. Ignored when no password is given.
   */
  readonly spinCount?: number;
}
```

---

### `WorkbookProtection`

<sub>interface</sub>

A workbook's structure/window protection. The three lock flags each default to `false` (absent),
matching OOXML: an omitted attribute leaves that aspect unlocked. The optional `credentials`
bag carries the opaque password attributes verbatim: the library never verifies a password, it
only refuses to lose one.

```ts
interface WorkbookProtection {
  /** Lock the workbook structure: no adding, deleting, reordering, or unhiding sheets. */
  readonly lockStructure?: boolean;
  /** Lock the workbook window geometry. */
  readonly lockWindows?: boolean;
  /** Lock the revision-tracking state. */
  readonly lockRevision?: boolean;
  /** Preserved password/agile-hash attributes, keyed by their OOXML attribute name. */
  readonly credentials?: Readonly<Partial<Record<WorkbookProtectionCredentialAttr, string>>>;
}
```

---

### `WorkbookProtectionCredentialAttr`

<sub>type</sub>

One of the attribute names `WORKBOOK_PROTECTION_CREDENTIAL_ATTRS` enumerates.

```ts
type WorkbookProtectionCredentialAttr =
  (typeof WORKBOOK_PROTECTION_CREDENTIAL_ATTRS)[number];
```
