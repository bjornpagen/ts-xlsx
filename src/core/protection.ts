// Worksheet-level protection: the `<sheetProtection>` state that makes the per-cell
// `locked`/`hidden` flags enforceable. A cell's protection does nothing until the sheet
// itself is protected; this module models that sheet-level switch and the optional
// password credential that guards lifting it.
//
// The option surface is stated in the AUTHOR's terms: each flag answers "may a user do
// this while the sheet is protected?" (`sort: true` = sorting stays available). OOXML
// encodes the inverse ("1" LOCKS an operation, "0"/omission PERMITS it) and its per-
// attribute defaults differ; that encoding table is {@link SHEET_PROTECTION_FLAGS} below,
// shared by the writer and reader, while the translation that consumes it lives in the io layer.

import {concat, toBase64, utf16leBytes} from '../bytes.ts';
import {sha512} from '../sha512.ts';
import type {AssertNever} from './internal.ts';

/**
 * Whether each protected-sheet operation stays available to a user. Every flag is an
 * *allow* flag: `true` keeps the operation permitted, `false` forbids it, and an absent
 * flag falls to Excel's default for that operation (most editing operations default to
 * forbidden once a sheet is protected; selecting cells defaults to permitted).
 */
export interface SheetProtectionFlags {
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

/** {@link SheetProtectionFlags} plus the password-hardening knob accepted by `protect`. */
export interface SheetProtectionOptions extends SheetProtectionFlags {
  /**
   * Iteration count for the password hash. Higher is slower to brute-force; Excel writes
   * 100000 by default. Ignored when no password is given.
   */
  readonly spinCount?: number;
}

/**
 * A password-derived credential, in OOXML's agile form: the hash algorithm, the salted
 * iterated hash of the password, the salt, and the iteration count: everything a consumer
 * needs to verify a supplied password without the password ever being stored.
 */
export interface SheetProtectionCredential {
  readonly algorithmName: string;
  readonly hashValue: string;
  readonly saltValue: string;
  readonly spinCount: number;
}

/** A sheet's protection: which operations stay allowed, and the optional password guard. */
export interface SheetProtection {
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

const LEGACY_PASSWORD_HASH = /^[0-9A-Fa-f]{4}$/;

/**
 * Whether a string is a legacy sheet password hash as its schema type (`ST_UnsignedShortHex`) spells
 * one: exactly four hexadecimal digits. The reader keeps only such a value and the writer refuses
 * anything else, since the hash is interpolated into an attribute.
 */
export function isLegacyPasswordHash(value: string): boolean {
  return LEGACY_PASSWORD_HASH.test(value);
}

/**
 * The OOXML encoding table for the protection flags: each `<sheetProtection>` attribute paired
 * with whether that operation is *forbidden by default* once a sheet is protected. Both directions
 * key off this one list. The writer turns an author allow-flag into an attribute (omitting values
 * equal to the default) and the reader turns an attribute back into an allow-flag, so serialization
 * and deserialization can never fall out of step. Most editing operations default to forbidden
 * under protection; selecting cells and the object/scenario operations default to permitted.
 */
export const SHEET_PROTECTION_FLAGS = [
  {key: 'formatCells', defaultForbidden: true},
  {key: 'formatColumns', defaultForbidden: true},
  {key: 'formatRows', defaultForbidden: true},
  {key: 'insertColumns', defaultForbidden: true},
  {key: 'insertRows', defaultForbidden: true},
  {key: 'insertHyperlinks', defaultForbidden: true},
  {key: 'deleteColumns', defaultForbidden: true},
  {key: 'deleteRows', defaultForbidden: true},
  {key: 'sort', defaultForbidden: true},
  {key: 'autoFilter', defaultForbidden: true},
  {key: 'pivotTables', defaultForbidden: true},
  {key: 'objects', defaultForbidden: false},
  {key: 'scenarios', defaultForbidden: false},
  {key: 'selectLockedCells', defaultForbidden: false},
  {key: 'selectUnlockedCells', defaultForbidden: false},
] as const satisfies readonly {
  readonly key: keyof SheetProtectionFlags;
  readonly defaultForbidden: boolean;
}[];

/**
 * Compile-time proof that {@link SHEET_PROTECTION_FLAGS} covers every {@link SheetProtectionFlags}
 * operation. The table's doc says the writer and reader can never fall out of step; that holds only
 * while every flag is in it, and a flag added to the interface alone would be neither written nor read.
 */
export type EverySheetProtectionFlagIsDeclared = AssertNever<
  Exclude<keyof SheetProtectionFlags, (typeof SHEET_PROTECTION_FLAGS)[number]['key']>
>;

// OOXML's agile hashing (ECMA-376 / MS-OFFCRYPTO): the password is UTF-16LE, prefixed with
// the salt for the first hash, then re-hashed `spinCount` times with a little-endian uint32
// iteration counter mixed in. SHA-512 is the modern choice Excel writes.
//
// Both primitives are deliberately platform-neutral: `src/sha512.ts` rather than `node:crypto`,
// and the Web Crypto `getRandomValues` that browsers and Node both carry as a global. Protecting
// a sheet is therefore something a browser can do, and nothing on the package's entry graph
// imports a Node built-in to make it possible.
const ALGORITHM_NAME = 'SHA-512';
const DEFAULT_SPIN_COUNT = 100000;
const SALT_BYTES = 16;

/**
 * Derive a fresh {@link SheetProtectionCredential} for a password. Each call generates a new
 * random salt, so protecting two sheets with the same password yields different credentials:
 * the salt is real randomness, not a stub.
 *
 * The spin loop is the cost of the scheme rather than of this implementation: Excel's default
 * 100000 iterations is 100000 chained SHA-512 digests, and no shortcut through them exists.
 */
export function deriveCredential(
  password: string,
  spinCount: number = DEFAULT_SPIN_COUNT,
): SheetProtectionCredential {
  const salt = crypto.getRandomValues(new Uint8Array(SALT_BYTES));
  let hash = sha512(concat([salt, utf16leBytes(password)]));
  // One buffer for the whole loop, holding the previous digest followed by the counter: at the
  // default spin count the alternative is a hundred thousand throwaway 68-byte allocations.
  const spun = new Uint8Array(hash.length + 4);
  const counter = new DataView(spun.buffer, hash.length, 4);
  for (let i = 0; i < spinCount; i++) {
    spun.set(hash, 0);
    counter.setUint32(0, i, true);
    hash = sha512(spun);
  }
  return {
    algorithmName: ALGORITHM_NAME,
    hashValue: toBase64(hash),
    saltValue: toBase64(salt),
    spinCount,
  };
}
