// A worksheet's protection in BIFF12: the binary counterpart of `<sheetProtection>`.
//
// Two records carry it, and Excel writes both for a sheet protected with a password. `BrtSheetProtection`
// is the legacy form: a 16-bit password hash, then sixteen `Bool32`s. `BrtSheetProtectionIso` is the
// agile form: a spin count, the same sixteen flags, then the hash, the salt and the algorithm name the
// XML twin spells as `hashValue`, `saltValue` and `algorithmName`. A sheet protected without a password
// carries only the first.
//
// The layout was read off a pair of files Excel saved from one workbook, as `.xlsb` and as `.xlsx`,
// once per protection setting (Excel 16.0 build 20326), rather than taken from memory. Two findings a
// reader has to get right:
//
// - The first flag says the sheet is protected, and the other fifteen say an operation is *allowed*,
//   which is the author's sense the model uses and the inverse of the XML attributes: `formatCells="0"`
//   in the XML twin is a set `fFormatCells` here.
// - Every flag is always present. The XML twin states only the attributes that depart from their schema
//   default, and the model records a flag only when the XML states it, so a flag is recorded here only
//   when it departs from that default too. Otherwise the two readings of one workbook would differ.

import {toBase64} from '../../bytes.ts';
import type {AssertNever} from '../../core/internal.ts';
import {
  SHEET_PROTECTION_FLAGS,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
} from '../../core/protection.ts';
import {hex} from '../../hex.ts';
import type {RecordReader} from './primitives.ts';

// The fifteen operation flags in the order both records store them, after `fLocked`.
const FLAGS_IN_RECORD_ORDER = [
  'objects',
  'scenarios',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'selectLockedCells',
  'sort',
  'autoFilter',
  'pivotTables',
  'selectUnlockedCells',
] as const satisfies readonly (keyof SheetProtectionFlags)[];

/**
 * The half of the proof an ordered list owes: `satisfies` covers "no invented flag", this covers "no
 * omission". A flag left out here would be read at the wrong offset along with every flag after it.
 */
export type EveryProtectionFlagHasARecordSlot = AssertNever<
  Exclude<keyof SheetProtectionFlags, (typeof FLAGS_IN_RECORD_ORDER)[number]>
>;

// Whether each operation is allowed when the XML twin says nothing about it.
const ALLOWED_BY_DEFAULT = new Map<keyof SheetProtectionFlags, boolean>(
  SHEET_PROTECTION_FLAGS.map(({key, defaultForbidden}) => [key, !defaultForbidden]),
);

/** What a worksheet's protection records say, gathered as they arrive. */
export class SheetProtectionRecords {
  #locked = false;
  #flags: SheetProtectionFlags = {};
  #legacyPasswordHash: string | undefined;
  #credential: SheetProtectionCredential | undefined;

  /** Read a `BrtSheetProtection`. */
  legacy(reader: RecordReader): void {
    const protpwd = reader.u16();
    this.#readFlags(reader);
    // Zero is "no password": the hash of an empty password is not zero, and Excel writes zero beside
    // an agile credential.
    this.#legacyPasswordHash = protpwd === 0 ? undefined : hex(protpwd, 4);
  }

  /** Read a `BrtSheetProtectionIso`. */
  iso(reader: RecordReader): void {
    const spinCount = reader.u32();
    this.#readFlags(reader);
    const hashValue = toBase64(reader.bytes(reader.u32()));
    const saltValue = toBase64(reader.bytes(reader.u32()));
    this.#credential = {algorithmName: reader.wideString(), hashValue, saltValue, spinCount};
  }

  /** The protection the records described, or `undefined` for a sheet that is not protected. */
  result(): SheetProtection | undefined {
    if (!this.#locked) return undefined;
    return {
      flags: this.#flags,
      ...(this.#credential === undefined ? {} : {credential: this.#credential}),
      ...(this.#legacyPasswordHash === undefined
        ? {}
        : {legacyPasswordHash: this.#legacyPasswordHash}),
    };
  }

  #readFlags(reader: RecordReader): void {
    this.#locked = reader.u32() !== 0;
    const flags: {-readonly [K in keyof SheetProtectionFlags]?: boolean} = {};
    for (const key of FLAGS_IN_RECORD_ORDER) {
      const allowed = reader.u32() !== 0;
      if (allowed !== ALLOWED_BY_DEFAULT.get(key)) flags[key] = allowed;
    }
    this.#flags = flags;
  }
}
