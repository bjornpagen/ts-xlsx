// The BIFF12 record types this reader consumes, by their [MS-XLSB] "Record Enumeration" (section 2.3)
// numbers. Named constants rather than magic numbers at the switch sites: a record number carries no
// hint of what it is, so `BRT.CellRk` is the difference between a readable parser and a table of
// unexplained integers.
//
// Deliberately partial. [MS-XLSB] enumerates ~760 record types; listing the ones we do not decode
// would claim coverage we do not have. A record whose type is absent here is skipped by the parsers,
// which is the correct behaviour for a forward-compatible reader: the format grows, and an unknown
// record is always safely framed (its size is in the stream) even when its meaning is not known.

/** BIFF12 record numbers, grouped by the part whose stream they appear in. */
export const BRT = {
  // Cell table (worksheet part). The single-byte record space: these are the hot path.
  RowHdr: 0,
  CellBlank: 1,
  CellRk: 2,
  CellError: 3,
  CellBool: 4,
  CellReal: 5,
  CellSt: 6,
  CellIsst: 7,
  FmlaString: 8,
  FmlaNum: 9,
  FmlaBool: 10,
  FmlaError: 11,
  CellRString: 62,
  ColInfo: 60,
  MergeCell: 176,

  // Shared-string table.
  SSTItem: 19,
  EndSst: 160,

  // Style sheet: each collection is a Begin/End pair around its entries, and `XF` appears inside two
  // of them (the named-style layer and the direct-format layer), so the style parser tracks which
  // collection it is in rather than keying on the record number alone.
  Fmt: 44,
  Font: 43,
  Fill: 45,
  Border: 46,
  XF: 47,
  Style: 48,
  BeginFmts: 615,
  EndFmts: 616,
  BeginFonts: 611,
  EndFonts: 612,
  BeginFills: 603,
  EndFills: 604,
  BeginBorders: 613,
  EndBorders: 614,
  BeginCellStyleXFs: 626,
  EndCellStyleXFs: 627,
  BeginCellXFs: 617,
  EndCellXFs: 618,
  BeginStyles: 619,
  EndStyles: 620,

  // Workbook part.
  WbProp: 153,
  BundleSh: 156,
  BeginBundleShs: 143,
  EndBundleShs: 144,
  Name: 39,

  // The externals block: which workbooks a formula can reach, and which sheets of them each `ixti` a
  // 3-D reference carries names. `SupSelf` declares a supporting book that is *this* workbook.
  BeginExternals: 353,
  EndExternals: 354,
  SupSelf: 357,
  ExternSheet: 362,

  // Worksheet structure.
  WsFmtInfo: 485,
  ArrFmla: 426,
  SheetProtection: 535,
  SheetProtectionIso: 678,
} as const;
