// The workbook's metadata part as a cell reaches it, independent of how the part is spelled.
//
// A cell carries up to two indices into the part. `cm` points at cell metadata, which is how Excel marks
// a dynamic-array formula; `vm` points at value metadata, which is how it names the error behind a
// `#VALUE!` it could not store literally. Each is a 1-based index into a list of blocks, each block a
// list of records `(t, v)`: `t` a 1-based index into the part's metadata types, `v` a 0-based index
// into that type's own future-metadata blocks. The `XLDAPR` type's blocks hold dynamic-array properties;
// the `XLRICHVALUE` type's blocks each name one rich value, by its index in the rich-value part.
//
// `.xlsx` spells the part as `xl/metadata.xml`, and that codec reads it into {@link WorkbookMetadata};
// what a cell's indices then resolve to is decided here, once.

import type {ErrorCode} from '../../core/value.ts';

/** The metadata type whose blocks mark a dynamic-array formula. */
export const DYNAMIC_ARRAY_TYPE = 'XLDAPR';

/** The metadata type whose blocks name a rich value. */
export const RICH_VALUE_TYPE = 'XLRICHVALUE';

/** One record of a metadata block: `type` indexes the metadata types from 1, `value` that type's blocks from 0. */
export interface MetadataRecord {
  readonly type: number;
  readonly value: number;
}

/** The metadata part, read as far as a cell's `cm` and `vm` can reach into it. */
export interface WorkbookMetadata {
  /** Each metadata type's name, in order. */
  readonly typeNames: readonly string[];
  /** Whether each dynamic-array block, in order, sets `fDynamic`. */
  readonly dynamicArrayBlocks: readonly boolean[];
  /** The rich value each rich-value block names, in order; -1 for a block naming none it can read. */
  readonly richValueBlocks: readonly number[];
  /** Each cell-metadata block's records, in order: a `cm` of 1 is the first. */
  readonly cellBlocks: readonly (readonly MetadataRecord[])[];
  /** Each value-metadata block's records, in order: a `vm` of 1 is the first. */
  readonly valueBlocks: readonly (readonly MetadataRecord[])[];
}

/** A workbook with no metadata part. */
export const NO_METADATA: WorkbookMetadata = {
  typeNames: [],
  dynamicArrayBlocks: [],
  richValueBlocks: [],
  cellBlocks: [],
  valueBlocks: [],
};

/** What a cell's metadata indices resolve to, for the readers of every sheet in the workbook. */
export interface CellMetadataIndex {
  /** The `cm` values marking a dynamic-array formula. */
  readonly dynamicArrayCells: ReadonlySet<number>;
  /** The error each `vm` names through a rich value, where it names one. */
  readonly valueErrors: ReadonlyMap<number, ErrorCode>;
}

/** The index of a workbook with no metadata part: no cell is a dynamic array, and no value names an error. */
export const NO_CELL_METADATA: CellMetadataIndex = {
  dynamicArrayCells: new Set(),
  valueErrors: new Map(),
};

/**
 * Resolve every `cm` and `vm` the metadata part can answer. `richValueErrors` is the error each rich
 * value holds, by its index in the rich-value part.
 *
 * An index resolving to no block, to a block of another type, or to a rich value that is not an error
 * marks nothing, so a cell carrying it reads as its own `<v>` says.
 */
export function indexCellMetadata(
  metadata: WorkbookMetadata,
  richValueErrors: ReadonlyMap<number, ErrorCode>,
): CellMetadataIndex {
  const typeOf = (record: MetadataRecord): string | undefined =>
    metadata.typeNames[record.type - 1];
  const dynamicArrayCells = new Set<number>();
  metadata.cellBlocks.forEach((records, index) => {
    const dynamic = records.some(
      (record) =>
        typeOf(record) === DYNAMIC_ARRAY_TYPE && metadata.dynamicArrayBlocks[record.value] === true,
    );
    if (dynamic) dynamicArrayCells.add(index + 1);
  });
  const valueErrors = new Map<number, ErrorCode>();
  metadata.valueBlocks.forEach((records, index) => {
    for (const record of records) {
      if (typeOf(record) !== RICH_VALUE_TYPE) continue;
      const error = richValueErrors.get(metadata.richValueBlocks[record.value] ?? -1);
      if (error !== undefined) {
        valueErrors.set(index + 1, error);
        break;
      }
    }
  });
  return {dynamicArrayCells, valueErrors};
}
