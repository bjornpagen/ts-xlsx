// The workbook's metadata part as `.xlsb` spells it, `xl/metadata.bin`, read into the codec-neutral model
// a cell's `BrtCellMeta` and `BrtValueMeta` resolve through (`../cell-metadata/metadata.ts`).
//
// The records mirror `xl/metadata.xml` one for one: a `BrtMdtinfo` per metadata type, a
// `BrtBeginEsfmd`/`BrtEndEsfmd` list of future-metadata blocks per type, each block opened by a
// `BrtBeginFmd`, and two `BrtBeginEsmdb`/`BrtEndEsmdb` lists of `BrtMdb` blocks, whose flag says whether
// they are the cell or the value metadata. The two records inside a block that carry what the XML form's
// extensions do were read as Excel 16.0 (build 20326) writes and reads them
// (`test/corpus/fixtures/excel-oracle/xlsb-cell-metadata.json`): a dynamic-array block holds its flags in
// record 4097, `fDynamic` in bit 0, and a rich-value block names its rich value in the second word of
// record 5003.

import {
  DYNAMIC_ARRAY_TYPE,
  type MetadataRecord,
  NO_METADATA,
  RICH_VALUE_TYPE,
  type WorkbookMetadata,
} from '../cell-metadata/metadata.ts';
import {RecordReader} from './primitives.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

/**
 * Read `xl/metadata.bin` as far as a cell's metadata indices reach into it.
 *
 * @throws {XlsbParseError} if a record the model reads is shorter than its fields.
 */
export function parseMetadataPart(part: Uint8Array | undefined): WorkbookMetadata {
  if (part === undefined) return NO_METADATA;
  const typeNames: string[] = [];
  const dynamicArrayBlocks: boolean[] = [];
  const richValueBlocks: number[] = [];
  const cellBlocks: MetadataRecord[][] = [];
  const valueBlocks: MetadataRecord[][] = [];
  // The type whose future-metadata list is open, and the list a `BrtMdb` joins.
  let futureType: string | undefined;
  let blocks: MetadataRecord[][] | undefined;
  for (const record of readRecords(part)) {
    const reader = new RecordReader(record.data);
    switch (record.type) {
      case BRT.Mdtinfo:
        reader.skip(8); // The type's copy and paste behaviour, and the lowest version supporting it.
        typeNames.push(reader.wideString());
        break;
      case BRT.BeginEsfmd:
        reader.skip(4); // The block count, which the blocks that follow restate.
        futureType = reader.wideString();
        break;
      case BRT.EndEsfmd:
        futureType = undefined;
        break;
      case BRT.BeginFmd:
        if (futureType === DYNAMIC_ARRAY_TYPE) dynamicArrayBlocks.push(false);
        else if (futureType === RICH_VALUE_TYPE) richValueBlocks.push(-1);
        break;
      case BRT.DynamicArrayProperties:
        if (futureType === DYNAMIC_ARRAY_TYPE && dynamicArrayBlocks.length > 0) {
          dynamicArrayBlocks[dynamicArrayBlocks.length - 1] = (reader.u8() & 1) !== 0;
        }
        break;
      case BRT.RichValueBlock:
        if (futureType === RICH_VALUE_TYPE && richValueBlocks.length > 0) {
          reader.skip(4);
          richValueBlocks[richValueBlocks.length - 1] = reader.u32();
        }
        break;
      case BRT.BeginEsmdb:
        reader.skip(4); // The block count, as above.
        blocks = (reader.u32() & 1) !== 0 ? cellBlocks : valueBlocks;
        break;
      case BRT.EndEsmdb:
        blocks = undefined;
        break;
      case BRT.Mdb:
        blocks?.push(metadataRecords(reader));
        break;
      default:
        break;
    }
  }
  return {typeNames, dynamicArrayBlocks, richValueBlocks, cellBlocks, valueBlocks};
}

// A `BrtMdb`: a record count, then each record's type and value. The count is the file's, so it bounds
// nothing: the records read are the ones the payload holds.
function metadataRecords(reader: RecordReader): MetadataRecord[] {
  const count = reader.u32();
  const records: MetadataRecord[] = [];
  while (records.length < count && reader.remaining >= 8) {
    records.push({type: reader.u32(), value: reader.u32()});
  }
  return records;
}
