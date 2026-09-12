// The workbook's cell metadata (`xl/metadata.xml`), both directions: which `<c cm>` indices mark a
// dynamic-array formula, and the part a written one points at.
//
// A `cm` is a 1-based index into `<cellMetadata>`, whose `<bk>` blocks each list `<rc t v>` records:
// `t` a 1-based index into `<metadataTypes>`, `v` a 0-based index into that type's own blocks. The
// dynamic-array type is `XLDAPR`, whose blocks sit in `<futureMetadata name="XLDAPR">` and carry an
// `<xda:dynamicArrayProperties fDynamic>` extension. Excel saves one such block and one cell-metadata
// block pointing at it, shared by every dynamic-array formula in the workbook (Excel 16.0 build 20326,
// `test/corpus/fixtures/excel-oracle/array-formulas.json`).
//
// The part holds more than this: a rich value's `vm` resolves through `<valueMetadata>` the same way.
// Nothing reads that half yet, and the writer builds the part from the model, so it does not survive a
// save.

import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml} from '../../xml/xml-read.ts';
import {boolStrict, localName} from '../../xml/xml-scan.ts';
import {XML_DECLARATION} from '../../xml/xml.ts';
import {
  DYNAMIC_ARRAY_NS,
  DYNAMIC_ARRAY_PROPERTIES_EXT_URI,
  SPREADSHEETML_NS,
} from './namespaces.ts';

const DYNAMIC_ARRAY_TYPE = 'XLDAPR';

/**
 * The `cm` indices whose cell metadata marks a dynamic-array formula. An index resolving to no block,
 * to a block of another type, or to properties with `fDynamic` off marks nothing, so a cell carrying
 * it reads as a legacy array formula.
 */
export function parseDynamicArrayCellMetadata(xml: string): ReadonlySet<number> {
  const marked = new Set<number>();
  if (xml === '') return marked;
  const typeNames: string[] = [];
  // Whether each `XLDAPR` block, in order, sets `fDynamic`.
  const dynamicBlocks: boolean[] = [];
  // Each cell-metadata block's records, in order.
  const cellBlocks: {readonly t: number; readonly v: number}[][] = [];
  // Which of the two block lists a `<bk>` belongs to. A self-closing container holds no block, so it
  // opens none: left open, it would claim the blocks of the list after it.
  let container: 'dynamicArrays' | 'cells' | undefined;
  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      switch (localName(name)) {
        case 'metadataType':
          typeNames.push(attrs.name ?? '');
          break;
        case 'futureMetadata':
          container =
            !selfClosing && attrs.name === DYNAMIC_ARRAY_TYPE ? 'dynamicArrays' : undefined;
          break;
        case 'cellMetadata':
          container = selfClosing ? undefined : 'cells';
          break;
        case 'valueMetadata':
          container = undefined;
          break;
        case 'bk':
          if (container === 'dynamicArrays') dynamicBlocks.push(false);
          else if (container === 'cells') cellBlocks.push([]);
          break;
        case 'dynamicArrayProperties':
          if (container === 'dynamicArrays' && dynamicBlocks.length > 0) {
            dynamicBlocks[dynamicBlocks.length - 1] = boolStrict(attrs.fDynamic);
          }
          break;
        case 'rc': {
          const block = cellBlocks.at(-1);
          const t = numInteger(attrs.t, 1);
          const v = numInteger(attrs.v, 0);
          if (container === 'cells' && block !== undefined && t !== undefined && v !== undefined) {
            block.push({t, v});
          }
          break;
        }
        default:
          break;
      }
    },
    onClose(name) {
      const local = localName(name);
      if (local === 'futureMetadata' || local === 'cellMetadata') container = undefined;
    },
  });
  cellBlocks.forEach((records, index) => {
    const dynamic = records.some(
      ({t, v}) => typeNames[t - 1] === DYNAMIC_ARRAY_TYPE && dynamicBlocks[v] === true,
    );
    if (dynamic) marked.add(index + 1);
  });
  return marked;
}

/**
 * Records whether any cell written so far points into the cell metadata, so the part is emitted only
 * for a workbook that needs it. Filled during the sheet pass, as the shared-strings table is, because
 * the streaming writer renders a row long before the package around it is assembled.
 */
export class CellMetadataTable {
  #dynamicArrays = false;

  /** The `cm` a dynamic-array formula's cell is written with. Every such cell shares the one block. */
  markDynamicArray(): number {
    this.#dynamicArrays = true;
    return 1;
  }

  /** Whether no cell points into the part; the writer omits it entirely when so. */
  get isEmpty(): boolean {
    return !this.#dynamicArrays;
  }

  /** Serialise the `xl/metadata.xml` part, in the form Excel saves it. */
  toXml(): string {
    return (
      XML_DECLARATION +
      `<metadata xmlns="${SPREADSHEETML_NS}" xmlns:xda="${DYNAMIC_ARRAY_NS}">` +
      '<metadataTypes count="1">' +
      `<metadataType name="${DYNAMIC_ARRAY_TYPE}" minSupportedVersion="120000" copy="1" pasteAll="1" ` +
      'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" ' +
      'assign="1" coerce="1" cellMeta="1"/>' +
      '</metadataTypes>' +
      `<futureMetadata name="${DYNAMIC_ARRAY_TYPE}" count="1"><bk><extLst>` +
      `<ext uri="${DYNAMIC_ARRAY_PROPERTIES_EXT_URI}">` +
      '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>' +
      '</ext></extLst></bk></futureMetadata>' +
      '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
      '</metadata>'
    );
  }
}
