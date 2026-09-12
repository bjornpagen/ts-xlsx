// The workbook's metadata part as `.xlsx` spells it, `xl/metadata.xml`, both directions: read into the
// codec-neutral model a cell's `cm` and `vm` resolve through (`../cell-metadata/metadata.ts`), and
// written, with the two rich-value parts a `vm` reaches, from what the cells written pointed at.
//
// The part's blocks sit in `<futureMetadata name>` lists, one per metadata type: an `XLDAPR` block
// carries an `<xda:dynamicArrayProperties fDynamic>` extension, an `XLRICHVALUE` block an `<xlrd:rvb i>`
// naming a rich value. `<cellMetadata>` and `<valueMetadata>` hold the blocks a `cm` and a `vm` index.
//
// Excel 16.0 (build 20326) saves one dynamic-array block and one cell-metadata block pointing at it,
// shared by every dynamic-array formula (`test/corpus/fixtures/excel-oracle/array-formulas.json`). It
// saves a rich value per error cell, and shares one between a pasted value and the cell it came from
// (`test/corpus/fixtures/excel-oracle/rich-value-errors.json`); the writer shares one between every
// cell holding the same error. Excel also saves `xl/richData/rdRichValueTypes.xml`, flags for keys no
// error has; it opened a package without that part clean and read every error in it back the same, so
// the writer does not emit it.

import type {ErrorCode} from '../../core/value.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml} from '../../xml/xml-read.ts';
import {boolStrict, localName} from '../../xml/xml-scan.ts';
import {XML_DECLARATION} from '../../xml/xml.ts';
import {
  type CellMetadataIndex,
  DYNAMIC_ARRAY_TYPE,
  indexCellMetadata,
  type MetadataRecord,
  RICH_VALUE_TYPE,
  type WorkbookMetadata,
} from '../cell-metadata/metadata.ts';
import {
  ERROR_STRUCTURE,
  parseRichValueErrors,
  RICH_VALUE_ERROR_TYPES,
} from '../cell-metadata/rich-values.ts';
import type {PartRelationships} from '../opc/read-opc.ts';
import {
  DYNAMIC_ARRAY_NS,
  DYNAMIC_ARRAY_PROPERTIES_EXT_URI,
  RICH_DATA_NS,
  RICH_VALUE_BLOCK_EXT_URI,
  SPREADSHEETML_NS,
} from './namespaces.ts';

/**
 * What every `cm` and `vm` in the workbook resolves to: the metadata part and the two rich-value parts,
 * each found through the workbook's relationships by type. Both worksheet readers resolve their cells
 * through this one answer.
 */
export function readCellMetadata(workbookRels: PartRelationships): CellMetadataIndex {
  return indexCellMetadata(
    parseMetadataPart(workbookRels.relatedText('sheetMetadata') ?? ''),
    parseRichValueErrors(
      workbookRels.relatedText('rdRichValueStructure') ?? '',
      workbookRels.relatedText('rdRichValue') ?? '',
    ),
  );
}

/** Read `xl/metadata.xml` as far as a cell's `cm` and `vm` reach into it. */
export function parseMetadataPart(xml: string): WorkbookMetadata {
  const typeNames: string[] = [];
  const dynamicArrayBlocks: boolean[] = [];
  const richValueBlocks: number[] = [];
  const cellBlocks: MetadataRecord[][] = [];
  const valueBlocks: MetadataRecord[][] = [];
  // Which of the four block lists a `<bk>` belongs to. A self-closing container holds no block, so it
  // opens none: left open, it would claim the blocks of the list after it.
  let container: 'dynamicArrays' | 'richValues' | 'cells' | 'values' | undefined;
  const records = (): MetadataRecord[] | undefined =>
    container === 'cells'
      ? cellBlocks.at(-1)
      : container === 'values'
        ? valueBlocks.at(-1)
        : undefined;
  if (xml === '') return {typeNames, dynamicArrayBlocks, richValueBlocks, cellBlocks, valueBlocks};
  parseXml(xml, {
    onOpen(name, attrs, selfClosing) {
      switch (localName(name)) {
        case 'metadataType':
          typeNames.push(attrs.name ?? '');
          break;
        case 'futureMetadata':
          container = selfClosing
            ? undefined
            : attrs.name === DYNAMIC_ARRAY_TYPE
              ? 'dynamicArrays'
              : attrs.name === RICH_VALUE_TYPE
                ? 'richValues'
                : undefined;
          break;
        case 'cellMetadata':
          container = selfClosing ? undefined : 'cells';
          break;
        case 'valueMetadata':
          container = selfClosing ? undefined : 'values';
          break;
        case 'bk':
          if (container === 'dynamicArrays') dynamicArrayBlocks.push(false);
          else if (container === 'richValues') richValueBlocks.push(-1);
          else if (container === 'cells') cellBlocks.push([]);
          else if (container === 'values') valueBlocks.push([]);
          break;
        case 'dynamicArrayProperties':
          if (container === 'dynamicArrays' && dynamicArrayBlocks.length > 0) {
            dynamicArrayBlocks[dynamicArrayBlocks.length - 1] = boolStrict(attrs.fDynamic);
          }
          break;
        case 'rvb':
          if (container === 'richValues' && richValueBlocks.length > 0) {
            richValueBlocks[richValueBlocks.length - 1] = numInteger(attrs.i, 0) ?? -1;
          }
          break;
        case 'rc': {
          const block = records();
          const type = numInteger(attrs.t, 1);
          const value = numInteger(attrs.v, 0);
          if (block !== undefined && type !== undefined && value !== undefined) {
            block.push({type, value});
          }
          break;
        }
        default:
          break;
      }
    },
    onClose(name) {
      const local = localName(name);
      if (local === 'futureMetadata' || local === 'cellMetadata' || local === 'valueMetadata') {
        container = undefined;
      }
    },
  });
  return {typeNames, dynamicArrayBlocks, richValueBlocks, cellBlocks, valueBlocks};
}

const METADATA_TYPE_FLAGS =
  'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" ' +
  'rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"';

/**
 * Records what the cells written so far point into, so the metadata part and the rich-value parts are
 * emitted only for a workbook that needs them. Filled during the sheet pass, as the shared-strings table
 * is, because the streaming writer renders a row long before the package around it is assembled.
 */
export class WorkbookMetadataTable {
  #dynamicArrays = false;
  // Each error a cell pointed at, in first-use order, with the `vm` its cells are written with: the n-th
  // is rich value n - 1 and value-metadata block n.
  readonly #errors = new Map<ErrorCode, number>();

  /** The `cm` a dynamic-array formula's cell is written with. Every such cell shares the one block. */
  markDynamicArray(): number {
    this.#dynamicArrays = true;
    return 1;
  }

  /**
   * The `vm` a cell holding `error` is written with, beside a `<v>` of `#VALUE!`, or `undefined` for
   * an error Excel reads literally, which needs no metadata.
   */
  markRichValueError(error: ErrorCode): number | undefined {
    if (!RICH_VALUE_ERROR_TYPES.has(error)) return undefined;
    const known = this.#errors.get(error);
    if (known !== undefined) return known;
    const vm = this.#errors.size + 1;
    this.#errors.set(error, vm);
    return vm;
  }

  /** Whether no cell points into the metadata part; the writer omits it entirely when so. */
  get isEmpty(): boolean {
    return !this.#dynamicArrays && this.#errors.size === 0;
  }

  /** Whether a cell points at a rich value, so the rich-value parts are emitted beside the metadata. */
  get hasRichValues(): boolean {
    return this.#errors.size > 0;
  }

  /** Serialise the `xl/metadata.xml` part, in the form Excel saves it. */
  toXml(): string {
    const dynamic = this.#dynamicArrays;
    const errorCount = this.#errors.size;
    const types = [
      ...(dynamic
        ? [`<metadataType name="${DYNAMIC_ARRAY_TYPE}" ${METADATA_TYPE_FLAGS} cellMeta="1"/>`]
        : []),
      ...(errorCount > 0
        ? [`<metadataType name="${RICH_VALUE_TYPE}" ${METADATA_TYPE_FLAGS}/>`]
        : []),
    ];
    // The rich-value type follows the dynamic-array one when both are declared, as Excel orders them.
    const richValueType = types.length;
    const indices = Array.from({length: errorCount}, (_, index) => index);
    return (
      XML_DECLARATION +
      `<metadata xmlns="${SPREADSHEETML_NS}"` +
      (errorCount > 0 ? ` xmlns:xlrd="${RICH_DATA_NS}"` : '') +
      (dynamic ? ` xmlns:xda="${DYNAMIC_ARRAY_NS}"` : '') +
      '>' +
      `<metadataTypes count="${types.length}">${types.join('')}</metadataTypes>` +
      (dynamic
        ? `<futureMetadata name="${DYNAMIC_ARRAY_TYPE}" count="1"><bk><extLst>` +
          `<ext uri="${DYNAMIC_ARRAY_PROPERTIES_EXT_URI}">` +
          '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>' +
          '</ext></extLst></bk></futureMetadata>'
        : '') +
      (errorCount > 0
        ? `<futureMetadata name="${RICH_VALUE_TYPE}" count="${errorCount}">` +
          indices
            .map(
              (index) =>
                `<bk><extLst><ext uri="${RICH_VALUE_BLOCK_EXT_URI}"><xlrd:rvb i="${index}"/></ext></extLst></bk>`,
            )
            .join('') +
          '</futureMetadata>'
        : '') +
      (dynamic ? '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' : '') +
      (errorCount > 0
        ? `<valueMetadata count="${errorCount}">` +
          indices.map((index) => `<bk><rc t="${richValueType}" v="${index}"/></bk>`).join('') +
          '</valueMetadata>'
        : '') +
      '</metadata>'
    );
  }

  /**
   * Serialise `xl/richData/rdrichvalue.xml`: one rich value per error, in the order the `vm`s were handed
   * out, each naming its `errorType` and the unknown sub-type, 0.
   */
  richValuesXml(): string {
    const values = [...this.#errors.keys()].map(
      (error) => `<rv s="0"><v>${RICH_VALUE_ERROR_TYPES.get(error)}</v><v>0</v></rv>`,
    );
    return (
      XML_DECLARATION +
      `<rvData xmlns="${RICH_DATA_NS}" count="${values.length}">${values.join('')}</rvData>`
    );
  }
}

/** `xl/richData/rdrichvaluestructure.xml`: the one structure every rich value the writer emits has. */
export const RICH_VALUE_STRUCTURES_XML =
  XML_DECLARATION +
  `<rvStructures xmlns="${RICH_DATA_NS}" count="1">` +
  `<s t="${ERROR_STRUCTURE}"><k n="errorType" t="i"/><k n="subType" t="i"/></s>` +
  '</rvStructures>';
