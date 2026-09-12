// The rich values a cell's value metadata names, as far as they are errors.
//
// Excel stores an error it has no literal for as `<v>#VALUE!</v>` and names the real one in a rich value:
// an `<rv s>` in `xl/richData/rdrichvalue.xml` whose structure, the `s`-th `<s t="_error">` of
// `xl/richData/rdrichvaluestructure.xml`, lists its keys in the order the `<rv>`'s `<v>` children give
// their values. The key that matters is `errorType`. Both parts are XML in either serialisation of a
// workbook.
//
// Which `errorType` is which error came from Excel 16.0 (build 20326): one constant cell per number,
// opened over COM, read back as `Range.Text` and `Value2`
// (`test/corpus/fixtures/excel-oracle/rich-value-errors.json`). Only the numbers whose text and CVErr
// code both name the same error are here. The six Excel has no literal for are the ones a writer must
// store this way, and the three classic ones it also reads back are listed so a file storing them this
// way reads as Excel shows it.

import type {ErrorCode} from '../../core/value.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';

/** The structure type of a rich value naming an error. */
export const ERROR_STRUCTURE = '_error';

/** The `errorType` a rich value names each error with, for the errors Excel stores only this way. */
export const RICH_VALUE_ERROR_TYPES: ReadonlyMap<ErrorCode, number> = new Map<ErrorCode, number>([
  ['#SPILL!', 8],
  ['#CONNECT!', 9],
  ['#BLOCKED!', 10],
  ['#UNKNOWN!', 11],
  ['#FIELD!', 12],
  ['#CALC!', 13],
]);

const ERROR_OF_TYPE: ReadonlyMap<number, ErrorCode> = new Map<number, ErrorCode>([
  [3, '#REF!'],
  [4, '#NAME?'],
  [6, '#N/A'],
  ...[...RICH_VALUE_ERROR_TYPES].map(([code, type]): [number, ErrorCode] => [type, code]),
]);

/**
 * The error each rich value holds, keyed by its 0-based index in the rich-value part. A value whose
 * structure is not `_error`, has no `errorType` key, or names an `errorType` Excel was not seen to read
 * back as one error holds none.
 */
export function parseRichValueErrors(
  structuresXml: string,
  valuesXml: string,
): ReadonlyMap<number, ErrorCode> {
  const errors = new Map<number, ErrorCode>();
  if (structuresXml === '' || valuesXml === '') return errors;
  const errorTypeKeys = parseErrorTypeKeys(structuresXml);

  // Only the one `<v>` holding the value's `errorType` is gathered; every other key's text is skipped,
  // so a part of long strings costs a scan and no copies.
  let index = -1;
  let wanted = -1;
  let position = -1;
  let capturing = false;
  let text = '';
  parseXml(valuesXml, {
    onOpen(name, attrs, selfClosing) {
      const local = localName(name);
      if (local === 'rv') {
        index++;
        position = -1;
        const structure = numInteger(attrs.s, 0);
        wanted = structure === undefined ? -1 : (errorTypeKeys[structure] ?? -1);
      } else if (local === 'v') {
        position++;
        capturing = !selfClosing && position === wanted;
        text = '';
      }
    },
    onText(chunk) {
      if (capturing) text += chunk;
    },
    onClose(name) {
      if (localName(name) !== 'v' || !capturing) return;
      capturing = false;
      const error = ERROR_OF_TYPE.get(numInteger(text.trim(), 0) ?? -1);
      if (error !== undefined) errors.set(index, error);
    },
  });
  return errors;
}

// Where each structure keeps its `errorType`, by structure index: -1 for a structure that is not an
// error or names no such key.
function parseErrorTypeKeys(xml: string): number[] {
  const keys: number[] = [];
  let isError = false;
  let key = -1;
  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      if (local === 's') {
        isError = attrs.t === ERROR_STRUCTURE;
        key = -1;
        keys.push(-1);
      } else if (local === 'k' && keys.length > 0) {
        key++;
        if (isError && attrs.n === 'errorType' && keys[keys.length - 1] === -1) {
          keys[keys.length - 1] = key;
        }
      }
    },
  });
  return keys;
}
