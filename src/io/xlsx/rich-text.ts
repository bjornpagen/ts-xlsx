// Rich-text runs: a cell value composed of independently-formatted text runs.
//
// OOXML stores rich text as a sequence of `<r>` runs, each an optional `<rPr>` (the run's font, in
// the CT_RPrElt shape, identical to a styles `<font>` except the face element is `<rFont>`, not
// `<name>`) followed by a `<t>` text element. The writer serialises a rich-text value where it writes
// every other string value: inline (`t="inlineStr"`) by default, or as a rich `<si>` in the
// shared-strings table under `useSharedStrings`. The same runs render both, and the reader
// reconstructs them from either.

import type {RichTextRun} from '../../core/value.ts';
import {textElement} from '../../xml/xml.ts';
import {fontXml} from './font-xml.ts';

/**
 * Serialise a rich-text value's runs as the inner content of an `<is>` element. A zero-length run
 * is dropped: an empty `<t/>` is schema-invalid (Excel flags the file as corrupt) and an empty
 * run contributes nothing to the rendered text, so omitting it is loss-free.
 */
export function richTextRunsXml(runs: readonly RichTextRun[]): string {
  return runs
    .filter((run) => run.text !== '')
    .map((run) => {
      const rPr = run.font !== undefined ? `<rPr>${fontXml(run.font, 'rFont')}</rPr>` : '';
      return `<r>${rPr}${textElement(run.text)}</r>`;
    })
    .join('');
}
