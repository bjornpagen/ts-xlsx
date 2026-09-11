// The references a preserved part spells, moved through the row and column splices made to the
// workbook since it was read.
//
// A chart's series and a pivot cache's source name cells on the workbook's sheets, and Excel moves
// them on an insert or a delete as it moves a cell formula (Excel 16.0, recorded in
// `test/corpus/fixtures/excel-oracle/references-outside-cells-through-a-splice.json`). The model keeps
// both parts as bytes it does not read, and `core/` may not parse XML, so the workbook journals each
// splice and the writer replays the journal here, over the part as it was read. Each edit is made at
// the offsets the scanner found, so everything else in the part rides through byte for byte, and a
// part nothing moved in is handed back as the very bytes it came in as.
//
// What a text rewrite cannot reproduce: Excel re-derives a series whose name or categories a delete
// takes, where this writes `#REF!` into the reference the delete took.

import {type SheetSplice, spliceFormula} from '../../core/formula-references.ts';
import {splicePivotSource} from '../../core/pivot-table.ts';
import {XmlParseError} from '../../xml/errors.ts';
import {elementRange, elementRanges, openElements} from '../../xml/xml-read.ts';
import {decodeEntities} from '../../xml/xml-scan.ts';
import {emptyElement, escapeText} from '../../xml/xml.ts';
import {relAttr} from '../opc/namespaces.ts';

interface PartEdit {
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

type Rewriter = (xml: string, splices: readonly SheetSplice[]) => PartEdit[];

// Every reference in a chart part is the text of an element named `f`: a series' values, categories
// and name (`c:f`), a chartex data dimension (`cx:f`), and an extension's data label range (`c15:f`).
// A chart spells each one with its sheet, so no reference in it has a home sheet.
const chartReferences: Rewriter = (xml, splices) => {
  const edits: PartEdit[] = [];
  for (const range of elementRanges(xml, 'f')) {
    const content = xml.slice(range.contentStart, range.contentEnd);
    // Markup inside a reference (a CDATA section, say) is not a spelling this edits: leaving it keeps
    // the part as it was rather than escaping the markup into text.
    if (content.includes('<')) continue;
    const formula = decodeEntities(content);
    const moved = splices.reduce((text, edit) => spliceFormula(text, undefined, edit), formula);
    if (moved !== formula) {
      edits.push({start: range.contentStart, end: range.contentEnd, text: escapeText(moved)});
    }
  }
  return edits;
};

// A pivot cache drawing from a range of this workbook names it in `<worksheetSource sheet ref>`. One
// naming a range of another workbook, through a relationship, is not this workbook's to move, and one
// naming a table or a defined name carries no `ref` at all.
const pivotCacheSource: Rewriter = (xml, splices) => {
  for (const source of openElements(xml, 'worksheetSource')) {
    if (relAttr(source.scope, source.attrs, 'id') !== undefined) return [];
    break;
  }
  const range = elementRange(xml, ['cacheSource', 'worksheetSource']);
  const {sheet, ref} = range?.attrs ?? {};
  if (range === undefined || sheet === undefined || ref === undefined) return [];
  const moved = splices.reduce((text, edit) => splicePivotSource(sheet, text, edit), ref);
  if (moved === ref) return [];
  // A `<worksheetSource>` has no children, so an element written with a close tag is replaced whole
  // by its empty form.
  return [
    {
      start: range.start,
      end: range.end,
      text: emptyElement(range.name, {...range.attrs, ref: moved}),
    },
  ];
};

// Keyed by the media type in lower case, which is how a lookup spells it.
const REWRITERS: ReadonlyMap<string, Rewriter> = new Map(
  (
    [
      ['application/vnd.openxmlformats-officedocument.drawingml.chart+xml', chartReferences],
      ['application/vnd.ms-office.chartex+xml', chartReferences],
      [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
        pivotCacheSource,
      ],
    ] as const
  ).map(([type, rewrite]) => [type.toLowerCase(), rewrite]),
);

/**
 * A preserved part's bytes with the references it spells moved through `splices`, in order, or the
 * same bytes when the part is of a kind that spells none or nothing in it moved.
 */
export function splicePreservedPart(
  bytes: Uint8Array,
  contentType: string,
  splices: readonly SheetSplice[],
): Uint8Array {
  if (splices.length === 0) return bytes;
  // A media type compares ignoring case (RFC 9110), and a package is free to spell one in capitals.
  const rewrite = REWRITERS.get(contentType.toLowerCase());
  if (rewrite === undefined) return bytes;
  const source = new TextDecoder().decode(bytes);
  let edits: PartEdit[];
  try {
    edits = rewrite(source, splices);
  } catch (error) {
    // A part the scanner cannot read was carried verbatim when it was read, and a splice is no reason
    // to start refusing it at the save.
    if (error instanceof XmlParseError) return bytes;
    throw error;
  }
  if (edits.length === 0) return bytes;
  let xml = source;
  // Last to first, so an earlier edit cannot move a later one's offsets.
  for (const edit of edits.toSorted((a, b) => b.start - a.start)) {
    xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end);
  }
  return new TextEncoder().encode(xml);
}
