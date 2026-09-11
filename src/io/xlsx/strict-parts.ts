// A part of an ISO/IEC 29500 Strict package, carried into the Transitional package every write produces.
//
// Strict is the same vocabulary as Transitional under other names: its namespaces live under
// `http://purl.oclc.org/ooxml/`, and the percentages DrawingML stores as thousandths (`78000`) are
// written as strings (`78%`). The modelled parts are read by local name and written from the model, so
// none of that reaches them. A part the model does not read rides through as bytes, though, and one
// still spelled in Strict, inside a package whose relationships and content types say Transitional, is
// a part a consumer cannot place: the Open XML SDK could not load a Strict workbook's theme after a
// round trip at all (`PackageOpenError`), and once its namespace was translated it refused every `78%`
// as not an `Int32`. So each preserved part is translated as the reader captures it. Relationship
// types are translated one layer down, as the container reads them (`io/opc/strict-relationships.ts`).
//
// Both tables are the schema's own, read out of the ECMA-376 graph the `ooxml-lookup` skill queries:
// the namespace pairs are every vocabulary declared under both profiles, and the percentages every
// attribute whose Strict type is a DrawingML, chart or diagram percentage. Excel 16.0 saving as Strict
// writes the DrawingML ones as `N%` and the chart ones as plain integers, which is the Transitional
// spelling and not the Strict one; a chart's `N%` is translated all the same, since it is the spelling
// the Strict schema asks a producer for.

import {XmlParseError} from '../../xml/errors.ts';
import {NamespaceScope} from '../../xml/xml-namespaces.ts';
import {tagRanges} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {startTag} from '../../xml/xml.ts';

const STRICT = 'http://purl.oclc.org/ooxml/';
const TRANSITIONAL = 'http://schemas.openxmlformats.org/';

// Each Strict namespace and the Transitional one of the same vocabulary, as `strict transitional`
// under the two roots.
const NAMESPACES: ReadonlyMap<string, string> = new Map(
  [
    'drawingml/chart drawingml/2006/chart',
    'drawingml/chartDrawing drawingml/2006/chartDrawing',
    'drawingml/diagram drawingml/2006/diagram',
    'drawingml/lockedCanvas drawingml/2006/lockedCanvas',
    'drawingml/main drawingml/2006/main',
    'drawingml/picture drawingml/2006/picture',
    'drawingml/spreadsheetDrawing drawingml/2006/spreadsheetDrawing',
    'drawingml/wordprocessingDrawing drawingml/2006/wordprocessingDrawing',
    'officeDocument/bibliography officeDocument/2006/bibliography',
    'officeDocument/characteristics officeDocument/2006/characteristics',
    'officeDocument/customProperties officeDocument/2006/custom-properties',
    'officeDocument/customXml officeDocument/2006/customXml',
    'officeDocument/docPropsVTypes officeDocument/2006/docPropsVTypes',
    'officeDocument/extendedProperties officeDocument/2006/extended-properties',
    'officeDocument/math officeDocument/2006/math',
    'officeDocument/relationships officeDocument/2006/relationships',
    'officeDocument/sharedTypes officeDocument/2006/sharedTypes',
    'presentationml/main presentationml/2006/main',
    'schemaLibrary/main schemaLibrary/2006/main',
    'spreadsheetml/main spreadsheetml/2006/main',
    'wordprocessingml/main wordprocessingml/2006/main',
  ].map((pair) => {
    const [strict, transitional] = pair.split(' ');
    return [`${STRICT}${strict}`, `${TRANSITIONAL}${transitional}`];
  }),
);

const DRAWINGML = `${TRANSITIONAL}drawingml/2006/main`;

// The elements whose attribute holds a percentage, keyed by that attribute, per vocabulary, with what
// one percent is in Transitional: thousandths in DrawingML and diagrams, whole percent in a chart.
// `a:buSzPct` is absent on purpose: Transitional spells it `N%` too.
const PERCENTAGES: ReadonlyMap<string, PercentageTable> = new Map([
  [
    DRAWINGML,
    percentageTable(1000, {
      a: 'alphaRepl',
      amt: 'alphaModFix tint',
      b: 'fillRect fillToRect scrgbClr srcRect tileRect',
      baseline: 'defRPr endParaRPr rPr',
      bright: 'lum',
      contrast: 'lum',
      d: 'ds',
      endA: 'reflection',
      endPos: 'reflection',
      fontScale: 'normAutofit',
      g: 'scrgbClr',
      l: 'fillRect fillToRect srcRect tileRect',
      lim: 'miter',
      lnSpcReduction: 'normAutofit',
      lum: 'hsl hslClr',
      pos: 'gs',
      r: 'fillRect fillToRect scrgbClr srcRect tileRect',
      sat: 'hsl hslClr',
      sp: 'ds',
      stA: 'reflection',
      stPos: 'reflection',
      sx: 'outerShdw reflection tile xfrm',
      sy: 'outerShdw reflection tile xfrm',
      t: 'fillRect fillToRect srcRect tileRect',
      thresh: 'alphaBiLevel biLevel',
      tx: 'relOff',
      ty: 'relOff',
      val:
        'alpha alphaMod alphaOff blue blueMod blueOff green greenMod greenOff hueMod lum lumMod ' +
        'lumOff red redMod redOff sat satMod satOff shade spcPct tint',
      zoom: 'camera',
    }),
  ],
  [
    `${TRANSITIONAL}drawingml/2006/chart`,
    percentageTable(1, {
      val:
        'bubbleScale depthPercent gapDepth gapWidth hPercent holeSize lblOffset overlap ' +
        'secondPieSize thickness',
    }),
  ],
  [
    `${TRANSITIONAL}drawingml/2006/diagram`,
    percentageTable(
      1000,
      Object.fromEntries(
        [
          'custLinFactNeighborX',
          'custLinFactNeighborY',
          'custLinFactX',
          'custLinFactY',
          'custRadScaleInc',
          'custRadScaleRad',
          'custScaleX',
          'custScaleY',
        ].map((attribute) => [attribute, 'prSet']),
      ),
    ),
  ],
]);

interface PercentageTable {
  /** One percent, in the integer Transitional stores. */
  readonly scale: number;
  /** `element attribute`, for each attribute of an element that holds a percentage. */
  readonly pairs: ReadonlySet<string>;
}

function percentageTable(
  scale: number,
  elementsByAttribute: Record<string, string>,
): PercentageTable {
  const pairs = Object.entries(elementsByAttribute).flatMap(([attribute, elements]) =>
    elements.split(' ').map((element) => `${element} ${attribute}`),
  );
  return {scale, pairs: new Set(pairs)};
}

const PERCENTAGE = /^-?[0-9]+(?:\.[0-9]+)?%$/;

const STRICT_BYTES = new TextEncoder().encode(STRICT);

/**
 * A package part's bytes in Transitional: an XML part declaring a Strict namespace has its
 * declarations and its percentages translated, and any other part is handed back as the same bytes.
 * Only the opening tags that change are re-rendered, so everything else in the part rides through byte
 * for byte. A part the scanner cannot read is left as it came, as it would have been carried.
 */
export function transitionalPart(bytes: Uint8Array, contentType: string): Uint8Array {
  if (!isXmlContentType(contentType) || !includesBytes(bytes, STRICT_BYTES)) return bytes;
  const source = new TextDecoder().decode(bytes);
  let xml = '';
  let copied = 0;
  const scope = new NamespaceScope();
  try {
    for (const tag of tagRanges(source)) {
      if (tag.close) {
        scope.close();
        continue;
      }
      scope.open(tag.attrs);
      const translated = translatedAttributes(tag.name, tag.attrs, scope);
      if (tag.selfClosing) scope.close();
      if (translated === undefined) continue;
      xml += source.slice(copied, tag.start) + startTag(tag.name, translated, tag.selfClosing);
      copied = tag.end;
    }
  } catch (error) {
    if (error instanceof XmlParseError) return bytes;
    throw error;
  }
  return copied === 0 ? bytes : new TextEncoder().encode(xml + source.slice(copied));
}

// The attributes of one opening tag with its namespace declarations and its percentages translated,
// or `undefined` when nothing in it is Strict. The element's namespace is resolved against the scope as
// the source declared it, then translated, so a Strict and a Transitional spelling of one vocabulary
// select the same percentages.
function translatedAttributes(
  name: string,
  attrs: {readonly [attribute: string]: string},
  scope: NamespaceScope,
): Record<string, string> | undefined {
  let translated: Record<string, string> | undefined;
  const set = (attribute: string, value: string): void => {
    translated ??= {...attrs};
    translated[attribute] = value;
  };
  const declared = scope.elementNamespace(name);
  const namespace = (declared === undefined ? undefined : NAMESPACES.get(declared)) ?? declared;
  const local = localName(name);
  const percentages = PERCENTAGES.get(namespace ?? '');
  for (const attribute in attrs) {
    const value = attrs[attribute] ?? '';
    if (attribute === 'xmlns' || attribute.startsWith('xmlns:')) {
      const uri = NAMESPACES.get(value);
      if (uri !== undefined) set(attribute, uri);
    } else if (attribute === 'uri' && local === 'graphicData' && namespace === DRAWINGML) {
      // A graphic frame names the vocabulary of what it holds by namespace URI, in an attribute value
      // rather than a declaration: `<a:graphicData uri="…/chart">` over a chart.
      const uri = NAMESPACES.get(value);
      if (uri !== undefined) set(attribute, uri);
    } else if (percentages?.pairs.has(`${local} ${attribute}`) === true && PERCENTAGE.test(value)) {
      set(attribute, String(Math.round(Number.parseFloat(value) * percentages.scale)));
    }
  }
  return translated;
}

function isXmlContentType(contentType: string): boolean {
  const type = contentType.toLowerCase();
  return type.endsWith('+xml') || type === 'application/xml' || type === 'text/xml';
}

function includesBytes(haystack: Uint8Array, needle: Uint8Array): boolean {
  const [first = 0] = needle;
  const last = haystack.length - needle.length;
  for (let i = haystack.indexOf(first); i !== -1 && i <= last; i = haystack.indexOf(first, i + 1)) {
    let matched = 1;
    while (matched < needle.length && haystack[i + matched] === needle[matched]) matched += 1;
    if (matched === needle.length) return true;
  }
  return false;
}
