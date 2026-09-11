// The XML scanner: source text in, events and attribute values out.
//
// OOXML uses a small, regular subset of XML, so the reader does not need, and must not pay for, a
// general-purpose DOM library (see ADR 0004). This scans the source in a single O(n) pass with no
// recursion, emitting open/text/close events; the OOXML reader consumes them and builds only the
// model, so peak memory tracks real content rather than document structure.
//
// Security posture: entities are *decoded, never expanded*. Only the five predefined entities and
// numeric character references are recognised; DTDs and `<!ENTITY>` definitions are skipped, so
// entity-expansion (billion-laughs) and external-entity (XXE) attacks are structurally impossible
// here, not merely mitigated.
//
// Everything here answers "what does this piece of markup say", plus the three OOXML booleans, which
// are readings of a value but are the ones every OOXML dialect asks for rather than only a
// spreadsheet. Two neighbours hold the rest. `xml-read.ts` holds the ways of *driving* a scan (the
// pull generators, the push adapter and its multi-pass driver, the verbatim subtree capture, the text
// gatherer), and nothing here holds a traversal. `xml-attrs.ts` holds the spreadsheet-flavoured
// readings of an attribute's value: the numbers with their floors, the numeric-literal coercion, the
// enumerated token, the `_xHHHH_` cell-text escape.
//
// Both seams are load-bearing rather than tidiness. `/customui` is a ribbon reader that wants the
// scan, an OOXML boolean, and nothing else; while its neighbours lived in one module every helper
// added to any of them was charged to that entry's bundle budget, which is how it drifted 3 KB over
// one before the first split and was left with 0.7 KB of headroom before the second.

import {XmlParseError} from './errors.ts';
import {isRepresentableCodePoint, stripUnrepresentable} from './xml-chars.ts';

export interface XmlAttributes {
  readonly [name: string]: string;
}

/**
 * One parse event from {@link xmlEvents}. The payloads are what the push adapter in `xml-read.ts`
 * hands a handler triple, unchanged: `text` is already entity-decoded (or verbatim CDATA), and a
 * `<x/>` yields one `open` with `selfClosing: true` and no matching `close`. The discriminated
 * `kind` lets a *pull* consumer drive the parse: the shape the streaming reader needs, where a push
 * callback cannot `yield`.
 */
export type XmlEvent =
  | {
      readonly kind: 'open';
      readonly name: string;
      readonly attrs: XmlAttributes;
      readonly selfClosing: boolean;
    }
  | {readonly kind: 'text'; readonly text: string}
  | {readonly kind: 'close'; readonly name: string};

// A Map, not an object literal, because the key comes straight out of the file: an object would
// answer `constructor`, `toString` and a dozen other attacker-chosen names out of Object.prototype,
// and a function stringified into the decoded text is engine source in the model.
const PREDEFINED_ENTITIES: ReadonlyMap<string, string> = new Map([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
]);

const ENTITY = /&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z][a-zA-Z0-9]*);/g;

/**
 * Decode XML character references and the five predefined entities. An unrecognised
 * `&name;` is left verbatim rather than expanded: there is no DTD, so there is nothing
 * to expand it to, and refusing to invent one is what makes entity-expansion attacks
 * impossible.
 *
 * A reference naming a code point XML 1.0 has no representation for is left verbatim too, on the
 * same grounds as one naming no code point at all. `&#1;` is not an escape for U+0001, it is another
 * way of spelling an ill-formed document, and decoding it puts in the model a character the writer
 * is *guaranteed* to refuse: a hostile file would then read cleanly and fail on the next save with
 * an `AuthoringError` blaming the caller, several layers from the input that caused it. The bound
 * comes from `./xml-chars.ts` rather than from the writer's own guard so that this half of the
 * codec keeps importing none of the serialisation vocabulary.
 */
export function decodeEntities(value: string): string {
  if (!value.includes('&')) return value;
  return value.replace(ENTITY, (match, body: string) => {
    if (body.charCodeAt(0) === 0x23 /* # */) {
      const codePoint =
        body.charCodeAt(1) === 0x78 /* x */
          ? Number.parseInt(body.slice(2), 16)
          : Number.parseInt(body.slice(1), 10);
      if (!isRepresentableCodePoint(codePoint)) return match;
      return String.fromCodePoint(codePoint);
    }
    return PREDEFINED_ENTITIES.get(body) ?? match;
  });
}

/**
 * Every string this scanner hands out: entity-decoded, then stripped of anything XML 1.0 could not
 * have carried in the first place.
 *
 * A *reference* to an unrepresentable character keeps the spelling the file used, because there is
 * one to keep. A raw one has none, so the only choices are to drop it or to admit a value the writer
 * must refuse, and admitting it reports a corrupt input as the caller's mistake. Such a character
 * makes the document ill-formed by the `Char` production, so it was never legally there; the cell
 * values that legitimately carry one spell it `_x0001_`, which is a different convention entirely
 * and is decoded by {@link decodeSpreadsheetText} well after this.
 */
function admitText(value: string): string {
  return stripUnrepresentable(decodeEntities(value));
}

const EQUALS = 0x3d;
const SLASH = 0x2f;
const GREATER_THAN = 0x3e;
const DOUBLE_QUOTE = 0x22;
const SINGLE_QUOTE = 0x27;

const WIDE_WHITESPACE = /\s/;

// JavaScript's `\s`, not XML's four whitespace characters: the regex this scan replaced split names on
// `\s`, and keeping its set keeps every tag reading exactly as it did. Beyond ASCII the set is asked of
// `\s` itself, one character at a time, rather than copied into a table that could drift from it.
function isScanWhitespace(code: number): boolean {
  if (code < 0x80) return code === 0x20 || (code >= 0x09 && code <= 0x0d);
  return WIDE_WHITESPACE.test(String.fromCharCode(code));
}

function skipScanWhitespace(source: string, from: number): number {
  let i = from;
  while (i < source.length && isScanWhitespace(source.charCodeAt(i))) i++;
  return i;
}

function endsName(code: number): boolean {
  // Everything between `>` and NBSP ends nothing, and that range holds the letters most names are
  // made of, so the common character is answered by one comparison instead of the whitespace table.
  if (code > GREATER_THAN && code < 0xa0) return false;
  return isScanWhitespace(code) || code === EQUALS || code === SLASH || code === GREATER_THAN;
}

/**
 * A tag's attributes: each a name, then `=`, then a value in either quote style. Matching the quotes
 * is what lets a delimiter-respecting scan find a tag's end even when a value holds a `>` (legal but
 * rare). Names may carry a namespace prefix (`r:id`, `xml:space`). Anything that is not an attribute,
 * such as a name nothing assigns or a quote never closed, is skipped rather than refused.
 *
 * Linear in the source, and that is the point of it being hand-written. The regex it replaced
 * backtracked over a run of name characters no `=` followed from every position inside the run, so a
 * single junk token in a tag cost the square of its length: a megabyte of one letter, a kilobyte
 * zipped, cost minutes inside either reader. Here a dropped token resumes the scan where the token
 * ended, and a value's closing quote is found with one `indexOf`, so no character is visited more
 * than a constant number of times.
 */
export function parseAttributes(source: string): XmlAttributes {
  // Null-prototype: attribute names are file-derived, and `XmlAttributes` is an index signature
  // every reader reads through, so an inherited `constructor` would read as a present attribute.
  const attrs: Record<string, string> = Object.create(null) as Record<string, string>;
  const length = source.length;
  let i = 0;
  while (i < length) {
    const nameStart = i;
    while (i < length && !endsName(source.charCodeAt(i))) i++;
    if (i === nameStart) {
      i++;
      continue;
    }
    let cursor = skipScanWhitespace(source, i);
    if (source.charCodeAt(cursor) !== EQUALS) {
      // A name nothing assigns, like the `jjj` in `<c jjj r="A1">`: dropped, resuming at whatever
      // ended it.
      i = cursor;
      continue;
    }
    cursor = skipScanWhitespace(source, cursor + 1);
    const quote = source.charCodeAt(cursor);
    const close =
      quote === DOUBLE_QUOTE || quote === SINGLE_QUOTE
        ? source.indexOf(quote === DOUBLE_QUOTE ? '"' : "'", cursor + 1)
        : -1;
    if (close === -1) {
      // No value, or one never closed. What stands at the cursor starts the next name, a stray quote
      // included, since a quote is a name character to this scan as it was to the regex.
      i = cursor;
      continue;
    }
    attrs[source.slice(nameStart, i)] = admitText(
      normalizeAttributeWhitespace(source.slice(cursor + 1, close)),
    );
    i = close + 1;
  }
  return attrs;
}

// XML 1.0 §3.3.3: a literal tab, LF or CR in an attribute value reads as a space, and end-of-line
// handling (§2.11) has already made a CRLF pair one LF, so the pair is one space. Applied to the raw
// value, before entities are decoded, because a character reference is exempt: `&#10;` is how a file
// spells a line break an attribute keeps, and the writer escapes one that way for this very rule. Kept
// raw, a newline in a foreign file's validation prompt read as a line break where Excel shows a space,
// and a save then wrote it back as `&#10;`, a real line break.
const ATTRIBUTE_WHITESPACE = /\r\n|[\t\n\r]/g;

function normalizeAttributeWhitespace(raw: string): string {
  // The scan runs on every attribute of every tag and almost no value holds one of the three, so the
  // replace is reached only past a loop over the value's few characters; asking the regex costs every
  // attribute a third again.
  for (let i = 0; i < raw.length; i++) {
    const code = raw.charCodeAt(i);
    if (code === 0x09 || code === 0x0a || code === 0x0d) {
      return raw.replace(ATTRIBUTE_WHITESPACE, ' ');
    }
  }
  return raw;
}

// Scan to the tag's closing `>`, honouring quoted attribute values so a `>` inside a
// value does not end the tag prematurely.
function findTagEnd(source: string, start: number): number {
  let quote = '';
  for (let i = start + 1; i < source.length; i++) {
    const ch = source[i];
    if (quote !== '') {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return i;
    }
  }
  throw new XmlParseError('unterminated tag: missing ">"');
}

// A `<!DOCTYPE …>` may contain a bracketed internal subset with its own `>`; balance the
// brackets so the declaration is skipped whole. We never act on its contents.
function skipDeclaration(source: string, start: number): number {
  let depth = 0;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    else if (ch === '>' && depth <= 0) return i + 1;
  }
  throw new XmlParseError('unterminated markup declaration: missing ">"');
}

// Non-tag markup: a comment, a CDATA section, a processing instruction, or a markup declaration
// (`<!DOCTYPE ...>`). `next` is where the scan resumes past it. A CDATA section additionally reports
// its content bounds, because that is the one form the two scanners below treat differently: the
// event stream yields the text, the verbatim capture steps over it.
export type Markup =
  | {readonly kind: 'comment' | 'pi' | 'declaration'; readonly next: number}
  | {
      readonly kind: 'cdata';
      readonly contentStart: number;
      readonly contentEnd: number;
      readonly next: number;
    };

// What counts as markup rather than element content, stated once for both scanners below.
//
// Both are run over the same part (`parseStyleTable` puts both over `xl/styles.xml`), so a form one
// of them recognised and the other did not would make them disagree about where an element ends: a
// `</dxf>` inside a comment would close a capture in one and not the other. Classifying here keeps
// that impossible rather than merely unlikely. Returns undefined at a tag, which is the one form
// each scanner handles for itself, since that is where they genuinely differ.
export function markupAt(source: string, lt: number): Markup | undefined {
  if (source.startsWith('<!--', lt)) {
    const end = source.indexOf('-->', lt + 4);
    if (end === -1) throw new XmlParseError('unterminated comment');
    return {kind: 'comment', next: end + 3};
  }
  if (source.startsWith('<![CDATA[', lt)) {
    const end = source.indexOf(']]>', lt + 9);
    if (end === -1) throw new XmlParseError('unterminated CDATA section');
    return {kind: 'cdata', contentStart: lt + 9, contentEnd: end, next: end + 3};
  }
  if (source.startsWith('<?', lt)) {
    const end = source.indexOf('?>', lt + 2);
    if (end === -1) throw new XmlParseError('unterminated processing instruction');
    return {kind: 'pi', next: end + 2};
  }
  if (source.startsWith('<!', lt)) return {kind: 'declaration', next: skipDeclaration(source, lt)};
  return undefined;
}

// A tag at `lt`, split into the pieces both scanners want: whether it is a close tag, the name as
// written (namespace prefix included), the body after the name for {@link parseAttributes}, whether
// it closed itself, and where the scan resumes. The counterpart to {@link markupAt} for the one form
// that is not markup, shared for the same reason: the split is identical in both, and only what each
// builds from the pieces differs.
export interface Tag {
  readonly close: boolean;
  readonly name: string;
  readonly attrSource: string;
  readonly selfClosing: boolean;
  /** One past the tag's `>`, so a caller capturing verbatim source can slice up to it. */
  readonly next: number;
}

export function tagAt(source: string, lt: number): Tag {
  const gt = findTagEnd(source, lt);
  const raw = source.slice(lt + 1, gt);
  const next = gt + 1;
  if (raw.charCodeAt(0) === 0x2f /* / */) {
    return {close: true, name: raw.slice(1).trim(), attrSource: '', selfClosing: false, next};
  }
  const selfClosing = raw.charCodeAt(raw.length - 1) === 0x2f;
  const body = selfClosing ? raw.slice(0, -1) : raw;
  const nameEnd = firstWhitespace(body);
  return {
    close: false,
    name: nameEnd === -1 ? body : body.slice(0, nameEnd),
    attrSource: nameEnd === -1 ? '' : body.slice(nameEnd),
    selfClosing,
    next,
  };
}

/**
 * Scan an XML document as a *pull* stream of {@link XmlEvent}s in a single O(n) pass with no
 * recursion. This is the parser's core; {@link parseXml} is a thin push adapter over it. A
 * consumer that must produce output incrementally (the streaming row reader) pulls events and
 * yields as it goes, holding only its own running state; a push callback cannot.
 *
 * Throws {@link XmlParseError} on malformed markup.
 */
export function* xmlEvents(source: string): Generator<XmlEvent> {
  const length = source.length;
  let i = 0;

  while (i < length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) {
      const chunk = source.slice(i);
      if (chunk.length > 0) yield {kind: 'text', text: admitText(normalizeLineEndings(chunk))};
      return;
    }
    if (lt > i) {
      const chunk = source.slice(i, lt);
      if (chunk.length > 0) yield {kind: 'text', text: admitText(normalizeLineEndings(chunk))};
    }

    const markup = markupAt(source, lt);
    if (markup !== undefined) {
      // The one place the two scanners part company: an event stream owes its consumer the CDATA
      // text, a verbatim capture owes it nothing and takes the bounds only to step over them.
      if (markup.kind === 'cdata') {
        yield {
          kind: 'text',
          text: stripUnrepresentable(source.slice(markup.contentStart, markup.contentEnd)),
        };
      }
      i = markup.next;
      continue;
    }

    const tag = tagAt(source, lt);
    yield tag.close
      ? {kind: 'close', name: tag.name}
      : {
          kind: 'open',
          name: tag.name,
          attrs: parseAttributes(tag.attrSource),
          selfClosing: tag.selfClosing,
        };
    i = tag.next;
  }
}

// XML end-of-line handling (spec §2.11): a literal CRLF or lone CR in character data is
// normalized to a single LF, so a value's in-cell line breaks read back identically whatever
// newline convention the producer wrote. Normalization precedes entity decoding, so a
// deliberately-encoded carriage return (&#13;) survives it: the escape hatch for a real CR.
// CDATA is delivered verbatim (it bypasses this), matching the reader's CDATA contract.
function normalizeLineEndings(chunk: string): string {
  if (!chunk.includes('\r')) return chunk;
  return chunk.replace(/\r\n?/g, '\n');
}

function firstWhitespace(source: string): number {
  for (let i = 0; i < source.length; i++) {
    const code = source.charCodeAt(i);
    if (code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d) return i;
  }
  return -1;
}

/** Strip a namespace prefix from a qualified name (`r:id` → `id`, `sheet` → `sheet`). */
export function localName(qualified: string): string {
  const colon = qualified.indexOf(':');
  return colon === -1 ? qualified : qualified.slice(colon + 1);
}

// OOXML spells booleans three ways, and the reader needs all three. A `<b/>`-style font flag
// defaults to on when present with no value, so its absence is meaningful (`boolPresent`). Most
// attributes are plain xsd:booleans that are off unless an explicit "1"/"true" turns them on
// (`boolStrict`). An optional attribute that must round-trip byte-clean has to distinguish absent
// from present-and-false and drop an unrecognised token rather than coerce it (`boolTristate`).
// There is no fourth variant: a hand-rolled `attr !== '0'` reads `"false"` as true, because
// xsd:boolean spells false both ways and only Excel's own files consistently pick the digit.

/** An OOXML boolean that is on when present with no value (`<b/>` is bold) and off only on an
 * explicit `"0"`/`"false"`; absence reads as on. */
export function boolPresent(val: string | undefined): boolean {
  return val === undefined || (val !== '0' && val !== 'false');
}

/** An OOXML boolean that is on only when explicitly `"1"`/`"true"`; anything else, including
 * absence and a truthy-looking `"0"`, is off. */
export function boolStrict(val: string | undefined): boolean {
  return val === '1' || val === 'true';
}

/** An optional OOXML boolean: `undefined` when the attribute is absent or carries an unrecognised
 * token, otherwise its `"1"`/`"true"` vs `"0"`/`"false"` value. Lets a caller store only the
 * attributes the source actually carried, so a re-write stays byte-clean. */
export function boolTristate(val: string | undefined): boolean | undefined {
  if (val === '1' || val === 'true') return true;
  if (val === '0' || val === 'false') return false;
  return undefined;
}
