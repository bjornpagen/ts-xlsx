import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XmlParseError} from './errors.ts';
import {
  capturedText,
  closeEmptyElements,
  elementRange,
  elementRanges,
  elementSubtrees,
  openElements,
  parseXml,
  parseXmlPasses,
  type SaxHandlers,
  TextCapture,
} from './xml-read.ts';
import {localName, type XmlAttributes, xmlEvents} from './xml-scan.ts';

// The selection every `elementSubtrees` test below that captures differential styles shares.
const DXFS = new Map([['dxfs', 'dxf']]);

interface Event {
  readonly kind: 'open' | 'text' | 'close';
  readonly name?: string;
  readonly attrs?: XmlAttributes;
  readonly selfClosing?: boolean;
  readonly text?: string;
}

function events(source: string): Event[] {
  const out: Event[] = [];
  parseXml(source, {
    onOpen(name, attrs, selfClosing) {
      // Spread onto a plain object: the parsed map is deliberately null-prototype (a file supplies
      // the attribute names), which strict deep equality counts as a difference from a literal.
      out.push({kind: 'open', name, attrs: {...attrs}, selfClosing});
    },
    onText(text) {
      out.push({kind: 'text', text});
    },
    onClose(name) {
      out.push({kind: 'close', name});
    },
  });
  return out;
}

test('parseXml reports open/text/close for a simple element', () => {
  assert.deepEqual(events('<a>hi</a>'), [
    {kind: 'open', name: 'a', attrs: {}, selfClosing: false},
    {kind: 'text', text: 'hi'},
    {kind: 'close', name: 'a'},
  ]);
});

test('parseXml runs an onOpen-only handler, ignoring the omitted onText/onClose', () => {
  const opened: string[] = [];
  parseXml('<a>hi<b/></a>', {
    onOpen(name) {
      opened.push(name);
    },
  });
  assert.deepEqual(opened, ['a', 'b']);
});

test('capturedText yields each named element text as it closes, in document order', () => {
  const source =
    '<cp:coreProperties><dc:title>T</dc:title><x>skip</x><dc:creator>C</dc:creator></cp:coreProperties>';
  assert.deepEqual(
    [...capturedText(source, ['title', 'creator'])],
    [
      {local: 'title', text: 'T'},
      {local: 'creator', text: 'C'},
    ],
  );
});

test('capturedText joins the chunks an entity splits a value into', () => {
  // A run of character data ends at an entity, so `a & b` arrives as three text events. The value a
  // caller reads must be the whole element text, not its first fragment.
  assert.deepEqual([...capturedText('<t>a &amp; b</t>', 't')], [{local: 't', text: 'a & b'}]);
});

test('capturedText yields an empty element as empty text, and a self-closing one not at all', () => {
  // `<x/>` fires no close, and carries no text to report. Distinguishing the two is the branch
  // TextCapture exists to make structural rather than something each caller remembers.
  assert.deepEqual([...capturedText('<t></t>', 't')], [{local: 't', text: ''}]);
  assert.deepEqual([...capturedText('<t/>', 't')], []);
});

test('capturedText takes a single name as well as a set, and ignores namespace prefixes', () => {
  assert.deepEqual(
    [...capturedText('<Properties><vt:Company>ACME</vt:Company></Properties>', 'Company')],
    [{local: 'Company', text: 'ACME'}],
  );
});

test('capturedText is lazy, so a caller can stop at the element it wanted', () => {
  let first: string | undefined;
  for (const {text} of capturedText('<r><t>one</t><t>two</t></r>', 't')) {
    first = text;
    break;
  }
  assert.equal(first, 'one');
});

test('openElements with no filter yields every start, skipping text and close', () => {
  const seen = [...openElements('<a><b>hi</b><c/></a>')].map((e) => e.name);
  assert.deepEqual(seen, ['a', 'b', 'c']);
});

test('openElements restricts to the given local names, ignoring namespace prefixes', () => {
  const source = '<r><a:sheet name="one"/><x/><b:sheet name="two"/></r>';
  const names = [...openElements(source, 'sheet')].map((e) => e.attrs.name);
  assert.deepEqual(names, ['one', 'two']);
});

test('openElements matches any of several local names and reports the matched local name', () => {
  const source = '<Types><Default Extension="xml"/><Override PartName="/x"/><Other/></Types>';
  assert.deepEqual(
    [...openElements(source, 'Override', 'Default')].map((e) => e.local),
    ['Default', 'Override'],
  );
});

test('openElements surfaces decoded attributes and can be stopped early by breaking', () => {
  let first: XmlAttributes | undefined;
  for (const {attrs} of openElements('<r><a v="x &amp; y"/><a v="z"/></r>', 'a')) {
    first = attrs;
    break;
  }
  assert.equal(first?.v, 'x & y');
});

test('closeEmptyElements expands a named self-closing element to open+close, as if <x></x>', () => {
  const events = [...closeEmptyElements(xmlEvents('<r><c/></r>'), new Set(['c']))];
  assert.deepEqual(
    events.map((e) =>
      e.kind === 'open'
        ? `open:${e.name}:${e.selfClosing}`
        : `${e.kind}:${'name' in e ? e.name : ''}`,
    ),
    ['open:r:false', 'open:c:false', 'close:c', 'close:r'],
  );
});

test('closeEmptyElements matches by local name and leaves un-named self-closing tags bare', () => {
  // `a:c` matches on local name `c` and is expanded; the un-named `d` stays a self-closing open
  // with no synthesized close.
  const events = [...closeEmptyElements(xmlEvents('<a:c/><d/>'), new Set(['c']))];
  assert.deepEqual(
    events.map((e) => (e.kind === 'open' ? `${e.name}:${e.selfClosing}` : e.kind)),
    ['a:c:false', 'close', 'd:true'],
  );
});

test('parseXml with closeEmptyElements fires onClose only for the named self-closing element', () => {
  const opens: string[] = [];
  const closes: string[] = [];
  parseXml(
    '<r><c/><v/></r>',
    {
      onOpen: (name) => opens.push(name),
      onClose: (name) => closes.push(name),
    },
    {closeEmptyElements: new Set(['c'])},
  );
  assert.deepEqual(opens, ['r', 'c', 'v']);
  // `<c/>` is expanded (fires a close); the un-named `<v/>` is not; `<r>` closes normally.
  assert.deepEqual(closes, ['c', 'r']);
});

// Both scanners in this file classify markup through one shared step, so each truncation is pinned
// from both sides: a form that stopped throwing in one of them would otherwise be caught in neither.
const TRUNCATED: readonly [string, RegExp][] = [
  ['<r><!-- oops', /unterminated comment/],
  ['<r><![CDATA[ oops', /unterminated CDATA section/],
  ['<r><?oops', /unterminated processing instruction/],
];

test('an unterminated comment, CDATA section or processing instruction throws in both scanners', () => {
  for (const [source, message] of TRUNCATED) {
    assert.throws(() => events(source), message, `xmlEvents on ${source}`);
    assert.throws(() => elementSubtrees(source, DXFS), message, `elementSubtrees on ${source}`);
  }
});

test('the two scanners agree on what is markup and what is element content', () => {
  // They run over the same part (the stylesheet goes through both), so a form one recognises and the
  // other does not would make them disagree about where an element ends. Every one of the four
  // non-tag forms here holds a `</dxf>` that is text, not a close tag.
  const xml =
    '<?xml version="1.0"?><!DOCTYPE s><s><dxfs>' +
    '<dxf><!-- </dxf> --><![CDATA[</dxf>]]><?pi </dxf> ?><b/></dxf>' +
    '</dxfs></s>';

  const {fragments} = elementSubtrees(xml, DXFS);
  assert.deepEqual(fragments.get('dxfs'), [
    '<dxf><!-- </dxf> --><![CDATA[</dxf>]]><?pi </dxf> ?><b/></dxf>',
  ]);

  // The same source through the event stream: the CDATA arrives as text (the one form the two
  // deliberately treat differently) and nothing else does, so no `</dxf>` is reported twice.
  const evs = events(xml);
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['</dxf>'],
  );
  assert.deepEqual(
    evs.filter((e) => e.kind === 'close').map((e) => e.name),
    ['dxf', 'dxfs', 's'],
  );
});

// Gather every text a capture yields over one document, so a test reads as the XML it is about.
function captured(xml: string, names: string | Iterable<string>): string[] {
  const capture = new TextCapture(names);
  const out: string[] = [];
  parseXml(xml, {
    onOpen(name, _attrs, selfClosing) {
      capture.open(localName(name), selfClosing);
    },
    onText(chunk) {
      capture.text(chunk);
    },
    onClose(name) {
      const text = capture.close(localName(name));
      if (text !== undefined) out.push(text);
    },
  });
  return out;
}

test('TextCapture gathers an element text across the chunks it arrives in', () => {
  assert.deepEqual(captured('<r><t>a &amp; b</t><t>second</t></r>', 't'), ['a & b', 'second']);
});

test('TextCapture ignores a self-closing element, which will never fire a close', () => {
  assert.deepEqual(captured('<r><t/><t>after</t></r>', 't'), ['after']);
  // The regression this exists to make impossible: a latch nothing closes, leaking the following
  // element's text into the empty one.
  const capture = new TextCapture('t');
  capture.open('t', true);
  capture.text('stray');
  assert.equal(capture.capturing, false);
  assert.equal(capture.close('t'), undefined);
});

test('TextCapture leaves a capture in progress alone when an unrelated element opens', () => {
  assert.deepEqual(captured('<t>before<b/>after</t>', 't'), ['beforeafter']);
});

test('TextCapture answers only for the element that started the capture', () => {
  const capture = new TextCapture(['f', 'sqref']);
  capture.open('f', false);
  capture.text('A1>0');
  assert.equal(capture.close('sqref'), undefined, 'a sibling name does not consume it');
  assert.equal(capture.close('f'), 'A1>0');
  assert.equal(capture.close('f'), undefined, 'and it unlatches');
});

test('TextCapture reset abandons a capture that truncated markup left armed', () => {
  const capture = new TextCapture('t');
  capture.open('t', false);
  capture.text('orphan');
  capture.reset();
  assert.equal(capture.capturing, false);
  capture.text('loose');
  assert.equal(capture.close('t'), undefined, 'neither text belongs to whoever closes next');
});

test('TextCapture over a set tells the caller which element it captured, by the close it answers', () => {
  assert.deepEqual(
    captured('<cp><title>T</title><ignored>X</ignored><creator>C</creator></cp>', [
      'title',
      'creator',
    ]),
    ['T', 'C'],
  );
});

test('parseXmlPasses delivers every event to every pass', () => {
  const seen: string[][] = [[], []];
  const record = (index: number): SaxHandlers => ({
    onOpen: (name, _attrs, selfClosing) =>
      seen[index]?.push(`open ${name}${selfClosing ? '/' : ''}`),
    onText: (text) => seen[index]?.push(`text ${text}`),
    onClose: (name) => seen[index]?.push(`close ${name}`),
  });
  parseXmlPasses('<a x="1">hi<b/></a>', [{handlers: record(0)}, {handlers: record(1)}]);
  assert.deepEqual(seen[0], ['open a', 'text hi', 'open b/', 'close a']);
  assert.deepEqual(seen[1], seen[0], 'both passes see the same events, in the same order');
});

test('parseXmlPasses unions the self-closing expansions its passes ask for', () => {
  const seen: string[] = [];
  // Only the second pass needs `<b/>` expanded, and only the first `<c/>`. Sharing one parse means
  // both expansions happen for both, which is the one way a pass can tell it is sharing.
  parseXmlPasses('<r><b/><c/></r>', [
    {
      handlers: {
        onOpen: (name, _attrs, selfClosing) => seen.push(`open ${name}${selfClosing ? '/' : ''}`),
        onClose: (name) => seen.push(`close ${name}`),
      },
      closeEmptyElements: new Set(['c']),
    },
    {handlers: {onOpen: () => {}}, closeEmptyElements: new Set(['b'])},
  ]);
  assert.deepEqual(seen, ['open r', 'open b', 'close b', 'open c', 'close c', 'close r']);
});

test('parseXmlPasses over no passes is a parse that reaches nobody, not a throw', () => {
  assert.doesNotThrow(() => parseXmlPasses('<a><b/></a>', []));
});

test('elementSubtrees captures a child element verbatim, tags included', () => {
  const {fragments} = elementSubtrees('<s><dxfs count="1"><dxf><b/></dxf></dxfs></s>', DXFS);
  assert.deepEqual(fragments.get('dxfs'), ['<dxf><b/></dxf>']);
});

test('elementSubtrees counts nesting, so a same-named descendant does not end the capture', () => {
  const {fragments} = elementSubtrees(
    '<s><dxfs><dxf>a<dxf>inner</dxf>b</dxf><dxf>second</dxf></dxfs></s>',
    DXFS,
  );
  assert.deepEqual(fragments.get('dxfs'), ['<dxf>a<dxf>inner</dxf>b</dxf>', '<dxf>second</dxf>']);
});

test('elementSubtrees captures a self-closing child as its own tag', () => {
  const {fragments} = elementSubtrees(
    '<s><indexedColors><rgbColor rgb="FF00FF00"/><rgbColor rgb="FF0000FF"/></indexedColors></s>',
    new Map([['indexedColors', 'rgbColor']]),
  );
  assert.deepEqual(fragments.get('indexedColors'), [
    '<rgbColor rgb="FF00FF00"/>',
    '<rgbColor rgb="FF0000FF"/>',
  ]);
});

test('elementSubtrees keeps CDATA and comments inside a subtree, and reads neither as markup', () => {
  // A `</dxf>` inside a comment or a CDATA section is text, not the element's end. A regex scanner
  // is exactly what cannot tell the difference.
  const xml = '<s><dxfs><dxf><![CDATA[</dxf>]]><!-- </dxf> --><b/></dxf></dxfs></s>';
  const {fragments} = elementSubtrees(xml, DXFS);
  assert.deepEqual(fragments.get('dxfs'), ['<dxf><![CDATA[</dxf>]]><!-- </dxf> --><b/></dxf>']);
});

test('elementSubtrees scopes the child to its container', () => {
  // `<color>` appears all over a stylesheet; only the ones inside `<mruColors>` are the swatches.
  const xml =
    '<s><fonts><font><color rgb="FFAAAAAA"/></font></fonts>' +
    '<colors><mruColors><color rgb="FF111111"/></mruColors></colors></s>';
  const {fragments} = elementSubtrees(xml, new Map([['mruColors', 'color']]));
  assert.deepEqual(fragments.get('mruColors'), ['<color rgb="FF111111"/>']);
});

test('elementSubtrees reports a container element own attributes', () => {
  const {attributes} = elementSubtrees(
    '<s><tableStyles count="0" defaultTableStyle="TableStyleMedium2"/></s>',
    new Map([['tableStyles', 'tableStyle']]),
  );
  assert.equal(attributes.get('tableStyles')?.defaultTableStyle, 'TableStyleMedium2');
  assert.equal(
    attributes.get('tableStyles')?.count,
    '0',
    'a self-closing container still reports what it declared',
  );
});

test('elementSubtrees reads a container once, ignoring a second block of the same name', () => {
  const {fragments} = elementSubtrees(
    '<s><dxfs><dxf>a</dxf></dxfs><dxfs><dxf>b</dxf></dxfs></s>',
    DXFS,
  );
  assert.deepEqual(fragments.get('dxfs'), ['<dxf>a</dxf>'], 'the first block, not the two merged');
});

test('elementSubtrees captures several containers in the one scan', () => {
  const xml =
    '<s><dxfs><dxf>d</dxf></dxfs><colors><indexedColors><rgbColor rgb="FF1"/></indexedColors>' +
    '<mruColors><color rgb="FF2"/></mruColors></colors></s>';
  const {fragments} = elementSubtrees(
    xml,
    new Map([
      ['dxfs', 'dxf'],
      ['indexedColors', 'rgbColor'],
      ['mruColors', 'color'],
    ]),
  );
  assert.deepEqual(fragments.get('dxfs'), ['<dxf>d</dxf>']);
  assert.deepEqual(fragments.get('indexedColors'), ['<rgbColor rgb="FF1"/>']);
  assert.deepEqual(fragments.get('mruColors'), ['<color rgb="FF2"/>']);
});

test('elementSubtrees throws on a subtree that never closes rather than slicing a partial one', () => {
  // Re-emitting half a `<dxf>` verbatim hands broken markup on as though it were content.
  assert.throws(
    () => elementSubtrees('<s><dxfs><dxf><b/>', DXFS),
    (error: unknown) => error instanceof XmlParseError,
  );
});

test('elementSubtrees yields nothing for a container the document does not carry', () => {
  const {fragments, attributes} = elementSubtrees('<s><fonts><font/></fonts></s>', DXFS);
  assert.equal(fragments.get('dxfs'), undefined);
  assert.equal(attributes.get('dxfs'), undefined);
});

test('elementRange finds an element by path and bounds its content', () => {
  const source = '<root><a><b x="1">body</b></a></root>';
  const found = elementRange(source, ['a', 'b']);
  assert.ok(found);
  assert.equal(source.slice(found.start, found.end), '<b x="1">body</b>');
  assert.equal(source.slice(found.contentStart, found.contentEnd), 'body');
  assert.deepEqual({...found.attrs}, {x: '1'});
});

test('elementRanges finds every element of a name, in order, whatever its prefix', () => {
  const source =
    '<c:chart><c:tx><c:f>S!$B$1</c:f></c:tx><!-- <c:f>no</c:f> --><c:val><x:f/><c:f>S!$B$2:$B$5</c:f></c:val></c:chart>';
  const found = [...elementRanges(source, 'f')].map((range) => ({
    name: range.name,
    content: source.slice(range.contentStart, range.contentEnd),
    element: source.slice(range.start, range.end),
  }));
  assert.deepEqual(found, [
    {name: 'c:f', content: 'S!$B$1', element: '<c:f>S!$B$1</c:f>'},
    {name: 'x:f', content: '', element: '<x:f/>'},
    {name: 'c:f', content: 'S!$B$2:$B$5', element: '<c:f>S!$B$2:$B$5</c:f>'},
  ]);
});

test('elementRanges keeps a nested same-named element inside its ancestor, and refuses one unclosed', () => {
  const source = '<a><f>1<f>2</f>3</f><f>4</f></a>';
  assert.deepEqual(
    [...elementRanges(source, 'f')].map((range) => source.slice(range.start, range.end)),
    ['<f>1<f>2</f>3</f>', '<f>4</f>'],
  );
  assert.throws(() => [...elementRanges('<a><f>1</a>', 'f')], XmlParseError);
});

test('elementRange skips a close tag hiding in a comment or a CDATA section', () => {
  const source = '<root><a><!-- </a> --><![CDATA[</a>]]>body</a></root>';
  const found = elementRange(source, ['a']);
  assert.ok(found);
  assert.equal(
    source.slice(found.contentStart, found.contentEnd),
    '<!-- </a> --><![CDATA[</a>]]>body',
  );
});

test('elementRange counts a same-named descendant rather than ending on it', () => {
  const source = '<root><a>x<a>inner</a>y</a></root>';
  const found = elementRange(source, ['a']);
  assert.ok(found);
  assert.equal(source.slice(found.contentStart, found.contentEnd), 'x<a>inner</a>y');
});

test('elementRange matches on the local name, whatever prefix the source binds', () => {
  const source = '<z:root xmlns:z="urn:x"><z:a><z:b/></z:a></z:root>';
  const found = elementRange(source, ['a', 'b']);
  assert.ok(found);
  assert.equal(found.name, 'z:b');
  assert.equal(found.start, source.indexOf('<z:b/>'));
  assert.equal(found.contentStart, found.contentEnd, 'an empty element has no content');
});

test('elementRange answers undefined for a path that is not there', () => {
  assert.equal(elementRange('<root><a/></root>', ['a', 'b']), undefined);
  assert.equal(elementRange('<root><b/></root>', ['a']), undefined);
});

test('elementRange refuses an unterminated element rather than returning half a range', () => {
  assert.throws(() => elementRange('<root><a>body</root>', ['a']), XmlParseError);
});
